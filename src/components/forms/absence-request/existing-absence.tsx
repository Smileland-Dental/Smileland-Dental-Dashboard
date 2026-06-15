"use client";

import { getPSTDate } from '@/lib/date';
import React, { useState } from "react";
import { doc, updateDoc, deleteDoc } from "firebase/firestore";
import { db, storage } from '@/lib/firebase.config';
import { ref, uploadBytes, getDownloadURL, listAll, deleteObject } from "firebase/storage";
import FeedbackModal from '@/components/ui/FeedbackModal';
import { AlertCircle, Lock } from 'lucide-react'; // Added icons for visual feedback
import { OFFICES } from '@/lib/constants';

const incidentTypes = ["Late In", "Early Out", "Absent", "Leave and Come Back", "Long Lunch", "Switch Shift"];
//const officeLocations = ["Corporate", "Ming", "Bernard", "California", "Ortho", "Delano", "Tulare", "Visalia", "Fresno"];

export default function ExistingAbsenceForm({ absence, onFormSubmit, onClose }: { absence: any, onFormSubmit: () => void, onClose: () => void }) {
  // --- 1. DEFINE DENIED STATE ---
  const isDenied = absence.manager_approval === 'denied' || absence.final_approval === 'denied';
  const isHRCallIn = absence.type_of_request === "HR Call In";

  const [formData, setFormData] = useState({
    type_of_request: absence.type_of_request || '',
    type_of_incident: absence.type_of_incident || '',
    office: absence.office || '',
    incident_start: absence.incident_start || '',
    incident_end: absence.incident_end || '',
    employee_comments: absence.employee_comments || '',
    eta: absence.eta || '',
    etd: absence.etd || '',
    excuse_note_submitted: absence.excuse_note_submitted || 'not_provided',
  });

  const [newExcuseNotes, setNewExcuseNotes] = useState<File[]>([]);
  const [notesToRemove, setNotesToRemove] = useState<string[]>([]);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [isChecked, setIsChecked] = useState(false);
  const [feedback, setFeedback] = useState<{ isOpen: boolean; type: 'success' | 'error'; message: string }>({
    isOpen: false,
    type: 'success',
    message: '',
  });

  const isDataChanged = () => {
    if (isDenied) return false; // Force false if denied
    // ... (rest of field changes logic remains same)
    const hasFieldChanges = 
      formData.type_of_request !== (absence.type_of_request || '') ||
      formData.type_of_incident !== (absence.type_of_incident || '') ||
      formData.office !== (absence.office || '') ||
      formData.incident_start !== (absence.incident_start || '') ||
      formData.incident_end !== (absence.incident_end || '') ||
      formData.employee_comments !== (absence.employee_comments || '') ||
      formData.eta !== (absence.eta || '') ||
      formData.etd !== (absence.etd || '') ||
      formData.excuse_note_submitted !== (absence.excuse_note_submitted || 'not_provided');

    const hasFileChanges = newExcuseNotes.length > 0 || notesToRemove.length > 0;
    return hasFieldChanges || hasFileChanges;
  };

const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>) => {
    if (isDenied) return; // Prevent state updates if denied
    const { name, value } = e.target;
    // ... (rest of change logic remains same)
    if (name === "type_of_incident") {
      if (value === absence.type_of_incident) {
        setFormData(prev => ({ ...prev, [name]: value, eta: absence.eta || '', etd: absence.etd || '' }));
      } else {
        setFormData(prev => ({ ...prev, [name]: value, eta: '', etd: '' }));
      }
    } else {
      setFormData(prev => ({ ...prev, [name]: value }));
    }
    setIsChecked(false);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setNewExcuseNotes(prev => [...prev, ...Array.from(e.target.files!)]);
      setFormData(prev => ({
        ...prev,
        excuse_note_submitted: 'submitted'
      }));
      setIsChecked(false);
    }
  };

  const handleToggleRemoveExistingNote = (url: string) => {
    setNotesToRemove(prev => {
      const isRemoving = !prev.includes(url);
      const nextNotesToRemove = isRemoving ? [...prev, url] : prev.filter(u => u !== url);
      
      const existingNotesCount = (absence.excuse_note || []).length - nextNotesToRemove.length;
      
      // If we have files staged or remaining, keep it 'submitted'. Otherwise, wipe it.
      if (newExcuseNotes.length === 0 && existingNotesCount === 0) {
        setFormData(f => ({ ...f, excuse_note_submitted: 'not_provided' }));
      } else {
        setFormData(f => ({ ...f, excuse_note_submitted: 'submitted' }));
      }

      return nextNotesToRemove;
    });
    setIsChecked(false);
  };

  const handleDelete = async () => {
    if (!window.confirm("Are you sure you want to PERMANENTLY delete this request?")) return;
    setIsSubmitting(true);
    try {
      // --- 1. CLEAN UP STORAGE ---
      // Create a reference to the folder containing this absence's notes
      const folderRef = ref(storage, `excuse-notes/${absence.id}`);

      // List all files in that folder and delete them
      const fileList = await listAll(folderRef);

      // Delete every file in the folder
      const deletePromises = fileList.items.map(fileRef => deleteObject(fileRef));
      await Promise.all(deletePromises);

      // --- 2. DELETE DATABASE RECORD ---
      //console.log("Deleting absence with ID:", absence.id);
      await deleteDoc(doc(db, "absences", absence.id));
      setFeedback({ isOpen: true, type: 'success', message: "Request deleted successfully." });
    } catch (error) {
      setFeedback({ isOpen: true, type: 'error', message: "Error deleting request." });
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleSubmit = async (event: React.FormEvent) => {
    event.preventDefault();
    if (!isDataChanged() || !isChecked) return;
    setIsSubmitting(true);

    try {
      const submissionData = { ...formData };
      if (!["Late In", "Leave and Come Back"].includes(formData.type_of_incident)) submissionData.eta = "";
      if (!["Early Out", "Leave and Come Back"].includes(formData.type_of_incident)) submissionData.etd = "";

      // --- 1. DELETE REMOVED FILES FROM STORAGE ---
      if (notesToRemove.length > 0) {
        const deletionPromises = notesToRemove.map((url) => {
          const fileRef = ref(storage, url);
          return deleteObject(fileRef);
        });
        // We use allSettled so that if one file fails (e.g., already deleted), 
        // the whole form update doesn't crash.
        await Promise.allSettled(deletionPromises);
      }

      // --- 2. UPLOAD NEW FILES ---
      let newNoteUrls: string[] = [];
      if (newExcuseNotes.length > 0) {
        for (const file of newExcuseNotes) {
          const storageRef = ref(storage, `excuse-notes/${absence.id}/${file.name}`);
          await uploadBytes(storageRef, file);
          const url = await getDownloadURL(storageRef);
          newNoteUrls.push(url);
        }
      }
      
      // --- 3. FILTER LOCAL STATE ---
      const existingNotes = absence.excuse_note || [];
      const remainingExistingNotes = existingNotes.filter((url: string) => !notesToRemove.includes(url));
      const finalNotes = [...remainingExistingNotes, ...newNoteUrls];

      const managerApprovalStatus = isHRCallIn ? 'not_required' : 'pending';

      // --- 4. UPDATE FIRESTORE ---
      await updateDoc(doc(db, "absences", absence.id), {
        ...submissionData,
        excuse_note: finalNotes,
        manager_approval: managerApprovalStatus,
        manager_approval_name: '',
        final_approval: 'pending',
        final_approval_name: '', 
        updatedAt: new Date(), 
      });

      setFeedback({ isOpen: true, type: 'success', message: "Request updated successfully!" });
      // Reset removal tracking after success
      setNotesToRemove([]);
      setNewExcuseNotes([]);
    } catch (error) {
      setFeedback({ isOpen: true, type: 'error', message: "Error updating request." });
    } finally {
      setIsSubmitting(false);
    }
  };

  const changed = isDataChanged();

  return (
    <>
      <div className="fixed inset-0 bg-black/50 flex justify-center items-center z-50" onClick={onClose}>
        <div className="bg-white p-8 rounded-lg shadow-xl w-full max-w-2xl max-h-[90vh] overflow-y-auto" onClick={(e) => e.stopPropagation()}>
          
          <div className="flex justify-between items-center mb-6">
            <h2 className="text-2xl font-bold text-gray-800">
              {isDenied ? 'View Absence Request' : 'Edit Absence Request'}
            </h2>
            
            
            {/* 2. HIDE DELETE IF DENIED */}
            {!isDenied && (
              <button type="button" onClick={handleDelete} className="text-xs font-bold text-red-500 hover:bg-red-50 px-3 py-1 rounded border border-red-200 uppercase tracking-tighter disabled:opacity-50 disabled:cursor-not-allowed disabled:border-gray-200 disabled:text-gray-400"
                disabled={absence.type_of_request === "HR Call In" || absence.final_approval === "approved" || absence.final_approval === "denied" || isSubmitting}
              >
                Delete Request
              </button>
            )}
          </div>

          {/* 3. DENIED STATUS BANNER */}
          {isDenied && (
            <div className="mb-6 p-4 bg-rose-50 border border-rose-200 rounded-xl flex items-center gap-3">
              <Lock className="h-5 w-5 text-rose-500" />
              <div>
                <p className="text-sm font-bold text-rose-800">Request Locked</p>
                <p className="text-xs text-rose-600">This request has been denied and can no longer be edited.</p>
              </div>
            </div>
          )}
          
          <form onSubmit={handleSubmit} className="space-y-5">
            {/* Request Type */}
            <div className="p-4 bg-amber-50 rounded-lg border border-amber-200">
              <label className="block text-sm font-semibold text-amber-900 mb-2">Type Of Request</label>
              <div className="flex gap-6 text-sm">
                {["Incident Notice", "Time Off Request", "HR Call In"].map(val => (
                  <label key={val} className="flex items-center gap-2 cursor-pointer">
                    <input 
                      type="radio" 
                      name="type_of_request" 
                      value={val} 
                      checked={formData.type_of_request === val} 
                      onChange={handleChange} 
                      disabled={isDenied || isHRCallIn} // Disable
                      className="accent-amber-600 disabled:opacity-50" 
                    />
                    <span className={`text-amber-800 ${isDenied ? 'opacity-70' : ''}`}>{val}</span>
                  </label>
                ))}
              </div>
            </div>

            {/* Incident & Office */}
            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-1">
                <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">Incident Type</label>
                <select name="type_of_incident" value={formData.type_of_incident} onChange={handleChange} disabled={isDenied || isHRCallIn} className="w-full border border-gray-300 p-2.5 rounded-md outline-none disabled:bg-gray-50 disabled:text-gray-500" required>
                  {incidentTypes.map(t => <option key={t} value={t}>{t}</option>)}
                </select>
              </div>
              <div className="space-y-1">
                <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">Office Location</label>
                <select name="office" value={formData.office} onChange={handleChange} disabled={isDenied || isHRCallIn} className="w-full border border-gray-300 p-2.5 rounded-md outline-none disabled:bg-gray-50 disabled:text-gray-500" required>
                  {OFFICES.map(o => <option key={o} value={o}>{o}</option>)}
                </select>
              </div>
            </div>

            {/* Dates */}
            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-1">
                <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">Start Date</label>
                <input name="incident_start" type="date" value={formData.incident_start} onChange={handleChange} disabled={isDenied || isHRCallIn} className="w-full border border-gray-300 p-2.5 rounded-md outline-none disabled:bg-gray-50" required />
              </div>
              <div className="space-y-1">
                <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">End Date</label>
                <input name="incident_end" type="date" value={formData.incident_end} onChange={handleChange} disabled={isDenied || isHRCallIn} className="w-full border border-gray-300 p-2.5 rounded-md outline-none disabled:bg-gray-50" required />
              </div>
            </div>

            {/* Times */}
            <div className="grid grid-cols-2 gap-4">
              {["Late In", "Leave and Come Back"].includes(formData.type_of_incident) && (
                <div className="space-y-1">
                  <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">ETA (Arrival)</label>
                  <input name="eta" type="time" value={formData.eta} onChange={handleChange} disabled={isDenied || isHRCallIn} className="w-full border border-gray-300 p-2.5 rounded-md outline-none disabled:bg-gray-50" required />
                </div>
              )}
              {["Early Out", "Leave and Come Back"].includes(formData.type_of_incident) && (
                <div className="space-y-1">
                  <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">ETD (Departure)</label>
                  <input name="etd" type="time" value={formData.etd} onChange={handleChange} disabled={isDenied || isHRCallIn} className="w-full border border-gray-300 p-2.5 rounded-md outline-none disabled:bg-gray-50" required />
                </div>
              )}
            </div>

            {/* Comments */}
            <div className="space-y-1">
              <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">Comments</label>
              <textarea name="employee_comments" value={formData.employee_comments} onChange={handleChange} disabled={isDenied} className="w-full border border-gray-300 p-2.5 rounded-md outline-none disabled:bg-gray-50" rows={2} />
            </div>

          {/* File Management Container */}
          <div className={`p-4 rounded-lg space-y-3 mt-1 ${isDenied ? 'bg-slate-50 border border-slate-200' : 'bg-blue-50 border border-blue-200'}`}>
            <p className={`text-xs font-bold uppercase ${isDenied ? 'text-slate-500' : 'text-blue-800'}`}>Notes & Files</p>

            {/* 1. Existing Uploaded Files Array */}
            {absence.excuse_note?.map((url: string, i: number) => {
              const decodedPath = decodeURIComponent(url.split('/o/')[1].split('?')[0]);
              const fileName = decodedPath.split('/').pop();
              const isMarkedForDeletion = notesToRemove.includes(url);

              return (
                <div key={i} className={`flex justify-between items-center bg-white p-2 rounded border border-slate-100 ${isMarkedForDeletion ? 'bg-rose-50/60 border-rose-100 text-slate-400' : 'bg-white border-slate-100 text-slate-700'}`}>
                  <a 
                    href={url} 
                    target="_blank" 
                    rel="noreferrer" 
                    className={`text-xs truncate max-w-[250px] font-medium text-blue-600 hover:underline ${isMarkedForDeletion ? 'line-through text-slate-400' : ''}`}
                  >
                    {fileName}
                  </a>
                  {!isDenied && (
                    <button type="button" onClick={() => handleToggleRemoveExistingNote(url)} className={`text-[10px] font-bold uppercase ml-2 px-2 py-1 rounded ${isMarkedForDeletion ? 'text-green-600 hover:bg-green-50' : 'text-red-500 hover:bg-red-50'}`}>
                      {isMarkedForDeletion ? 'Undo Delete' : 'Remove'}
                    </button>
                  )}
                </div>
              );
            })}

            {/* 2. Brand New Staged Files Preview List (Crucial Missing Feature!) */}
            {newExcuseNotes.length > 0 && (
              <div className="space-y-1.5 pt-1 border-t border-blue-100/50">
                <p className="text-[10px] font-bold text-blue-500 uppercase tracking-wider">Staged for Upload:</p>
                {newExcuseNotes.map((file, idx) => (
                  <div key={idx} className="flex justify-between items-center bg-blue-50/50 p-2 rounded border border-blue-100 animate-in fade-in-50 duration-200">
                    <span className="text-xs font-medium text-slate-700 truncate max-w-[250px]">{file.name}</span>
                    <button 
                      type="button" 
                      onClick={() => {
                        setNewExcuseNotes(prev => {
                          const updatedFiles = prev.filter((_, i) => i !== idx);
                          
                          // 💡 If there are no more new files staged, and no surviving old files, revert back
                          const existingNotesCount = (absence.excuse_note || []).length - notesToRemove.length;
                          if (updatedFiles.length === 0 && existingNotesCount === 0) {
                            setFormData(f => ({ ...f, excuse_note_submitted: 'not_provided' }));
                          }
                          
                          return updatedFiles;
                        });
                        setIsChecked(false);
                      }}
                      className="text-[10px] font-bold text-red-500 hover:bg-red-50 px-2 py-1 rounded uppercase"
                    >
                      Cancel
                    </button>
                  </div>
                ))}
              </div>
            )}

            {/* 3. Standard File Selector Input Trigger */}
            {!isDenied && (
            <div className="relative mt-2">
                <label 
                  htmlFor="file-upload" 
                  className="inline-block text-xs font-bold uppercase bg-blue-600 text-white py-1.5 px-4 rounded cursor-pointer shadow-sm hover:bg-blue-700 transition"
                >
                  Browse Files
                </label>
                <input 
                  id="file-upload" // 👈 Tied to the label above
                  type="file" 
                  multiple 
                  onChange={handleFileChange}
                  className="hidden" // 👈 Hides the ugly native browser text and "Choose File" button entirely!
                />
              </div>
            )}
          </div>

            {changed && !isDenied && (
              <div className="flex items-start gap-2 px-1 pt-2 animate-in fade-in slide-in-from-top-1 duration-300">
                <input 
                  type="checkbox" 
                  id="confirm-update" 
                  checked={isChecked} 
                  onChange={(e) => setIsChecked(e.target.checked)} 
                  className="mt-1 h-4 w-4 text-blue-600 rounded border-gray-300 focus:ring-blue-500 cursor-pointer" 
                />
                <label htmlFor="confirm-update" className="text-xs text-gray-600 leading-tight cursor-pointer">
                  I confirm these updates are accurate. <span className="font-bold text-gray-800">Note: This will reset the approval status to pending.</span>
                </label>
              </div>
            )}

            {/* Footer Buttons */}
            <div className="flex items-center justify-between mt-6 border-t pt-5">
              {/* ID on the left */}
              <div className="text-[10px] font-mono text-gray-400 uppercase tracking-widest">
                ID: <span className="select-all font-bold text-gray-500">{absence.id}</span>
              </div>

              {/* Buttons on the right */}
              <div className="flex gap-3">
                <button 
                  type="button" 
                  onClick={onClose} 
                  className="px-5 py-2 text-sm font-medium text-gray-600 hover:bg-gray-100 rounded-md transition"
                >
                  {isDenied ? 'Close' : 'Cancel'}
                </button>
                
                {!isDenied && (
                  <button 
                    type="submit" 
                    // Logic: Button is enabled ONLY if data changed AND user checked the box
                    disabled={!changed || !isChecked || isSubmitting} 
                    className="px-5 py-2 text-sm font-medium bg-blue-600 text-white rounded-md hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition shadow-sm"
                  >
                    {isSubmitting ? 'Saving...' : 'Save Changes'}
                  </button>
                )}
              </div>
            </div>
          </form>
        </div>
      </div>

      <FeedbackModal 
        isOpen={feedback.isOpen} 
        type={feedback.type} 
        message={feedback.message} 
        onClose={() => { 
          setFeedback(f => ({ ...f, isOpen: false })); 
          if (feedback.type === 'success') onFormSubmit(); 
        }} 
      />
    </>
  );
}