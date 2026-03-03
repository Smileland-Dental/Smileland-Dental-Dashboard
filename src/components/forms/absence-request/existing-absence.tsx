"use client";

import { getPSTDate } from '@/lib/date';
import React, { useState } from "react";
import { doc, updateDoc, deleteDoc } from "firebase/firestore";
import { db, storage } from '@/lib/firebase.config';
import { ref, uploadBytes, getDownloadURL, listAll, deleteObject } from "firebase/storage";
import FeedbackModal from '@/components/ui/FeedbackModal';

const incidentTypes = ["Late In", "Early Out", "Absent", "Leave and Come Back", "Long Lunch", "Switch Shift"];
const officeLocations = ["Corporate", "Ming", "Bernard", "California", "Ortho", "Delano", "Tulare", "Visalia", "Fresno"];

export default function ExistingAbsenceForm({ absence, onFormSubmit, onClose }: { absence: any, onFormSubmit: () => void, onClose: () => void }) {
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

  // Helper to determine if the form is different from the original database record
  const isDataChanged = () => {
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
    const { name, value } = e.target;

    if (name === "type_of_incident") {
      // Restore original times if switching back to original type
      if (value === absence.type_of_incident) {
        setFormData(prev => ({
          ...prev,
          [name]: value,
          eta: absence.eta || '',
          etd: absence.etd || ''
        }));
      } else {
        setFormData(prev => ({ ...prev, [name]: value, eta: '', etd: '' }));
      }
    } else {
      setFormData(prev => ({ ...prev, [name]: value }));
    }
    setIsChecked(false);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      setNewExcuseNotes(prev => [...prev, ...Array.from(e.target.files!)]);
      setIsChecked(false);
    }
  };

  const handleToggleRemoveExistingNote = (url: string) => {
    setNotesToRemove(prev => prev.includes(url) ? prev.filter(u => u !== url) : [...prev, url]);
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

      // --- 4. UPDATE FIRESTORE ---
      await updateDoc(doc(db, "absences", absence.id), {
        ...submissionData,
        excuse_note: finalNotes,
        manager_approval: 'pending',
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
            <h2 className="text-2xl font-bold text-gray-800">Edit Absence Request</h2>
            <button type="button" onClick={handleDelete} className="text-xs font-bold text-red-500 hover:bg-red-50 px-3 py-1 rounded border border-red-200 uppercase tracking-tighter">
              Delete Request
            </button>
          </div>
          
          <form onSubmit={handleSubmit} className="space-y-5">
            {/* Request Type */}
            <div className="p-4 bg-amber-50 rounded-lg border border-amber-200">
              <label className="block text-sm font-semibold text-amber-900 mb-2">Type Of Request *</label>
              <div className="flex gap-6 text-sm">
                {["Incident Notice", "Time Off Request"].map(val => (
                  <label key={val} className="flex items-center gap-2 cursor-pointer">
                    <input type="radio" name="type_of_request" value={val} checked={formData.type_of_request === val} onChange={handleChange} className="accent-amber-600" />
                    <span className="text-amber-800">{val}</span>
                  </label>
                ))}
              </div>
            </div>

            {/* Incident & Office */}
            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-1">
                <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">Incident Type</label>
                <select name="type_of_incident" value={formData.type_of_incident} onChange={handleChange} className="w-full border border-gray-300 p-2.5 rounded-md outline-none" required>
                  {incidentTypes.map(t => <option key={t} value={t}>{t}</option>)}
                </select>
              </div>
              <div className="space-y-1">
                <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">Office Location</label>
                <select name="office" value={formData.office} onChange={handleChange} className="w-full border border-gray-300 p-2.5 rounded-md outline-none" required>
                  {officeLocations.map(o => <option key={o} value={o}>{o}</option>)}
                </select>
              </div>
            </div>

            {/* Dates */}
            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-1">
                <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">Start Date</label>
                <input name="incident_start" type="date" value={formData.incident_start} onChange={handleChange} className="w-full border border-gray-300 p-2.5 rounded-md outline-none" required />
              </div>
              <div className="space-y-1">
                <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">End Date</label>
                <input name="incident_end" type="date" value={formData.incident_end} onChange={handleChange} className="w-full border border-gray-300 p-2.5 rounded-md outline-none" required />
              </div>
            </div>

            {/* Times */}
            <div className="grid grid-cols-2 gap-4">
              {["Late In", "Leave and Come Back"].includes(formData.type_of_incident) && (
                <div className="space-y-1">
                  <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">ETA (Arrival)</label>
                  <input name="eta" type="time" value={formData.eta} onChange={handleChange} className="w-full border border-gray-300 p-2.5 rounded-md outline-none" required />
                </div>
              )}
              {["Early Out", "Leave and Come Back"].includes(formData.type_of_incident) && (
                <div className="space-y-1">
                  <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">ETD (Departure)</label>
                  <input name="etd" type="time" value={formData.etd} onChange={handleChange} className="w-full border border-gray-300 p-2.5 rounded-md outline-none" required />
                </div>
              )}
            </div>

            {/* Comments */}
            <div className="space-y-1">
              <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">Comments</label>
              <textarea name="employee_comments" value={formData.employee_comments} onChange={handleChange} className="w-full border border-gray-300 p-2.5 rounded-md outline-none" rows={2} />
            </div>

            {/* File Management */}
            <div className="p-4 bg-blue-50 border border-blue-200 rounded-lg space-y-3 mt-1">
              <p className="text-xs font-bold text-blue-800 uppercase">Notes & Files</p>
        
              {/* List Existing Files from Database */}
              {absence.excuse_note?.map((url: string, i: number) => {
                // Extract filename from Firebase URL
                // 1. Get everything between /o/ and ?alt=media
                // 2. Decode the URI (converts %20 back to spaces, etc)
                // 3. Take the last part after the last "/"
                const decodedPath = decodeURIComponent(url.split('/o/')[1].split('?')[0]);
                const fileName = decodedPath.split('/').pop();

                return (
                  <div key={i} className="flex justify-between items-center bg-white p-2 rounded border border-blue-100">
                    <a 
                      href={url} 
                      target="_blank" 
                      rel="noreferrer" 
                      className={`text-xs truncate max-w-[250px] font-medium ${
                        notesToRemove.includes(url) ? 'line-through text-red-400' : 'text-blue-600 hover:underline'
                      }`}
                    >
                      {fileName}
                    </a>
                    <button 
                      type="button" 
                      onClick={() => handleToggleRemoveExistingNote(url)} 
                      className="text-[10px] font-bold text-red-500 uppercase ml-2 hover:bg-red-50 px-2 py-1 rounded"
                    >
                      {notesToRemove.includes(url) ? 'Undo' : 'Remove'}
                    </button>
                  </div>
                );
              })}

              {/* Show names of NEW files selected but not yet uploaded */}
              {newExcuseNotes.length > 0 && (
                <div className="space-y-1">
                  <p className="text-[10px] font-bold text-gray-500 uppercase">Pending Upload:</p>
                  {newExcuseNotes.map((file, idx) => (
                    <div key={idx} className="text-xs text-green-600 flex items-center justify-between bg-green-50 p-1 px-2 rounded">
                      <span className="truncate max-w-[200px]">{file.name}</span>
                      <button 
                        type="button" 
                        onClick={() => setNewExcuseNotes(prev => prev.filter((_, i) => i !== idx))}
                        className="text-red-500 text-[10px]"
                      >
                        ✕
                      </button>
                    </div>
                  ))}
                </div>
              )}

              <input 
                type="file" 
                multiple 
                onChange={handleFileChange} 
                className="text-xs w-full file:mr-4 file:py-1 file:px-3 file:rounded file:border-0 file:bg-blue-600 file:text-white file:cursor-pointer" 
              />
            </div>

            {/* Confirmation Section */}
            {changed && (
              <div className="flex items-start gap-2 pt-2 animate-in fade-in duration-300">
                <input 
                  type="checkbox" 
                  id="confirm" 
                  checked={isChecked} 
                  onChange={(e) => setIsChecked(e.target.checked)} 
                  className="mt-1 h-4 w-4 text-blue-600 rounded border-gray-300" 
                  required 
                />
                <label htmlFor="confirm" className="text-xs text-gray-600 leading-tight">
                  I confirm these updates are accurate. This will restart the approval process.
                </label>
              </div>
            )}

            {/* Footer Buttons */}
            <div className="flex justify-end gap-3 mt-6 border-t pt-5">
              <button type="button" onClick={onClose} className="px-5 py-2 text-sm font-medium text-gray-600 hover:bg-gray-100 rounded-md transition">
                Cancel
              </button>
              <button 
                type="submit" 
                // Button is disabled if: nothing changed, OR if changed but not checked, OR currently submitting
                disabled={!changed || (changed && !isChecked) || isSubmitting} 
                className="px-5 py-2 text-sm font-medium bg-blue-600 text-white rounded-md hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition"
              >
                {isSubmitting ? 'Saving...' : 'Save Changes'}
              </button>
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