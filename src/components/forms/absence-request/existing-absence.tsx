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

export default function ExistingAbsenceForm({ absence, onFormSubmit, onClose }: { absence: any, onFormSubmit: () => void, onClose: () => void }) {
  if (!absence) return null;
  // --- 1. DEFINE DENIED STATE ---
  const isDenied = absence.manager_approval === 'denied' || absence.final_approval === 'denied';
  const isApproved = absence.final_approval === 'approved' || absence.final_approval === 'approved_with_note';
  const isHRCallIn = absence.type_of_request === "HR Call In";
  const isPendingAction = absence.status === 'pending_action';

  // Master disabled rules flags
  const disableCoreInputs = isDenied || isApproved || isHRCallIn;
  const disableCommentsAndFiles = isDenied;

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
    final_notes: absence.final_notes || '',
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

  const isDateInvalid = Boolean(
    formData.incident_start && 
    formData.incident_end && 
    new Date(formData.incident_end) < new Date(formData.incident_start)
  );

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

      // CHECK HERE

    const hasFileChanges = newExcuseNotes.length > 0 || notesToRemove.length > 0;
    return hasFieldChanges || hasFileChanges;
  };

const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>) => {
    if (isDenied) return; // Prevent state updates if denied
    const { name, value } = e.target;

    if (isApproved && name !== "employee_comments") return;

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
    if (disableCommentsAndFiles) return;
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
    if (!window.confirm("Are you sure you want to cancel this request?")) return;
    setIsSubmitting(true);
    try {
      await updateDoc(doc(db, "absences", absence.id), {
        status: 'archived',
        updatedAt: new Date(),
      });

      setFeedback({ isOpen: true, type: 'success', message: "Request has been canceled and archived successfully!" });
    } catch (error) {
      setFeedback({ isOpen: true, type: 'error', message: "Error canceling request." });
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleSubmit = async (event: React.FormEvent) => {
    event.preventDefault();
    if (!isDataChanged() || !isChecked) return;

    if (isDateInvalid) {
      setFeedback({ isOpen: true, type: 'error', message: "Dates Invalid" });
      return;
    }

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

      // --- 4. CONDITIONAL APPROVAL STATUS UPDATE ---
      // 💡 If already approved, do not overwrite the approvals with pending!
      const approvalPayload = isApproved ? {
        manager_approval: absence.manager_approval,
        manager_approval_name: absence.manager_approval_name || '',
        final_approval: absence.final_approval,
        final_approval_name: absence.final_approval_name || '',
      } : {
        manager_approval: isHRCallIn ? 'not_required' : 'pending',
        manager_approval_name: '',
        final_approval: 'pending',
        final_approval_name: '',
      };

      // --- 5. UPDATE FIRESTORE ---
      await updateDoc(doc(db, "absences", absence.id), {
        ...submissionData,
        excuse_note: finalNotes,
        ...approvalPayload,
        status: isPendingAction ? 'active' : (absence.status || 'active'), // 💡 Automatically activates it
        updatedAt: new Date(), 
      });

      setFeedback({ isOpen: true, type: 'success', message: isPendingAction ? "Notice acknowledged successfully!" : "Request updated successfully!" });
      // Reset removal tracking after success
      setNotesToRemove([]);
      setNewExcuseNotes([]);
    } catch (error) {
      setFeedback({ isOpen: true, type: 'error', message: "Error updating request." });
    } finally {
      setIsSubmitting(false);
    }
  };

  const changed = isPendingAction ? true : isDataChanged();
  const activeExistingNotesCount = (absence.excuse_note || []).length - notesToRemove.length;

  return (
    <>
      <div className="fixed inset-0 bg-black/50 flex justify-center items-center z-50" onClick={onClose}>
        <div className="bg-white p-8 rounded-lg shadow-xl w-full max-w-2xl max-h-[90vh] overflow-y-auto" onClick={(e) => e.stopPropagation()}>
          
          <div className="flex justify-between items-center mb-6">
            <h2 className="text-2xl font-bold text-gray-800">
              {(isDenied || isApproved) ? 'View Absence Request' : 'Edit Absence Request'}
            </h2>
            {absence.status === 'archived' && (
              <h3 className="text-lg font-bold text-red-700 uppercase">Archived</h3>
            )}

            {/* 2. HIDE DELETE IF DENIED */}
            {!isDenied && absence.status !== 'archived' && (
              <div className="flex flex-col gap-1.5">
              <button type="button" onClick={handleDelete} className="text-xs font-bold text-red-500 hover:bg-red-50 px-3 py-1 rounded border border-red-200 uppercase tracking-tighter disabled:opacity-50 disabled:cursor-not-allowed disabled:border-gray-200 disabled:text-gray-400"
                disabled={absence.type_of_request === "HR Call In" || absence.final_approval === "approved" || absence.final_approval === "denied" || isSubmitting}
              >
                Cancel Request
              </button>

              {(absence.type_of_request === "HR Call In" || isApproved || absence.final_approval === "denied" || isSubmitting) && (
                <p className="text-[10px] text-slate-400 italic text-right">Contact EXEC to cancel.</p>
              )}
              </div>
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

          {isApproved && (
            <div className="mb-6 p-4 bg-emerald-50 border border-emerald-200 rounded-xl flex items-center gap-3">
              <Lock className="h-5 w-5 text-emerald-500" />
              <div>
                <p className="text-sm font-bold text-emerald-800">Request Approved</p>
                <p className="text-xs text-emerald-600">The request is locked. You may still write notes or append additional files.</p>
              </div>
            </div>
          )}
          
          <form onSubmit={handleSubmit} className="space-y-5">
            {/* Request Type */}
            <div className="p-4 bg-amber-50 rounded-lg border border-amber-200">
              <label className="block text-sm font-semibold text-amber-900 mb-2">Type Of Request</label>
              <div className="flex flex-wrap gap-6 text-sm">
                {["Incident Notice", "Time Off Request", "HR Call In", "No Call", "Call In After Shift", "Previously Not Approved"].map(val => {
                  const specialTypes = ["HR Call In", "No Call", "Call In After Shift", "Previously Not Approved"]
                  const isSpecialType = specialTypes.includes(absence.type_of_request);
                  const isRadioDisabled = disableCoreInputs || isSpecialType || (specialTypes.includes(val) && !specialTypes.includes(formData.type_of_request));
                  return (
                  <label key={val} className="flex items-center gap-2 cursor-pointer">
                    <input 
                      type="radio" 
                      name="type_of_request" 
                      value={val} 
                      checked={formData.type_of_request === val} 
                      onChange={handleChange} 
                      disabled={isRadioDisabled} // Disable
                      className="accent-amber-600 disabled:opacity-50" 
                    />
                    <span className={`text-amber-800 ${(isRadioDisabled) ? 'opacity-40' : ''}`}>{val}</span>
                  </label>
                );
                })}
              </div>
            </div>

            {/* Incident & Office */}
            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-1">
                <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">Incident Type</label>
                <select name="type_of_incident" value={formData.type_of_incident} onChange={handleChange} disabled={disableCoreInputs} className="w-full border border-gray-300 p-2.5 rounded-md outline-none disabled:bg-gray-50 disabled:text-gray-500" required>
                  {incidentTypes.map(t => <option key={t} value={t}>{t}</option>)}
                </select>
              </div>
              <div className="space-y-1">
                <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">Office Location</label>
                <select name="office" value={formData.office} onChange={handleChange} disabled={disableCoreInputs} className="w-full border border-gray-300 p-2.5 rounded-md outline-none disabled:bg-gray-50 disabled:text-gray-500" required>
                  {OFFICES.map(o => <option key={o} value={o}>{o}</option>)}
                </select>
              </div>
            </div>

            {/* Dates */}
            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-1">
                <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">Start Date</label>
                <input name="incident_start" type="date" value={formData.incident_start} onChange={handleChange} disabled={disableCoreInputs} className="w-full border border-gray-300 p-2.5 rounded-md outline-none disabled:bg-gray-50" required />
              </div>
              <div className="space-y-1">
                <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">End Date</label>
                <input name="incident_end" type="date" min={formData.incident_start || ""} value={formData.incident_end} onChange={handleChange} disabled={disableCoreInputs} className={`w-full border border-gray-300 p-2.5 rounded-md outline-none disabled:bg-gray-50 ${isDateInvalid ? 'bg-rose-50 text-rose-600 ring-2 ring-rose-500' : 'bg-white'}`} required />
              </div>
            </div>

            {/* Date Validation Error Message */}
            {isDateInvalid && (
              <div className="flex items-center gap-2 p-3 bg-rose-50 rounded-xl text-rose-600 text-xs font-bold border border-rose-200 animate-in fade-in">
                <AlertCircle className="w-4 h-4 shrink-0" />
                <span>End date cannot be earlier than the start date.</span>
              </div>
            )}

            {/* Times */}
            <div className="grid grid-cols-2 gap-4">
              {["Late In", "Leave and Come Back"].includes(formData.type_of_incident) && (
                <div className="space-y-1">
                  <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">ETA (Arrival)</label>
                  <input name="eta" type="time" value={formData.eta} onChange={handleChange} disabled={disableCoreInputs} className="w-full border border-gray-300 p-2.5 rounded-md outline-none disabled:bg-gray-50" required />
                </div>
              )}
              {["Early Out", "Leave and Come Back"].includes(formData.type_of_incident) && (
                <div className="space-y-1">
                  <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">ETD (Departure)</label>
                  <input name="etd" type="time" value={formData.etd} onChange={handleChange} disabled={disableCoreInputs} className="w-full border border-gray-300 p-2.5 rounded-md outline-none disabled:bg-gray-50" required />
                </div>
              )}
            </div>

            {/* Comments */}
            <div className="space-y-1">
              <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">
                Employee Comments
                {isPendingAction && <span className="text-red-500 font-bold"> *Required</span>}
              </label>
              <textarea 
                name="employee_comments" 
                value={formData.employee_comments} 
                onChange={handleChange} 
                disabled={disableCommentsAndFiles} 
                required={isPendingAction} // 💡 Dynamically requires input
                placeholder={'Notes and Comments...'}
                className="w-full border border-gray-300 p-2.5 rounded-md outline-none disabled:bg-gray-50 focus:ring-1 focus:ring-blue-500" 
                rows={2} 
              />
            </div>

          {/* File Management Container */}
          <div className={`p-4 rounded-lg space-y-3 mt-1 ${disableCommentsAndFiles ? 'bg-slate-50 border border-slate-200' : 'bg-blue-50 border border-blue-200'}`}>
            <p className={`text-xs font-bold uppercase ${disableCommentsAndFiles ? 'text-slate-500' : 'text-blue-800'}`}>Notes & Files</p>

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
                  {!disableCommentsAndFiles && (
                    <button type="button" onClick={() => handleToggleRemoveExistingNote(url)} className={`text-[10px] font-bold uppercase ml-2 px-2 py-1 rounded ${isMarkedForDeletion ? 'text-green-600 hover:bg-green-50' : 'text-red-500 hover:bg-red-50'}`}>
                      {isMarkedForDeletion ? 'Undo Delete' : 'Remove'}
                    </button>
                  )}
                </div>
              );
            })}

            {/* 2. Brand New Staged Files Preview List*/}
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
            {!disableCommentsAndFiles && (    
              <div className="flex items-center gap-3 mt-2">
                <label 
                  htmlFor="file-upload" 
                  className={`inline-block text-xs font-bold uppercase py-1.5 px-4 rounded shadow-sm transition ${
                    formData.excuse_note_submitted === 'not_provided'
                      ? 'bg-gray-300 text-gray-500 cursor-not-allowed pointer-events-none shadow-none'
                      : 'bg-blue-600 text-white cursor-pointer hover:bg-blue-700'
                  }`}
                >
                  Browse Files
                </label>
                <input 
                  id="file-upload" 
                  type="file" 
                  multiple 
                  disabled={formData.excuse_note_submitted === 'not_provided'}
                  onChange={handleFileChange}
                  className="hidden" 
                />
                {/* The "Do not submit an excuse note" markup from Step 1 goes right here */}
                {activeExistingNotesCount === 0 && (
                <label 
                  htmlFor="opt-out-notes" 
                  className={`flex items-center gap-2 px-3 py-1.5 rounded-md border cursor-pointer select-none text-xs font-bold transition shadow-xs ${
                    formData.excuse_note_submitted === 'not_provided' 
                      ? 'bg-blue-100 border-blue-300 text-blue-900' 
                      : 'bg-white/80 border-blue-200/60 text-slate-700 hover:bg-white'
                  }`}
                >
                  <input 
                    id="opt-out-notes"
                    type="checkbox" 
                    checked={formData.excuse_note_submitted === 'not_provided'}
                    onChange={(e) => {
                      const checked = e.target.checked;
                      setFormData(prev => ({
                        ...prev,
                        excuse_note_submitted: checked ? 'not_provided' : 'pending'
                      }));
                      // Clear out any new files if they opt out
                      if (checked) {
                        setNewExcuseNotes([]);
                      }
                      setIsChecked(false);
                    }}
                    className="h-3.5 w-3.5 rounded text-blue-600 border-gray-300 focus:ring-blue-500 cursor-pointer accent-blue-600"
                  />
                  <span>Do not submit an excuse note</span>
                </label> )}
              </div>
            )}
          </div>

          {formData.final_notes && (
            <div className="space-y-1">
              <label className="text-xs font-bold text-rose-700 uppercase tracking-wider">
                Final Notes
              </label>
              <p className="w-full border font-semibold bg-slate-50 border-gray-300 p-2.5 rounded-md outline-none disabled:bg-gray-50 focus:ring-1 focus:ring-blue-500">
                {formData.final_notes} 
              </p>
            </div>
          )}

            {/* Checkbox text modification */}
            {changed && !isDenied && (
              <div className="flex items-start gap-2 px-1 pt-2 animate-in fade-in slide-in-from-top-1 duration-300">
                <input 
                  type="checkbox" 
                  id="confirm-update" 
                  checked={isChecked} 
                  onChange={(e) => setIsChecked(e.target.checked)} 
                  className="h-4 w-4 text-blue-600 rounded border-gray-300 focus:ring-blue-500 cursor-pointer" 
                />
                <label htmlFor="confirm-update" className="text-xs text-gray-600 leading-tight cursor-pointer">
                  {isPendingAction ? (
                    <span>I have read and acknowledge this absence record.</span>
                  ) : isApproved ? (
                    <span>I confirm these supplemental notes/files are accurate. <span className="font-bold text-gray-800">Note: This will preserve your current approval status.</span></span>
                  ) : (
                    <span>I confirm these updates are accurate. <span className="font-bold text-gray-800">Note: This will reset the approval status to pending.</span></span>
                  )}
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
                
                {/* Submit Button modification */}
                {!isDenied && (
                  <button 
                    type="submit" 
                    disabled={!changed || !isChecked || isSubmitting || isDateInvalid} 
                    className="px-5 py-2 text-sm font-medium bg-blue-600 text-white rounded-md hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition shadow-sm"
                  >
                    {isSubmitting ? 'Saving...' : isPendingAction ? 'Acknowledge Notice' : 'Save Changes'}
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