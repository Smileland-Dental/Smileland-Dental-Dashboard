"use client";
import { getPSTDate } from '@/lib/date';
import React, { useMemo, useState } from "react";
import { doc, updateDoc, addDoc, collection } from "firebase/firestore";
import { db, storage } from '@/lib/firebase.config';
import { ref, uploadBytes, getDownloadURL } from "firebase/storage";
import FeedbackModal from '@/components/ui/FeedbackModal'; 
import { Button } from '@/components/ui/button';
import { OFFICES } from '@/lib/constants';
import { AbsenceRequest } from '@/lib/types';
import { TriangleAlert, AlertCircle } from 'lucide-react';

const incidentTypes = ["Late In", "Early Out", "Absent", "Leave and Come Back", "Long Lunch", "Switch Shift"];

interface NewAbsenceFormProps {
  employeeFirestore: string;
  employeeID: string;
  employeeTitle: string;
  employeeName: string;
  employeeSkipManagerApproval: boolean;
  employeeExistingRequests: AbsenceRequest[]; // Pass existing requests to check for overlaps
  onFormSubmit: () => void;
  onClose: () => void;
}

export default function NewAbsenceForm({employeeFirestore, employeeID, employeeTitle, employeeName, employeeSkipManagerApproval, employeeExistingRequests, onFormSubmit, onClose }: NewAbsenceFormProps) {
  const [formData, setFormData] = useState({
    employeeFirestoreID: employeeFirestore || '',
    employee_id: employeeID || '',
    employee_title: employeeTitle || '',
    employee_name: employeeName || '',
    type_of_request: '',
    type_of_incident: '',
    office: '',
    incident_start: '',
    incident_end: '',
    employee_comments: '',
    eta: '',
    etd: '',
    excuse_note_submitted: 'not_provided',
    manager_approval: employeeSkipManagerApproval ? 'not_required' : 'pending',
    manager_approval_name: '',
    final_approval: 'pending',
    final_approval_name: '', 
    final_notes: '',
    DAP: 0,
    skipManagerApproval: employeeSkipManagerApproval || false,
    DOAPoints: 0,
    status: 'active',
  });

  const [excuse_notes, setExcuseNotes] = useState<File[]>([]);
  const [isSubmitting, setIsSubmitting] = useState(false);

  const [isChecked, setIsChecked] = useState(false);
  const [feedback, setFeedback] = useState<{ isOpen: boolean; type: 'success' | 'error'; message: string }>({
    isOpen: false,
    type: 'success',
    message: '',
  });

  const isOverlapping = useMemo(() => {
    if (!formData.type_of_incident || !formData.incident_start || !formData.incident_end) {
      return false;
    }

    const formStart = new Date(formData.incident_start + "T00:00:00").getTime();
    const formEnd = new Date(formData.incident_end + "T00:00:00").getTime();

    // Look through array instances currently stored in parent state
    return employeeExistingRequests.some((absence) => {
      if (absence.type_of_incident !== formData.type_of_incident) {
        return false;
      }
      const isConflictingStatus = ['active', 'pending_action'].includes(absence.status);
      if (!isConflictingStatus) {
        return false; // Skips 'archived'
      }

      const existingStart = new Date(absence.incident_start + "T00:00:00").getTime();
      const existingEnd = new Date(absence.incident_end + "T00:00:00").getTime();

      // Check date interval overlap logic: (StartA <= EndB) AND (EndA >= StartB)
      return formStart <= existingEnd && formEnd >= existingStart;
    });
  }, [formData.type_of_incident, formData.incident_start, formData.incident_end, employeeExistingRequests]);

  const isDateInvalid = Boolean(
    formData.incident_start && 
    formData.incident_end && 
    new Date(formData.incident_end) < new Date(formData.incident_start)
  );

  const isDateWindowValid = useMemo(() => {
    // If fields are empty, keep it valid (or fallback to HTML5 field requirements)
    if (!formData.type_of_request || !formData.incident_start) return true;

    // Set up standard Date objects set exactly to Midnight
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const startDate = new Date(formData.incident_start + "T00:00:00");

    // Calculate absolute difference in days
    const msPerDay = 24 * 60 * 60 * 1000;
    const dayDifference = (startDate.getTime() - today.getTime()) / msPerDay;

    if (formData.type_of_request === "Incident Notice") {
      // Must be within 30 days of today (including past dates or less 30 days into the future)
      return dayDifference < 30;
    }

    if (formData.type_of_request === "Time Off Request") {
      // Must be 30 days or more in the future
      return dayDifference >= 30;
    }
    
    return true; // Default to valid if type_of_request is not selected yet
  }, [formData.type_of_request, formData.incident_start]);

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>) => {
    const { name, value } = e.target;

    if (name === 'type_of_incident') {
      setFormData(prev => ({
        ...prev,
        [name]: value,

        // Reset time fields so old data doesn't persist in the database if user switches between incident types
        eta: '',
        etd: '',
      }));
    }
    else if (name === 'type_of_request') {
      setFormData(prev => ({
        ...prev,
        [name]: value,

        // Reset Excuse Note to 'not_provided' if switching requests so 'N/A' doesn't exist on Incident Notice
        excuse_note_submitted: 'not_provided',
      }))
    }
    else{
      setFormData(prev => ({ ...prev, [name]: value }));
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      setExcuseNotes(prev => [...prev, ...Array.from(e.target.files!)]);
    }
  };

  const handleRemoveNewFile = (fileName: string) => {
    setExcuseNotes(prev => prev.filter(file => file.name !== fileName));
  };

  const handleSubmit = async (event: React.FormEvent) => {
    event.preventDefault();
    if (!isChecked) return; 

    if (isDateInvalid) {
      console.error("End date cannot be before start date.");
      return;
    }

    setIsSubmitting(true);

    try {
      // 1. Create the doc
      const absenceRef = await addDoc(collection(db, "absences"), {
        ...formData,
        final_approval: 'pending',
        createdAt: new Date(),
        excuse_note: [], 
      });

      // 2. Handle File Uploads
      if (excuse_notes.length > 0 && formData.excuse_note_submitted === 'submitted') {
        const uploadedFileUrls: string[] = [];
        for (const file of excuse_notes) {
          const storageRef = ref(storage, `excuse-notes/${absenceRef.id}/${file.name}`);
          await uploadBytes(storageRef, file);
          const url = await getDownloadURL(storageRef);
          uploadedFileUrls.push(url);
        }
        // 3. Update doc with URLs
        await updateDoc(doc(db, "absences", absenceRef.id), { excuse_note: uploadedFileUrls });
      }

      setFeedback({
        isOpen: true,
        type: 'success',
        message: "Absence request submitted successfully!",
      });
    } catch (error) {
      console.error("Submission Error:", error);
      setFeedback({
        isOpen: true,
        type: 'error',
        message: "Error submitting request. Please try again.",
      });
    } finally {
      setIsSubmitting(false);
    }
  };

  //const showETA_ETDField = ["Late In", "Early Out"].includes(formData.type_of_incident);

  return (
    <>
      <div className="fixed inset-0 bg-black/50 flex justify-center items-center z-50" onClick={onClose}>
        <div className="bg-white p-8 rounded-lg shadow-xl w-full max-w-2xl max-h-[90vh] overflow-y-auto" onClick={(e) => e.stopPropagation()}>
          <h2 className="text-2xl font-bold mb-6 text-gray-800">New Absence Request</h2>
          {isOverlapping && (
            <div className="p-3.5 bg-red-50 border border-red-200 text-red-900 rounded-md text-xs font-medium mt-2">
              <TriangleAlert className="inline-block size-3 mr-1 mb-1" />
              <strong>Schedule Conflict:</strong> You already has an active <strong>{formData.type_of_incident}</strong> request registered within these dates.
            </div>
          )}
          
          <form onSubmit={handleSubmit} className="space-y-5">
            {/* Request Type Selection */}
            <div className="p-4 bg-amber-50 rounded-lg border border-amber-200">
              <label className="block text-sm font-semibold text-amber-900 mb-2">Type Of Request *</label>
              <div className="flex gap-6 text-sm">
                <label className="flex items-center gap-2 cursor-pointer">
                  <input type="radio" name="type_of_request" value="Incident Notice" onChange={handleChange} required className="accent-amber-600" />
                  <span className="text-amber-800">Incident Notice</span>
                </label>
                <label className="flex items-center gap-2 cursor-pointer">
                  <input type="radio" name="type_of_request" value="Time Off Request" onChange={handleChange} className="accent-amber-600" />
                  <span className="text-amber-800">Time-Off Request</span>
                </label>
              </div>
            </div>

            {/* Row 1: Incident & Office */}
            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-1">
                <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">Incident Type</label>
                <select name="type_of_incident" value={formData.type_of_incident} onChange={handleChange} className="w-full border border-gray-300 p-2.5 rounded-md focus:ring-2 focus:ring-blue-500 outline-none" required>
                  <option value="">Select Type</option>
                  {incidentTypes.map(t => <option key={t} value={t}>{t}</option>)}
                </select>
              </div>
              <div className="space-y-1">
                <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">Office Location</label>
                <select name="office" value={formData.office} onChange={handleChange} className="w-full border border-gray-300 p-2.5 rounded-md focus:ring-2 focus:ring-blue-500 outline-none" required>
                  <option value="">Select Office</option>
                  {OFFICES.map(o => <option key={o} value={o}>{o}</option>)}
                </select>
              </div>
            </div>

            {/* Row 2: Dates */}
            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-1">
                <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">Start Date</label>
                <input name="incident_start" type="date" onChange={handleChange} className="w-full border border-gray-300 p-2.5 rounded-md focus:ring-2 focus:ring-blue-500 outline-none" required />
                {!isDateWindowValid && (
                  <p className="text-xs text-red-500 font-medium mt-1">
                    {formData.type_of_request === "Incident Notice" 
                      ? "Incident Notice must be dated within 30 days of today."
                      : "Time-Off Requests require at least 30 days advance notice."}
                  </p>
                )}
              </div>
              <div className="space-y-1">
                <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">End Date</label>
                <input name="incident_end" type="date" min={formData.incident_start || ""} onChange={handleChange} className={`
                w-full border border-gray-300 p-2.5 rounded-md focus:ring-2 focus:ring-blue-500 outline-none ${isDateInvalid ? 'bg-rose-50 text-rose-600 ring-2 ring-rose-500' : 'bg-white'}`} required/>
              </div>
            </div>

                {/* Date Validation Error Message */}
                {isDateInvalid && (
                  <div className="flex items-center gap-2 p-3 bg-rose-50 rounded-xl text-rose-600 text-xs font-bold border border-rose-200 animate-in fade-in">
                    <AlertCircle className="w-4 h-4 shrink-0" />
                    <span>End date cannot be earlier than the start date.</span>
                  </div>
                )}

            {/* Row 3: Conditional Time Field for Late In*/}
            {(formData.type_of_incident === "Late In") && (
              <div className="space-y-1">
                <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">Estimated Arrival (ETA)</label>
                <input name="eta" type="time" onChange={handleChange} className="w-full border border-gray-300 p-2.5 rounded-md focus:ring-2 focus:ring-blue-500 outline-none" required />
              </div>
            )}

            {/* Row 3: Conditional Time Field for Early Out*/}
            {(formData.type_of_incident === "Early Out") && (
              <div className="space-y-1">
                <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">Estimated Departure (ETD)</label>
                <input name="etd" type="time" onChange={handleChange} className="w-full border border-gray-300 p-2.5 rounded-md focus:ring-2 focus:ring-blue-500 outline-none" required />
              </div>
            )}


            {/* Row 3: Conditional Time Field for Leave and Come Back */}
            {(formData.type_of_incident === "Leave and Come Back") && (
              <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">
                      Time Leaving (ETD)
                    </label>
                    <input 
                      name="etd" 
                      type="time" 
                      value={formData.etd}
                      onChange={handleChange} 
                      className="w-full border border-gray-300 p-2.5 rounded-md focus:ring-2 focus:ring-blue-500 outline-none" 
                      required 
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">
                      Time Returning (ETA)
                    </label>
                    <input 
                      name="eta" 
                      type="time" 
                      value={formData.eta}
                      onChange={handleChange} 
                      className="w-full border border-gray-300 p-2.5 rounded-md focus:ring-2 focus:ring-blue-500 outline-none" 
                      required 
                    />
                  </div>
                </div>
            )}

            {/* Comments */}
            <div className="space-y-1">
              <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">Additional Comments</label>
              <textarea name="employee_comments" onChange={handleChange} className="w-full border border-gray-300 p-2.5 rounded-md focus:ring-2 focus:ring-blue-500 outline-none" rows={2}/>
            </div>

            {/* --- EXCUSE NOTE SECTION --- */}
            <div className="p-4 bg-gray-50 border border-gray-200 rounded-lg space-y-4 mt-1">
              <label className="block text-sm font-semibold text-gray-700">Excuse Note Submission</label>
              <div className="flex flex-wrap gap-4 text-xs">
                {['submitted', 'pending', 'not_provided'].map((opt) => (
                  <label key={opt} className="flex items-center gap-1.5 cursor-pointer capitalize">
                    <input 
                      type="radio" 
                      name="excuse_note_submitted" 
                      value={opt} 
                      onChange={handleChange} 
                      checked={formData.excuse_note_submitted === opt}
                      className="accent-blue-600"
                    />
                    {opt.replace('not_provided', 'Not Providing')}
                  </label>
                ))}

                {/* Conditionally add the N/A option exclusively for Time Off Requests */}
                {formData.type_of_request === "Time Off Request" && (
                  <label className="flex items-center gap-1.5 cursor-pointer uppercase">
                    <input 
                      type="radio" 
                      name="excuse_note_submitted" 
                      value="na" 
                      onChange={handleChange} 
                      checked={formData.excuse_note_submitted === 'na'}
                      className="accent-blue-600"
                    />
                    N/A
                  </label>
                )}
              </div>

              {formData.excuse_note_submitted === 'submitted' && (
                <div className="mt-2 space-y-3">
                  <input
                    type="file"
                    multiple
                    onChange={handleFileChange}
                    className="text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-xs file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 w-full"
                  />
                  {excuse_notes.length > 0 && (
                    <ul className="space-y-2">
                      {excuse_notes.map((file, idx) => (
                        <li key={idx} className="flex justify-between items-center text-xs bg-white p-2 border rounded border-gray-200">
                          <span className="truncate max-w-[200px]">{file.name}</span>
                          <button type="button" onClick={() => handleRemoveNewFile(file.name)} className="text-red-500 hover:underline">Remove</button>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
              )}
            </div>

            {/* Confirmation */}
            <div className="flex items-start gap-2 pt-2">
              <input 
                type="checkbox" 
                id="confirm" 
                checked={isChecked} 
                onChange={(e) => setIsChecked(e.target.checked)} 
                className="mt-1 h-4 w-4 text-blue-600 rounded border-gray-300"
                required 
              />
              <label htmlFor="confirm" className="text-xs text-gray-600 leading-tight">
                I confirm that all information provided is accurate to the best of my knowledge.
              </label>
            </div>

            {/* Action Buttons */}
            <div className="flex justify-end gap-3 mt-6 border-t pt-5">
              <button type="button" onClick={onClose} className="px-5 py-2 text-sm font-medium text-gray-600 hover:bg-gray-100 rounded-md transition">
                Cancel
              </button>
              <button 
                type="submit" 
                disabled={isSubmitting || !isChecked || !isDateWindowValid ||  isOverlapping || isDateInvalid} 
                className="px-5 py-2 text-sm font-medium bg-blue-600 text-white rounded-md hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition"
              >
                {isSubmitting ? 'Submitting...' : 'Submit Request'}
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