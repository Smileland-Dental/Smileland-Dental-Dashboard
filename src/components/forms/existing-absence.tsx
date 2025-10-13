"use client";

import { getPSTDate } from '@/lib/date';
import React, { useState } from "react";
import { doc, updateDoc } from "firebase/firestore";
import { db, storage } from '@/lib/firebase.config';
import { ref, uploadBytes, getDownloadURL } from "firebase/storage";
import { Button } from '@/components/ui/button';

// Constants for select dropdowns
const incidentTypes = [
  "Late In", "Early Out", "Absent", "Leave and Come Back", "Long Lunch", "Switch Shift"
];

const officeLocations = [
  "Corporate", "Ming", "Bernard", "California", "Ortho", "Delano", "Tulare", "Visalia", "Fresno"
];

export default function ExistingAbsenceForm({ absence, onFormSubmit }: { absence: any, onFormSubmit: () => void }) {
  // State for editable fields
  const [comments, setComments] = useState(absence.employee_comments || "");
  const [newExcuseNotes, setNewExcuseNotes] = useState<File[]>([]);
  // NEW: State to track URLs of existing notes marked for removal
  const [notesToRemove, setNotesToRemove] = useState<string[]>([]);
  
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [submitStatus, setSubmitStatus] = useState("");

  const canAddNotes = absence.excuse_note_submitted === 'pending' || absence.excuse_note_submitted === 'submitted';

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setNewExcuseNotes(prev => [...prev, ...Array.from(e.target.files ?? [])]);
    }
  };

  // NEW: Handler to remove a newly selected file before upload
  const handleRemoveNewFile = (fileName: string) => {
    setNewExcuseNotes(prev => prev.filter(file => file.name !== fileName));
  };

  // NEW: Handler to mark an existing note for removal
  const handleToggleRemoveExistingNote = (url: string) => {
    setNotesToRemove(prev => 
      prev.includes(url) 
        ? prev.filter(u => u !== url) 
        : [...prev, url]
    );
  };

  // NEW: Handler for changing status to 'not_provided'
  const handleWillNotSubmit = async () => {
    if (window.confirm("Are you sure you want to mark this incident as having no excuse note? This action cannot be undone.")) {
      setIsSubmitting(true);
      setSubmitStatus("Updating status...");
      try {
        const absenceRef = doc(db, "absences", absence.id);
        await updateDoc(absenceRef, {
          excuse_note_submitted: 'not_provided',
          date_submitted: getPSTDate(),
        });
        setSubmitStatus("Status updated successfully!");
        setTimeout(() => onFormSubmit(), 1500);
      } catch (error) {
        console.error("Error updating status: ", error);
        setSubmitStatus("Failed to update status.");
        setIsSubmitting(false);
      }
    }
  };

  const handleSubmit = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setIsSubmitting(true);
    setSubmitStatus("Submitting...");

    try {
      let newNoteUrls: string[] = [];
      
      // 1. Upload new files if any have been selected
      if (newExcuseNotes.length > 0) {
        setSubmitStatus("Uploading new files...");
        const uploadPromises = newExcuseNotes.map(async (file) => {
          const storageRef = ref(storage, `excuse-notes/${absence.id}/${file.name}`);
          await uploadBytes(storageRef, file);
          return getDownloadURL(storageRef);
        });
        newNoteUrls = await Promise.all(uploadPromises);
      }
      
      // 2. Determine the final list of notes for Firestore
      const existingNotes = absence.excuse_note || [];
      const remainingNotes = existingNotes.filter((url: string) => !notesToRemove.includes(url));
      const finalNotes = [...remainingNotes, ...newNoteUrls];

      const dataToUpdate: { [key: string]: any } = {
        employee_comments: comments,
        date_submitted: getPSTDate(),
        excuse_note: finalNotes,
      };

      // 3. Update status to 'submitted' if it was pending and notes are now present
      if (absence.excuse_note_submitted === 'pending' && finalNotes.length > 0) {
        dataToUpdate.excuse_note_submitted = 'submitted';
      }

      setSubmitStatus("Updating record...");
      const absenceRef = doc(db, "absences", absence.id);
      await updateDoc(absenceRef, dataToUpdate);
      
      // Note: For a complete solution, you would also delete files from Storage here
      // using the `notesToRemove` array and Firebase's deleteObject function.

      setSubmitStatus("Report updated successfully!");
      setTimeout(() => onFormSubmit(), 1500);

    } catch (error) {
      console.error("Error updating document: ", error);
      setSubmitStatus("Error updating report. Please try again.");
      setIsSubmitting(false);
    }
  };

  const goBack = () => onFormSubmit();

  const showETA_ETDField = ["Late In", "Early Out", "Leave and Come Back"].includes(absence.type_of_incident);
  const inputClasses = "w-full p-3 bg-gray-50 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition";
  const disabledInputClasses = `${inputClasses} disabled:bg-gray-200 disabled:cursor-not-allowed disabled:text-gray-500`;

  return (
    <form onSubmit={handleSubmit} className="max-w-4xl mx-auto my-10 p-8 bg-[#dbeafe] rounded-xl shadow-lg border border-gray-200">
      <div className="flex items-center justify-between mb-6">
        <h2 className="text-3xl font-bold text-gray-800">
          Add to Incident Report
        </h2>
        <Button size="lg" className="text-l text-white hover:bg-red-700" onClick={goBack} variant="destructive">Go Back</Button>
      </div>

      {/* READ-ONLY Section */}
      <div className="p-4 border border-dashed border-gray-400 rounded-lg mb-6">
        <h3 className="font-bold text-lg text-gray-700 mb-4">Original Submission Details (Read-Only)</h3>
        
        {/* Type Of Request */}
        <div className="mb-4">
          <label className="block mb-2 font-semibold text-gray-700">Type Of Request</label>
          <div className="flex items-center space-x-6">
            <label className="flex items-center space-x-2">
              <input type="radio" value="Incident Notice" checked={absence.type_of_request === 'Incident Notice'} disabled className="disabled:cursor-not-allowed" />
              <span>Incident Notice</span>
            </label>
            <label className="flex items-center space-x-2">
              <input type="radio" value="Time Off Request" checked={absence.type_of_request === 'Time Off Request'} disabled className="disabled:cursor-not-allowed" />
              <span>Time-Off Request</span>
            </label>
          </div>
        </div>
        
        {/* Office, Incident Type */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-4">
          <div>
            <label className="block mb-2 font-semibold text-gray-700">Type Of Incident</label>
            <select value={absence.type_of_incident} className={disabledInputClasses} disabled>
              {incidentTypes.map(type => <option key={type} value={type}>{type}</option>)}
            </select>
          </div>
          <div>
            <label className="block mb-2 font-semibold text-gray-700">Office Location</label>
            <select value={absence.office} className={disabledInputClasses} disabled>
              {officeLocations.map(location => <option key={location} value={location}>{location}</option>)}
            </select>
          </div>
        </div>

        {/* Conditional ETA/ETD Field */}
        {showETA_ETDField && (
          <div className="mb-4">
            <label className="block mb-2 font-semibold text-gray-700">ETA / ETD</label>
            <input type="time" value={absence.eta_etd} className={disabledInputClasses} disabled />
          </div>
        )}

        {/* Incident Start and End Dates */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          <div>
            <label className="block mb-2 font-semibold text-gray-700">Start Date</label>
            <input type="date" value={absence.incident_start} className={disabledInputClasses} disabled />
          </div>
          <div>
            <label className="block mb-2 font-semibold text-gray-700">End Date</label>
            <input type="date" value={absence.incident_end} className={disabledInputClasses} disabled />
          </div>
        </div>
      </div>
      
        {/* EDITABLE Section */}
        <div className="p-4 border border-dashed border-green-500 bg-green-50 rounded-lg">
            <h3 className="font-bold text-lg text-green-800 mb-4">Update Information</h3>
            
            {/* Comments */}
            <div className="mb-6">
                <label htmlFor="employee_comments" className="block mb-2 font-semibold text-gray-700">Update or Add Comments</label>
                <textarea id="employee_comments" value={comments} onChange={(e) => setComments(e.target.value)} className={inputClasses} rows={3} />
            </div>

            {/* Excuse Note Section */}
            <div>
                <label className="block mb-2 font-semibold text-gray-700">Excuse Note Management</label>
                
                {/* Display existing notes with remove option */}
                {absence.excuse_note && absence.excuse_note.length > 0 && (
                    <div className="mb-4">
                        <p className="text-sm font-medium text-gray-600">Already Submitted:</p>
                        <ul className="list-none pl-1 mt-1 text-sm space-y-2">
                            {absence.excuse_note.map((url: string, index: number) => (
                                <li key={index} className={`flex items-center justify-between p-2 rounded-md ${notesToRemove.includes(url) ? 'bg-red-100' : 'bg-gray-100'}`}>
                                    <a href={url} target="_blank" rel="noopener noreferrer" className={`text-blue-600 hover:underline ${notesToRemove.includes(url) ? 'line-through' : ''}`}>
                                        Excuse Note {index + 1}
                                    </a>
                                    <Button type="button" size="sm" variant="destructive" onClick={() => handleToggleRemoveExistingNote(url)}>
                                        {notesToRemove.includes(url) ? 'Undo' : 'Remove'}
                                    </Button>
                                </li>
                            ))}
                        </ul>
                    </div>
                )}

                {canAddNotes ? (
                    <div>
                        <label htmlFor="new_excuse_notes" className="block mb-2 text-sm font-medium text-gray-600">Attach New Note(s)</label>
                        <input type="file" id="new_excuse_notes" multiple onChange={handleFileChange} className={`${inputClasses} file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100`} />
                        
                        {/* Display newly selected files with remove option */}
                        {newExcuseNotes.length > 0 && (
                            <div className="mt-4">
                                <p className="text-sm font-medium text-gray-600">Staged for Upload:</p>
                                <ul className="list-none pl-1 mt-1 text-sm space-y-2">
                                    {newExcuseNotes.map((file, index) => (
                                        <li key={index} className="flex items-center justify-between p-2 rounded-md bg-blue-50">
                                            <span className="text-gray-700">{file.name}</span>
                                            <Button type="button" size="sm" variant="ghost" className="text-red-500 hover:bg-red-100" onClick={() => handleRemoveNewFile(file.name)}>
                                                Remove
                                            </Button>
                                        </li>
                                    ))}
                                </ul>
                            </div>
                        )}
                    </div>
                ) : (
                    <div className="p-3 bg-yellow-100 border border-yellow-300 rounded-lg text-center text-sm text-yellow-800">
                        No note was required for this submission, so new files cannot be added.
                    </div>
                )}
            </div>
        </div>
      
        {/* "WILL NOT SUBMIT" Section */}
        {absence.excuse_note_submitted === 'pending' && (
            <div className="my-6 p-4 border border-dashed border-amber-500 bg-amber-50 rounded-lg text-center">
                <p className="mb-2 text-amber-800">If you do not plan to provide a note for this incident:</p>
                <Button type="button" variant="outline" className="border-amber-600 text-amber-800 hover:bg-amber-100" onClick={handleWillNotSubmit} disabled={isSubmitting}>
                    I will not submit an excuse note
                </Button>
            </div>
        )}
      
        {/* Submission Section */}
        <div className="text-center border-t mt-6 pt-6">
            <Button type="submit" size="lg" disabled={isSubmitting} className="bg-blue-600 hover:bg-blue-700 disabled:bg-gray-400">
              {isSubmitting ? submitStatus : 'Submit Updates'}
            </Button>
            {submitStatus && !isSubmitting && <p className="mt-4 text-center text-red-600">{submitStatus}</p>}
        </div>
    </form>
  );
}