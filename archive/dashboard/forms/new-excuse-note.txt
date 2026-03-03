"use client";

import { getPSTDate } from '@/lib/date';
import React, { useState, useEffect } from "react"; // Import useEffect
import { doc, updateDoc, addDoc, collection } from "firebase/firestore";
import { db, storage } from '@/lib/firebase.config';
import { ref, uploadBytes, getDownloadURL } from "firebase/storage";
import { Button } from '@/components/ui/button';

export default function NewIncidentReportForm({ employeeID, onFormSubmit }: { employeeID: String, onFormSubmit: () => void }) {
  // State variables for form fields
  const [excuse_note, setExcuseNote] = useState<File | null>(null);
  const [comments, setComments] = useState("");

  const [submitStatus, setSubmitStatus] = useState("");

  const handleSubmit = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setSubmitStatus("Submitting...");

    /*
    let excuseNoteUrl = "";
    if (excuse_note) {
        const storageRef = ref(storage, `excuse-notes/${absence.id}/${excuse_note.name}`);
        await uploadBytes(storageRef, excuse_note);
        excuseNoteUrl = await getDownloadURL(storageRef);
    }
    */

    const formData = {
        employee_id: employeeID,
        office: '',
        type_of_incident: '',
        incident_start: '',
        incident_end: '',
        notes: '',
        comments,
        date_submitted: getPSTDate(),
        incident_submitted: true,
        excuse_note: "Test URL",
    };

    try {
        await addDoc(collection(db, "absences"), formData);
        setSubmitStatus("Data submitted successfully!");
        onFormSubmit(); // Notify parent component
    } 
    catch (error) {
        console.error("Error updating document: ", error);
        setSubmitStatus("Error submitting data.");
    }
  };

  const goBack = () => {
    setSubmitStatus("Discarding changes...");
    onFormSubmit(); 
  }

  const inputClasses = "w-full p-3 bg-gray-50 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition";

   return (
        <form 
            onSubmit={handleSubmit} 
            className="max-w-4xl mx-auto my-10 p-8 bg-blue-200 rounded-xl shadow-lg"
        >
            <div className="flex items-center justify-between mb-6">
              <h2 className="text-3xl font-bold text-center text-gray-800">
                  Employee Incident Report
              </h2>
              <Button size="lg" className="text-l text-white hover:bg-red-700" onClick={goBack} variant="destructive">Go Back</Button>
            </div>

            {/* Additional Comments */}
            <div className="mb-6">
                <label htmlFor="comment" className="block mb-2 font-semibold text-gray-700">Additional Comments</label>
                <textarea id="comment" value={comments} onChange={(e) => setComments(e.target.value)} className={inputClasses} />
            </div>

            {/* Submission Section */}
            <div className="text-center">
                <button 
                    type="submit" 
                    className="py-3 px-8 bg-blue-600 text-white font-semibold rounded-lg shadow-md hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-opacity-75 disabled:bg-gray-400 disabled:cursor-not-allowed transition-colors"
                >
                    {submitStatus === 'Submitting...' ? 'Submitting...' : 'Submit Report'}
                </button>
            </div>
        </form>
    );
}