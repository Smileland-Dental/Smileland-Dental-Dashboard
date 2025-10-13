"use client";

import { getPSTDate } from '@/lib/date';
import React, { useState, useEffect } from "react"; // Import useEffect
import { doc, updateDoc, addDoc, collection } from "firebase/firestore";
import { db, storage } from '@/lib/firebase.config';
import { ref, uploadBytes, getDownloadURL } from "firebase/storage";
import { Button } from '@/components/ui/button';

const incidentTypes = [
  "Late In",
  "Early Out",
  "Absent",
  "Leave and Come Back",
  "Long Lunch",
  "Switch Shift"
];

const officeLocations = [
  "Corporate",
  "Ming",
  "Bernard",
  "California",
  "Ortho",
  "Delano",
  "Tulare",
  "Visalia",
  "Fresno"
]

export default function NewAbsenceForm({ employeeID, employeeTitle, employeeName, onFormSubmit } : { employeeID: String, employeeTitle: String, employeeName: String, onFormSubmit: () => void }) {
  // State variables for form fields

  const [formData, setFormData] = useState({
    employee_id: employeeID, // get from login
    employee_title: employeeTitle, // get from login
    employee_name: employeeName,
    date_submitted: getPSTDate(), // today's date
    type_of_request: '', // Incident Notice, Time-Off Request
    type_of_incident: '', // Late In, Early Out, Absent, Leave and Come Back, Long Lunch, Switch Shift
    office: '', // Corporate, Ming, Bernard, California, Ortho, Delano, Tulare, Visalia, Fresno 
    incident_start: '', // Start Date of Absence
    incident_end: '', // End Date of Absence
    employee_comments: '', // Additional comments from employee
    eta_etd: "", // Hour:Minutes | For Late In/Early Out, when they expect to arrive/leave
    excuse_note_submitted: '', // pending, submitted, not_provided
    excuse_note: [], // File URL(s) for uploaded excuse notes
  });

  const [excuse_notes, setExcuseNotes] = useState<File[]>([]);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [submitStatus, setSubmitStatus] = useState("");
  const [isChecked, setIsChecked] = useState(false);

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>) => {
    const { name, value } = e.target;
    setFormData(field => ({ ...field, [name]: value }));
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setExcuseNotes(prev => [...prev, ...Array.from(e.target.files ?? [])]);
    }
  };

  const handleRemoveNewFile = (fileName: string) => {
    setExcuseNotes(prev => prev.filter(file => file.name !== fileName));
  };

  const handleSubmit = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setIsSubmitting(true);
    setSubmitStatus("Submitting...");

    try {
      

      // 1. Create the document in Firestore first to get an ID
      const absenceRef = await addDoc(collection(db, "absences"), {
          ...formData,
          manager_approval: 'pending',
          manager_name: '',
          final_approval: 'pending',
          final_name: '',
          manager_notes: '',
          excuse_note: [], // Start with an empty URL
      });

      const uploadedFileUrls: string[] = [];
      // 2. If there's a file, upload it using the new document's ID
      if (excuse_notes.length > 0 && formData.excuse_note_submitted === 'submitted') {
        setSubmitStatus("Uploading file(s)...");
        for (const file of excuse_notes) {
          const storageRef = ref(storage, `excuse-notes/${absenceRef.id}/${file.name}`);
          await uploadBytes(storageRef, file);
          const downloadUrl = await getDownloadURL(storageRef);
          uploadedFileUrls.push(downloadUrl);
        }

        // 3. Update the document with the array of file URLs
        await updateDoc(doc(db, "absences", absenceRef.id), {
          excuse_note: uploadedFileUrls,
        });
      }

      setSubmitStatus("Request submitted successfully!");
      setTimeout(() => onFormSubmit(), 1500); // Give user time to see success message
    } catch (error) {
      console.error("Error submitting request: ", error);
      setSubmitStatus("Error submitting data. Please try again.");
      setIsSubmitting(false);
    }
  };

  const goBack = () => {
    setSubmitStatus("Discarding changes...");
    onFormSubmit(); 
  };

  const showETA_ETDField = ["Late In", "Early Out", "Leave and Come Back"].includes(formData.type_of_incident);
  const inputClasses = "w-full p-3 bg-gray-50 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition";

  return (
    <form onSubmit={handleSubmit} className="max-w-4xl mx-auto my-10 p-8 bg-[#c1e8db] rounded-xl shadow-lg border border-gray-200">
      <div className="flex items-center justify-between mb-6">
        <h2 className="text-3xl font-bold text-center text-gray-800">
          Absence Request / Report Form
        </h2>
        <Button size="lg" className="text-l text-white hover:bg-red-700" onClick={goBack} variant="destructive">Go Back</Button>
      </div>

      {/* Type Of Request */}
      <div className="mb-6 bg-amber-100 p-2 rounded-lg border border-amber-200">
        <legend className="block mb-2 font-semibold text-gray-700">Type Of Request <span className="text-red-500">*</span></legend>
        <div className="flex items-center space-x-6">
          <label className="flex items-center space-x-2 cursor-pointer">
            <input type="radio" name="type_of_request" value="Incident Notice" onChange={handleChange} required />
            <span>Incident Notice (Less than 30 days)</span>
          </label>
          <label className="flex items-center space-x-2 cursor-pointer">
            <input type="radio" name="type_of_request" value="Time Off Request" onChange={handleChange} />
            <span>Time-Off Request (More than 30 days)</span>
          </label>
        </div>
      </div>

      {/* Office, Incident Type*/}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-6">
        <div>
          <label htmlFor="type_of_incident" className="block mb-2 font-semibold text-gray-700">Type Of Incident <span className="text-red-500">*</span></label>
          <select id="type_of_incident" name="type_of_incident" value={formData.type_of_incident} onChange={handleChange} className={inputClasses} required>
            <option value="" disabled>--Select Incident Type--</option>
            {incidentTypes.map(type => <option key={type} value={type}>{type}</option>)}
          </select>
        </div>
        <div>
          <label htmlFor="office" className="block mb-2 font-semibold text-gray-700">Office Location <span className="text-red-500">*</span></label>
          <select id="office" name="office" value={formData.office} onChange={handleChange} className={inputClasses} required>
            <option value="" disabled>--Select Office--</option>
            {officeLocations.map(location => <option key={location} value={location}>{location}</option>)}
          </select>
        </div>
      </div>

      {/* Conditionally Rendered ETA/ETD Field based on Incident Type */}
      {showETA_ETDField && (
        <div className="mb-6 p-4 bg-blue-50 border border-blue-200 rounded-lg">
          <label htmlFor="eta_etd" className="block mb-2 font-semibold text-gray-700">Estimated Time of Arrival / Departure (ETA/ETD) <span className="text-red-500">*</span></label>
          <input type="time" id="eta_etd" name="eta_etd" value={formData.eta_etd} onChange={handleChange} className={inputClasses} required/>
        </div>
      )}

      {/* Incident Start and End */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-6">
        <div>
          <label htmlFor="incident_start" className="block mb-2 font-semibold text-gray-700">Start Date & Time <span className="text-red-500">*</span></label>
          <input type="date" id="incident_start" name="incident_start" value={formData.incident_start} onChange={handleChange} className={inputClasses} required/>
        </div>
        <div>
          <label htmlFor="incident_end" className="block mb-2 font-semibold text-gray-700">End Date & Time <span className="text-red-500">*</span></label>
          <input type="date" id="incident_end" name="incident_end" value={formData.incident_end} onChange={handleChange} className={inputClasses} required/>
        </div>
      </div>
      
      {/* Additional Comments */}
      <div className="mb-6">
        <label htmlFor="employee_comments" className="block mb-2 font-semibold text-gray-700">Additional Comments <span className="text-red-500">*</span></label>
        <textarea id="employee_comments" name="employee_comments" value={formData.employee_comments} onChange={handleChange} className={inputClasses} rows={2} required/>
      </div>

      {/* Excuse Note File Upload Section */}
      <div className="mb-8 p-4 bg-gray-50 border border-gray-200 rounded-lg">
        <legend className="block mb-3 font-semibold text-gray-700">Excuse Note Submission <span className="text-red-500">*</span></legend>
        <div className="flex flex-col sm:flex-row sm:items-center sm:space-x-6 space-y-2 sm:space-y-0">
          <label className="flex items-center space-x-2 cursor-pointer">
            <input type="radio" name="excuse_note_submitted" value="submitted" onChange={handleChange} checked={formData.excuse_note_submitted === 'submitted'} required />
            <span>Submit Note Now</span>
          </label>
          <label className="flex items-center space-x-2 cursor-pointer">
            <input type="radio" name="excuse_note_submitted" value="pending" onChange={handleChange} checked={formData.excuse_note_submitted === 'pending'} />
            <span>Will Provide Note Later</span>
          </label>
          <label className="flex items-center space-x-2 cursor-pointer">
            <input type="radio" name="excuse_note_submitted" value="not_provided" onChange={handleChange} checked={formData.excuse_note_submitted === 'not_provided'} />
            <span>No Note Will Be Provided</span>
          </label>
        </div>

        {/* Conditionally render file input when 'Submit Note' is selected */}
        {formData.excuse_note_submitted === 'submitted' && (
          <div className="mt-4">
            <label htmlFor="excuse_note_files" className="block mb-2 text-sm font-medium text-gray-600">Attach Excuse Note(s)</label>
            <input
              type="file"
              id="excuse_note_files"
              multiple
              required
              onChange={handleFileChange}
              className={`${inputClasses} file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 disabled:bg-gray-200 disabled:cursor-not-allowed`}
              disabled={formData.excuse_note_submitted !== 'submitted'}
            />
            {excuse_notes.length > 0 && (
              <div className="mt-2 text-sm text-gray-600">
                <p className="text-sm font-medium text-gray-600">Staged for Upload:</p>
                <ul className="list-none pl-1 mt-1 text-sm space-y-2">
                    {excuse_notes.map((file, index) => (
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
        )}
      </div>

      {/* Submission Section */}
      <div className="text-center border-t pt-6">
         <div className="flex items-center justify-center mb-4">
            <input type="checkbox" id="submitCheck" checked={isChecked} onChange={(e) => setIsChecked(e.target.checked)} className="h-4 w-4 text-blue-600 border-gray-300 rounded focus:ring-blue-500" required/>
            <label htmlFor="submitCheck" className="ml-2 block text-sm text-gray-900">I confirm the details are correct and wish to submit.</label>
         </div>
        <button 
          disabled={!isChecked || isSubmitting}
          type="submit"
          className="py-3 px-8 bg-blue-600 text-white font-semibold rounded-lg shadow-md hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-opacity-75 disabled:bg-gray-400 disabled:cursor-not-allowed transition-colors"
        >
          {isSubmitting ? submitStatus : 'Submit Report'}
        </button>
        {submitStatus && <p className="mt-4 text-center text-gray-600">{submitStatus}</p>}
      </div>
    </form>
  );
}