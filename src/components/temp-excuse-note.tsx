"use client";

import { getPSTDate } from '@/lib/date';
import React, { use, useState } from "react";
import Form from 'next/form';
import { Button } from '@/components/ui/button';
import { collection, doc, setDoc, query, where, getDocs, addDoc } from "firebase/firestore";
import { db, storage } from '@/lib/firebase.config';
import { exec } from 'child_process';
import { ref, uploadBytes, getDownloadURL } from "firebase/storage";

export default function IncidentReportForm() {
    // State variables for form fields
    const [name, setName] = useState("");
    const [employee_id, setEmployeeID] = useState("");
    const [office, setOffice] = useState("");
    const [email, setEmail] = useState("");
    const [excuse_note, setExcuseNote] = useState<File | null>(null);
    const [department, setDepartment] = useState("");
    const [start_date, setStartDate] = useState("");
    const [end_date, setEndDate] = useState("");
    const [type_of_incident, setTypeOfIncident] = useState("");
    const [ETA, setETA] = useState("");
    const [ETD, setETD] = useState("");
    const [comment, setComment] = useState("");
    const pstDate = getPSTDate();
    
    // State for submission status feedback
    const [submitStatus, setSubmitStatus] = useState(""); // e.g., "loading", "success", "error"

    const handleSubmit = async (event: React.FormEvent<HTMLFormElement>) => {
        event.preventDefault();
        /*if (!excuse_note) {
            setSubmitStatus("Please upload an excuse note.");
            return
        }*/
        setSubmitStatus("Submitting...");

        const formData = {name, employee_id, office, email, department, start_date, end_date, type_of_incident, ETA, ETD, comment, pstDate};

        try {
            await setDoc(doc(db, "myCollection"), 
              formData
            , { merge: true });
            setSubmitStatus("Data submitted successfully!");
            setName("");
            setEmail("");
        } catch (error) {
            console.error("Error adding document: ", error);
            setSubmitStatus("Error submitting data.");
        }
    };

    // Reusable classes for input fields
    const inputClasses = "w-full p-3 bg-gray-50 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition";

    return (
        <form 
            onSubmit={handleSubmit} 
            className="max-w-4xl mx-auto my-10 p-8 bg-blue-200 rounded-xl shadow-lg"
        >
            <h2 className="text-3xl font-bold text-center text-gray-800 mb-8">
                Employee Incident Report
            </h2>
            
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-6">

                {/* Employee ID */}
                <div>
                    <label htmlFor="employee_id" className="block mb-2 font-semibold text-gray-700">Employee ID</label>
                    <input type="text" id="employee_id" value={employee_id} onChange={(e) => setEmployeeID(e.target.value)} className={inputClasses} required />
                </div>

                {/* Email */}
                <div>
                    <label htmlFor="email" className="block mb-2 font-semibold text-gray-700">Email</label>
                    <input type="email" id="email" value={email} onChange={(e) => setEmail(e.target.value)} className={inputClasses} required />
                </div>
                <br />
                {/* Office Location */}
                <div>
                    <label htmlFor="office" className="block mb-2 font-semibold text-gray-700">Office Location</label>
                    <input type="text" id="office" value={office} onChange={(e) => setOffice(e.target.value)} className={inputClasses} />
                </div>

                {/* Department */}
                <div>
                    <label htmlFor="department" className="block mb-2 font-semibold text-gray-700">Department</label>
                    <select id="department" value={department} onChange={(e) => setDepartment(e.target.value)} className={inputClasses} required>
                        <option value="" disabled>Select Department</option>
                        <option value="Engineering">Engineering</option>
                        <option value="Human Resources">Human Resources</option>
                        <option value="Marketing">Marketing</option>
                        <option value="Sales">Sales</option>
                    </select>
                </div>

                {/* Type of Incident */}
                <div>
                    <label htmlFor="type_of_incident" className="block mb-2 font-semibold text-gray-700">Type of Incident</label>
                    <select id="type_of_incident" value={type_of_incident} onChange={(e) => setTypeOfIncident(e.target.value)} className={inputClasses} required>
                        <option value="" disabled>Select Type</option>
                        <option value="Tardiness">Tardiness</option>
                        <option value="Absence">Absence</option>
                        <option value="Early Departure">Early Departure</option>
                        <option value="Other">Other</option>
                    </select>
                </div>

                {/* Incident Start Date */}
                <div>
                    <label htmlFor="start_date" className="block mb-2 font-semibold text-gray-700">Incident Start Date</label>
                    <input type="date" id="start_date" value={start_date} onChange={(e) => setStartDate(e.target.value)} className={inputClasses} required />
                </div>

                {/* Incident End Date */}
                <div>
                    <label htmlFor="end_date" className="block mb-2 font-semibold text-gray-700">Incident End Date</label>
                    <input type="date" id="end_date" value={end_date} onChange={(e) => setEndDate(e.target.value)} className={inputClasses} />
                </div>

                {/* Estimated Time of Arrival */}
                <div>
                    <label htmlFor="ETA" className="block mb-2 font-semibold text-gray-700">Estimated Time of Arrival</label>
                    <input type="time" id="ETA" value={ETA} onChange={(e) => setETA(e.target.value)} className={inputClasses} />
                </div>
                
                {/* Estimated Time of Departure */}
                <div>
                    <label htmlFor="ETD" className="block mb-2 font-semibold text-gray-700">Estimated Time of Departure</label>
                    <input type="time" id="ETD" value={ETD} onChange={(e) => setETD(e.target.value)} className={inputClasses} />
                </div>
            </div>

            {/* Excuse Note */}
            <div className="mb-6">
                <label htmlFor="excuse_note" className="block mb-2 font-semibold text-gray-700">Excuse Note</label>
                <input 
                  type="file" 
                  id="excuse_note"
                  onChange={(e) => setExcuseNote(e.target.files ? e.target.files[0] : null)}
                  className={inputClasses}
                />
            </div>

            {/* Additional Comments */}
            <div className="mb-6">
                <label htmlFor="comment" className="block mb-2 font-semibold text-gray-700">Additional Comments</label>
                <textarea id="comment" value={comment} onChange={(e) => setComment(e.target.value)} className={inputClasses} />
            </div>

            {/* Submission Section */}
            <div className="text-center">
                <button 
                    type="submit" 
                    className="py-3 px-8 bg-blue-600 text-white font-semibold rounded-lg shadow-md hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-opacity-75 disabled:bg-gray-400 disabled:cursor-not-allowed transition-colors"
                    disabled={submitStatus === 'loading'}
                >
                    {submitStatus === 'loading' ? 'Submitting...' : 'Submit Report'}
                </button>
                
                {submitStatus === 'success' && <p className="mt-4 text-green-600 font-medium">Form submitted successfully!</p>}
                {submitStatus === 'error' && <p className="mt-4 text-red-600 font-medium">Submission failed. Please try again.</p>}
            </div>
        </form>
    );
}