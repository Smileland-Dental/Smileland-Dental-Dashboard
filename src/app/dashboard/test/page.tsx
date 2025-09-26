/* "use client";
import { getPSTDate } from '@/lib/date';
import React, { use, useState } from "react";
import Form from 'next/form';
import { Button } from '@/components/ui/button';
import { collection, doc, setDoc, query, where, getDocs, addDoc } from "firebase/firestore";
import { db } from '@/lib/firebase.config';

export default function testPage() {
    const [name, setName] = useState("");
    const [employee_id, setEmployeeID] = useState("");
    const [office, setOffice] = useState("");
    const [email, setEmail] = useState("");
    const [execuse_note, setExcuseNote] = useState("");
    const [department, setDepartment] = useState("");
    const [start_date, setStartDate] = useState("");
    const [end_date, setEndDate] = useState("");
    const [type_of_incident, setTypeOfIncident] = useState("");
    const [ETA, setETA] = useState("");
    const [ETD, setETD] = useState("");
    const [comment, setComment] = useState("");
    const [submitStatus, setSubmitStatus] = useState("");
    const pstDate = getPSTDate();

    const handleSubmit = async (event: React.FormEvent<HTMLFormElement>) => {
        event.preventDefault();
        setSubmitStatus("Submitting...");

        try {
            await setDoc(doc(db, "myCollection", name), {
                name,
                today_date: pstDate,
                email,
                office: "Corporate",
            }, { merge: true });
            setSubmitStatus("Data submitted successfully!");
            setName("");
            setEmail("");
        } catch (error) {
            console.error("Error adding document: ", error);
            setSubmitStatus("Error submitting data.");
        }
    };

  return (
    <div>
      <h1>This is a test page. {pstDate}</h1>
          <form className="bg-red-400" onSubmit={handleSubmit}>
            <div>
                <label htmlFor="name">Name:</label>
                <input
                  className="bg-amber-200"
                    type="text"
                    id="name"
                    value={name}
                    onChange={(e) => setName(e.target.value)}
                    required
                />
            </div>
            <div>
                <label htmlFor="email">Email:</label>
                <input
                  className="bg-amber-200"
                    type="email"
                    id="email"
                    value={email}
                    onChange={(e) => setEmail(e.target.value)}
            
                />
            </div>
            <button className="bg-blue-200" type="submit">Submit</button>
            {submitStatus && <p>{submitStatus}</p>}
        </form>
    </div>
  )
} */

  /*
import IncidentReportForm from '@/components/excuse-note';

export default function Home() {
  return (
    // The bg-gray-100 provides a nice contrast for the white form
    <main className="white min-h-screen p-4 sm:p-6 md:p-8">
      <IncidentReportForm />
    </main>
  );
}
*/

"use client";

import IncidentReportForm from '@/components/forms/excuse-note';
import NewIncidentReportForm from '@/components/forms/new-excuse-note';
import AbsenceForm from '@/components/forms/absence';

import { useState } from 'react';

import { collection, query, where, getDocs, doc, getDoc } from "firebase/firestore";
import { db } from '@/lib/firebase.config';
import { Button } from '@/components/ui/button';

export default function Page() {
  const [employee_id, setEmployeeID] = useState("");
  const [year, setYear] = useState("");
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [absences, setAbsences] = useState<any[]>([]);
  const [selectedAbsence, setSelectedAbsence] = useState<any>(null);
  const [newAbsence, setNewAbsence] = useState(false);
  const [error, setError] = useState("");
  const [employeeInfo, setEmployeeInfo] = useState<any>(null);

  const handleLogin = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setError("");

    try {
      // Create a direct reference to the employee document using the employee_id
      const employeeDocRef = doc(db, "employees", employee_id);

      // Fetch the document snapshot
      const employeeDocSnap = await getDoc(employeeDocRef);

      // Check if the document exists and if the birth_year field matches
      if (employeeDocSnap.exists() && employeeDocSnap.data().year === year) {
        setIsAuthenticated(true);
        fetchAbsences(employee_id);
        setEmployeeInfo(employeeDocSnap.data());
      }
      else {
        setError("Invalid credentials. Please try again.");
      }
    }
    catch (err) {
      console.error("Login Error: ", err);
      setError("An error occurred during login. Please try again.");
    }
  };

  const handleLogout = () => {
    setIsAuthenticated(false);
    setSelectedAbsence(null);
    setEmployeeID('') 
    setYear('')
  };

  const fetchAbsences = async (id: string) => {
    try {
      // This query remains the same as 'employee_id' is a field in the 'absences' collection
      const q = query(collection(db, "absences"), where("employee_id", "==", employee_id));
      const querySnapshot = await getDocs(q);
      const absenceData = querySnapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
      setAbsences(absenceData);
    } 
    catch (err) {
      setError("Failed to fetch absences.");
    }
  };

  const handleFormSubmit = () => {
    setSelectedAbsence(null);
    setNewAbsence(false);
    fetchAbsences(employee_id); // Refresh the list of absences
  };

  if (selectedAbsence) {
    return <IncidentReportForm absence={selectedAbsence} onFormSubmit={handleFormSubmit} />;
  };

  if (newAbsence) {
    return <AbsenceForm employeeID={employee_id} onFormSubmit={handleFormSubmit} />;
  };

  return (
    <main className="white min-h-screen p-4 sm:p-6 md:p-8">
      {!isAuthenticated ? (
        <div className="max-w-md mx-auto mt-20">
          <form onSubmit={handleLogin} className="bg-white p-8 rounded-lg shadow-md">
            <h2 className="text-2xl font-bold text-center mb-6">Employee Login</h2>
            <div className="mb-4">
              <label htmlFor="employeeId" className="block text-gray-700">Employee ID</label>
              <input type="text" id="employeeId" value={employee_id} onChange={(e) => setEmployeeID(e.target.value)} className="w-full p-2 border rounded" required />
            </div>
            <div className="mb-6">
              <label htmlFor="birthYear" className="block text-gray-700">Birth Year</label>
              <input type="text" id="birthYear" value={year} onChange={(e) => setYear(e.target.value)} className="w-full p-2 border rounded" required />
            </div>
            <button type="submit" className="w-full bg-blue-500 text-white p-2 rounded cursor-pointer">Login</button>
            {error && <p className="text-red-500 text-center mt-4">{error}</p>}
          </form>
        </div>
      ) : (
        <div>
          <div className="flex flex-row gap-5">
            <h2 className="text-2xl font-bold mb-4">{employeeInfo.name} Absences</h2> 
            <Button onClick={() => setNewAbsence(true)}  className="bg-green-500 hover:bg-green-700 text-white p-2 rounded cursor-pointer">New Incident Report</Button>
            <Button onClick={handleLogout} className="hover:bg-red-700 rounded cursor-pointer" variant="destructive">Sign Out</Button>
          </div>
          {absences.length > 0 ? (
            <ul className="space-y-4">
               {[...absences]
                .sort((a, b) => new Date(a.incident_start).getTime() - new Date(b.incident_start).getTime())
                .map(absence => {
                  const submittedColor = absence.incident_submitted ? 'bg-green-100' : 'bg-red-100';
                  return (
                    <li 
                      key={absence.id} 
                      onClick={() => setSelectedAbsence(absence)} 
                      className={`p-4 ${submittedColor} sm:w-2/3 rounded-lg cursor-pointer hover:bg-gray-200`}
                    >
                      <div className="columns-2">
                        <p><strong>Incident Start:</strong> {absence.incident_start}</p>
                        <p><strong>Type Of Incident:</strong> {absence.type_of_incident}</p>
                        <p><strong>Office:</strong> {absence.office}</p>
                        <p><strong>Note Provided:</strong> {absence.incident_submitted ? 'Yes' : 'No'}</p>
                      </div>
                      <p className='col-span-2'><strong>Comments:</strong> {absence.comments}</p>
                    </li>
                  );
                })
              }
            </ul>
          ) : (
            <p>No pending absences to report.</p>
          )}
          {/* Create a sorted copy before mapping */}
        </div>
      )}
    </main>
  );
}