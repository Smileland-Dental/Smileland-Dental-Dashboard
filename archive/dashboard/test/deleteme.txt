"use client";
import IncidentReportForm from '@/components/forms/excuse-note';
import { useState } from 'react';
// Import 'doc' and 'getDoc' for fetching a single document by ID
import { collection, query, where, getDocs, doc, getDoc } from "firebase/firestore";
import { db } from '@/lib/firebase.config';

export default function Home() {
  const [employeeId, setEmployeeId] = useState("");
  const [birthYear, setBirthYear] = useState("");
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [absences, setAbsences] = useState<any[]>([]);
  const [selectedAbsence, setSelectedAbsence] = useState<any | null>(null);
  const [error, setError] = useState("");

  const handleLogin = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setError("");

    try {
      // Create a direct reference to the employee document using the employeeId
      const employeeDocRef = doc(db, "employees", employeeId);
      
      // Fetch the document snapshot
      const employeeDocSnap = await getDoc(employeeDocRef);

      // Check if the document exists and if the birth_year field matches
      if (employeeDocSnap.exists() && employeeDocSnap.data().year === birthYear) {
        setIsAuthenticated(true);
        fetchAbsences(employeeId);
      } else {
        setError("Invalid credentials. Please try again.");
      }
    } catch (err) {
      console.error("Login Error: ", err);
      setError("An error occurred during login. Please try again.");
    }
  };

  const fetchAbsences = async (id: string) => {
    try {
      // This query remains the same as 'employee_id' is a field in the 'absences' collection
      const q = query(collection(db, "absences"), where("employee_id", "==", id), where("incident_submitted", "==", false));
      const querySnapshot = await getDocs(q);
      const absenceData = querySnapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
      setAbsences(absenceData);
    } catch (err) {
      setError("Failed to fetch absences.");
    }
  };

  const handleFormSubmit = () => {
    setSelectedAbsence(null);
    fetchAbsences(employeeId); // Refresh the list of absences
  };

  if (selectedAbsence) {
    return <IncidentReportForm absence={selectedAbsence} onFormSubmit={handleFormSubmit} />;
  }

  return (
    <main className="white min-h-screen p-4 sm:p-6 md:p-8">
      {!isAuthenticated ? (
        <div className="max-w-md mx-auto mt-20">
          <form onSubmit={handleLogin} className="bg-white p-8 rounded-lg shadow-md">
            <h2 className="text-2xl font-bold text-center mb-6">Employee Login</h2>
            <div className="mb-4">
              <label htmlFor="employeeId" className="block text-gray-700">Employee ID</label>
              <input type="text" id="employeeId" value={employeeId} onChange={(e) => setEmployeeId(e.target.value)} className="w-full p-2 border rounded" required />
            </div>
            <div className="mb-6">
              <label htmlFor="birthYear" className="block text-gray-700">Birth Year</label>
              <input type="text" id="birthYear" value={birthYear} onChange={(e) => setBirthYear(e.target.value)} className="w-full p-2 border rounded" required />
            </div>
            <button type="submit" className="w-full bg-blue-500 text-white p-2 rounded">Login</button>
            {error && <p className="text-red-500 text-center mt-4">{error}</p>}
          </form>
        </div>
      ) : (
        <div>
          <h2 className="text-2xl font-bold mb-4">Your Absences</h2>
          {absences.length > 0 ? (
            <ul className="space-y-4">
              {absences.map(absence => (
                <li key={absence.id} onClick={() => setSelectedAbsence(absence)} className="p-4 bg-gray-100 rounded-lg cursor-pointer hover:bg-gray-200">
                  <p><strong>Date:</strong> {absence.date}</p>
                  <p><strong>Reason:</strong> {absence.reason}</p>
                </li>
              ))}
            </ul>
          ) : (
            <p>No pending absences to report.</p>
          )}
        </div>
      )}
    </main>
  );
}

