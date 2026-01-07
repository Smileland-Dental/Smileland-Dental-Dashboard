"use client";

import NewAbsenceForm from '@/components/forms/new-absence';
import ExistingAbsenceForm from '@/components/forms/existing-absence';

import { useState, useMemo } from 'react';

import { collection, query, where, getDocs, doc, getDoc } from "firebase/firestore";
import { db } from '@/lib/firebase.config';
import { Button } from '@/components/ui/button';

import { AbsenceRequest } from '@/lib/types';
import { AbsenceTable } from '@/components/employee-absence-table';

export default function Page() {
  const [employee_id, setEmployeeID] = useState("");
  const [employeeTitle, setEmployeeTitle] = useState("");
  const [employeeName, setEmployeeName] = useState("");
  const [year, setYear] = useState("");
  const [isAuthenticated, setIsAuthenticated] = useState(false);

  const [allAbsences, setAbsences] = useState<AbsenceRequest[]>([]);
  const [selectedAbsence, setSelectedAbsence] = useState<AbsenceRequest | null>(null);
  const [newAbsence, setNewAbsence] = useState(false);

  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [employeeInfo, setEmployeeInfo] = useState<any>(null);
  const [activeTab, setActiveTab] = useState('needed');

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
        setEmployeeTitle(employeeDocSnap.data().title);
        setEmployeeName(employeeDocSnap.data().name);
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
      setLoading(true);
      // This query remains the same as 'employee_id' is a field in the 'absences' collection
      const q = query(collection(db, "absences"), where("employee_id", "==", employee_id));
      const querySnapshot = await getDocs(q);
      const absenceData = querySnapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
      setAbsences(absenceData as AbsenceRequest[]);
    } 
    catch (err) {
      setAbsences([]);
      setError("Failed to fetch absences.");
    }
    setLoading(false);
  };

  const handleFormSubmit = () => {
    setSelectedAbsence(null);
    setNewAbsence(false);
    fetchAbsences(employee_id); // Refresh the list of absences
  };

  const handleViewDetails = (absence: AbsenceRequest) => {
    setSelectedAbsence(absence);
  };

  const noteSubmisisonAbsences = useMemo(() => {
    return allAbsences.filter(absence => absence.excuse_note_submitted === 'pending');
  }, [allAbsences]);

  const pendingAbsences = useMemo(() => {
    if (['Dentist', 'Supervisor', 'Exempt', 'EXEC'].includes(employeeTitle)) {
      return allAbsences.filter(absence => absence.final_approval === 'pending');
    }
    return allAbsences.filter(absence => absence.manager_approval === 'pending' && absence.final_approval === 'pending');
  }, [allAbsences]);

  const approvedAbsences = useMemo(() => {
    if (['Dentist', 'Supervisor', 'Exempt', 'EXEC'].includes(employeeTitle)) {
      return allAbsences.filter(absence => absence.final_approval === 'approved');
    }
    return allAbsences.filter(absence => absence.manager_approval === 'approved' && absence.final_approval === 'approved');
  }, [allAbsences]);

  const deniedAbsences = useMemo(() => {
    return allAbsences.filter(absence => absence.manager_approval === 'denied' || absence.final_approval === 'denied');
  }, [allAbsences]);

  if (loading) {
    return <div>Loading...</div>;
  }

  if (selectedAbsence) {
    return <ExistingAbsenceForm absence={selectedAbsence} onFormSubmit={handleFormSubmit} />;
  };

  if (newAbsence) {
    return <NewAbsenceForm employeeID={employee_id} employeeTitle={employeeTitle} employeeName={employeeName} onFormSubmit={handleFormSubmit} />;
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

          <div className="border-b border-gray-200 mb-2">
            <nav className="-mb-px flex space-x-8" aria-label="Tabs">
              <button
                onClick={() => setActiveTab('needed')}
                className={`${activeTab === 'needed' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm`}
              >
                Note Submission ({noteSubmisisonAbsences.length})
              </button>
              <button
                onClick={() => setActiveTab('pending')}
                className={`${activeTab === 'pending' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm`}
              >
                Pending Approval ({pendingAbsences.length})
              </button>
              <button
                onClick={() => setActiveTab('approved')}
                className={`${activeTab === 'approved' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm`}
              >
                Approved ({approvedAbsences.length})
              </button>
              <button
                onClick={() => setActiveTab('denied')}
                className={`${activeTab === 'denied' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm`}
              >
                Denied ({deniedAbsences.length})
              </button>
              <button onClick={() => setActiveTab('all')}
                className={`${activeTab === 'all' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm`}
              >
                All Requests ({allAbsences.length})
              </button>
            </nav>
          </div>

          {/* Absence Table */}
          
          <div className="mt-4">
            {activeTab === 'needed' && (
              <AbsenceTable requests={noteSubmisisonAbsences} status={'needed'} onViewDetails={handleViewDetails} />
            )}
            {activeTab === 'pending' && (
              <AbsenceTable requests={pendingAbsences} status={'pending'}  onViewDetails={handleViewDetails} />
            )}
            {activeTab === 'approved' && (
                <AbsenceTable requests={approvedAbsences} status={'approved'}  onViewDetails={handleViewDetails} />
            )}
              {activeTab === 'denied' && (
                <AbsenceTable requests={deniedAbsences} status={'denied'}  onViewDetails={handleViewDetails} />
            )}
            {activeTab === 'all' && (
                <AbsenceTable requests={allAbsences} status={'all'}  onViewDetails={handleViewDetails} />
            )}
          </div>

          {/*}
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
                        <p><strong>Note Provided:</strong> {absence.excuse_note_submitted}</p>
                      </div>
                      <p className='col-span-2'><strong>Comments:</strong> {absence.comments}</p>
                    </li>
                  );
                })
              }
            </ul>
          ) : (
            <p>No pending absences to report.</p>
          )} 8*/}


          {/* Create a sorted copy before mapping */}
        </div>
      )}
    </main>
  );
}