"use client";

import NewAbsenceForm from '@/components/forms/absence-request/new-absence';
import ExistingAbsenceForm from '@/components/forms/absence-request/existing-absence';

import { useState, useMemo } from 'react';

import { collection, query, where, getDocs, doc, getDoc } from "firebase/firestore";
import { db } from '@/lib/firebase.config';
import { Button } from '@/components/ui/button';

import { AbsenceRequest } from '@/lib/types';
import { AbsenceTable } from '@/components/forms/absence-request/employee-absence-table';

export default function Page() {
  const [employeeFirestoreID, setEmployeeFirestoreID] = useState("");
  const [employeeID, setEmployeeID] = useState("");
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
      // 1. Create a query to find the employee document where the 'employeeID' field matches
      const employeesRef = collection(db, "employees");
      console.log("Attempting login for Employee ID:", employeeID);
      const q = query(employeesRef, where("employeeID", "==", employeeID));

      // 2. Execute the query
      const querySnapshot = await getDocs(q);

      // 3. Check if we found a matching employee
      if (!querySnapshot.empty) {
        // Get the first matching document
        const employeeDoc = querySnapshot.docs[0];
        const data = employeeDoc.data();
      
        // 4. Extract and validate Birth Year
        const dob = data.dateOfBirth;
        const employee_birthYear = dob.split('-')[0];
        console.log("Employee Name from DB:", data.firstName, data.lastName);
        console.log("Employee Birth Year from DB:", employee_birthYear);
        if (employee_birthYear === year) {
          setIsAuthenticated(true);
        
          // 5. IMPORTANT: Set the Firestore auto-generated ID (e.g., "7zK9...") 
          // as the employeeFirestoreID so the NewAbsenceForm saves correctly.
          setEmployeeFirestoreID(employeeDoc.id); 
        
          // 6. Set other local states
          setEmployeeInfo(data);
          setEmployeeTitle(data.jobTitle);
          setEmployeeName(`${data.firstName} ${data.lastName}`);
        
          // 7. Fetch history using the employee's ID string
          fetchAbsences(employeeID);
        } 
        else {
          setError("Invalid credentials. Please enter correct Employee ID and Birth Year.");
        }
      } 
      else {
        setError("Invalid credentials. Please enter correct Employee ID and Birth Year.");
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
      const q = query(collection(db, "absences"), where("employee_id", "==", employeeID));
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
    fetchAbsences(employeeID); // Refresh the list of absences
  };

  const handleViewDetails = (absence: AbsenceRequest) => {
    setSelectedAbsence(absence);
  };

  const noteSubmisisonAbsences = useMemo(() => {
    return allAbsences.filter(absence => absence.excuse_note_submitted === 'pending');
  }, [allAbsences]);

  const pendingAbsences = useMemo(() => {
    return allAbsences.filter(absence => absence.manager_approval === 'pending' || absence.final_approval === 'pending');
  }, [allAbsences]);

  const approvedAbsences = useMemo(() => {
    return allAbsences.filter(absence => (absence.manager_approval === 'approved' && absence.final_approval === 'approved') || (absence.manager_approval === 'not_required' && absence.final_approval === 'approved'));
  }, [allAbsences]);

  const deniedAbsences = useMemo(() => {
    return allAbsences.filter(absence => absence.manager_approval === 'denied' || absence.final_approval === 'denied');
  }, [allAbsences]);

  if (loading) {
    return <div>Loading...</div>;
  }

  /*
  if (selectedAbsence) {
    return <ExistingAbsenceForm absence={selectedAbsence} onFormSubmit={handleFormSubmit} />;
  };

  if (newAbsence) {
    return <NewAbsenceForm employeeID={employeeID} employeeTitle={employeeTitle} employeeName={employeeName} onFormSubmit={handleFormSubmit} />;
  };
  */

  return (
    <main className="white min-h-screen p-4 sm:p-6 md:p-8">
      {!isAuthenticated ? (
        <div className="max-w-md mx-auto mt-20">
          <form onSubmit={handleLogin} className="bg-white p-8 rounded-lg shadow-md">
            <h2 className="text-2xl font-bold text-center mb-6">Employee Login</h2>
            <div className="mb-4">
              <label htmlFor="employeeId" className="block text-gray-700">Employee ID</label>
              <input type="text" id="employeeId" value={employeeID} onChange={(e) => setEmployeeID(e.target.value)} className="w-full p-2 border rounded" required autoComplete="off"/>
            </div>
            <div className="mb-6">
              <label htmlFor="birthYear" className="block text-gray-700">Birth Year</label>
              <input type="text" id="birthYear" value={year} onChange={(e) => setYear(e.target.value)} className="w-full p-2 border rounded" required autoComplete="off"/>
            </div>
            <button type="submit" className="w-full bg-blue-500 text-white p-2 rounded cursor-pointer">Login</button>
            {error && <p className="text-red-500 text-center mt-4">{error}</p>}
          </form>
        </div>
      ) : (
        <div>
          <div className="flex flex-row gap-5">
            <h2 className="text-2xl font-bold mb-4">{employeeInfo.firstName} {employeeInfo.lastName} Absences</h2> 
            <Button onClick={() => setNewAbsence(true)}  className="bg-green-500 hover:bg-green-700 text-white p-2 rounded cursor-pointer shadow-sm">New Incident Report</Button>
            <Button onClick={handleLogout} className="hover:bg-red-700 rounded cursor-pointer shadow-sm" variant="destructive">Sign Out</Button>
          </div>

          <div className="border-b border-gray-200 mb-2">
            <nav className="-mb-px flex space-x-8" aria-label="Tabs">
              <button
                onClick={() => setActiveTab('needed')}
                className={`${activeTab === 'needed' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm`}
              >
                Notes Required ({noteSubmisisonAbsences.length})
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

          {newAbsence && (
          <NewAbsenceForm 
            employeeFirestore={employeeFirestoreID}
            employeeID={employeeID} 
            employeeTitle={employeeTitle} 
            employeeName={employeeName} 
            employeeSkipManagerApproval={employeeInfo.skipManagerApproval}
            onFormSubmit={handleFormSubmit} 
            onClose={() => setNewAbsence(false)}
            />
          )}

          {selectedAbsence && (
            <ExistingAbsenceForm 
              absence={selectedAbsence} 
              onFormSubmit={handleFormSubmit} 
              onClose={() => setSelectedAbsence(null)}
            />
          )}
        </div>
      )}
    </main>
  );
}