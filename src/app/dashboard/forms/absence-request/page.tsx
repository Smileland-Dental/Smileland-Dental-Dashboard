"use client";

import NewAbsenceForm from '@/components/forms/absence-request/new-absence';
import ExistingAbsenceForm from '@/components/forms/absence-request/existing-absence';

import { useState, useMemo, useEffect } from 'react';

import { collection, query, where, getDocs, doc, getDoc } from "firebase/firestore";
import { db } from '@/lib/firebase.config';
import { Button } from '@/components/ui/button';

import { AbsenceRequest } from '@/lib/types';
import { AbsenceTable } from '@/components/forms/absence-request/employee-absence-table';

import { EarlyOutForm } from '@/components/forms/absence-request/volunteer-early-out';

export default function Page() {
  const [employeeFirestoreID, setEmployeeFirestoreID] = useState("");
  const [employeeID, setEmployeeID] = useState("");
  const [employeeTitle, setEmployeeTitle] = useState("");
  const [employeeName, setEmployeeName] = useState("");
  const [year, setYear] = useState("");
  const [isAuthenticated, setIsAuthenticated] = useState(false);

  const [allAbsences, setAbsences] = useState<AbsenceRequest[]>([]);
  const [archivedAbsences, setArchivedAbsences] = useState<AbsenceRequest[]>([]);
  const [pendingEmployeeAbsences, setPendingEmployeeAbsences] = useState<AbsenceRequest[]>([]);
  const [selectedAbsence, setSelectedAbsence] = useState<AbsenceRequest | null>(null);
  const [newAbsence, setNewAbsence] = useState(false);

  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [employeeInfo, setEmployeeInfo] = useState<any>(null);
  const [activeTab, setActiveTab] = useState('pending_employee');

  const [selectedEarlyOutForm, setSelectedEarlyOutForm] = useState<boolean>(false);

  const handleLogin = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setError("");

    try {
      // 1. Create a query to find the employee document where the 'employeeID' field matches
      const employeesRef = collection(db, "employees");
      //console.log("Attempting login for Employee ID:", employeeID);
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
        //console.log("Employee Name from DB:", data.firstName, data.lastName);
        //console.log("Employee Birth Year from DB:", employee_birthYear);
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

  useEffect(() => {
    if (!isAuthenticated) return;

    let timeoutId: NodeJS.Timeout;
    let lastActivityTime = Date.now(); // Store exact timestamp of last activity
    const INACTIVITY_LIMIT = 5 * 60 * 1000; // 5 minutes

    const resetTimer = () => {
      clearTimeout(timeoutId);
      lastActivityTime = Date.now(); // Update timestamp

      timeoutId = setTimeout(() => {
        handleLogout();
        alert("You have been signed out due to 5 minutes of inactivity.");
      }, INACTIVITY_LIMIT);
    };

    // Fix for iOS: Check elapsed time when the user returns to the tab
    const handleVisibilityChange = () => {
      if (document.visibilityState === 'visible') {
        const timeElapsedSinceLastActivity = Date.now() - lastActivityTime;
          
        // If they were away longer than 5 minutes, log them out instantly
        if (timeElapsedSinceLastActivity >= INACTIVITY_LIMIT) {
          handleLogout();
          alert("You have been signed out due to 5 minutes of inactivity.");
        } else {
          // Otherwise, adjust the timer for the remaining time left
          clearTimeout(timeoutId);
          timeoutId = setTimeout(() => {
            handleLogout();
            alert("You have been signed out due to 5 minutes of inactivity.");
          }, INACTIVITY_LIMIT - timeElapsedSinceLastActivity);
        }
      }
    };

    const activityEvents = ['mousemove', 'keydown', 'click', 'scroll', 'touchstart'];

    activityEvents.forEach(event => {
      window.addEventListener(event, resetTimer);
    });

    // Listen for when the user locks/unlocks their phone or switches tabs
    document.addEventListener('visibilitychange', handleVisibilityChange);

    resetTimer();

    return () => {
      clearTimeout(timeoutId);
      activityEvents.forEach(event => {
        window.removeEventListener(event, resetTimer);
      });
      document.removeEventListener('visibilitychange', handleVisibilityChange);
    };
  }, [isAuthenticated]);

  const fetchAbsences = async (id: string) => {
    try {
      setLoading(true);
      // This query remains the same as 'employee_id' is a field in the 'absences' collection and status is used to filter active requests
      const q = query(collection(db, "absences"), where("employee_id", "==", employeeID));
      const querySnapshot = await getDocs(q);
      const absenceData = querySnapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as AbsenceRequest));

      const activeAbsences = absenceData.filter(absence => (absence.status !== 'archived' && absence.status !== 'pending_action'));
      const archivedAbsences = absenceData.filter(absence => absence.status === 'archived');
      const pendingActionAbsences = absenceData.filter(absence => absence.status === 'pending_action')
      setAbsences(activeAbsences);
      setArchivedAbsences(archivedAbsences);
      setPendingEmployeeAbsences(pendingActionAbsences);
    } 
    catch (err) {
      setAbsences([]);
      setError("Failed to fetch absences.");
    }
    setLoading(false);
  };

  // Combine both live active groups into an evaluation array for checking conflicts
  const conflictingRequestsEvaluation = useMemo(() => {
    return [...allAbsences, ...pendingEmployeeAbsences];
  }, [allAbsences, pendingEmployeeAbsences]);

  const handleFormSubmit = () => {
    setSelectedAbsence(null);
    setNewAbsence(false);
    fetchAbsences(employeeID); // Refresh the list of absences
  };

  const handleViewDetails = (absence: AbsenceRequest) => {
    setSelectedAbsence(absence);
  };

  const pendingEmployeeActionAbsences = useMemo (() => {
    return pendingEmployeeAbsences;
  }, [allAbsences])

  const noteSubmisisonAbsences = useMemo(() => {
    return allAbsences.filter(absence => absence.excuse_note_submitted === 'pending');
  }, [allAbsences]);

  const pendingAbsences = useMemo(() => {
    return allAbsences.filter(absence => absence.manager_approval === 'pending' || absence.final_approval === 'pending');
  }, [allAbsences]);

  const approvedAbsences = useMemo(() => {
    return allAbsences.filter(absence => (absence.manager_approval === 'approved' && (absence.final_approval === 'approved' || absence.final_approval === 'approved_with_note')) || (absence.manager_approval === 'not_required' && (absence.final_approval === 'approved' || absence.final_approval === 'approved_with_note')));
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
            <input type="password" id="employeeId" value={employeeID} onChange={(e) => setEmployeeID(e.target.value)} className="w-full p-2 border rounded" required autoComplete="off"/>
          </div>
          <div className="mb-6">
            <label htmlFor="birthYear" className="block text-gray-700">Birth Year</label>
            <input type="password" id="birthYear" value={year} onChange={(e) => setYear(e.target.value)} className="w-full p-2 border rounded" required autoComplete="off"/>
          </div>
          <button type="submit" className="w-full bg-blue-500 text-white p-2 rounded cursor-pointer hover:bg-indigo-500 hover:border-indigo-400">Login</button>
          {error && <p className="text-red-500 text-center mt-4">{error}</p>}
          
          <div className="relative my-6 flex items-center justify-center">
            <div className="absolute inset-0 flex items-center"><span className="w-full border-t border-slate-200"></span></div>
            <span className="relative bg-white px-3 text-xs text-slate-400 font-bold uppercase">Or</span>
          </div>

          <button
            type="button"
            onClick={() => setSelectedEarlyOutForm(true)}
            className="w-full border-2 border-dashed border-indigo-200 text-indigo-600 p-2.5 rounded-lg font-bold hover:bg-indigo-50/50 hover:border-indigo-400 transition-all cursor-pointer text-center text-sm"
          >
            Submit Volunteer Early Out
          </button>
        </form>

        {selectedEarlyOutForm && <EarlyOutForm onClose={() => setSelectedEarlyOutForm(false)} />}
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
                onClick={() => setActiveTab('pending_employee')}
                className={`${activeTab === 'pending_employee' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm`}
              >
                Pending Employee ({pendingEmployeeActionAbsences.length})
              </button>
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
                All Active Requests ({allAbsences.length})
              </button>
              <button onClick={() => setActiveTab('archived')}
                className={`${activeTab === 'archived' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm`}
              >
                Archived ({archivedAbsences.length})
              </button>
            </nav>
          </div>

          {/* Absence Table */}
          
          <div className="mt-4">
            {activeTab === 'pending_employee' && (
              <AbsenceTable requests={pendingEmployeeActionAbsences} status={'pending_employee'} onViewDetails={handleViewDetails} />
            )}
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
            {activeTab === 'archived' && archivedAbsences.length > 0 && (
                <AbsenceTable requests={archivedAbsences} status={'archived'} onViewDetails={handleViewDetails} />
            )}
          </div>

          {newAbsence && (
          <NewAbsenceForm 
            employeeFirestore={employeeFirestoreID}
            employeeID={employeeID} 
            employeeTitle={employeeTitle} 
            employeeName={employeeName} 
            employeeSkipManagerApproval={employeeInfo.skipManagerApproval}
            employeeExistingRequests={conflictingRequestsEvaluation}
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