'use client';

import { db } from '@/lib/firebase.config'; // Adjust path if necessary
import { collection, query, where, getDoc, doc, updateDoc, serverTimestamp } from "firebase/firestore";
import React, { useState, useEffect, useMemo } from 'react';
import { onAuthStateChanged, User, getAuth } from "firebase/auth";
import { auth } from '@/lib/firebase.config';
import { getAbsenceRequestsByOffices } from '@/components/api/approval';
import { RequestTable } from '@/components/request-table';
import { AbsenceRequest } from '@/lib/types';

export default function Page() {
  const [user, setUser] = useState<User | null>(null);
  const [userName, setUserName] = useState<string | "">("");
  const [userRole, setUserRole] = useState<string | "">("");
  //const [userTitles, ]
  const [userOffices, setUserOffices] = useState<string[]>([]);

  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const [allRequests, setAllRequests] = useState<AbsenceRequest[]>([]);
  const [activeTab, setActiveTab] = useState('all');

  const [selectedRequest, setSelectedRequest] = useState<AbsenceRequest | null>(null);
  const [managerNotes, setManagerNotes] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      setLoading(true);
      setError(null);
      if (currentUser) {
        setUser(currentUser);

        const userDocRef = doc(db, 'users', currentUser.uid);

        try{
          const userDocSnap = await getDoc(userDocRef);
          if (userDocSnap.exists()) {
            const userData = userDocSnap.data();

            const name = currentUser.displayName || "";
            const role = userData.role || null;
            const offices = userData.offices || [];

            setUserName(name);
            setUserRole(role);
            setUserOffices(offices);
            // Fetch absence requests based on the user's offices and role
            if ((userData.offices).length > 0) {
              const requests = await getAbsenceRequestsByOffices(offices, userName, role);
              setAllRequests(requests as AbsenceRequest[]);
            }
            else{
              setAllRequests([]);
            }
          }
          else{
            setUserRole("");
          }
        }
        catch (err:any){
          console.error("Error fetching user data or requests:", err);
          setError("Failed to load data. Please try again later.");
        }
      }
      else {
        setUser(null);
        setUserRole("");
        setAllRequests([]);
      }
      setLoading(false);
    });
    return () => unsubscribe();
  }, []); // The empty array ensures this effect runs only once on component mount

  /*  Pending Requests
      HR: Will get requests that have been approved by Managers and are awaiting final approval
      Managers: Will get requests that are pending their approval, first stage of approval for most employees
      Director: Will get the same requests as HR, but will approve Doctors, Supervisors, Exempt, and EXEC directly
  */
  const pendingRequests = useMemo(() => {
    if(userRole === 'HR'){
      return allRequests.filter(r => r.manager_approval === 'approved' && r.final_approval === 'pending');
    }
    else if(userRole === 'Manager'){
      return allRequests.filter(r => r.manager_approval === 'pending'  && r.final_approval === 'pending');
    }
    else if(userRole === 'Director'){
      return allRequests.filter(r => r.manager_approval === 'approved' && r.final_approval === 'pending');
    }
    else {
      return [];
    }
  }, [allRequests, userRole]);

  /*  Doctor, Supervisors, Exempt, EXEC Requests should go straight to Maria for final approval
      Does not need manager approval
      This will be a separate tab from the normal pending requests
  */
  const pendingExemptRequests = useMemo(() => {
    if(userRole === 'Director'){
      return allRequests.filter(r => r.final_approval === 'pending' && ['Dentist', 'Supervisor', 'Exempt', 'EXEC'].includes(r.employee_title));
    }
    else {
      return [];
    }
  }, [allRequests, userRole]);

  /*  Approval Requests
      HR: Will see all requests that are fully approved
      Managers: Will see all requests they have approved
      Director: Will see all requests that are fully approved
  */
  const approvedRequests = useMemo(() => {
    if(userRole === 'HR' || userRole === 'Director'){
      return allRequests.filter(r => r.final_approval === 'approved');
    }
    else if(userRole === 'Manager'){
      return allRequests.filter(r => r.manager_approval === 'approved');
    }
    else {
      return [];
    }
  }, [allRequests, userRole]);

  /*  Denied Requests
      HR: Will see all requests that have been denied at any stage
      Managers: Will see all requests they have denied
      Director: Will see all requests that have been denied at any stage
  */
  const deniedRequests = useMemo(() => {
    if(userRole === 'HR' || userRole === 'Director'){
      return allRequests.filter(r => r.final_approval === 'denied' || r.manager_approval === 'denied');
    }
    else if(userRole === 'Manager'){
      return allRequests.filter(r => r.manager_approval === 'denied');
    }
    else {
      return [];
    }
  }, [allRequests, userRole]);

  // Event Handlers for Request Actions
  const handleOpenModal = (request: AbsenceRequest) => {
    setSelectedRequest(request);
    // Pre-fill notes if they already exist from a previous edit
    setManagerNotes(request.manager_notes || '');
  };

  const handleCloseModal = () => {
    setSelectedRequest(null);
    setManagerNotes('');
  };

  const handleApproval = async (decision: 'approved' | 'denied') => {
    if (!selectedRequest || !user) return;

    setIsSubmitting(true);
    const docRef = doc(db, 'absences', selectedRequest.id);

    try {
      let updateData: any = {};
      if (userRole === 'Manager') { // Manager View
        updateData = {
          manager_approval: decision,
          manager_notes: managerNotes,
          manager_name: user.displayName,
          // You might want a timestamp for the action
          // manager_action_date: serverTimestamp() 
        };
      } else { // HR or Director View
        updateData = {
          final_approval: decision,
          final_name: user.displayName,
          // HR can add to the notes, or you could have a separate `final_notes` field
          manager_notes: managerNotes,
          // final_action_date: serverTimestamp()
        };
      }

      await updateDoc(docRef, updateData);

      // Update local state for immediate UI feedback
      setAllRequests(prevRequests =>
        prevRequests.map(req =>
          req.id === selectedRequest.id ? { ...req, ...updateData } : req
        )
      );

      handleCloseModal();
    } catch (err) {
      console.error("Failed to update document: ", err);
      alert("An error occurred. Could not process the request.");
    } finally {
      setIsSubmitting(false);
    }
  };

  // Only render while loading or if there's an error
  if (loading) {
    return <div>Loading...</div>;
  }
  if (error) {
    return <div className="p-4 text-red-500">{error}</div>;
  }

  return (
    <div className="p-2 md:p-4">
      <h1 className="text-2xl font-bold mb-4">Approval Dashboard</h1>
      <p className="mb-4 text-gray-600">Role: <span className="font-semibold">{userRole}</span></p>
      
      {/* Tabs */}
      <div className="border-b border-gray-200">
        <nav className="-mb-px flex space-x-8" aria-label="Tabs">
          <button onClick={() => setActiveTab('all')}
            className={`${activeTab === 'all' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm`}
          >
            All Requests ({allRequests.length})
          </button>
          <button
            onClick={() => setActiveTab('pending')}
            className={`${activeTab === 'pending' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm`}
          >
            Pending ({pendingRequests.length})
          </button>

          {/* Director-only Tab */}
          {userRole === 'Director' && (
             <button
                onClick={() => setActiveTab('pending_exempt')}
                className={`${activeTab === 'pending_exempt' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm`}
             >
                Pending Exempt ({pendingExemptRequests.length})
             </button>
          )}

          <button
            onClick={() => setActiveTab('approved')}
            className={`${activeTab === 'approved' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm`}
          >
            Approved ({approvedRequests.length})
          </button>
          <button
            onClick={() => setActiveTab('denied')}
            className={`${activeTab === 'denied' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm`}
          >
            Denied ({deniedRequests.length})
          </button>
        </nav>
      </div>

      {/* Request Table */}
      <div className="mt-4">
        {activeTab === 'all' && (
          <RequestTable requests={allRequests} userRole={userRole} onViewDetails={handleOpenModal} />
        )}
        {activeTab === 'pending' && (
          <RequestTable requests={pendingRequests} userRole={userRole}  onViewDetails={handleOpenModal} />
        )}
        {activeTab === 'pending_exempt' && userRole === 'Director' && (
           <RequestTable requests={pendingExemptRequests} userRole={userRole}  onViewDetails={handleOpenModal} />
        )}
        {activeTab === 'approved' && (
           <RequestTable requests={approvedRequests} userRole={userRole}  onViewDetails={handleOpenModal} />
        )}
         {activeTab === 'denied' && (
           <RequestTable requests={deniedRequests} userRole={userRole}  onViewDetails={handleOpenModal} />
        )}
      </div>

      {/* Modal for Request Details and Approval */}
      {selectedRequest && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 backdrop-blur-sm">
          <div className="bg-white rounded-lg shadow-lg max-w-3xl w-full p-6 relative" onClick={(e) => e.stopPropagation()}>

            {/* Modal Header */}
            <div className="flex justify-between items-center p-5 border-b sticky top-0 bg-white">
              <h2 className="text-xl font-bold text-gray-800">{selectedRequest.date_submitted} {selectedRequest.employee_name}'s Request</h2>
              <button onClick={handleCloseModal} className="absolute top-4 right-4 text-gray-500 hover:text-gray-700">
                <svg className="w-5 h-5" fill="currentColor" viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg"><path fillRule="evenodd" d="M4.293 4.293a1 1 0 011.414 0L10 8.586l4.293-4.293a1 1 0 111.414 1.414L11.414 10l4.293 4.293a1 1 0 01-1.414 1.414L10 11.414l-4.293 4.293a1 1 0 01-1.414-1.414L8.586 10 4.293 5.707a1 1 0 010-1.414z" clipRule="evenodd"></path></svg>
              </button>
            </div>

            {/* Modal Content Body*/}
            <div className="p-6 space-y-6">
              {/* Request Details */}
              <div>
                <h3 className="text-lg font-semibold text-gray-700 mb-2">Request Details</h3>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-x-4 gap-y-2 text-sm">
                  <p><strong>Employee:</strong> {selectedRequest.employee_name}</p>
                  <p><strong>Employee ID:</strong> {selectedRequest.employee_id}</p>
                  <p><strong>Office:</strong> {selectedRequest.office}</p>
                  <p><strong>Title:</strong> {selectedRequest.employee_title}</p>
                  <p><strong>Request Type:</strong> {selectedRequest.type_of_request}</p>
                  <p><strong>Leave Type:</strong> {selectedRequest.type_of_incident}</p>

                  {selectedRequest.eta_etd && (
                    <p className="md:col-span-2"><strong>ETA/ETD:</strong> <span className="text-gray-600 italic">{selectedRequest.eta_etd}</span></p>
                  )}

                  <p className="md:col-span-2 align-items:center"><strong>Date:</strong> {new Date(selectedRequest.incident_start).toLocaleDateString()} - {new Date(selectedRequest.incident_end).toLocaleDateString()}</p>
                  {selectedRequest.employee_comments && (
                    <p className="md:col-span-2"><strong>Employee Notes:</strong> <span className="text-gray-600 italic">{selectedRequest.employee_comments}</span></p>
                  )}

                  <p className="md:col-span-2">
                    <strong>Excuse Note:</strong>{" "}
                    {selectedRequest.excuse_note_submitted === "submitted" ? (
                      <span className="text-green-600 italic font-semibold">Submitted</span>
                    ) : selectedRequest.excuse_note_submitted === "not_provided" ? (
                      <span className="text-red-600 italic font-semibold">Not Provided</span>
                    ) : (
                      <span className="text-yellow-600 italic font-semibold">Pending</span>
                    )}
                  </p>

                  {selectedRequest.excuse_note_submitted === "submitted" && selectedRequest.excuse_note && selectedRequest.excuse_note.length > 0 && (
                    <div>
                    <p className="text-sm font-medium text-gray-600">Already Submitted:</p>
                    <ul className="list-none pl-3 mt-1 text-sm space-y-2">
                        {selectedRequest.excuse_note.map((url: string, index: number) => (
                            <li key={index}>
                                <a href={url} target="_blank" rel="noopener noreferrer" className={`text-blue-600 hover:underline`}>
                                    Excuse Note {index + 1}
                                </a>
                            </li>
                        ))}
                    </ul>
                    </div>
                  )}
                </div>
              </div>

                {/* Modal Approval History*/}
              <div>
                <h3 className="text-lg font-semibold text-gray-700 mb-2">Approval History</h3>
                <div className="space-y-3 text-sm">

                  {/* Manager Approval */}
                  {!(['Dentist', 'Supervisor', 'Exempt', 'EXEC'].includes(selectedRequest.employee_title)) && (
                  <div className="p-3 bg-gray-50 rounded-md">
                    <p><strong>Manager Approval:</strong> <span className={`font-semibold italic capitalize ${
                      selectedRequest.manager_approval === 'approved' ? 'text-green-600' :
                      selectedRequest.manager_approval === 'denied' ? 'text-red-600' :
                      'text-yellow-600'
                    }`}>{selectedRequest.manager_approval}</span></p>
                    {selectedRequest.manager_name && <p><strong>Manager:</strong> {selectedRequest.manager_name}</p>}
                  </div>
                  )}

                  {/* Final Approval */}
                  {userRole !== 'Manager' && selectedRequest.manager_approval !== 'denied' && (
                  <div className="p-3 bg-gray-50 rounded-md">
                  <p><strong>Final Approval:</strong> <span className={`font-semibold italic capitalize ${
                    selectedRequest.final_approval === 'approved' ? 'text-green-600' :
                    selectedRequest.final_approval === 'denied' ? 'text-red-600' :
                    'text-yellow-600'
                  }`}>{selectedRequest.final_approval}</span></p>
                  {selectedRequest.final_name && <p><strong>Final Approver:</strong> {selectedRequest.final_name}</p>}
                  </div>
                  )}

                  {/* Manager Notes */}
                  {selectedRequest.manager_notes && 
                  ((selectedRequest.final_approval === 'approved' && ['Dentist', 'Supervisor', 'Exempt', 'EXEC'].includes(selectedRequest.employee_title)) ||
                  (selectedRequest.final_approval === 'approved' && selectedRequest.manager_approval === 'approved')) ? (
                    <div className="p-3 bg-gray-50 rounded-md">
                      <p><strong>Notes:</strong> <span className="text-gray-600 italic">{selectedRequest.manager_notes}</span></p>
                    </div>
                  ) : (
                    <div className="p-3 bg-gray-50 rounded-md">
                      <p><strong>Notes:</strong> <span className="text-gray-600 italic">No notes provided.</span></p>
                    </div>
                  )}

                  {(userRole === 'Manager' && selectedRequest.manager_approval === 'pending') ||
                  ((userRole === 'HR' || userRole === 'Director') && 
                  selectedRequest.final_approval === 'pending' && 
                  (selectedRequest.manager_approval === 'approved' || ['Dentist', 'Supervisor', 'Exempt', 'EXEC'].includes(selectedRequest.employee_title))) ? (
                <div>
                  <h3 className="text-lg font-semibold text-gray-700 mb-2">Take Action</h3>
                  <label htmlFor="managerNotes" className="block mb-2 text-sm font-medium text-gray-900">Notes (Optional):</label>
                  <textarea
                    id="managerNotes"
                    rows={3}
                    className="block p-2.5 w-full text-sm text-gray-900 bg-gray-50 rounded-lg border border-gray-300 focus:ring-blue-500 focus:border-blue-500"
                    placeholder="Add notes for your decision..."
                    value={managerNotes}
                    onChange={(e) => setManagerNotes(e.target.value)}
                  ></textarea>
                </div>
              ) : null}

                </div>
              </div>
            </div>


            {/* Modal Footer - Approval Actions */}
            <div className="flex items-center justify-end p-5 border-t space-x-3 sticky bottom-0 bg-white">
              <button 
                onClick={handleCloseModal}
                disabled={isSubmitting}
                className="text-gray-500 bg-white hover:bg-gray-100 focus:ring-4 focus:outline-none focus:ring-blue-300 rounded-lg border border-gray-200 text-sm font-medium px-5 py-2.5 hover:text-gray-900 focus:z-10 disabled:opacity-50"
              >
                Close
              </button>
              {/* Action Buttons - Conditionally Rendered */}
              {(userRole === 'Manager' && selectedRequest.manager_approval === 'pending') ||
               ((userRole === 'HR' || userRole === 'Director') && 
                selectedRequest.final_approval === 'pending' && 
                (selectedRequest.manager_approval === 'approved' || ['Dentist', 'Supervisor', 'Exempt', 'EXEC'].includes(selectedRequest.employee_title))) ? (
                <>
                  <button
                    onClick={() => handleApproval('denied')}
                    disabled={isSubmitting}
                    className="text-white bg-red-600 hover:bg-red-700 focus:ring-4 focus:outline-none focus:ring-red-300 font-medium rounded-lg text-sm px-5 py-2.5 text-center disabled:bg-red-400"
                  >
                    {isSubmitting ? 'Submitting...' : 'Deny'}
                  </button>
                  <button
                    onClick={() => handleApproval('approved')}
                    disabled={isSubmitting}
                    className="text-white bg-green-600 hover:bg-green-700 focus:ring-4 focus:outline-none focus:ring-green-300 font-medium rounded-lg text-sm px-5 py-2.5 text-center disabled:bg-green-400"
                  >
                    {isSubmitting ? 'Submitting...' : 'Approve'}
                  </button>
                </>
              ) : null}
            </div>

          </div>
        </div>
      )}
    </div>
  );
}