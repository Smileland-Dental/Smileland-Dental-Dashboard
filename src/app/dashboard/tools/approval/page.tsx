'use client';

import { db } from '@/lib/firebase.config';
import { doc, updateDoc } from "firebase/firestore";
import React, { useState, useEffect, useMemo } from 'react';
import { useAuth } from '@/contexts/AuthContext'; // Using your existing context

import { User } from '@/lib/types';
import { getAbsenceRequestsByUser } from '@/components/approval/approval';
import { RequestTable } from '@/components/approval/request-table';
import { AbsenceRequest } from '@/lib/types';
import RequestDetailsModal from '@/components/approval/request-details-modal';
import ProtectedRoute from '@/components/auth/ProtectedRoute';

export default function Page() {
  // Destructure from useAuth context
  const { user, loading: authLoading } = useAuth();
  
  const [loadingRequests, setLoadingRequests] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [allRequests, setAllRequests] = useState<AbsenceRequest[]>([]);
  const [pendingEmployeeAbsences, setPendingEmployeeAbsences] = useState<AbsenceRequest[]>([]);
  const [activeTab, setActiveTab] = useState('all');

  // --- Date Filter State ---
  const [startDate, setStartDate] = useState(() => {
    const d = new Date();
    d.setDate(d.getDate() - 30); // Default to last 30 days
    return d.toISOString().split('T')[0];
  });
  const [endDate, setEndDate] = useState(() => {
    const d = new Date();
    d.setDate(d.getDate() + 30); // Default to next 30 days
    return d.toISOString().split('T')[0];
  });

  const [selectedRequest, setSelectedRequest] = useState<AbsenceRequest | null>(null);
  const [managerNotes, setManagerNotes] = useState('');
  const [finalNotes, setFinalNotes] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);

  // Derived values from context to keep code readable
  const userRole = user?.role || "";
  const userOffices = user?.offices || [];
  const userName = user?.username || "";
  console.log("Fetched requests for user:", userName, "Role:", userRole, "Offices:", userOffices, "Managed Employees:", user?.managedEmployeeIds);

  useEffect(() => {
    const fetchRequests = async () => {
      // Only fetch if we have a user and they have assigned offices
      if (user) {
        setLoadingRequests(true);
        setError(null);
        try {
          const requests = await getAbsenceRequestsByUser(
            user, startDate, endDate
          ); // Filter out archived requests at this stage
          setAllRequests((requests as AbsenceRequest[]).filter(r => (r.status !== 'archived' && r.status !== 'pending_action')));
          setPendingEmployeeAbsences((requests as AbsenceRequest[]).filter(r => r.status === 'pending_action'))
        } catch (err: any) {
          console.error("Error fetching requests:", err);
          setError("Failed to load requests. Please try again later.");
        } finally {
          setLoadingRequests(false);
        }
      } 
    };

    if (!authLoading) {
      fetchRequests();
      console.log("All Requests:", allRequests);
    }
  }, [user, userOffices, userRole, userName, authLoading, startDate, endDate]);

  const finalizedAllRequests = useMemo(() => {
    return allRequests.filter(r => {
      if (userRole === 'Manager') {
        return r.manager_approval !== 'not_required';
      }
      return allRequests;
    });
  }, [allRequests, userRole]);

  const pendingEmployeeActionAbsences = useMemo (() => {
    return pendingEmployeeAbsences;
  }, [allRequests, userRole])

  // Logic for filtering requests (remain unchanged but use derived userRole)
  const pendingRequests = useMemo(() => {
    return allRequests.filter(r => {
        // 1. If Manager: Show if they haven't acted yet
        if (userRole === 'Manager') {
          return r.manager_approval === 'pending' && r.final_approval === 'pending';
        }
        
        // 2. If HR/Director: Show if it's ready for final sign-off 
        // Show pending if manager approval is done but final approval is pending
        if (userRole === 'HR' || userRole === 'Director') {;
          return r.manager_approval === 'pending' && r.final_approval === 'pending';
        }
        return false;
      });
  }, [allRequests, userRole]);

  const pendingFinalApproval = useMemo(() => {
    // Director and HR only: Show requests that are pending and marked as exempt from manager approval (skipManagerApproval = true)
    return allRequests.filter(r => {
      if (userRole === 'Director' || userRole === 'HR') {
        return (r.manager_approval === 'not_required' && r.final_approval === 'pending') || (r.manager_approval === 'approved' && r.final_approval === 'pending'); // For now, show all pending requests for Directors. Adjust logic as needed for exempt filtering.
      }
      return false;
    });
  }, [allRequests, userRole]);

  const approvedRequests = useMemo(() => {
    return allRequests.filter(r => {
      if (userRole === 'Manager') {
        return (r.manager_approval === 'approved' && (r.final_approval === 'pending' || r.final_approval === 'approved' || r.final_approval === 'approved_with_note'));
      }
      if (userRole === 'HR' || userRole === 'Director') {
        return r.final_approval === 'approved' || r.final_approval === 'approved_with_note';
      }
      return false;
    });
  }, [allRequests, userRole]);

  const deniedRequests = useMemo(() => {
    if (userRole === 'HR' || userRole === 'Director') {
      return allRequests.filter(r => r.final_approval === 'denied' || r.manager_approval === 'denied');
    } else if (userRole === 'Manager') {
      return allRequests.filter(r => r.manager_approval === 'denied' || r.final_approval === 'denied');
    }
    return [];
  }, [allRequests, userRole]);

  const handleOpenModal = (request: AbsenceRequest) => {
    setSelectedRequest(request);
    setManagerNotes(request.manager_notes || '');
    setFinalNotes(request.final_notes || '');
  };

  const handleCloseModal = () => {
    setSelectedRequest(null);
    setManagerNotes('');
    setFinalNotes('');
  };

  const handleApproval = async (decision: 'approved' | 'denied' | 'approved_with_note') => {
    if (!selectedRequest || !user) return;

    setIsSubmitting(true);
    const docRef = doc(db, 'absences', selectedRequest.id);

    try {
      let updateData: any = {};
        if (userRole === 'Manager') {
          updateData = {
            manager_approval: decision,
            manager_notes: managerNotes,
            manager_approval_name: userName,
          };
        } else {
          updateData = {
            final_approval: decision,
            final_approval_name: userName, // Fixed the typo from your snippet
            manager_notes: managerNotes,
            final_notes: finalNotes
          };
      }

      await updateDoc(docRef, updateData);

      setAllRequests(prevRequests =>
        prevRequests.map(req =>
          (req.id === selectedRequest.id ? { ...req, ...updateData } : req)
        ).filter(req => req.status !== 'archived')
      );

      handleCloseModal();
    } catch (err) {
      console.error("Failed to update document: ", err);
      alert("An error occurred. Could not process the request.");
    } finally {
      setIsSubmitting(false);
    }
  };

  // Show loading state if either Auth or the initial Request fetch is happening
  if (authLoading || (loadingRequests && allRequests.length === 0)) {
    return (
      <div className="flex items-center justify-center min-h-screen">
        <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-indigo-600"></div>
        <span className="ml-2">Loading Dashboard...</span>
      </div>
    );
  }

  if (error) {
    return <div className="p-4 text-red-500">{error}</div>;
  }

  if (!user) {
    return <div className="p-4">Please log in to view the approval dashboard.</div>;
  }

  return (
    <ProtectedRoute allowedRoles={['HR', 'Director', 'Manager']}>
    <div className="p-2 md:p-4">
      <h1 className="text-3xl font-black mb-4">Approval Dashboard</h1>

      <div className="flex flex-wrap items-end gap-4 mb-6 bg-white p-4 rounded-lg shadow-sm border">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">Start Date</label>
          <input 
            type="date" 
            value={startDate} 
            onChange={(e) => setStartDate(e.target.value)}
            className="block w-full border-gray-300 rounded-md shadow-sm focus:ring-indigo-500 focus:border-indigo-500 sm:text-sm border p-2"
          />
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">End Date</label>
          <input 
            type="date" 
            value={endDate} 
            onChange={(e) => setEndDate(e.target.value)}
            className="block w-full border-gray-300 rounded-md shadow-sm focus:ring-indigo-500 focus:border-indigo-500 sm:text-sm border p-2"
          />
        </div>
        <p className="mb-4 text-gray-600">Role: <span className="font-semibold">{userRole || 'Unassigned'}</span></p>
      </div>
      
      {/* Tabs */}
      {/* Scrollable Tabs Container */}
      <div className="border-b border-gray-200">
        <nav 
          className="-mb-px flex space-x-8 overflow-x-auto scrollbar-hide px-4 md:px-0" 
          aria-label="Tabs"
          style={{ 
            WebkitOverflowScrolling: 'touch',
            // This ensures the last item has space on the right after scrolling
            paddingRight: '1rem' 
          }} 
        >
          <button 
            onClick={() => setActiveTab('all')}
            className={`${activeTab === 'all' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm transition-colors flex-shrink-0`}
          >
            All Active Requests ({finalizedAllRequests.length})
          </button>

          <button
            onClick={() => setActiveTab('pending')}
            className={`${activeTab === 'pending' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm transition-colors flex-shrink-0`}
          >
            {userRole === 'Manager' ? 'Pending' : 'Pending Manager Approval'} ({pendingRequests.length})
          </button>

          {(userRole === 'Director'  || userRole === 'HR') && (
             <button
                onClick={() => setActiveTab('pending_final')}
                className={`${activeTab === 'pending_final' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm transition-colors flex-shrink-0`}
             >
                Pending Final Approval ({pendingFinalApproval.length})
             </button>
          )}

          <button
            onClick={() => setActiveTab('approved')}
            className={`${activeTab === 'approved' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm transition-colors flex-shrink-0`}
          >
            Approved ({approvedRequests.length})
          </button>

          <button
            onClick={() => setActiveTab('denied')}
            className={`${activeTab === 'denied' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm transition-colors flex-shrink-0`}
          >
            Denied ({deniedRequests.length})
          </button>

          <button
            onClick={() => setActiveTab('pending_employee')}
            className={`${activeTab === 'pending_employee' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm transition-colors flex-shrink-0`}
          >
            Pending Employee Action ({pendingEmployeeActionAbsences.length})
          </button>
          
          {/* Invisible spacer to ensure padding on the far right when scrolled to the end */}
          <div className="flex-shrink-0 w-4 md:hidden" aria-hidden="true" />
        </nav>
      </div>
           
      <div className="mt-4">
        {activeTab === 'all' && (
          <RequestTable requests={finalizedAllRequests} userRole={userRole} onViewDetails={handleOpenModal} />
        )}
        {activeTab === 'pending' && (
          <RequestTable requests={pendingRequests} userRole={userRole} onViewDetails={handleOpenModal} />
        )}
        {activeTab === 'pending_final' && (userRole === 'Director' || userRole === 'HR') && (
           <RequestTable requests={pendingFinalApproval} userRole={userRole} onViewDetails={handleOpenModal} />
        )}
        {activeTab === 'approved' && (
           <RequestTable requests={approvedRequests} userRole={userRole} onViewDetails={handleOpenModal} />
        )}
         {activeTab === 'denied' && (
           <RequestTable requests={deniedRequests} userRole={userRole} onViewDetails={handleOpenModal} />
        )}
        {activeTab === 'pending_employee' && (
           <RequestTable requests={pendingEmployeeActionAbsences} userRole={userRole} onViewDetails={handleOpenModal} />
        )}
      </div>

      {selectedRequest && (
        <RequestDetailsModal
          selectedRequest={selectedRequest}
          userRole={userRole}
          managerNotes={managerNotes}
          finalNotes={finalNotes}
          isSubmitting={isSubmitting}
          setManagerNotes={setManagerNotes}
          setFinalNotes={setFinalNotes}
          handleCloseModal={handleCloseModal}
          handleApproval={handleApproval}
        />
      )}
    </div>
    </ProtectedRoute>
  );
}
