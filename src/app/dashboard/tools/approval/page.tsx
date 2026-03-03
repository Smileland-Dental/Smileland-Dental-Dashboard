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

export default function Page() {
  // Destructure from useAuth context
  const { user, loading: authLoading } = useAuth();
  
  const [loadingRequests, setLoadingRequests] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [allRequests, setAllRequests] = useState<AbsenceRequest[]>([]);
  const [activeTab, setActiveTab] = useState('all');

  const [selectedRequest, setSelectedRequest] = useState<AbsenceRequest | null>(null);
  const [managerNotes, setManagerNotes] = useState('');
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
          user,
        );
          setAllRequests(requests as AbsenceRequest[]);
        } catch (err: any) {
          console.error("Error fetching requests:", err);
          setError("Failed to load requests. Please try again later.");
        } finally {
          setLoadingRequests(false);
        }
      } else if (user) {
        // User exists but has no offices assigned
        setAllRequests([]);
      }
    };

    if (!authLoading) {
      fetchRequests();
      console.log("All Requests:", allRequests);
    }
  }, [user, userOffices, userRole, userName, authLoading]);

  // Logic for filtering requests (remain unchanged but use derived userRole)
  const pendingRequests = useMemo(() => {
    if (userRole === 'HR') {
      return allRequests.filter(r => r.manager_approval === 'approved' && r.final_approval === 'pending');
    } else if (userRole === 'Manager') {
      return allRequests.filter(r => r.manager_approval === 'pending' && r.final_approval === 'pending');
    } else if (userRole === 'Director') {
      return allRequests.filter(r => r.manager_approval === 'approved' && r.final_approval === 'pending');
    }
    return [];
  }, [allRequests, userRole]);

  const pendingExemptRequests = useMemo(() => {
    if (userRole === 'Director') {
      return allRequests.filter(r => r.final_approval === 'pending' && ['Dentist', 'Supervisor', 'Exempt', 'EXEC'].includes(r.employee_title));
    }
    return [];
  }, [allRequests, userRole]);

  const approvedRequests = useMemo(() => {
    if (userRole === 'HR' || userRole === 'Director') {
      return allRequests.filter(r => r.final_approval === 'approved');
    } else if (userRole === 'Manager') {
      return allRequests.filter(r => r.manager_approval === 'approved');
    }
    return [];
  }, [allRequests, userRole]);

  const deniedRequests = useMemo(() => {
    if (userRole === 'HR' || userRole === 'Director') {
      return allRequests.filter(r => r.final_approval === 'denied' || r.manager_approval === 'denied');
    } else if (userRole === 'Manager') {
      return allRequests.filter(r => r.manager_approval === 'denied');
    }
    return [];
  }, [allRequests, userRole]);

  const handleOpenModal = (request: AbsenceRequest) => {
    setSelectedRequest(request);
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
        };
      }

      await updateDoc(docRef, updateData);

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
    <div className="p-2 md:p-4">
      <h1 className="text-2xl font-bold mb-4">Approval Dashboard</h1>
      <p className="mb-4 text-gray-600">Role: <span className="font-semibold">{userRole || 'Unassigned'}</span></p>
      
      {/* Tabs */}
      {/* Scrollable Tabs Container */}
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
            All Requests ({allRequests.length})
          </button>

          <button
            onClick={() => setActiveTab('pending')}
            className={`${activeTab === 'pending' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm transition-colors flex-shrink-0`}
          >
            Pending ({pendingRequests.length})
          </button>

          {userRole === 'Director' && (
             <button
                onClick={() => setActiveTab('pending_exempt')}
                className={`${activeTab === 'pending_exempt' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm transition-colors flex-shrink-0`}
             >
                Pending Exempt ({pendingExemptRequests.length})
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
          
          {/* Invisible spacer to ensure padding on the far right when scrolled to the end */}
          <div className="flex-shrink-0 w-4 md:hidden" aria-hidden="true" />
        </nav>
      </div>
           
      <div className="mt-4">
        {activeTab === 'all' && (
          <RequestTable requests={allRequests} userRole={userRole} onViewDetails={handleOpenModal} />
        )}
        {activeTab === 'pending' && (
          <RequestTable requests={pendingRequests} userRole={userRole} onViewDetails={handleOpenModal} />
        )}
        {activeTab === 'pending_exempt' && userRole === 'Director' && (
           <RequestTable requests={pendingExemptRequests} userRole={userRole} onViewDetails={handleOpenModal} />
        )}
        {activeTab === 'approved' && (
           <RequestTable requests={approvedRequests} userRole={userRole} onViewDetails={handleOpenModal} />
        )}
         {activeTab === 'denied' && (
           <RequestTable requests={deniedRequests} userRole={userRole} onViewDetails={handleOpenModal} />
        )}
      </div>

      {selectedRequest && (
        <RequestDetailsModal
          selectedRequest={selectedRequest}
          userRole={userRole}
          managerNotes={managerNotes}
          isSubmitting={isSubmitting}
          setManagerNotes={setManagerNotes}
          handleCloseModal={handleCloseModal}
          handleApproval={handleApproval}
        />
      )}
    </div>
  );
}