import React, { useEffect, useState, useCallback, use } from 'react';
import { Loader2, AlertCircle, FileText } from 'lucide-react';
import { AbsenceRequest, Employee } from '@/lib/types';
import { db } from '@/lib/firebase.config';
import { collection, query, where, getDocs, orderBy, limit, 
  startAfter, 
  limitToLast, 
  endBefore,
  QueryDocumentSnapshot,
  DocumentData,
  getCountFromServer } from "firebase/firestore";

interface TabProps {
  employee: Employee;
}

const itemsPerPage = 5;

export const IncidentNoticesTab = ({ employee }: TabProps) => {
  const [allUserRequests, setAllUserRequests] = useState<AbsenceRequest[]>([]);
  const [loadingRequests, setLoadingRequests] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Pagination State
  const [currentPage, setCurrentPage] = useState(1);
  const [totalCount, setTotalCount] = useState(0);
  const [firstVisible, setFirstVisible] = useState<QueryDocumentSnapshot<DocumentData> | null>(null);
  const [lastVisible, setLastVisible] = useState<QueryDocumentSnapshot<DocumentData> | null>(null);

  useEffect(() => {
    const fetchCount = async () => {
      if (!employee?.id) return;
      const q = query(collection(db, "absences"), where("employeeFirestoreID", "==", employee.id));
      const snapshot = await getCountFromServer(q);
      setTotalCount(snapshot.data().count);
    };
    fetchCount();
  }, [employee.id]); // Use ID here, not the whole object

  const fetchRequests = useCallback(async (direction?: 'next' | 'prev') => {
    if (!employee?.id) return;

    setLoadingRequests(true);
    setError(null);

    try {
      const absencesRef = collection(db, "absences");
      let q = query(
        absencesRef,
        where("employeeFirestoreID", "==", employee.id),
        orderBy("incident_start", "desc")
      );

      // Apply pagination cursors
      if (direction === 'next' && lastVisible) {
        q = query(q, startAfter(lastVisible), limit(itemsPerPage));
      } else if (direction === 'prev' && firstVisible) {
        q = query(q, endBefore(firstVisible), limitToLast(itemsPerPage));
      } else {
        q = query(q, limit(itemsPerPage));
      }

      const snap = await getDocs(q);

      if (!snap.empty) {
        const results = snap.docs.map(doc => ({ id: doc.id, ...doc.data() } as AbsenceRequest));
        setAllUserRequests(results);
        setFirstVisible(snap.docs[0]);
        setLastVisible(snap.docs[snap.docs.length - 1]);
      }
    } catch (err: any) {
      console.error("Firebase Query Error:", err);
      setError("Could not load incident notices. Check Firestore indexes.");
    } finally {
      setLoadingRequests(false);
    }
  }, [employee.id, firstVisible, lastVisible]);

  // Initial load
  useEffect(() => {
    fetchRequests();
  }, [employee.id]);

  const handleNext = () => {
    setCurrentPage(prev => prev + 1);
    fetchRequests('next');
  };

  const handlePrev = () => {
    setCurrentPage(prev => prev - 1);
    fetchRequests('prev');
  };

  return (
    <div className="relative">
      <div className="sticky top-0 z-10 bg-white pb-4 border-b mb-4 flex justify-between items-center">
        <h3 className="text-2xl font-bold text-gray-800">Incident Notices</h3>
      <div className="flex items-center gap-2">
          <button
            onClick={handlePrev}
            disabled={currentPage === 1 || loadingRequests}
            className="px-3 py-1 text-sm font-semibold rounded-lg border border-gray-200 hover:bg-gray-50 transition-colors text-gray-600 
             disabled:opacity-50 
             disabled:bg-red-50 
             disabled:text-red-400 
             disabled:border-red-100 
             disabled:cursor-not-allowed"
          >
            Previous
          </button>
          
          <span className="text-xs font-bold text-gray-400 uppercase tabular-nums">
            Page {currentPage} of {Math.ceil(totalCount / itemsPerPage) || 1}
          </span>

          <button
            onClick={handleNext}
            disabled={currentPage * itemsPerPage >= totalCount || loadingRequests}
            className="px-3 py-1 text-sm font-semibold rounded-lg border border-gray-200 hover:bg-gray-50 transition-colors text-gray-600 
             disabled:opacity-50 
             disabled:bg-red-50 
             disabled:text-red-400 
             disabled:border-red-100 
             disabled:cursor-not-allowed"
          >
            Next
          </button>
        </div>
      </div>

      {/* Loading State */}
      {loadingRequests && (
        <div className="flex flex-col items-center justify-center py-20 text-blue-600">
          <Loader2 className="animate-spin mb-2" size={32} />
          <p className="text-gray-500 animate-pulse">Fetching records...</p>
        </div>
      )}

      {/* Error State */}
      {error && (
        <div className="flex items-center gap-3 p-4 bg-red-50 text-red-700 rounded-lg border border-red-200">
          <AlertCircle size={20} />
          <p className="text-sm font-medium">{error}</p>
        </div>
      )}

      {/* Empty State */}
      {!loadingRequests && !error && allUserRequests.length === 0 && (
        <div className="flex flex-col items-center justify-center py-20 border-2 border-dashed border-gray-200 rounded-xl bg-gray-50">
          <FileText className="text-gray-300 mb-2" size={40} />
          <p className="text-gray-500">No incident records found for this employee.</p>
        </div>
      )}

      {/* Data List */}
      <div className="space-y-3">
        {!loadingRequests && allUserRequests.map((request) => (
          <div key={request.id} className="p-4 border border-gray-200 rounded-lg transition-colors bg-white">
            <div className="flex justify-between items-start">
              <div>
                <p className="font-bold text-gray-900">{request.type_of_request || 'General Incident'}</p>
                <p className="text-sm text-gray-500">{request.type_of_incident || 'No description provided.'}</p>
              </div>
              <div className="text-right">
                <StatusBadge req={request} />
                <p className="text-[10px] text-gray-400 mt-2">{(request.incident_start === request.incident_end) ? request.incident_start : `${request.incident_start} - ${request.incident_end}`}</p>
              </div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

// Logic-heavy Status Badge (Manager + Final)
const StatusBadge = ({ req }: { req: AbsenceRequest }) => {
  if (req.manager_approval === 'denied' || req.final_approval === 'denied') {
    return <Badge color="red" text={req.manager_approval === 'denied' && req.final_approval === 'pending' ? "Manager Denied" : "Denied"} />;
  }
  if (req.final_approval === 'approved') return <Badge color="green" text="Fully Approved" />;
  if (req.manager_approval === 'approved') return <Badge color="yellow" text="Manager Approved" />;
  if (req.manager_approval === 'not_required' && req.final_approval === 'pending') return <Badge color="yellow" text="Pending Corp" />;
  return <Badge color="gray" text="Pending" />;
};

const Badge = ({ color, text }: { color: string; text: string }) => {
  const styles: Record<string, string> = {
    green: 'bg-emerald-100 text-emerald-700 ring-emerald-600/20',
    red: 'bg-rose-100 text-rose-700 ring-rose-600/20',
    yellow: 'bg-amber-100 text-amber-700 ring-amber-600/20',
    gray: 'bg-slate-100 text-slate-600 ring-slate-600/10',
  };
  return (
    <span className={`inline-flex items-center rounded-full px-2.5 py-1 text-[10px] font-black uppercase ring-1 ring-inset ${styles[color]}`}>
      {text}
    </span>
  );
};