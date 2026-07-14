'use client';
import React, { useState, useEffect, useMemo } from "react";
import { db, storage } from '@/lib/firebase.config';
import { collection, doc, updateDoc, addDoc, Timestamp, deleteDoc } from 'firebase/firestore';
import { ref, deleteObject } from "firebase/storage";
import { useAuth } from '@/contexts/AuthContext';
import { AbsenceRequest } from "@/lib/types";
import { getAbsenceRequestsByUser, getRequestStatusText } from '@/components/approval/approval';
import { StatusBadge } from "@/components/ui/status-badge";
import ProtectedRoute from '@/components/auth/ProtectedRoute';

// UI Components
import { Search, ChevronLeft, ChevronRight, Edit3, Calendar, FileDown, Plus } from 'lucide-react';
import { HRRequestDetailsModal } from "@/components/approval/hr-request-details-modal";
import { HRCreateAbsenceModal } from "@/components/forms/absence-request/hr-absence";
import * as XLSX from 'xlsx';

import { useSort } from '@/hooks/custom-hooks'
import { SortIcon } from '@/components/ui/table-sort'

import { OFFICES } from '@/lib/constants';
const itemsPerPage = 50;
//const incidentTypes = ["Late In", "Early Out", "Absent", "Leave and Come Back", "Long Lunch", "Switch Shift", "Cancel Cell"];

export default function Page() {
  const { user, loading: authLoading } = useAuth();
  
  // --- Date Filter State (Default to last 30 days) ---
  const [startDate, setStartDate] = useState(() => {
    const d = new Date();
    d.setDate(d.getDate() - 30);
    return d.toISOString().split('T')[0];
  });
  const [endDate, setEndDate] = useState(() => {
    const d = new Date();
    d.setDate(d.getDate() + 30);
    return d.toISOString().split('T')[0];
  });

  const [absences, setAbsences] = useState<AbsenceRequest[]>([]);
  const [searchTerm, setSearchTerm] = useState('');
  const [officeFilter, setOfficeFilter] = useState('all');
  const [activeFilter, setActiveFilter] = useState('active_and_pending');
  const [pendingNotesOnly, setPendingNotesOnly] = useState(false);

  const [currentPage, setCurrentPage] = useState(1);
  const [selectedAbsence, setSelectedAbsence] = useState<AbsenceRequest | null>(null);

  const [isAdding, setIsAdding] = useState(false);
  const [loading, setLoading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);

  // --- Server-Side Fetch Logic ---
  const fetchDocs = async (resetPages = false) => {
    if (!user) return;
    setLoading(true);
    try {
      // This calls your provided getAbsenceRequestsByUser function
      // which performs the server-side where() queries
      const data = await getAbsenceRequestsByUser(user, startDate, endDate);
      setAbsences(data);
      if (resetPages) {
        setCurrentPage(1); // Reset to page 1 on new fetch
      }
    } catch (err) {
      console.error("Fetch error:", err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (!authLoading) {
      fetchDocs(true);
    }
  }, [user, authLoading, startDate, endDate]);

  
  const { sortConfig, handleSort } = useSort();

  // --- Client-Side Search (Operating on the server-filtered set) ---
  const filteredRequests = useMemo(() => {
    const term = searchTerm.toLowerCase().trim();
  
    let result = absences.filter((a) => {
      // 1. Resolve fallback status context
      const currentStatus = a.status || 'active';

      // 2. Evaluate the status matching rule based on the selected option
      let statusMatches = currentStatus === activeFilter;
      if (activeFilter === 'active_and_pending') {
        statusMatches = currentStatus === 'active' || currentStatus === 'pending_action';
      }

      // 3. Return combined verification array checks
      return (
        statusMatches &&
        (officeFilter === 'all' || a.office === officeFilter) && 
        (
          a.employee_name.toLowerCase().includes(term) || 
          String(a.employee_id || '').toLowerCase().includes(term)
        )
      );
    });

    if (sortConfig.direction !== 'none' && sortConfig.key) {
      result = [...result].sort((a, b) => {
        let valA = a[sortConfig.key as keyof AbsenceRequest];
        let valB = b[sortConfig.key as keyof AbsenceRequest];

        // Format custom fields safely for evaluation
        if (sortConfig.key === 'createdAt') {
          valA = a.createdAt?.toDate ? a.createdAt.toDate().getTime() : 0;
          valB = b.createdAt?.toDate ? b.createdAt.toDate().getTime() : 0;
        }

        // --- ADDED TRAP DOOR FOR STATUS BADGE ---
        if (sortConfig.key === 'statusBadge') {
          valA = getRequestStatusText(a);
          valB = getRequestStatusText(b);
        }

        // Handle fallback conversions for missing types or empty values safely
        const strA = String(valA ?? '').toLowerCase();
        const strB = String(valB ?? '').toLowerCase();

        if (strA < strB) return sortConfig.direction === 'asc' ? -1 : 1;
        if (strA > strB) return sortConfig.direction === 'asc' ? 1 : -1;
        return 0;
      });
    }

    return result;
  }, [absences, searchTerm, officeFilter, activeFilter, sortConfig]);

  const paginated = useMemo(() => {
    const start = (currentPage - 1) * itemsPerPage;
    return filteredRequests.slice(start, start + itemsPerPage);
  }, [filteredRequests, currentPage]);

  const totalPages = Math.ceil(filteredRequests.length / itemsPerPage) || 1;

  // --- Actions ---
  const downloadExcel = () => {
    const worksheet = XLSX.utils.json_to_sheet(filteredRequests);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Absences");
    XLSX.writeFile(workbook, `Absence_Report_${startDate}_to_${endDate}.xlsx`);
  };

  const handleUpdate = async (updated: AbsenceRequest) => {
    // await updateDoc(doc(db, "absences", updated.id), { ...updated, updatedAt: Timestamp.now() });
    // setSelectedAbsence(null);
    // fetchDocs();
    setIsSaving(true);
    try {
      // 1. Separate 'id' from the rest of the document data
      const { id, ...dataToSave } = updated;

      // 2. Use the isolated 'id' to point to the correct document reference
      const docRef = doc(db, "absences", id);

      // 3. Save only the pure data back to the database
      await updateDoc(docRef, { 
        ...dataToSave, 
        updatedAt: Timestamp.now() 
      });
      setSelectedAbsence(null);
      await fetchDocs(false); // Refresh table
    } catch (error) {
      console.error("Update failed:", error);
      alert("Error updating record.");
    } finally {
      setIsSaving(false);
    }
  };

  const handleArchive = async (id: string) => {
    const currentAbsenceRequest = absences.find(req => req.id === id);
    const currentStatus = currentAbsenceRequest?.status || 'active';
    const isArchived = currentStatus === 'archived';

    const nextStatus = isArchived ? 'active' : 'archived';

    const confirmArchive = window.confirm('Are you sure? This will ' + (isArchived ? 'restore' : 'archive') + ' this request.');
    if (!confirmArchive) return;

    try {
      const newTime = new Date();
      await updateDoc(doc(db, "absences", id), {
        status: nextStatus,
        updatedAt: newTime,
      });

      setAbsences((prev) =>
        prev.map((item) =>
          item.id === id 
            ? { ...item, status: nextStatus, updatedAt: Timestamp.fromDate(newTime) } 
            : item
        )
      );
    
    // Close the modal
    setSelectedAbsence(null);
    alert("Request has been " + (isArchived ? 'restored' : 'archived') + " successfully!");
    } 
    catch (error) {
      console.error("Critical error during " + (isArchived ? 'restore' : 'archive') + ":", error);
      alert("An error occurred. The status change may not have saved.");
    }
  }

  if (authLoading) return <div className="p-8 text-center font-bold">Verifying Permissions...</div>;

  return (
    <ProtectedRoute allowedRoles={['HR', 'Director']}>
    <div className="p-4 md:p-8 max-w-screen mx-auto space-y-6 bg-white min-h-screen">
      <div className="flex justify-between items-center">
        <div>
          <h1 className="text-3xl font-black text-slate-900 tracking-tight">HR Absence Table</h1>
          {/*<p className="text-slate-400 text-sm font-bold uppercase tracking-widest">Server-Side Filtering Enabled</p>*/}
        </div>
        <div className="flex gap-3">
          <button onClick={downloadExcel} className="bg-white border border-slate-200 text-slate-600 px-4 py-2 rounded-xl font-bold flex items-center gap-2 hover:bg-slate-50 transition-all shadow-sm">
            <FileDown className="h-4 w-4" /> Export
          </button>
          <button onClick={() => setIsAdding(true)} className="bg-indigo-600 hover:bg-indigo-700 text-white px-6 py-2 rounded-xl font-bold transition-all shadow-lg shadow-indigo-100 flex items-center gap-2">
            <Plus className="h-4 w-4" /> New Record
          </button>
        </div>
      </div>

      {/* Control Bar: Search + Offices + Active Status + Dates */}
      <div className="grid grid-cols-2 md:grid-cols-7 gap-4 bg-white p-4 rounded-[2rem] shadow-sm border border-slate-100">
        <div className="col-span-2 md:col-span-2 relative">
          <Search className="absolute left-4 top-3 text-gray-400 h-5 w-5" />
          <input 
            className="w-full pl-12 pr-4 py-2.5 rounded-xl border-none bg-slate-50 focus:ring-2 focus:ring-indigo-500 text-sm font-medium" 
            placeholder="Search filtered results by Name or ID..." 
            value={searchTerm}
            onChange={e => setSearchTerm(e.target.value)} 
          />
        </div>
        <div className="flex col-span-1 items-center gap-2 bg-slate-50 px-3 py-1 rounded-xl border border-slate-100">
          <select
            className="bg-transparent border-none text-s font-bold focus:ring-0 p-0"
            value={activeFilter}
            onChange={e => setActiveFilter(e.target.value)}
          >
            <option value="active_and_pending">All</option>
            <option value="active">Active</option>
            <option value="pending_action">Pending Action</option>
            <option value="archived">Archived</option>
          </select>
        </div>
        <div className="flex col-span-1 items-center gap-2 bg-slate-50 px-3 py-1 rounded-xl border border-slate-100">
          <select 
            className="bg-transparent border-none text-s font-bold focus:ring-0 p-0" 
            value={officeFilter} 
            onChange={e => setOfficeFilter(e.target.value)}
          >
            <option value="all">All Offices</option>
            {OFFICES.map(office => (
              <option key={office} value={office}>
                {office}
              </option>
            ))}
          </select>
        </div>
        <div className="flex col-span-1 items-center gap-2 bg-slate-50 px-3 py-1 rounded-xl border border-slate-100">
          {/*<Calendar className="h-4 w-4 text-slate-400" />*/}
          <div className="flex flex-col w-full">
            <span className="text-[9px] font-black text-slate-400 uppercase">From</span>
            <input type="date" value={startDate} onChange={e => setStartDate(e.target.value)} className="bg-transparent border-none text-xs font-bold focus:ring-0 p-0" />
          </div>
        </div>
        <div className="flex col-span-1 items-center gap-2 bg-slate-50 px-3 py-1 rounded-xl border border-slate-100">
          {/*<Calendar className="h-4 w-4 text-slate-400" />*/}
          <div className="flex flex-col w-full">
            <span className="text-[9px] font-black text-slate-400 uppercase">To</span>
            <input type="date" value={endDate} onChange={e => setEndDate(e.target.value)} className="bg-transparent border-none text-xs font-bold focus:ring-0 p-0" />
          </div>
        </div>
        {/*<label 
          htmlFor="pending-notes-toggle"
          className={`flex col-span-2 md:col-span-1 items-center justify-between gap-3 bg-slate-50 px-4 py-2.5 rounded-xl border border-slate-100 cursor-pointer select-none transition-colors hover:bg-slate-100/80 ${
            pendingNotesOnly ? "bg-amber-50/50 border-amber-100" : ""
          }`}
        >
          <span className="text-xs font-bold text-slate-700">Pending Notes</span>
          <input
            id="pending-notes-toggle"
            type="checkbox"
            checked={pendingNotesOnly}
            onChange={e => setPendingNotesOnly(e.target.checked)}
            className="h-4 w-4 rounded text-indigo-600 border-gray-300 focus:ring-indigo-500 cursor-pointer accent-indigo-600"
          />
        </label>*/}
      </div>

      {/* Results Table */}
      <div className="bg-white rounded-[2rem] shadow-xl overflow-hidden border border-slate-200">
        <div className="overflow-x-auto">
        <table className="w-full text-left">
          <thead className="bg-slate-50 text-[10px] font-black uppercase text-slate-400">
            <tr>
              <th onClick={() => handleSort('employee_name')} className="px-6 py-4 text-left text-xs font-bold text-gray-500 uppercase tracking-wider cursor-pointer">Employee<SortIcon columnKey="employee_name" currentSortKey={sortConfig.key} direction={sortConfig.direction}/></th>
              <th onClick={() => handleSort('type_of_incident')} className="px-6 py-4 text-left text-xs font-bold text-gray-500 uppercase tracking-wider whitespace-nowrap cursor-pointer">Type<SortIcon columnKey="type_of_incident" currentSortKey={sortConfig.key} direction={sortConfig.direction}/></th>
              <th onClick={() => handleSort('office')} className="px-6 py-4 text-left text-xs font-bold text-gray-500 uppercase tracking-wider whitespace-nowrap cursor-pointer">Office<SortIcon columnKey="office" currentSortKey={sortConfig.key} direction={sortConfig.direction}/></th>
              <th onClick={() => handleSort('incident_start')} className="px-6 py-4 text-left text-xs font-bold text-gray-500 uppercase tracking-wider whitespace-nowrap cursor-pointer">Dates<SortIcon columnKey="incident_start" currentSortKey={sortConfig.key} direction={sortConfig.direction}/></th>
              <th onClick={() => handleSort('createdAt')} className="px-6 py-4 text-left text-xs font-bold text-gray-500 uppercase tracking-wider whitespace-nowrap cursor-pointer">Submitted<SortIcon columnKey="createdAt" currentSortKey={sortConfig.key} direction={sortConfig.direction}/></th>
              <th onClick={() => handleSort('statusBadge')} className="px-6 py-4 text-left text-xs font-bold text-gray-500 uppercase tracking-wider whitespace-nowrap cursor-pointer">Status<SortIcon columnKey="statusBadge" currentSortKey={sortConfig.key} direction={sortConfig.direction}/></th>
              <th onClick={() => handleSort('pendingDOAPoints')} className="text-left text-xs font-bold text-gray-500 uppercase tracking-wider whitespace-normal max-w-[80px] w-20 leading-tight select-none cursor-pointer">Pending DOA<SortIcon columnKey="pendingDOAPoints" currentSortKey={sortConfig.key} direction={sortConfig.direction}/></th>
              <th onClick={() => handleSort('DOAPoints')} className="text-left text-xs font-bold text-gray-500 uppercase tracking-wider whitespace-normal max-w-[60px] w-20 leading-tight select-none cursor-pointer">Final DOA<SortIcon columnKey="DOAPoints" currentSortKey={sortConfig.key} direction={sortConfig.direction}/></th>
              <th onClick={() => handleSort('DAP')} className="text-right pr-2 text-xs font-bold text-gray-500 uppercase tracking-wider whitespace-normal max-w-[60px] w-20 leading-tight select-none cursor-pointer">DAP<SortIcon columnKey="DAP" currentSortKey={sortConfig.key} direction={sortConfig.direction}/></th>
              {/*<th className="px-6 py-4 text-right text-xs font-bold text-gray-500 uppercase tracking-wider">Edit</th>*/}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-50">
            {loading ? (
              <tr><td colSpan={4} className="px-8 py-20 text-center font-bold text-slate-400 animate-pulse">Fetching from Server...</td></tr>
            ) : paginated.length > 0 ? (
              paginated.map(a => (
              <tr key={a.id} className="hover:bg-slate-50/50 group transition-colors">
                <td className="px-6 py-4 hover:cursor-pointer" onClick={() => setSelectedAbsence(a)}>
                  <div className="font-bold text-slate-900 group-hover:text-indigo-600 transition-colors">{a.employee_name}</div>
                  <div className="text-[10px] font-mono text-slate-400">ID: {a.employee_id}</div>
                </td>
                <td className="px-4 py-4 text-sm font-medium text-slate-600 whitespace-nowrap">
                  <span className="bg-slate-100 text-slate-500 px-2 py-1 rounded-md">
                    {a.type_of_incident === "Leave and Come Back" ? "Leave & CB" : a.type_of_incident}  
                    {/* a.type_of_incident */}
                  </span>  
                </td>
                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">{a.office}</td>
                <td className="px-6 py-4 whitespace-nowrap">
                  <div className="flex items-center gap-2 text-sm text-gray-600">
                    <Calendar className="h-3.5 w-3.5 text-gray-400" />
                    {(a.incident_start === a.incident_end) ? (<span>{(a.incident_start).replace(/^(\d{4})-(\d{2})-(\d{2})$/, '$2/$3/$1')}</span>) : (<span>{(a.incident_start).replace(/^(\d{4})-(\d{2})-(\d{2})$/, '$2/$3/$1')} - {(a.incident_end).replace(/^(\d{4})-(\d{2})-(\d{2})$/, '$2/$3/$1')}</span>)}
                  </div>
                </td>
                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                    {a.createdAt?.toDate ? a.createdAt.toDate().toLocaleDateString() : 'N/A'}
                </td>
                <td className="px-5 py-4 whitespace-nowrap">
                  <StatusBadge req={a} />
                </td>
                <td className="whitespace-nowrap">
                  <div className="px-1 text-sm font-bold text-slate-900">{a.pendingDOAPoints ? a.pendingDOAPoints : '0'}</div>
                </td>
                <td className="whitespace-nowrap">
                  <div className="px-1 text-sm font-bold text-slate-900">{a.DOAPoints ? a.DOAPoints : '0'}</div>
                </td>
                <td className="pr-4 text-right whitespace-nowrap">
                  <div className="px-1 text-sm font-bold text-slate-900">{a.DAP ? a.DAP : '0'}</div>
                  {/*<button onClick={() => setSelectedAbsence(a)} className="p-2 text-black hover:text-indigo-600 hover:bg-indigo-50 rounded-lg transition-all">
                    <Edit3 className="h-5 w-5"/>
                  </button>*/}
                </td>
              </tr>
              ))) : (
              <tr><td colSpan={8} className="px-8 py-15 text-center font-bold text-black">
                <div className="flex flex-col items-center gap-2">No records found for this date range.</div></td></tr>              
            )}
          </tbody>
        </table>

        {/* Server-Side Meta / Client-Side Pagination */}
        <div className="p-6 bg-slate-50/50 flex justify-between items-center border-t border-slate-100">
          {/*<p className="text-xs font-bold text-slate-400 uppercase tracking-widest">
            Results in Range: <span className="text-slate-900">{filteredRequests.length}</span>
          </p>*/}
            <p className="text-sm text-gray-700 font-bold">
              {filteredRequests.length > 0 ? (
                <>
                Showing <span className="font-bold">{(currentPage - 1) * itemsPerPage + 1}</span> to{' '}
                <span className="font-bold">
                  {Math.min(currentPage * itemsPerPage, filteredRequests.length)}
                </span> of{' '}
                <span className="font-bold">{filteredRequests.length}</span> results
              </>
              ): ("Showing 0 to 0 of 0 results")}
            </p>
          <div className="flex gap-2">
            <button
              onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
              disabled={currentPage === 1}
              className="font-bold px-3 py-1 text-sm bg-white border rounded-md disabled:opacity-50 hover:bg-gray-50"
            >
              Previous
            </button>
            <button
              onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))}
              disabled={currentPage === totalPages}
              className="font-bold px-3 py-1 text-sm bg-white border rounded-md disabled:opacity-50 hover:bg-gray-50"
            >
              Next
            </button>
          </div>
        </div>
      </div>

      {/* Modals */}
      {selectedAbsence && <HRRequestDetailsModal absence={selectedAbsence} userName={user?.username || ""} isSaving={isSaving} onClose={() => setSelectedAbsence(null)} onUpdate={handleUpdate} onArchive={handleArchive} />}
      <HRCreateAbsenceModal isOpen={isAdding} onClose={() => setIsAdding(false)} onSave={fetchDocs} />
      </div>
    </div>
    </ProtectedRoute>
  );
};