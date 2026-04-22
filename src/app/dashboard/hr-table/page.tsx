'use client';
import React, { useState, useEffect, useMemo } from "react";
import { db, storage } from '@/lib/firebase.config';
import { collection, doc, updateDoc, addDoc, Timestamp, deleteDoc } from 'firebase/firestore';
import { ref, deleteObject } from "firebase/storage";
import { useAuth } from '@/contexts/AuthContext';
import { AbsenceRequest } from "@/lib/types";
import { getAbsenceRequestsByUser } from '@/components/approval/approval';

// UI Components
import { Search, ChevronLeft, ChevronRight, Edit3, Calendar, FileDown, Plus } from 'lucide-react';
import { HRRequestDetailsModal } from "@/components/approval/hr-request-details-modal";
import { HRCreateAbsenceModal } from "@/components/forms/absence-request/hr-absence";
import * as XLSX from 'xlsx';

const itemsPerPage = 10;
const incidentTypes = ["Late In", "Early Out", "Absent", "Leave and Come Back", "Long Lunch", "Switch Shift", "Cancel Cell"];

export default function Page() {
  const { user, loading: authLoading } = useAuth();
  
  // --- Date Filter State (Default to last 30 days) ---
  const [startDate, setStartDate] = useState(() => {
    const d = new Date();
    d.setDate(d.getDate() - 30);
    return d.toISOString().split('T')[0];
  });
  const [endDate, setEndDate] = useState(() => new Date().toISOString().split('T')[0]);

  const [absences, setAbsences] = useState<AbsenceRequest[]>([]);
  const [searchTerm, setSearchTerm] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [selectedAbsence, setSelectedAbsence] = useState<AbsenceRequest | null>(null);
  const [isAdding, setIsAdding] = useState(false);
  const [loading, setLoading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);

  // --- Server-Side Fetch Logic ---
  const fetchDocs = async () => {
    if (!user) return;
    setLoading(true);
    try {
      // This calls your provided getAbsenceRequestsByUser function
      // which performs the server-side where() queries
      const data = await getAbsenceRequestsByUser(user, startDate, endDate);
      setAbsences(data);
      setCurrentPage(1); // Reset to page 1 on new fetch
    } catch (err) {
      console.error("Fetch error:", err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (!authLoading) {
      fetchDocs();
    }
  }, [user, authLoading, startDate, endDate]);

  // --- Client-Side Search (Operating on the server-filtered set) ---
  const filteredRequests = useMemo(() => {
    const term = searchTerm.toLowerCase().trim();
    return absences.filter((a) => 
      a.employee_name.toLowerCase().includes(term) || 
      String(a.employee_id || '').toLowerCase().includes(term)
    );
  }, [absences, searchTerm]);

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
      const docRef = doc(db, "absences", updated.id);
      await updateDoc(docRef, { 
        ...updated, 
        updatedAt: Timestamp.now() 
      });
      setSelectedAbsence(null);
      await fetchDocs(); // Refresh table
    } catch (error) {
      console.error("Update failed:", error);
      alert("Error updating record.");
    } finally {
      setIsSaving(false);
    }
  };

  const handleDelete = async (id: string) => {
    const confirmDelete = window.confirm("Are you sure? This will permanently remove the request and all attached files.");
    if (!confirmDelete) return;

    try {
    // 2. Identify the target (the request being deleted)
    // We use selectedAbsence because it contains the list of file URLs
    if (selectedAbsence?.excuse_note && selectedAbsence.excuse_note.length > 0) {
      console.log("Cleaning up storage files...");
      
      // 3. Loop through and delete each physical file from Storage
      const deletePromises = selectedAbsence.excuse_note.map((url) => {
        const fileRef = ref(storage, url);
        return deleteObject(fileRef).catch(err => {
          console.warn(`Could not delete file at ${url}:`, err);
          // We catch inside the map so one missing file doesn't stop the whole process
        });
      });

      await Promise.all(deletePromises);
    }

    // 4. Delete the document from Firestore
    await deleteDoc(doc(db, "absences", id));

    // 5. Update UI State
    // Remove it from the local list so the table updates
    setAbsences((prev) => prev.filter((item) => item.id !== id));
    
    // Close the modal
    setSelectedAbsence(null);

    alert("Request and all associated files deleted.");
  } catch (error) {
    console.error("Critical error during deletion:", error);
    alert("An error occurred. The database record may still exist.");
  }
  }

  if (authLoading) return <div className="p-8 text-center font-bold">Verifying Permissions...</div>;

  return (
    <div className="p-4 md:p-8 max-w-7xl mx-auto space-y-6 bg-white min-h-screen">
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

      {/* Control Bar: Search + Dates */}
      <div className="grid grid-cols-1 md:grid-cols-4 gap-4 bg-white p-4 rounded-[2rem] shadow-sm border border-slate-100">
        <div className="md:col-span-2 relative">
          <Search className="absolute left-4 top-3 text-gray-400 h-5 w-5" />
          <input 
            className="w-full pl-12 pr-4 py-2.5 rounded-xl border-none bg-slate-50 focus:ring-2 focus:ring-indigo-500 text-sm font-medium" 
            placeholder="Search filtered results by Name or ID..." 
            value={searchTerm}
            onChange={e => setSearchTerm(e.target.value)} 
          />
        </div>
        <div className="flex items-center gap-2 bg-slate-50 px-3 py-1 rounded-xl border border-slate-100">
          {/*<Calendar className="h-4 w-4 text-slate-400" />*/}
          <div className="flex flex-col w-full">
            <span className="text-[9px] font-black text-slate-400 uppercase">From</span>
            <input type="date" value={startDate} onChange={e => setStartDate(e.target.value)} className="bg-transparent border-none text-xs font-bold focus:ring-0 p-0" />
          </div>
        </div>
        <div className="flex items-center gap-2 bg-slate-50 px-3 py-1 rounded-xl border border-slate-100">
          {/*<Calendar className="h-4 w-4 text-slate-400" />*/}
          <div className="flex flex-col w-full">
            <span className="text-[9px] font-black text-slate-400 uppercase">To</span>
            <input type="date" value={endDate} onChange={e => setEndDate(e.target.value)} className="bg-transparent border-none text-xs font-bold focus:ring-0 p-0" />
          </div>
        </div>
      </div>

      {/* Results Table */}
      <div className="bg-white rounded-[2rem] shadow-xl overflow-hidden border border-slate-200">
        <div className="overflow-x-auto">
        <table className="w-full text-left">
          <thead className="bg-slate-50 text-[10px] font-black uppercase text-slate-400">
            <tr>
              <th className="px-6 py-4 text-left text-xs font-bold text-gray-500 uppercase tracking-wider">Employee</th>
              <th className="px-6 py-4 text-left text-xs font-bold text-gray-500 uppercase tracking-wider">Type</th>
              <th className="px-6 py-4 text-left text-xs font-bold text-gray-500 uppercase tracking-wider">Dates</th>
              <th className="px-6 py-4 text-left text-xs font-bold text-gray-500 uppercase tracking-wider">Submitted</th>
              <th className="px-6 py-4 text-left text-xs font-bold text-gray-500 uppercase tracking-wider">Status</th>
              <th className="text-left text-xs font-bold text-gray-500 uppercase tracking-wider">DOA</th>
              <th className="px-6 py-4 text-right text-xs font-bold text-gray-500 uppercase tracking-wider">Edit</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-50">
            {loading ? (
              <tr><td colSpan={4} className="px-8 py-20 text-center font-bold text-slate-400 animate-pulse">Fetching from Server...</td></tr>
            ) : paginated.length > 0 ? (
              paginated.map(a => (
              <tr key={a.id} className="hover:bg-slate-50/50 group transition-colors">
                <td className="px-8 py-4">
                  <div className="font-bold text-slate-900 group-hover:text-indigo-600 transition-colors">{a.employee_name}</div>
                  <div className="text-[10px] font-mono text-slate-400">ID: {a.employee_id}</div>
                </td>
                <td className="px-8 py-4 text-sm font-medium text-slate-600">
                  <span className="bg-slate-100 text-slate-500 px-2 py-1 rounded-md">
                    {/*a.type_of_incident === "Leave and Come Back" ? "LACB" : */}  
                    {a.type_of_incident}
                  </span>  
                </td>
                <td className="px-8 py-4 whitespace-nowrap">
                  <div className="flex items-center gap-2 text-sm text-gray-600">
                    <Calendar className="h-3.5 w-3.5 text-gray-400" />
                    {(a.incident_start === a.incident_end) ? (<span>{a.incident_start}</span>) : (<span>{a.incident_start} - {a.incident_end}</span>)}
                  </div>
                </td>
                <td className="px-8 py-4 whitespace-nowrap text-sm text-gray-500">
                    {a.createdAt?.toDate ? a.createdAt.toDate().toLocaleDateString() : 'N/A'}
                </td>
                <td className="px-6 py-4 whitespace-nowrap">
                  <StatusBadge req={a} />
                </td>
                <td className="whitespace-nowrap">
                  <div className="px-2 text-sm font-bold text-slate-900">{a.DOAPoints ? a.DOAPoints : '0'}</div>
                </td>
                <td className="px-6 py-4 text-right">
                  <button onClick={() => setSelectedAbsence(a)} className="p-2 text-black hover:text-indigo-600 hover:bg-indigo-50 rounded-lg transition-all">
                    <Edit3 className="h-5 w-5"/>
                  </button>
                </td>
              </tr>
              ))) : (
              <tr><td colSpan={6} className="px-8 py-20 text-center font-bold text-black">
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
      {selectedAbsence && <HRRequestDetailsModal absence={selectedAbsence} userName={user?.username || ""} isSaving={isSaving} onClose={() => setSelectedAbsence(null)} onUpdate={handleUpdate} onDelete={handleDelete} />}
      <HRCreateAbsenceModal isOpen={isAdding} onClose={() => setIsAdding(false)} onSave={fetchDocs} />
      </div>
    </div>
  );
};

const StatusBadge = ({ req }: { req: AbsenceRequest }) => {
  if (req.manager_approval === 'denied' && req.final_approval === 'pending') {
    return <Badge color="red" text="Manager Denied" />;
  }
  else if (req.manager_approval === 'denied' || req.final_approval === 'denied') {
    return <Badge color="red" text="Denied" />;
  }
  else if (req.final_approval === 'approved') {
    return <Badge color="green" text="Fully Approved" />;
  }
  else if (req.manager_approval === 'approved') {
    return <Badge color="yellow" text="Manager Approved" />;
  }
  else if (req.manager_approval === 'not_required' && req.final_approval === 'pending') {
    return <Badge color="yellow" text="Pending Final Approval" />;
  }
  else{
    return <Badge color="gray" text="Pending" />;
  }
};

const Badge = ({ color, text }: { color: string; text: string }) => {
  const styles: Record<string, string> = {
    green: 'bg-emerald-100 text-emerald-700 ring-emerald-600/20',
    red: 'bg-rose-100 text-rose-700 ring-rose-600/20',
    yellow: 'bg-amber-100 text-amber-700 ring-amber-600/20',
    gray: 'bg-slate-100 text-slate-600 ring-slate-600/10',
  };
  return (
    <span className={`inline-flex items-center rounded-full px-2 py-1 text-xs font-black text-[10px] uppercase ring-1 ring-inset ${styles[color]}`}>
      {text}
    </span>
  );
};