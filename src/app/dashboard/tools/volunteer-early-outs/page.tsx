"use client";

import React, { useState, useEffect, useMemo } from "react";
import { db } from "@/lib/firebase.config";
import { collection, query, where, getDocs, orderBy, Timestamp } from "firebase/firestore";
import { useAuth } from "@/contexts/AuthContext";
import { Search, Calendar, FileDown, Clock, Building } from "lucide-react";
import * as XLSX from "xlsx";
import { VolunteerEarlyOutRequest } from "@/lib/types";
import ProtectedRoute from '@/components/auth/ProtectedRoute';

import { OFFICES } from "@/lib/constants";

import { useSort } from '@/hooks/use-sort'
import { SortIcon } from '@/components/ui/table-sort'

const itemsPerPage = 50;

function formatTo12Hour(time24: string): string {
  if (!time24) return '';
  const [hours, minutes] = time24.split(':').map(Number);
  const period = hours >= 12 ? 'PM' : 'AM';
  const hours12 = hours % 12 || 12; // Converts 0/12/24 appropriately
  return `${hours12}:${String(minutes).padStart(2, '0')} ${period}`;
}

export default function Page() {
  const { user, loading: authLoading } = useAuth();
  const { sortConfig, handleSort } = useSort();

  // --- Date Filters (Default: Last 30 days to Next 30 days) ---
  const [startDate, setStartDate] = useState(() => {
    const d = new Date();
    d.setDate(d.getDate()-1);
    return d.toLocaleDateString('en-CA');
  });
  const [endDate, setEndDate] = useState(() => {
    const d = new Date();
    d.setDate(d.getDate()+1);
    return d.toLocaleDateString('en-CA');
  });

  const [submissions, setSubmissions] = useState<VolunteerEarlyOutRequest[]>([]);
  const [searchTerm, setSearchTerm] = useState("");
  const [officeFilter, setOfficeFilter] = useState("all");
  const [currentPage, setCurrentPage] = useState(1);
  const [loading, setLoading] = useState(false);

  // --- Firestore Data Fetching ---
  const fetchEarlyOuts = async (resetPages = false) => {
    if (!user) return;
    setLoading(true);
    try {
      const earlyOutsRef = collection(db, "volunteer-early-outs");
      
      // Perform date range queries using sequential string sorting on 'incident_date'
      const q = query(
        earlyOutsRef,
        where("incident_date", ">=", startDate),
        where("incident_date", "<=", endDate)
      );

      const querySnapshot = await getDocs(q);
      const fetchedData = querySnapshot.docs.map((doc) => ({
        id: doc.id,
        ...doc.data(),
      })) as VolunteerEarlyOutRequest[];

      // Sort programmatically by incident date descending (newest first)
      fetchedData.sort((a, b) => b.incident_date.localeCompare(a.incident_date));

      setSubmissions(fetchedData);
      if (resetPages) setCurrentPage(1);
    } catch (err) {
      console.error("Error fetching volunteer early outs:", err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (!authLoading) {
      fetchEarlyOuts(true);
    }
  }, [user, authLoading, startDate, endDate]);

  // --- Client-Side Search and Office Filters ---
  const filteredRequests = useMemo(() => {
    const term = searchTerm.toLowerCase().trim();
    let result = submissions.filter((r) => {
      const officeMatches = officeFilter === "all" || r.office === officeFilter;
      const searchMatches =
        r.employee_name.toLowerCase().includes(term) ||
        String(r.employee_id || "").toLowerCase().includes(term);

      return officeMatches && searchMatches;
    });

    if (sortConfig.direction !== 'none' && sortConfig.key) {
      result = [...result].sort((a, b) => {
        let valA = a[sortConfig.key as keyof VolunteerEarlyOutRequest];
        let valB = b[sortConfig.key as keyof VolunteerEarlyOutRequest];

        // Handle fallback conversions for missing types or empty values safely
        const strA = String(valA ?? '').toLowerCase();
        const strB = String(valB ?? '').toLowerCase();

        if (strA < strB) return sortConfig.direction === 'asc' ? -1 : 1;
        if (strA > strB) return sortConfig.direction === 'asc' ? 1 : -1;
        return 0;
      });
    }

    return result;
  }, [submissions, searchTerm, officeFilter, sortConfig]);

  // --- Pagination Logic ---
  const paginated = useMemo(() => {
    const start = (currentPage - 1) * itemsPerPage;
    return filteredRequests.slice(start, start + itemsPerPage);
  }, [filteredRequests, currentPage]);

  const totalPages = Math.ceil(filteredRequests.length / itemsPerPage) || 1;

  if (authLoading) return <div className="p-8 text-center font-bold">Verifying Permissions...</div>;

  return (
    <ProtectedRoute allowedRoles={['HR', 'Director']}>
    <div className="w-full p-2 sm:p-3 lg:p-4 space-y-6 min-h-screen flex flex-col">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h1 className="text-3xl font-black text-gray-900 tracking-tight">Volunteer Early Outs</h1>
        </div>
      </div>

      {/* Filter and Control Bars */}
      <div className="grid grid-cols-1 md:grid-cols-6 gap-4 bg-white p-4 rounded-[2rem] shadow-sm border border-slate-100">
        <div className="col-span-3 relative">
          <Search className="absolute left-4 top-3 text-gray-400 h-5 w-5" />
          <input
            className="w-full pl-12 pr-4 py-2.5 rounded-xl border-none bg-slate-50 focus:ring-2 focus:ring-indigo-500 text-sm font-medium"
            placeholder="Search by Name or ID..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </div>

        <div className="flex col-span-1 items-center gap-2 bg-slate-50 px-3 py-1 rounded-xl border border-slate-100">
          <select
            className="bg-transparent border-none text-sm font-bold focus:ring-0 p-0 w-full"
            value={officeFilter}
            onChange={(e) => setOfficeFilter(e.target.value)}
          >
            <option value="all">All Offices</option>
            {OFFICES.map((office) => (
              <option key={office} value={office}>
                {office}
              </option>
            ))}
          </select>
        </div>

        <div className="flex col-span-1 items-center gap-2 bg-slate-50 px-3 py-1 rounded-xl border border-slate-100">
          <div className="flex flex-col w-full">
            <span className="text-[9px] font-black text-slate-400 uppercase">From</span>
            <input
              type="date"
              value={startDate}
              onChange={(e) => setStartDate(e.target.value)}
              className="bg-transparent border-none text-xs font-bold focus:ring-0 p-0"
            />
          </div>
        </div>

        <div className="flex col-span-1 items-center gap-2 bg-slate-50 px-3 py-1 rounded-xl border border-slate-100">
          <div className="flex flex-col w-full">
            <span className="text-[9px] font-black text-slate-400 uppercase">To</span>
            <input
              type="date"
              value={endDate}
              onChange={(e) => setEndDate(e.target.value)}
              className="bg-transparent border-none text-xs font-bold focus:ring-0 p-0"
            />
          </div>
        </div>
      </div>

      {/* Main Records Table */}
      <div className="bg-white rounded-[2rem] shadow-xl overflow-hidden border border-slate-200">
        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead className="bg-slate-50 text-[10px] font-black uppercase text-slate-400">
              <tr>
                <th onClick={() => handleSort('employee_name')} className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Employee<SortIcon columnKey="employee_name" currentSortKey={sortConfig.key} direction={sortConfig.direction}/></th>
                <th onClick={() => handleSort('office')} className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Office<SortIcon columnKey="office" currentSortKey={sortConfig.key} direction={sortConfig.direction}/></th>
                <th onClick={() => handleSort('incident_date')} className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Early Out Window<SortIcon columnKey="incident_date" currentSortKey={sortConfig.key} direction={sortConfig.direction}/></th>
                <th onClick={() => handleSort('supervisor_name')} className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Supervisor<SortIcon columnKey="supervisor_name" currentSortKey={sortConfig.key} direction={sortConfig.direction}/></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-50">
              {loading ? (
                <tr>
                  <td colSpan={5} className="px-8 py-20 text-center font-bold text-slate-400 animate-pulse">
                    Loading Volunteer Records...
                  </td>
                </tr>
              ) : paginated.length > 0 ? (
                paginated.map((r) => (
                  <tr key={r.id} className="hover:bg-slate-50/50 group transition-colors">
                    <td className="px-6 py-4">
                      <div className="font-bold text-slate-900">{r.employee_name}</div>
                      <div className="text-[10px] font-mono text-slate-400">ID: {r.employee_id}</div>
                    </td>
                    <td className="px-6 py-4">
                      <div className="text-sm font-semibold text-slate-700 flex items-center gap-1.5">
                        {r.office}
                      </div>
                      <div className="text-xs text-slate-400">{r.employee_title}</div>
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap">
                      <div className="flex items-center gap-2 text-sm font-medium text-slate-900">
                        <Calendar className="h-4 w-4 text-indigo-500" />
                        {r.incident_date.replace(/^(\d{4})-(\d{2})-(\d{2})$/, "$2/$3/$1")}
                      </div>
                      <div className="text-xs text-slate-500 flex items-center gap-1 mt-0.5">
                        <Clock className="h-3 w-3" /> {formatTo12Hour(r.incident_time)}
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <div className="text-sm font-semibold text-slate-700">{r.supervisor_name}</div>
                      <div className="text-[10px] font-mono text-slate-400">ID: {r.supervisor_id}</div>
                    </td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={5} className="px-8 py-15 text-center font-bold text-slate-500">
                    No early out requests logged within this date range.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>

        {/* Pagination Footer */}
        <div className="p-6 bg-slate-50/50 flex justify-between items-center border-t border-slate-100">
          <p className="text-sm text-gray-700 font-bold">
            {filteredRequests.length > 0 ? (
              <>
                Showing <span className="font-bold">{(currentPage - 1) * itemsPerPage + 1}</span> to{" "}
                <span className="font-bold">
                  {Math.min(currentPage * itemsPerPage, filteredRequests.length)}
                </span>{" "}
                of <span className="font-bold">{filteredRequests.length}</span> results
              </>
            ) : (
              "Showing 0 to 0 of 0 results"
            )}
          </p>
          <div className="flex gap-2">
            <button
              onClick={() => setCurrentPage((prev) => Math.max(prev - 1, 1))}
              disabled={currentPage === 1}
              className="font-bold px-3 py-1 text-sm bg-white border rounded-md disabled:opacity-50 hover:bg-gray-50 cursor-pointer"
            >
              Previous
            </button>
            <button
              onClick={() => setCurrentPage((prev) => Math.min(prev + 1, totalPages))}
              disabled={currentPage === totalPages}
              className="font-bold px-3 py-1 text-sm bg-white border rounded-md disabled:opacity-50 hover:bg-gray-50 cursor-pointer"
            >
              Next
            </button>
          </div>
        </div>
      </div>
    </div>
    </ProtectedRoute>
  );
}