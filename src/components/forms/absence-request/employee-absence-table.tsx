'use client';

import React, { useState, useMemo, useEffect } from 'react';
import { AbsenceRequest } from '@/lib/types';
import { ChevronRight, Calendar, FileText, ArrowUp, ArrowDown, ArrowUpDown, Filter } from 'lucide-react';

type RequestTableProps = {
  requests: AbsenceRequest[];
  status: 'needed' | 'pending' | 'approved' | 'denied' | 'all' | 'archived';
  onViewDetails: (request: AbsenceRequest) => void;
};

export const AbsenceTable: React.FC<RequestTableProps> = ({ requests, status, onViewDetails }) => {
  // --- Date Filter State (Default to -30 to +30 days) ---
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

  const [typeFilter, setTypeFilter] = useState('all');
  const [sortOrder, setSortOrder] = useState<'asc' | 'desc' | null>(null);
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 5;

  // Unique types for the dropdown
  const uniqueTypes = useMemo(() => {
    const types = new Set(requests.map(r => r.type_of_incident));
    return Array.from(types).sort();
  }, [requests]);

  // Reset pagination on filter change
  useEffect(() => {
    setCurrentPage(1);
  }, [typeFilter, sortOrder, startDate, endDate]);

  const processedRequests = useMemo(() => {
    let result = requests.filter((req) => {
      // 1. Type Filter
      const matchesType = typeFilter === 'all' || req.type_of_incident === typeFilter;
      
      // 2. Date Range Filter (Check if incident_start falls within range)
      const reqDate = req.incident_start; // Assuming YYYY-MM-DD string format
      const isWithinDate = reqDate >= startDate && reqDate <= endDate;

      return matchesType && isWithinDate;
    });

    // 3. Sort Logic
    result.sort((a, b) => {
      const dateA = new Date(a.incident_start).getTime();
      const dateB = new Date(b.incident_start).getTime();
      return sortOrder === 'desc' ? dateB - dateA : dateA - dateB;
    });

    return result;
  }, [requests, typeFilter, sortOrder, startDate, endDate]);

  const handleSort = () => {
    if (sortOrder === 'asc') {
      setSortOrder('desc');
    } else if (sortOrder === 'desc') {
      // Cycle back to neutral no-sort state
      setSortOrder(null);
    } else {
      setSortOrder('asc');
    }
    setCurrentPage(1); // Go to page 1 on layout mutations
  };

  const paginatedRequests = useMemo(() => {
    const startIndex = (currentPage - 1) * itemsPerPage;
    return processedRequests.slice(startIndex, startIndex + itemsPerPage);
  }, [processedRequests, currentPage]);

  const totalPages = Math.ceil(processedRequests.length / itemsPerPage) || 1;

  return (
    <div className="space-y-4">
      {/* Control Bar: Date Pickers + Incident Filter */}
      <div className="grid grid-cols-1 lg:grid-cols-4 gap-3 bg-white p-4 rounded-2xl border border-gray-100 shadow-sm">
        
        {/* Date From */}
        <div className="flex flex-col bg-slate-50 px-3 py-1.5 rounded-xl border border-slate-100">
          <span className="text-[9px] font-black text-slate-400 uppercase">From</span>
          <input 
            type="date" 
            value={startDate} 
            onChange={e => setStartDate(e.target.value)} 
            className="bg-transparent border-none text-xs font-bold focus:ring-0 p-0 text-slate-900" 
          />
        </div>

        {/* Date To */}
        <div className="flex flex-col bg-slate-50 px-3 py-1.5 rounded-xl border border-slate-100">
          <span className="text-[9px] font-black text-slate-400 uppercase">To</span>
          <input 
            type="date" 
            value={endDate} 
            onChange={e => setEndDate(e.target.value)} 
            className="bg-transparent border-none text-xs font-bold focus:ring-0 p-0 text-slate-900" 
          />
        </div>

        {/* Incident Type Filter */}
        <div className="flex flex-col bg-slate-50 px-3 py-1.5 rounded-xl border border-slate-100">
          <span className="text-[9px] font-black text-slate-400 uppercase">Incident Type</span>
          <select
            className="bg-transparent border-none focus:ring-0 p-0 text-xs font-bold text-slate-900 cursor-pointer"
            value={typeFilter}
            onChange={(e) => setTypeFilter(e.target.value)}
          >
            <option value="all">All Categories</option>
            {uniqueTypes.map(t => (
              <option key={t} value={t}>{t}</option>
            ))}
          </select>
        </div>

        {/* Pending Note Indicator */}
        {status === 'needed' && (
          <div className="flex items-center justify-center gap-2 bg-amber-50 rounded-xl border border-amber-100">
            <FileText className="h-4 w-4 text-amber-600" />
            <span className="text-[10px] font-black text-amber-800 uppercase">Note Needed</span>
          </div>
        )}
      </div>

      {/* Main Table */}
      <div className="bg-white rounded-[2rem] shadow-xl overflow-hidden border border-slate-200">
        <div className="overflow-x-auto">
          <table className="min-w-full divide-y divide-slate-100">
            <thead className="bg-slate-50 text-[10px] font-black uppercase text-slate-400">
              <tr>
                <th className="px-6 py-4 text-left tracking-wider">
                  <button 
                    onClick={() => handleSort()}
                    className="flex items-center gap-1 hover:text-slate-900 transition-colors"
                  >
                    Incident Date
                    {sortOrder === 'asc' && <ArrowUp className="h-3 w-3 text-indigo-600" />}
                    {sortOrder === 'desc' && <ArrowDown className="h-3 w-3 text-indigo-600" />}
                    {sortOrder === null && <ArrowUpDown className="h-3 w-3 opacity-40" />}
                  </button>
                </th>
                <th className="px-6 py-4 text-left tracking-wider">Type</th>
                <th className="px-6 py-4 text-left tracking-wider">Excuse Note</th>
                <th className="px-6 py-4 text-left tracking-wider">Status</th>
                <th className="px-6 py-4 text-right tracking-wider">Action</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-50 bg-white">
              {paginatedRequests.length > 0 ? (
                paginatedRequests.map((req) => (
                  <tr key={req.id} className="group hover:bg-slate-50/50 transition-colors">
                    <td className="px-6 py-4 whitespace-nowrap">
                      <div className="flex items-center gap-2 text-sm font-bold text-slate-900">
                        <Calendar className="h-3.5 w-3.5 text-slate-400" />
                        {req.incident_start === req.incident_end 
                          ? (req.incident_start.replace(/^(\d{4})-(\d{2})-(\d{2})$/, '$2/$3/$1')) 
                          : `${req.incident_start.replace(/^(\d{4})-(\d{2})-(\d{2})$/, '$2/$3/$1')} - ${req.incident_end.replace(/^(\d{4})-(\d{2})-(\d{2})$/, '$2/$3/$1')}`}
                      </div>
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap">
                      <span className="text-[11px] font-bold text-slate-600 bg-slate-100 px-2 py-1 rounded-lg uppercase">
                        {req.type_of_incident}
                      </span>
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap">
                      <NoteBadge status={req.excuse_note_submitted} />
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap">
                      <StatusBadge req={req} />
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap text-right">
                      <button
                        onClick={() => onViewDetails(req)}
                        className="inline-flex items-center gap-1 text-[11px] font-black text-indigo-600 hover:text-indigo-800 uppercase tracking-tighter"
                      >
                        Details
                        <ChevronRight className="h-4 w-4 transition-transform group-hover:translate-x-0.5" />
                      </button>
                    </td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={5} className="px-6 py-16 text-center text-slate-400 font-bold uppercase text-xs italic">
                    No records found in this date range
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>

        {/* Pagination */}
        <div className="px-8 py-5 bg-slate-50/50 border-t border-slate-100 flex flex-col sm:flex-row items-center justify-between gap-4">
          <p className="text-xs font-bold text-slate-500 uppercase tracking-widest">
                            Showing <span className="font-medium">{(currentPage - 1) * itemsPerPage + 1}</span> to{' '}
                <span className="font-medium">
                  {Math.min(currentPage * itemsPerPage, processedRequests.length)}
                </span> of{' '}
                <span className="font-medium">{processedRequests.length}</span> results
          </p>
          <div className="flex items-center gap-2">
            <button
              onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
              disabled={currentPage === 1}
              className="px-4 py-2 text-[11px] font-black uppercase tracking-tighter bg-white border border-slate-200 rounded-xl shadow-sm disabled:opacity-50 hover:bg-slate-50 transition-all"
            >
              Prev
            </button>
            <button
              onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))}
              disabled={currentPage === totalPages}
              className="px-4 py-2 text-[11px] font-black uppercase tracking-tighter bg-white border border-slate-200 rounded-xl shadow-sm disabled:opacity-50 hover:bg-slate-50 transition-all"
            >
              Next
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

// ... (NoteBadge, StatusBadge, and Badge components remain the same as previously defined)

// Sub-component for the Excuse Note Status
const NoteBadge = ({ status }: { status: string }) => {
  const configs: Record<string, { bg: string, text: string, label: string }> = {
    pending: { bg: 'bg-yellow-50', text: 'text-yellow-700', label: 'Pending' },
    not_provided: { bg: 'bg-rose-50', text: 'text-rose-700', label: 'No Note' },
    submitted: { bg: 'bg-emerald-50', text: 'text-emerald-700', label: 'Submitted' },
  };

  const config = configs[status] || configs.not_provided;

  return (
    <span className={`px-2.5 py-1 rounded-lg text-[10px] font-black uppercase tracking-wide border shadow-sm ${config.bg} ${config.text} border-current/10`}>
      {config.label}
    </span>
  );
};

// Logic-heavy Status Badge (Manager + Final)
const StatusBadge = ({ req }: { req: AbsenceRequest }) => {
  if (req.manager_approval === 'denied' || req.final_approval === 'denied') {
    return <Badge color="red" text={req.manager_approval === 'denied' && req.final_approval === 'pending' ? "Manager Denied" : "Denied"} />;
  }
  if (req.final_approval === 'approved') return <Badge color="green" text="Approved" />;
  if (req.manager_approval === 'approved') return <Badge color="yellow" text="Manager Approved" />;
  if (req.manager_approval === 'not_required' && req.final_approval === 'pending') return <Badge color="yellow" text="Pending Corp" />;
  if (req.manager_approval === 'pending' && req.final_approval === 'pending') return <Badge color="gray" text="Pending Manager Approval" />;
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