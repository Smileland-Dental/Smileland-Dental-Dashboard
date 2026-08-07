'use client';

import React, { useState, useMemo, useEffect } from 'react';
import { AbsenceRequest } from '@/lib/types';
import { Search, Filter, ChevronRight, Calendar, User as UserIcon } from 'lucide-react';
import { useSort } from '@/hooks/use-sort';
import { SortIcon } from '@/components/ui/table-sort';
import { getRequestStatusText } from './approval';
import { StatusBadge } from '@/components/ui/status-badge'

type RequestTableProps = {
  requests: AbsenceRequest[];
  userRole: string;
  onViewDetails: (request: AbsenceRequest) => void;
};

export const RequestTable: React.FC<RequestTableProps> = ({ requests, userRole, onViewDetails }) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState('all');

  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 50; // Adjust this number as needed

  // Initialize your shared custom hook
  const { sortConfig, handleSort } = useSort();

  // Reset to page 1 whenever filters or search change
  useEffect(() => {
    setCurrentPage(1);
  }, [searchTerm, statusFilter, sortConfig]);

  const filteredRequests = useMemo(() => {
    let result = requests.filter((req) => {
      const searchLower = searchTerm.toLowerCase();
      const matchesSearch = 
        req.employee_name.toLowerCase().includes(searchLower) ||
        req.employee_id.includes(searchTerm);
      
      const status = 
        req.manager_approval === 'denied' || req.final_approval === 'denied' ? 'denied' :
        ((req.final_approval === 'approved') || (req.final_approval === 'approved_with_note')) ? 'approved' : 'pending';

      const matchesStatus = statusFilter === 'all' || status === statusFilter;

      return matchesSearch && matchesStatus;
    });

    if (sortConfig.direction !== 'none' && sortConfig.key) {
      result = [...result].sort((a, b) => {
        let valA = a[sortConfig.key as keyof AbsenceRequest];
        let valB = b[sortConfig.key as keyof AbsenceRequest];

        // Format dates safely
        if (sortConfig.key === 'createdAt') {
          valA = a.createdAt?.toDate ? a.createdAt.toDate().getTime() : 0;
          valB = b.createdAt?.toDate ? b.createdAt.toDate().getTime() : 0;
        }

        // Process status computed values for sorting
        if (sortConfig.key === 'statusBadge') {
          valA = getRequestStatusText(a);
          valB = getRequestStatusText(b);
        }

        const strA = String(valA ?? '').toLowerCase();
        const strB = String(valB ?? '').toLowerCase();

        if (strA < strB) return sortConfig.direction === 'asc' ? -1 : 1;
        if (strA > strB) return sortConfig.direction === 'asc' ? 1 : -1;
        return 0;
      });
    }
    return result;
  }, [requests, searchTerm, statusFilter, sortConfig]);

  const paginatedRequests = useMemo(() => {
    const startIndex = (currentPage - 1) * itemsPerPage;
    return filteredRequests.slice(startIndex, startIndex + itemsPerPage);
  }, [filteredRequests, currentPage]);

  const totalPages = Math.ceil(filteredRequests.length / itemsPerPage);

  return (
    <div className="space-y-4">
      {/* Search and Filter Control Bar */}
      <div className="flex flex-col md:flex-row gap-3 items-center justify-between bg-white p-4 rounded-xl border border-gray-200 shadow-sm">
        <div className="relative w-full md:w-1/3">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-gray-400" />
          <input
            type="text"
            placeholder="Search by name or ID..."
            className="w-full pl-10 pr-4 py-2 text-sm border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 transition-all"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </div>

        <div className="flex items-center gap-3 w-full md:w-auto">
          <div className="flex items-center gap-2 text-sm font-medium text-gray-600 bg-gray-50 px-3 py-2 rounded-lg border border-gray-200">
            <Filter className="h-4 w-4" />
            <span>Filter:</span>
            <select
              className="bg-transparent border-none focus:ring-0 p-0 text-sm font-semibold cursor-pointer"
              value={statusFilter}
              onChange={(e) => setStatusFilter(e.target.value)}
            >
              <option value="all">All Statuses</option>
              <option value="pending">Pending</option>
              <option value="approved">Approved</option>
              <option value="denied">Denied</option>
            </select>
          </div>
        </div>
      </div>

      {/* Table Section */}
      <div className="bg-white rounded-[2rem] shadow-xl overflow-hidden border border-slate-200">
        <div className="overflow-x-auto">
          <table className="min-w-full divide-y divide-gray-200">
            <thead className="bg-slate-50 text-[10px] font-black uppercase text-slate-400">
              <tr>
                <th onClick={() => handleSort('employee_name')} className="px-6 py-4 text-left text-xs font-bold text-gray-500 uppercase tracking-wider whitespace-nowrap">Employee<SortIcon columnKey="employee_name" currentSortKey={sortConfig.key} direction={sortConfig.direction}/></th>
                <th onClick={() => handleSort('type_of_incident')} className="px-6 py-4 text-left text-xs font-bold text-gray-500 uppercase tracking-wider whitespace-nowrap">Type<SortIcon columnKey="type_of_incident" currentSortKey={sortConfig.key} direction={sortConfig.direction}/></th>
                <th onClick={() => handleSort('office')} className="px-6 py-4 text-left text-xs font-bold text-gray-500 uppercase tracking-wider whitespace-nowrap">Office<SortIcon columnKey="office" currentSortKey={sortConfig.key} direction={sortConfig.direction}/></th>
                <th onClick={() => handleSort('incident_start')} className="px-6 py-4 text-left text-xs font-bold text-gray-500 uppercase tracking-wider whitespace-nowrap">Incident Dates<SortIcon columnKey="incident_start" currentSortKey={sortConfig.key} direction={sortConfig.direction}/></th>
                <th onClick={() => handleSort('createdAt')} className="px-6 py-4 text-left text-xs font-bold text-gray-500 uppercase tracking-wider whitespace-nowrap">Submitted<SortIcon columnKey="createdAt" currentSortKey={sortConfig.key} direction={sortConfig.direction}/></th>
                <th onClick={() => handleSort('statusBadge')} className="px-6 py-4 text-left text-xs font-bold text-gray-500 uppercase tracking-wider whitespace-nowrap">Status<SortIcon columnKey="statusBadge" currentSortKey={sortConfig.key} direction={sortConfig.direction}/></th>
                <th className="px-6 py-4 text-right text-xs font-bold text-gray-500 uppercase tracking-wider">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-50">
              {paginatedRequests.length > 0 ? (
                paginatedRequests.map((req) => (
                  <tr key={req.id} className="group hover:bg-gray-50/80 transition-all cursor-default">
                    <td className="px-6 py-4 whitespace-nowrap">
                      <div className="flex items-center gap-3">
                        <div className="flex flex-col">
                          <span className="text-sm font-semibold text-gray-900">{req.employee_name}</span>
                          <span className="text-xs text-gray-500 font-mono tracking-tighter">#{req.employee_id}</span>
                        </div>
                      </div>
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap">
                      <span className="text-sm text-gray-600 bg-gray-100 px-2 py-1 rounded-md">
                        {req.type_of_incident}
                      </span>
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">{req.office}</td>
                    <td className="px-6 py-4 whitespace-nowrap">
                      <div className="flex items-center gap-2 text-sm text-gray-600">
                        <Calendar className="h-3.5 w-3.5 text-gray-400" />
                        {(req.incident_start === req.incident_end) ? (<span>{(req.incident_start).replace(/^(\d{4})-(\d{2})-(\d{2})$/, '$2/$3/$1')}</span>) : (<span>{(req.incident_start).replace(/^(\d{4})-(\d{2})-(\d{2})$/, '$2/$3/$1')} - {(req.incident_end).replace(/^(\d{4})-(\d{2})-(\d{2})$/, '$2/$3/$1')}</span>)}
                      </div>
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                      {req.createdAt?.toDate ? req.createdAt.toDate().toLocaleDateString() : 'N/A'}
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap">
                      <StatusBadge req={req} />
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap text-right">
                      <button
                        onClick={() => onViewDetails(req)}
                        className="inline-flex items-center gap-1 text-sm font-semibold text-indigo-600 hover:text-indigo-700 transition-colors"
                      >
                        Details
                        <ChevronRight className="h-4 w-4 transition-transform group-hover:translate-x-0.5" />
                      </button>
                    </td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={6} className="px-6 py-20 text-center">
                    <div className="flex flex-col items-center gap-2">
                      <Search className="h-8 w-8 text-gray-300" />
                      <p className="text-gray-500 font-medium">No requests found matching your filters.</p>
                    </div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
          <div className="px-6 py-4 bg-gray-50 border-t border-gray-200 flex items-center justify-between">
            <p className="text-sm text-gray-700">
              {filteredRequests.length > 0 ? (
                <>
                Showing <span className="font-medium">{(currentPage - 1) * itemsPerPage + 1}</span> to{' '}
                <span className="font-medium">
                  {Math.min(currentPage * itemsPerPage, filteredRequests.length)}
                </span> of{' '}
                <span className="font-medium">{filteredRequests.length}</span> results
              </>
              ): ("Showing 0 to 0 of 0 results")}
            </p>
            <div className="flex gap-2">
              <button
                onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
                disabled={currentPage === 1}
                className="px-3 py-1 text-sm bg-white border rounded-md disabled:opacity-50 hover:bg-gray-50"
              >
                Previous
              </button>
              <button
                onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))}
                disabled={currentPage === totalPages}
                className="px-3 py-1 text-sm bg-white border rounded-md disabled:opacity-50 hover:bg-gray-50"
              >
                Next
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
