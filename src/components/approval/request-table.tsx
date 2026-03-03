'use client';

import React, { useState, useMemo } from 'react';
import { AbsenceRequest } from '@/lib/types';
import { Search, Filter, ChevronRight, Calendar, User as UserIcon } from 'lucide-react';

type RequestTableProps = {
  requests: AbsenceRequest[];
  userRole: string;
  onViewDetails: (request: AbsenceRequest) => void;
};

export const RequestTable: React.FC<RequestTableProps> = ({ requests, userRole, onViewDetails }) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState('all');

  const filteredRequests = useMemo(() => {
    return requests.filter((req) => {
      const searchLower = searchTerm.toLowerCase();
      const matchesSearch = 
        req.employee_name.toLowerCase().includes(searchLower) ||
        req.employee_id.includes(searchTerm);
      
      const status = 
        req.manager_approval === 'denied' || req.final_approval === 'denied' ? 'denied' :
        req.final_approval === 'approved' ? 'approved' : 'pending';

      const matchesStatus = statusFilter === 'all' || status === statusFilter;

      return matchesSearch && matchesStatus;
    });
  }, [requests, searchTerm, statusFilter]);

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
      <div className="bg-white rounded-xl border border-gray-200 shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="min-w-full divide-y divide-gray-200">
            <thead className="bg-gray-50/50">
              <tr>
                <th className="px-6 py-4 text-left text-xs font-bold text-gray-500 uppercase tracking-wider">Employee</th>
                <th className="px-6 py-4 text-left text-xs font-bold text-gray-500 uppercase tracking-wider">Type</th>
                <th className="px-6 py-4 text-left text-xs font-bold text-gray-500 uppercase tracking-wider">Incident Dates</th>
                <th className="px-6 py-4 text-left text-xs font-bold text-gray-500 uppercase tracking-wider">Submitted</th>
                <th className="px-6 py-4 text-left text-xs font-bold text-gray-500 uppercase tracking-wider">Status</th>
                <th className="px-6 py-4 text-right text-xs font-bold text-gray-500 uppercase tracking-wider">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filteredRequests.length > 0 ? (
                filteredRequests.map((req) => (
                  <tr key={req.id} className="group hover:bg-gray-50/80 transition-all cursor-default">
                    <td className="px-6 py-4 whitespace-nowrap">
                      <div className="flex items-center gap-3">
                        <div className="h-8 w-8 rounded-full bg-indigo-50 flex items-center justify-center text-indigo-600">
                          <UserIcon className="h-4 w-4" />
                        </div>
                        <div className="flex flex-col">
                          <span className="text-sm font-semibold text-gray-900">{req.employee_name}</span>
                          <span className="text-xs text-gray-500 font-mono tracking-tighter">#{req.employee_id}</span>
                        </div>
                      </div>
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap">
                      <span className="text-sm text-gray-600 bg-gray-100 px-2 py-1 rounded-md">
                        {req.type_of_request}
                      </span>
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap">
                      <div className="flex items-center gap-2 text-sm text-gray-600">
                        <Calendar className="h-3.5 w-3.5 text-gray-400" />
                        <span>{req.incident_start}</span>
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
        </div>
      </div>
    </div>
  );
};

// Refined StatusBadge Component
const StatusBadge = ({ req }: { req: AbsenceRequest }) => {
  if (req.manager_approval === 'denied' || req.final_approval === 'denied') {
    return <Badge color="red" text="Denied" />;
  }
  if (req.final_approval === 'approved') {
    return <Badge color="green" text="Fully Approved" />;
  }
  if (req.manager_approval === 'approved') {
    return <Badge color="yellow" text="Manager Approved" />;
  }
  return <Badge color="gray" text="Pending" />;
};

const Badge = ({ color, text }: { color: string; text: string }) => {
  const styles: Record<string, string> = {
    green: 'bg-green-50 text-green-700 ring-green-600/20',
    red: 'bg-red-50 text-red-700 ring-red-600/20',
    yellow: 'bg-yellow-50 text-yellow-800 ring-yellow-600/20',
    gray: 'bg-gray-50 text-gray-600 ring-gray-500/10',
  };
  return (
    <span className={`inline-flex items-center rounded-full px-2 py-1 text-xs font-medium ring-1 ring-inset ${styles[color]}`}>
      {text}
    </span>
  );
};