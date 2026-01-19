'use client'; // Add this if using Next.js App Router and the component has client-side interactions

import React from 'react';
import { AbsenceRequest } from '@/lib/types'; // Import the type from our new central file

type RequestTableProps = {
  requests: AbsenceRequest[];
  userRole: string;
  onViewDetails: (request: AbsenceRequest) => void;
};

export const RequestTable: React.FC<RequestTableProps> = ({ requests, userRole, onViewDetails }) => {
  if (requests.length !== 0) {
    if (userRole === 'Director') {
      return (
        <div className="overflow-x-auto">
          <table className="min-w-full bg-white divide-y divide-gray-200 mt-4">
            <thead className="bg-gray-50">
              <tr>
                <th className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Employee ID</th>
                <th className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Name</th>
                <th className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Type</th>
                <th className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Dates</th>
                <th className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Submitted</th>
                <th className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Status</th>
                <th className="px-3 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">Action</th>
              </tr>
            </thead>
            <tbody className="bg-white divide-y divide-gray-200">
              {requests.map((req) => (
                <tr key={req.id} className="hover:bg-gray-100">
                  <td className="px-3 py-4 whitespace-nowrap">{req.employee_id}</td>
                  <td className="px-3 py-4 whitespace-nowrap">{req.employee_name}</td>
                  <td className="px-3 py-4 whitespace-nowrap">{req.type_of_request}</td>
                  <td className="px-6 py-4 whitespace-nowrap">{req.incident_start} to {req.incident_end}</td>
                  <td className="px-6 py-4 whitespace-nowrap">{req.date_submitted}</td>
                  <td className="px-6 py-4 whitespace-nowrap">
                    {['Dentist', 'Supervisor', 'Exempt', 'EXEC'].includes(req.employee_title) && req.final_approval === 'approved' && (
                      <span className="px-2 inline-flex text-xs leading-5 font-semibold rounded-full bg-green-100 text-green-800">
                        Fully Approved
                      </span>
                    )}
                    {req.manager_approval === 'approved' && req.final_approval === 'approved' && (
                      <span className="px-2 inline-flex text-xs leading-5 font-semibold rounded-full bg-green-100 text-green-800">
                        Fully Approved
                      </span>
                    )}
                    {req.manager_approval === 'approved' && req.final_approval === 'pending' && !['Dentist', 'Supervisor', 'Exempt', 'EXEC'].includes(req.employee_title) && (  
                      <span className="px-2 inline-flex text-xs leading-5 font-semibold rounded-full bg-yellow-100 text-yellow-800">
                        Manager Approved
                      </span>
                    )}
                    {req.manager_approval === 'pending' && !['Dentist', 'Supervisor', 'Exempt', 'EXEC'].includes(req.employee_title)  && (
                      <span className="px-2 inline-flex text-xs leading-5 font-semibold rounded-full bg-gray-100 text-gray-800">
                        Pending
                      </span>
                    )}
                    {(req.manager_approval === 'denied' || req.final_approval === 'denied') && (
                      <span className="px-2 inline-flex text-xs leading-5 font-semibold rounded-full bg-red-100 text-red-800">
                        Denied
                      </span>
                    )}
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap text-right text-sm font-medium">
                                      <button
                      onClick={() => onViewDetails(req)}
                      className="text-indigo-600 hover:text-indigo-900 font-medium"
                    >
                      View Details
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      );
    }
    else if (userRole === 'Manager'){
      return (
        <div className="overflow-x-auto">
          <table className="min-w-full bg-white divide-y divide-gray-200 mt-4">
            <thead className="bg-gray-50">
              <tr>
                <th className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Employee ID</th>
                <th className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Name</th>
                <th className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Type</th>
                <th className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Dates</th>
                <th className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Submitted</th>
                <th className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Status</th>
                <th className="px-3 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">Action</th>
              </tr>
            </thead>
            <tbody className="bg-white divide-y divide-gray-200">
              {requests.map((req) => (
                <tr key={req.id} className="hover:bg-gray-100">
                  <td className="px-3 py-4 whitespace-nowrap">{req.employee_id}</td>
                  <td className="px-3 py-4 whitespace-nowrap">{req.employee_name}</td>
                  <td className="px-3 py-4 whitespace-nowrap">{req.type_of_request}</td>
                  <td className="px-6 py-4 whitespace-nowrap">{req.incident_start} to {req.incident_end}</td>
                  <td className="px-6 py-4 whitespace-nowrap">{req.date_submitted}</td>
                  <td className="px-6 py-4 whitespace-nowrap">
                    {req.manager_approval === 'approved' && req.final_approval === 'approved' && (
                      <span className="px-2 inline-flex text-xs leading-5 font-semibold rounded-full bg-green-100 text-green-800">
                        Fully Approved
                      </span>
                    )}
                    {req.manager_approval === 'approved' && req.final_approval === 'pending' && (  
                      <span className="px-2 inline-flex text-xs leading-5 font-semibold rounded-full bg-yellow-100 text-yellow-800">
                        Manager Approved
                      </span>
                    )}
                    {req.manager_approval === 'pending' && (
                      <span className="px-2 inline-flex text-xs leading-5 font-semibold rounded-full bg-gray-100 text-gray-800">
                        Pending
                      </span>
                    )}
                    {req.manager_approval === 'denied' && (
                      <span className="px-2 inline-flex text-xs leading-5 font-semibold rounded-full bg-red-100 text-red-800">
                        Denied
                      </span>
                    )}
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap text-right text-sm font-medium">
                    <button
                      onClick={() => onViewDetails(req)}
                      className="text-indigo-600 hover:text-indigo-900 font-medium"
                    >
                      View Details
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      );
    }
  }
  return <div className="p-4 text-center text-gray-500">No requests to display.</div>;
};
