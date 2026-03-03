'use client'; // Add this if using Next.js App Router and the component has client-side interactions

import React from 'react';
import { AbsenceRequest } from '@/lib/types'; // Import the type from our new central file

type RequestTableProps = {
  requests: AbsenceRequest[];
  status: 'needed' | 'pending' | 'approved' | 'denied' | 'all';
  onViewDetails: (request: AbsenceRequest) => void;
};

export const AbsenceTable: React.FC<RequestTableProps> = ({ requests, status, onViewDetails }) => {
  if (requests.length !== 0) {
    return (
      <div className="overflow-x-auto">
        {(status === 'needed') && (
          <h3 className='text-sm font-semibold'>EXCUSE NOTE PENDING:</h3>
        )}
        <table className="min-w-full bg-white divide-y divide-gray-200 mt-4">
          <thead className="bg-gray-50">
            <tr>
              <th className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Dates</th>
              <th className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Request Type</th>
              <th className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Incident Type</th>
              <th className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Office</th>
              <th className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Excuse Note</th>
              <th className="px-3 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">Action</th>
            </tr>
          </thead>
          <tbody className="bg-white divide-y divide-gray-200">
            {requests.map((req) => (
              <tr key={req.id} className="hover:bg-gray-100">
                <td className="px-3 py-4 whitespace-nowrap">{req.incident_start} to {req.incident_end}</td>
                <td className="px-3 py-4 whitespace-nowrap">{req.type_of_request}</td>
                <td className="px-3 py-4 whitespace-nowrap">{req.type_of_incident}</td>
                <td className="px-3 py-4 whitespace-nowrap">{req.office}</td>
                <td className="px-3 py-4 whitespace-nowrap capitalize">
                  {(req.excuse_note_submitted === 'pending') && (
                    <span className="px-2 inline-flex text-xs leading-5 font-semibold rounded-full bg-yellow-100 text-yellow-800">
                      Pending
                    </span>
                  )}
                  {(req.excuse_note_submitted === 'not_provided') && (
                  <span className="px-2 inline-flex text-xs leading-5 font-semibold rounded-full bg-red-100 text-red-800">
                    No Note
                  </span>
                  )}
                  {(req.excuse_note_submitted === 'submitted') && (
                    <span className="px-2 inline-flex text-xs leading-5 font-semibold rounded-full bg-green-100 text-green-800">
                      Submitted
                    </span>
                  )}

                </td>
                <td className="px-3 py-4 whitespace-nowrap text-right text-sm font-medium">
                  <button
                    onClick={() => onViewDetails(req)}
                    className="text-indigo-500 hover:text-indigo-900 font-medium cursor-pointer"
                  >
                    View Details
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    )
  }

  return <div className="p-4 text-center text-gray-500">No Request to Display</div>;
};
