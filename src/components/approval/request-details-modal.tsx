import React from 'react';
import { AbsenceRequest } from '@/lib/types'; // Import the type from our new central file

interface RequestDetailsModalProps {
  selectedRequest: AbsenceRequest;
  userRole: string;
  managerNotes: string;
  isSubmitting: boolean;
  setManagerNotes: (notes: string) => void;
  handleCloseModal: () => void;
  handleApproval: (status: 'approved' | 'denied') => void;
}

const RequestDetailsModal: React.FC<RequestDetailsModalProps> = ({
  selectedRequest,
  userRole,
  managerNotes,
  isSubmitting,
  setManagerNotes,
  handleCloseModal,
  handleApproval,
}) => {
  // Logic to determine if action buttons should be shown
  const canTakeAction = 
    (userRole === 'Manager' && selectedRequest.manager_approval === 'pending') ||
    ((userRole === 'HR' || userRole === 'Director') && 
     selectedRequest.final_approval === 'pending' && 
     (selectedRequest.manager_approval === 'approved' || ['Dentist', 'Supervisor', 'Exempt', 'EXEC'].includes(selectedRequest.employee_title)));

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 backdrop-blur-sm" onClick={handleCloseModal}>
      <div className="bg-white rounded-lg shadow-lg max-w-3xl w-full p-6 relative" onClick={(e) => e.stopPropagation()}>

        {/* Modal Header */}
        <div className="flex justify-between items-center p-5 border-b sticky top-0 bg-white">
          <h2 className="text-xl font-bold text-gray-800">{selectedRequest.createdAt?.toDate?.().toISOString().split('T')[0] || "No date"} {selectedRequest.employee_name}'s Request</h2>
          <button onClick={handleCloseModal} className="absolute top-4 right-4 text-gray-500 hover:text-gray-700">
            <svg className="w-5 h-5" fill="currentColor" viewBox="0 0 20 20"><path fillRule="evenodd" d="M4.293 4.293a1 1 0 011.414 0L10 8.586l4.293-4.293a1 1 0 111.414 1.414L11.414 10l4.293 4.293a1 1 0 01-1.414 1.414L10 11.414l-4.293 4.293a1 1 0 01-1.414-1.414L8.586 10 4.293 5.707a1 1 0 010-1.414z" clipRule="evenodd"></path></svg>
          </button>
        </div>

        {/* Modal Content Body*/}
        <div className="p-6 space-y-6 overflow-y-auto max-h-[70vh]">
          {/* Request Details */}
          <div>
            <h3 className="text-lg font-semibold text-gray-700 mb-2">{selectedRequest.type_of_request} Details</h3>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-x-4 gap-y-2 text-sm">
              <p><strong>Employee:</strong> {selectedRequest.employee_name}</p>
              <p><strong>Employee ID:</strong> {selectedRequest.employee_id}</p>
              <p><strong>Office:</strong> {selectedRequest.office}</p>
              <p><strong>Title:</strong> {selectedRequest.employee_title}</p>
              <p><strong>Request Type:</strong> {selectedRequest.type_of_request}</p>
              <p><strong>Incident Type:</strong> {selectedRequest.type_of_incident}</p>

              {(selectedRequest.etd || selectedRequest.eta) && (
                <div className="md:col-span-2 grid grid-cols-1 md:grid-cols-2 gap-4">
                  {selectedRequest.etd && (
                    <p><strong>ETD:</strong> <span className="text-gray-600">{selectedRequest.etd}</span></p>
                  )}

                  {selectedRequest.eta && (
                    <p><strong>ETA:</strong> <span className="text-gray-600">{selectedRequest.eta}</span></p>
                  )}
                </div>
              )}

              <p className="md:col-span-2"><strong>Date:</strong> {new Date(selectedRequest.incident_start).toDateString()} - {new Date(selectedRequest.incident_end).toDateString()}</p>
              
              {selectedRequest.employee_comments && (
                <p className="md:col-span-2"><strong>Employee Comments:</strong> <span className="text-gray-600 italic">{selectedRequest.employee_comments}</span></p>
              )}

              <p className="md:col-span-2">
                <strong>Excuse Note:</strong>{" "}
                {selectedRequest.excuse_note_submitted === "submitted" ? (
                  <span className="text-green-600 italic font-semibold">Submitted</span>
                ) : selectedRequest.excuse_note_submitted === "not_provided" ? (
                  <span className="text-red-600 italic font-semibold">Not Provided</span>
                ) : (
                  <span className="text-yellow-600 italic font-semibold">Pending</span>
                )}
              </p>

              {selectedRequest.excuse_note_submitted === "submitted" && selectedRequest.excuse_note && selectedRequest.excuse_note.length > 0 && (
                <div className="md:col-span-2">
                  <p className="text-sm font-medium text-gray-600">Attached Notes:</p>
                  <ul className="list-none pl-3 mt-1 text-sm space-y-2">
                    {selectedRequest.excuse_note.map((url: string, index: number) => (
                      <li key={index}>
                        <a href={url} target="_blank" rel="noopener noreferrer" className="text-blue-600 hover:underline">
                          Excuse Note {index + 1}
                        </a>
                      </li>
                    ))}
                  </ul>
                </div>
              )}

              <div className="md:col-span-2 grid grid-cols-1 md:grid-cols-2 gap-4">
                  {selectedRequest.createdAt && (
                    <p><strong>Created At:</strong> <span className="text-gray-600">{selectedRequest.createdAt?.toDate?.().toISOString().split('T')[0]}</span></p>
                  )}

                  {selectedRequest.updatedAt && (
                    <p><strong>Updated At:</strong> <span className="text-gray-600">{selectedRequest.updatedAt?.toDate?.().toISOString().split('T')[0]}</span></p>
                  )}
                </div>
            </div>
          </div>

          {/* Approval History */}
          <div>
            <h3 className="text-lg font-semibold text-gray-700 mb-2">Approval History</h3>
            <div className="space-y-3 text-sm">
            {/* Row 1: Manager Approval */}
              {!(['Dentist', 'Supervisor', 'Exempt', 'EXEC'].includes(selectedRequest.employee_title)) && (
                <div className="p-3 bg-gray-50 rounded-md grid grid-cols-2 gap-2 items-center text-sm">
                  <div>
                    <p>
                      <strong>Manager Approval:</strong>{" "}
                      <span className={`font-semibold italic capitalize ${
                        selectedRequest.manager_approval === 'approved' ? 'text-green-600' :
                        selectedRequest.manager_approval === 'denied' ? 'text-red-600' : 'text-yellow-600'
                      }`}>
                        {selectedRequest.manager_approval}
                      </span>
                    </p>
                  </div>

                  <div className="text-right">
                    {selectedRequest.manager_approval_name && (
                      <p className="text-gray-500 italic">By: {selectedRequest.manager_approval_name}</p>
                    )}
                  </div>
                </div>
              )}

              {/* Row 2: Final Approval */}
              {userRole !== 'Manager' && selectedRequest.manager_approval !== 'denied' && (
                <div className="p-3 bg-gray-50 rounded-md grid grid-cols-2 gap-2 items-center text-sm">
                  <div>
                    <p>
                      <strong>Final Approval:</strong>{" "}
                      <span className={`font-semibold italic capitalize ${
                        selectedRequest.final_approval === 'approved' ? 'text-green-600' :
                        selectedRequest.final_approval === 'denied' ? 'text-red-600' : 'text-yellow-600'
                      }`}>
                        {selectedRequest.final_approval}
                      </span>
                    </p>
                  </div>

                  <div className="text-right">
                    {selectedRequest.final_approval_name && (
                      <p className="text-gray-500 italic">By: {selectedRequest.final_approval_name}</p>
                    )}
                  </div>
                </div>
              )}

              {/* Decision Notes */}
              <div className="p-3 bg-gray-50 rounded-md">
                <p><strong>Notes:</strong> <span className="text-gray-600 italic">
                  {selectedRequest.manager_notes || "No notes provided."}
                </span></p>
              </div>

              {/* Action Input */}
              {canTakeAction && (
                <div className="mt-4">
                  <h3 className="text-lg font-semibold text-gray-700 mb-2">Take Action</h3>
                  <label htmlFor="managerNotes" className="block mb-2 text-sm font-medium text-gray-900">Notes (Optional):</label>
                  <textarea
                    id="managerNotes"
                    rows={3}
                    className="block p-2.5 w-full text-sm text-gray-900 bg-gray-50 rounded-lg border border-gray-300 focus:ring-blue-500 focus:border-blue-500"
                    placeholder="Add notes for your decision..."
                    value={managerNotes}
                    onChange={(e) => setManagerNotes(e.target.value)}
                  ></textarea>
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Modal Footer */}
        <div className="flex items-center justify-end p-5 border-t space-x-3 sticky bottom-0 bg-white">
          <button 
            onClick={handleCloseModal}
            disabled={isSubmitting}
            className="text-gray-500 bg-white hover:bg-gray-100 focus:ring-4 focus:outline-none focus:ring-blue-300 rounded-lg border border-gray-200 text-sm font-medium px-5 py-2.5 hover:text-gray-900 disabled:opacity-50"
          >
            Close
          </button>
          {canTakeAction && (
            <>
              <button
                onClick={() => handleApproval('denied')}
                disabled={isSubmitting}
                className="text-white bg-red-600 hover:bg-red-700 focus:ring-4 focus:outline-none focus:ring-red-300 font-medium rounded-lg text-sm px-5 py-2.5 text-center disabled:bg-red-400"
              >
                {isSubmitting ? 'Submitting...' : 'Deny'}
              </button>
              <button
                onClick={() => handleApproval('approved')}
                disabled={isSubmitting}
                className="text-white bg-green-600 hover:bg-green-700 focus:ring-4 focus:outline-none focus:ring-green-300 font-medium rounded-lg text-sm px-5 py-2.5 text-center disabled:bg-green-400"
              >
                {isSubmitting ? 'Submitting...' : 'Approve'}
              </button>
            </>
          )}
        </div>
      </div>
    </div>
  );
};

export default RequestDetailsModal;