import React from 'react';
import { AbsenceRequest } from '@/lib/types'; // Import the type from our new central file
import { ExternalLink, FileText, Check, X } from 'lucide-react';

interface RequestDetailsModalProps {
  selectedRequest: AbsenceRequest;
  userRole: string;
  managerNotes: string;
  isSubmitting: boolean;
  setManagerNotes: (notes: string) => void;
  handleCloseModal: () => void;
  handleApproval: (status: 'approved' | 'denied' | 'approved_with_note') => void;
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

  const higherRoleUser = (userRole === 'Director' || userRole === 'HR');
  
  const canTakeAction = 
    (userRole === 'Manager' && selectedRequest.manager_approval === 'pending' && !selectedRequest.skipManagerApproval) ||
    ((higherRoleUser) && selectedRequest.final_approval === 'pending' && (selectedRequest.manager_approval === 'approved' || selectedRequest.skipManagerApproval === true || selectedRequest.manager_approval === 'not_required'));

  // Helper to calculate days between dates
  //const getDuration = () => {
  //  const start = new Date(selectedRequest.incident_start);
  //  const end = new Date(selectedRequest.incident_end);
  //  const diffTime = Math.abs(end.getTime() - start.getTime());
  //  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
  //  return diffDays === 1 ? '1 Day' : `${diffDays} Days`;
  //};

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 backdrop-blur-sm p-4" onClick={handleCloseModal}>
      <div className="bg-white rounded-xl shadow-2xl max-w-3xl w-full flex flex-col max-h-[90vh] relative" onClick={(e) => e.stopPropagation()}>
        
        {/* Modal Header */}
        <div className="flex justify-between items-center pt-6 pl-6 pr-6 pb-3 border-b">
          <div>
            <h2 className="text-xl font-bold text-gray-900">{selectedRequest.employee_name}</h2>
            <p className="text-sm text-blue-500">ID: #{selectedRequest.employee_id} • {selectedRequest.employee_title} <br/>Request ID: {selectedRequest.id}</p>
          </div>
          <button onClick={handleCloseModal} className="p-2 hover:bg-gray-100 rounded-full transition-colors">
            <svg className="w-6 h-6 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12" /></svg>
          </button>
        </div>

        {/* Modal Content */}
        <div className="p-6 overflow-y-auto space-y-8">
          
          {/* Section: Request Overview */}
          <section>
            <div className="flex items-center gap-2 mb-4">
              {/* Type of Request*/}
              <span className={`px-3 py-1 rounded-full text-xs font-bold uppercase tracking-wider ${selectedRequest.type_of_request === 'Time Off Request' ? 'bg-blue-100 text-blue-700' : 'bg-amber-100 text-amber-700'}`}>
                {selectedRequest.type_of_request}
              </span>
              {/* Did they skip Manager Approval*/}
              {selectedRequest.skipManagerApproval && (
                <span className="px-3 py-1 rounded-full text-xs font-bold bg-pink-100 text-pink-700 uppercase tracking-wider">
                  Manager Exempt
                </span>
              )}
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-6 bg-gray-50 p-4 rounded-lg border border-gray-100">
              {/* Office | Incident Type | Date */}
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase">Office</label>
                <p className="text-sm font-semibold">{selectedRequest.office}</p>
              </div>
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase">Incident Type</label>
                <p className="text-sm font-semibold">{selectedRequest.type_of_incident}</p>
              </div>
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase">Dates</label>
                <p className="text-sm font-semibold">
                  {(selectedRequest.incident_start === selectedRequest.incident_end) ? (selectedRequest.incident_start).replace(/^(\d{4})-(\d{2})-(\d{2})$/, '$2/$3/$1') : `${(selectedRequest.incident_start).replace(/^(\d{4})-(\d{2})-(\d{2})$/, '$2/$3/$1')} — ${(selectedRequest.incident_end).replace(/^(\d{4})-(\d{2})-(\d{2})$/, '$2/$3/$1')}`}
                </p>
              </div>

              {/* ETA and ETD (Appears if provided for specific incidnet types */}
              {/* Row 2: ETA/ETD - Centered below the first row */}
              {(selectedRequest.eta || selectedRequest.etd) && (
                <div className="md:col-span-3 flex flex-row justify-center items-center gap-12 pt-2 border-t border-gray-200/50">
                  {selectedRequest.etd && (
                    <div className="text-center">
                      <label className="text-[10px] font-bold text-rose-400 uppercase block mb-0.5">Leaving At</label>
                      <p className="text-sm font-bold text-gray-700">{selectedRequest.etd}</p>
                    </div>
                  )}
                  
                  {/* Optional: Subtle divider between times if both exist */}
                  {selectedRequest.etd && selectedRequest.eta && (
                    <div className="h-4 w-px bg-gray-300 self-end mb-1" />
                  )}

                  {selectedRequest.eta && (
                    <div className="text-center">
                      <label className="text-[10px] font-bold text-emerald-500 uppercase block mb-0.5">Arriving At</label>
                      <p className="text-sm font-bold text-gray-700">{selectedRequest.eta}</p>
                    </div>
                  )}
                </div>
              )}
            </div>
          </section>

          {/* Section: Excuse Note (If applicable) */}
          <section>
            <h4 className="text-xs font-bold text-gray-400 uppercase mb-2">Excuse Notes</h4>

            {selectedRequest.excuse_note_submitted === 'submitted' && selectedRequest.excuse_note && selectedRequest.excuse_note.length > 0 ? (
              <div className="flex flex-wrap gap-2">
                {selectedRequest.excuse_note.map((noteUrl, index) => (
                  <a 
                    key={index} 
                    href={noteUrl} 
                    target="_blank" 
                    rel="noopener noreferrer"
                    className="flex items-center gap-2 text-sm font-bold text-emerald-700 bg-emerald-50 border border-emerald-200 px-4 py-2 rounded-xl hover:bg-emerald-100 transition-all shadow-sm"
                  >
                    <FileText className="h-4 w-4" />
                    <span>View Excuse Note {selectedRequest.excuse_note!.length > 1 ? index + 1 : ''}</span>
                    <ExternalLink className="h-3 w-3 opacity-50" />
                  </a>
                ))}
              </div>
            ) : selectedRequest.excuse_note_submitted === 'submitted' ? (
              /* Fallback if status is 'submitted' but array is empty/missing */
              <div className="text-sm text-gray-500 italic">Processing upload...</div>
            ) : null}

            {selectedRequest.excuse_note_submitted === 'pending' && (
              <div className="inline-flex items-center gap-2 text-sm font-bold text-amber-700 bg-amber-50 border border-amber-200 px-4 py-2 rounded-xl">
                Awaiting Documentation
              </div>
            )}

            {selectedRequest.excuse_note_submitted === 'not_provided' && (
              <div className="inline-flex items-center gap-2 text-sm font-bold text-red-700 bg-red-50 border border-red-200 px-4 py-2 rounded-xl">
                No Excuse Note Provided
              </div>
            )}
          </section>

          {/* Section: Comments & Notes */}
          <section className="grid md:grid-cols-2 gap-6">
            <div>
              <h4 className="text-xs font-bold text-gray-400 uppercase mb-2">Employee Comments</h4>
              <p className="text-sm text-gray-700 bg-white p-3 border rounded-md italic">
                {selectedRequest.employee_comments || "No comments provided."}
              </p>
            </div>
              {!(selectedRequest.type_of_request === "HR Call In" && userRole === 'Manager') && !((selectedRequest.final_approval === 'approved' || selectedRequest.final_approval === 'denied') && userRole === 'Manager') && (
                <div>
                  <h4 className="text-xs font-bold text-gray-400 uppercase mb-2">Admin Notes</h4>
                  <p className="text-sm text-gray-700 bg-white p-3 border rounded-md italic">
                    {selectedRequest.manager_notes || "No notes provided."}
                  </p>
                </div>
              )}
          </section>

          {/* Section: Approval Workflow */}
          <section>
            <h4 className="text-xs font-bold text-gray-400 uppercase mb-3">Approval Status</h4>
            <div className="space-y-3">
              {/* Manager Step */}
              <div className="flex items-center justify-between p-3 rounded-lg border border-gray-200">
                <span className="text-sm font-medium">Manager Approval</span>
                {selectedRequest.skipManagerApproval ? (
                  <span className="text-xs font-bold text-gray-400 uppercase italic">N/A (Exempt)</span>
                ) : (
                  <div className="text-right">
                    <span className={`text-xs font-bold uppercase ${selectedRequest.manager_approval === 'approved' ? 'text-green-600' : selectedRequest.manager_approval === 'denied' ? 'text-red-600' : selectedRequest.manager_approval === 'not_required' ? 'text-gray-800' :'text-amber-500'}`}>
                      {selectedRequest.manager_approval === 'not_required'  ? 'Not Required': selectedRequest.manager_approval}
                    </span>
                    {selectedRequest.manager_approval_name && <p className="text-[10px] text-gray-400">{selectedRequest.manager_approval_name}</p>}
                  </div>
                )}
              </div>

              {/* Final Step */}
              <div className="flex items-center justify-between p-3 rounded-lg border border-gray-200">
                <span className="text-sm font-medium">Final Approval</span>
                <div className="text-right">
                  <span className={`text-xs font-bold uppercase ${(selectedRequest.final_approval === 'approved' || selectedRequest.final_approval === 'approved_with_note') ? 'text-green-600' : selectedRequest.final_approval === 'denied' ? 'text-red-600' : 'text-amber-500'}`}>
                    {selectedRequest.final_approval === 'approved_with_note'  ? 'Approved With Note': selectedRequest.final_approval}
                  </span>
                  {selectedRequest.final_approval_name && <p className="text-[10px] text-gray-400">{selectedRequest.final_approval_name}</p>}
                </div>
              </div>
            </div>
          </section>

          {/* Decision Area */}
          {canTakeAction && !(selectedRequest.type_of_request === "HR Call In" && userRole === 'Manager') && !((selectedRequest.final_approval === 'approved' || selectedRequest.final_approval === 'denied') && userRole === 'Manager') && (
            <section className="bg-indigo-50 p-4 rounded-xl border border-indigo-100">
              <h3 className="text-sm font-bold text-indigo-900 mb-3 uppercase tracking-tighter">Admin Notes</h3>
              <textarea
                rows={3}
                className="block w-full text-sm p-3 rounded-lg border border-indigo-200 focus:ring-2 focus:ring-indigo-500 outline-none"
                placeholder="Add Notes (Optional)..."
                value={managerNotes}
                onChange={(e) => setManagerNotes(e.target.value)}
              />
            </section>
          )}
          {/*!canTakeAction && !(userRole === 'Manager') && (

            <div className="text-center text-sm italic text-gray-500">
              {managerNotes}
            </div>
            )
          */}
        </div>

        {/* Modal Footer */}
          <div className="p-6 border-t bg-gray-50 rounded-b-xl flex justify-between items-center">
            {/* Left Side: Date Stack */}
            <div className="flex flex-col gap-1">
              <p className="text-[10px] text-gray-400 font-medium italic uppercase tracking-tight">
                Submitted: {selectedRequest.createdAt?.toDate().toLocaleDateString()}
              </p>
              {selectedRequest.updatedAt && (
                <p className="text-[10px] text-indigo-400 font-medium italic uppercase tracking-tight">
                  Updated: {selectedRequest.updatedAt?.toDate().toLocaleDateString()}
                </p>
              )}
            </div>

            {/* Right Side: Action Buttons */}
            <div className="flex gap-3">
              {/*<button 
                onClick={handleCloseModal} 
                className="px-4 py-2 text-sm font-bold text-gray-500 hover:text-gray-700 transition-colors"
              >
                Cancel
              </button>{*/}
              
              {canTakeAction && (
                <>
                  {/* NEW: Save Notes Only Button */}

                  <button 
                    onClick={() => handleApproval('denied')} 
                    disabled={isSubmitting}
                    className="bg-red-600 hover:bg-red-700 text-white px-6 py-2 rounded-lg text-sm font-bold disabled:bg-gray-300 transition-colors"
                  >
                    <span className="block sm:hidden"><X /></span>
                    <span className="hidden sm:block">Deny</span>
                  </button>
                  
                  <button 
                    onClick={() => handleApproval('approved')} 
                    disabled={isSubmitting}
                    className="bg-green-600 hover:bg-green-700 text-white px-6 py-2 rounded-lg text-sm font-bold disabled:bg-gray-300 transition-colors shadow-lg shadow-green-600/20"
                  >
                    <span className="block sm:hidden"><Check /></span>
                    <span className="hidden sm:block">Approve</span>
                  </button>

                  {(higherRoleUser) && (
                    <button
                      onClick={() => handleApproval('approved_with_note')} 
                      disabled={isSubmitting}
                      className="bg-green-600 hover:bg-green-700 text-white px-6 py-2 rounded-lg text-sm font-bold disabled:bg-gray-300 transition-colors shadow-lg shadow-green-600/20"
                  >
                    <span className="block sm:hidden"><Check /> <FileText/></span>
                    <span className="hidden sm:block">Approve With Note</span>
                  </button>
                  )}
                </>
              )}
            </div>
          </div>
      </div>
    </div>
  );
};

export default RequestDetailsModal;