"use client";

import React, { useEffect, useMemo } from 'react';
import { X, Clock, MessageSquare, Save, Archive, ArchiveRestore, AlertCircle, FileText, User, Plus, Paperclip, FileEdit } from 'lucide-react';
import { AbsenceRequest } from "@/lib/types";
import { ref, uploadBytes, getDownloadURL, deleteObject} from "firebase/storage";
import { storage } from '@/lib/firebase.config';
import { OFFICES, getRequestColor } from '@/lib/constants';

interface HRDetailsProps {
  absence: AbsenceRequest;
  userName: string;
  onClose: () => void;
  onUpdate: (updated: AbsenceRequest) => void;
  onArchive: (id: string) => void;
  isSaving: boolean;
}

export const HRRequestDetailsModal = ({ absence, userName, onClose, onUpdate, onArchive, isSaving }: HRDetailsProps) => {
  const [tempData, setTempData] = React.useState<AbsenceRequest>(absence);
  const [isNotesForcedOpen, setIsNotesForcedOpen] = React.useState<boolean>(false);

  const [deleteFiles, setDeleteFiles] = React.useState<string[]>([]);
  const [newFiles, setNewFiles] = React.useState<File[]>([]); // Track files not yet uploaded
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  // 1. Change Detection Logic
  // Compares current state field-by-field to original database record
  const isChanged = (field: keyof AbsenceRequest) => {
    return JSON.stringify(tempData[field]) !== JSON.stringify(absence[field]);
  };

  // Determines if the "Commit Changes" button should be enabled
  const hasChanges = useMemo(() => {
    const dataChanged = JSON.stringify(tempData) !== JSON.stringify(absence);
    const filesAdded = newFiles.length > 0;
    const filesRemoved = deleteFiles.length > 0;
    return dataChanged || filesAdded || filesRemoved;
  }, [tempData, absence, newFiles, deleteFiles]);

  // 2. Documentation Logic
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      setNewFiles(prev => [...prev, ...Array.from(e.target.files!)]);
      // Set status to submitted if files are being added
      setTempData(prev => ({ ...prev, excuse_note_submitted: 'submitted' }));
    }
  };

  const removeNewFile = (index: number) => { // Changed to index for accuracy
    setNewFiles(prev => prev.filter((_, i) => i !== index));
  };

  // 3. Dynamic Enable/Disable Requirements
  const canEditETD = tempData.type_of_incident === 'Early Out' || tempData.type_of_incident === 'Leave and Come Back';
  const canEditETA = tempData.type_of_incident === 'Late In' || tempData.type_of_incident === 'Leave and Come Back';

  // Condition to show final notes text area (Automatically shows if it has text, or if button was clicked)
  const showFinalNotes = !!tempData.final_notes || isNotesForcedOpen;
    // Date Validation Logic
  const isDateInvalid = Boolean(
    tempData.incident_start && 
    tempData.incident_end && 
    new Date(tempData.incident_end) < new Date(tempData.incident_start)
  );

  return (
    <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-md z-50 flex items-center justify-center p-4">
      <div className="w-full max-w-4xl rounded-[2.5rem] shadow-2xl overflow-hidden animate-in zoom-in-95 duration-200">
        
        {/* Header */}
        <div className="py-4 px-6 border-b border-slate-50 flex justify-between items-center bg-slate-900">
          <div>
            <h2 className="text-xl font-black text-indigo-400 leading-none mb-1">
              {absence.createdAt.toDate().toLocaleDateString()} @ {absence.createdAt.toDate().toLocaleTimeString([], { 
                hour: '2-digit', 
                minute: '2-digit', 
                hour12: true 
              })} | Absence Request Details
            </h2>
            
            <div className="flex flex-col space-y-0.5">
              <div className="flex gap-3">
                <p className="text-[10px] font-bold text-indigo-500 uppercase tracking-widest leading-tight">
                  Absence ID: {absence.id}
                </p>
                
                {/* Conditional "Last Updated" Tag */}
                {absence.updatedAt && (
                  <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest leading-tight border-l border-slate-200 pl-3">
                    Last Updated: {absence.updatedAt.toDate().toLocaleString([], { 
                      month: '2-digit', 
                      day: '2-digit', 
                      hour: '2-digit', 
                      minute: '2-digit' 
                    })}
                  </p>
                )}
              </div>
              
              <p className="text-[10px] font-bold text-indigo-500 uppercase tracking-widest leading-tight">
                Employee Database ID: {absence.employeeFirestoreID}
              </p>
            </div>
          </div>
          
          <button onClick={onClose} className="p-2 hover:bg-white/10 rounded-full transition-all">
            <X className="h-5 w-5 text-white" />
          </button>
        </div>

        <form 
          className="bg-white py-2 px-8 space-y-6 max-h-[85vh] overflow-y-auto" 
          onSubmit={async (e) => {
            e.preventDefault();

            if (isDateInvalid) {
              console.error("End date cannot be before start date.");
              return;
            }
            
            try {
              if (deleteFiles.length > 0) {
                for (const url of deleteFiles) {
                  const fileRef = ref(storage, url);
                  await deleteObject(fileRef);
                }
              }
            // Start with current note URLs
            let uploadedFileUrls: string[] = [...(tempData.excuse_note || [])];

            // Upload new files if they exist
            if (newFiles.length > 0) {
              for (const file of newFiles) {
                const storageRef = ref(storage, `excuse-notes/${absence.id}/${file.name}`);
                await uploadBytes(storageRef, file);
                const url = await getDownloadURL(storageRef);
                uploadedFileUrls.push(url);
              }
            }

            const finalDataToSave = { ...tempData, excuse_note: uploadedFileUrls };

            // Your original signature cleanup logic
            if (finalDataToSave.manager_approval !== absence.manager_approval) {
              finalDataToSave.manager_approval_name = ""; 
            }
            if (finalDataToSave.final_approval !== absence.final_approval) {
              finalDataToSave.final_approval_name = "";
            }

            onUpdate(finalDataToSave);
            setDeleteFiles([]);
          }
          catch (err){
            console.error("Error uploading files or saving data:", err);
            alert("Error uploading Files")
          }
          }}
        >
          {/* Section 1: Read-Only Employee Info */}
          {/* employee_name, employee_id, employee_title (Read Only) */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4 p-4 rounded-3xl border bg-purple-200/50 border-purple-100">
            <div className="space-y-1">
              <label className="text-[9px] font-black text-slate-700 uppercase ml-1">Employee Name</label>
              <p className="text-sm font-bold text-indigo-600 px-1">{absence.employee_name}</p>
            </div>
            <div className="space-y-1 border-slate-200 md:border-l md:pl-4">
              <label className="text-[9px] font-black text-slate-700 uppercase ml-1">Employee ID</label>
              <p className="text-sm font-bold text-indigo-600 px-1">{absence.employee_id}</p>
            </div>
            <div className="space-y-1 border-slate-200 md:border-l md:pl-4">
              <label className="text-[9px] font-black text-slate-700 uppercase ml-1">Employee Title</label>
              <p className="text-sm font-bold text-indigo-600 px-1">{absence.employee_title}</p>
            </div>
          </div>
          
          {/* employee_comments */}
          <div className="space-y-3 gap-4 p-4 rounded-3xl border bg-amber-200/50 border-amber-100">
            <label className="text-[10px] font-black text-slate-700 uppercase ml-1 flex items-center gap-1">
              <MessageSquare className="h-3 w-3 text-slate-400"/> Employee Comments
            </label>
            <p className="text-sm font-bold text-indigo-600 px-1">{absence.employee_comments || "No Employee Comments Provided"}</p>
          </div>

          {/* Section 2: Type Selection (Light Red if changed) */}
          <div className= 'grid grid-cols-1 md:grid-cols-3 gap-6 p-4 rounded-3xl border transition-all bg-slate-200/50 border-slate-200 shadow-sm'>
            <div className="space-y-1">
              <label className="text-[10px] font-black text-black uppercase ml-1">Office</label>
              <select 
                value={tempData.office}
                onChange={(e) => setTempData({...tempData, office: e.target.value})}
                className={`w-full rounded-xl text-sm font-bold border-2 pl-1
                ${isChanged('office') ? 'bg-rose-50 border-rose-300 text-rose-900' : 'bg-white text-slate-700 border-indigo-200'}`}
              >
              {OFFICES.map((office) => (
                <option key={office} value={office}>
                  {office}
                </option>
              ))}
              </select>
            </div>

            <div className="space-y-1 border-slate-100 md:border-l md:pl-4">
              <label className="text-[10px] font-black text-black uppercase ml-1">Type of Incident</label>
              <select 
                value={tempData.type_of_incident}
                onChange={(e) => {
                  const newType = e.target.value as any;
      
                  // Define which types are allowed to have times
                  const needsETD = newType === 'Early Out' || newType === 'Leave and Come Back';
                  const needsETA = newType === 'Late In' || newType === 'Leave and Come Back';
                  setTempData({
                    ...tempData, 
                    type_of_incident: newType,
                    // Clear times if new type doesn't require them
                    etd: needsETD ? tempData.etd : "",
                    eta: needsETA ? tempData.eta : "",
                  });
                }}
                className={`w-full rounded-xl text-sm font-bold border-2 pl-1
                ${isChanged('type_of_incident') ? 'bg-rose-50 border-rose-300 text-rose-900' : 'bg-white text-slate-700 border-indigo-200'}`}
              >
                <option value="Late In">Late In</option>
                <option value="Early Out">Early Out</option>
                <option value="Leave and Come Back">Leave and Come Back</option>
                <option value="Absent">Absent</option>
                <option value="Long Lunch">Long Lunch</option>
                <option value="Switch Shift">Switch Shift</option>
              </select>
            </div>

            <div className="space-y-1 border-slate-100 md:border-l md:pl-4">
              <label className="text-[10px] font-black text-black uppercase ml-1">Request Type</label>
              <select 
                value={tempData.type_of_request}
                onChange={(e) => setTempData({...tempData, type_of_request: e.target.value as any})}
                className={`w-full rounded-xl text-sm font-bold border-2 transition-colors pl-1
                ${isChanged('type_of_request') ? 'bg-rose-50 border-rose-300 text-rose-900' : getRequestColor(tempData.type_of_request)}`}
              >
                <option value="HR Call In">HR Call In</option>
                <option value="Incident Notice">Incident Notice</option>
                <option value="Time Off Request">Time Off Request</option>
                <option value="No Call">No Call</option>
                <option value="Call In After Shift">Call In After Shift</option>
                <option value="Previously Not Approved">Previously Not Approved</option>
              </select>
            </div>
          </div>

          {/* Section 3: Dates & Times */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="space-y-3 p-4 rounded-3xl border transition-all bg-slate-200/50 border-slate-200 shadow-sm">
              <label className="text-[10px] font-black text-black uppercase ml-1">Incident Window</label>
              <div className="flex items-center gap-2 bg-white p-2 rounded-xl">
                <div className="flex-1 space-y-1">
                  <p className="text-[8px] font-black text-slate-400 uppercase ml-1">Start Date</p>
                  <input 
                    type="date" 
                    value={tempData.incident_start || ""} 
                    onChange={e => setTempData({...tempData, incident_start: e.target.value})}
                    className={`w-full p-2 rounded-lg border-none text-xs font-bold text-slate-700
                    ${isChanged('incident_start') ? 'bg-rose-50 text-rose-900' : 'bg-zinc-100 text-slate-700'}`}
                  />
                </div>
                <span className="text-slate-300 mt-4">→</span>
                <div className="flex-1 space-y-1">
                  <p className="text-[8px] font-black text-slate-400 uppercase ml-1">End Date</p>
                  <input 
                    type="date" 
                    value={tempData.incident_end || ""} 
                    min={tempData.incident_start || ""}
                    onChange={e => setTempData({...tempData, incident_end: e.target.value})}
                    className={`w-full p-2 rounded-lg border-none text-xs font-bold text-slate-700
                    ${isChanged('incident_end') ? 'bg-rose-50 text-rose-900' : 'bg-zinc-100 text-slate-700'}`}
                  />
                </div>
              </div>
              {/* Date Validation Error Message */}
              {isDateInvalid && (
                <div className="flex items-center gap-2 p-3 bg-rose-50 rounded-xl text-rose-600 text-xs font-bold border border-rose-200 animate-in fade-in">
                  <AlertCircle className="w-4 h-4 shrink-0" />
                  <span>End date cannot be earlier than the start date.</span>
                </div>
              )}
            </div>

            <div className="grid grid-cols-2 gap-3 bg-slate-200/50 px-4 py-8 rounded-2xl border border-indigo-100 shadow-sm">
              <div className="space-y-1">
                <label className={`text-[9px] font-black uppercase flex items-center gap-1 transition-colors ${canEditETD ? 'text-black' : 'text-slate-400'}`}>
                  <Clock className="h-3.5 w-3.5"/> ETD (Out)
                </label>
                <input 
                  type="time" 
                  disabled={!canEditETD}
                  value={tempData.etd || ""} 
                  onChange={e => setTempData({...tempData, etd: e.target.value})} 
                  className={`w-full p-2 rounded-lg border-2 text-sm font-bold transition-all ${
                    canEditETD 
                      ? isChanged('etd') ? 'bg-rose-50 border-rose-300 text-rose-900' : 'bg-white text-slate-900 border-indigo-200' 
                      : 'bg-slate-100/50 text-slate-400 border-transparent cursor-not-allowed opacity-50'
                  }`} 
                />
              </div>

              <div className="space-y-1">
                <label className={`text-[9px] font-black uppercase flex items-center gap-1 transition-colors ${canEditETA ? 'text-black' : 'text-slate-400'}`}>
                  <Clock className="h-3.5 w-3.5"/> ETA (In)
                </label>
                <input 
                  type="time" 
                  disabled={!canEditETA}
                  value={tempData.eta || ""} 
                  onChange={e => setTempData({...tempData, eta: e.target.value})} 
                  className={`w-full p-2 rounded-lg border-2 text-sm font-bold transition-all ${
                    canEditETA 
                      ? isChanged('eta') ? 'bg-rose-50 border-rose-300 text-rose-900' : 'bg-white text-slate-900 border-indigo-200' 
                      : 'bg-slate-100/50 text-slate-400 border-transparent cursor-not-allowed opacity-50'
                  }`} 
                />
              </div>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-1 gap-6">
            {/* Opening Comment */}
            <div className="space-y-2 p-4 rounded-3xl border transition-all bg-amber-200/50 border-amber-100 shadow-sm">
              <label className="text-s underline block text-center font-bold text-red-500 uppercase">For Key Holders or Supervisors</label>
              <div className="flex items-center">
                <label className="w-20 text-[10px] font-black text-black uppercase ml-1">Opening</label>
                <textarea name="comment_opening" value={tempData.comment_opening || ""} onChange={e => setTempData({ ...tempData, comment_opening: e.target.value })} 
                  className={`w-full p-1 rounded-md border text-xs font-medium transition-all outline-none resize-none ${
                  isChanged('comment_opening')
                    ? 'bg-rose-50 border-rose-300 text-rose-900 focus:ring-2 focus:ring-rose-200'
                    : 'bg-white border-slate-200 text-slate-800 focus:ring-2 focus:ring-indigo-100'
                }`}
                rows={1}/>
              </div>
              <div className="flex items-center">
                <label className="w-20 text-[10px] font-black text-black uppercase ml-1">Closing</label>
                <textarea name="comment_closing" value={tempData.comment_closing || ""} onChange={e => setTempData({ ...tempData, comment_closing: e.target.value })}
                className={`w-full p-1 rounded-md border text-xs font-medium transition-all outline-none resize-none ${
                  isChanged('comment_closing')
                    ? 'bg-rose-50 border-rose-300 text-rose-900 focus:ring-2 focus:ring-rose-200'
                    : 'bg-white border-slate-200 text-slate-800 focus:ring-2 focus:ring-indigo-100'
                }`}
                rows={1}/>
              </div>
            </div>
          </div>

          {/* Section 5: Excuse Notes Management */}
          <div className={`space-y-4 p-3 rounded-3xl border transition-all shadow-sm ${ (isChanged('excuse_note') || newFiles.length > 0 || isChanged('excuse_note_submitted')) ? 'bg-rose-50 border-rose-200' : 'bg-amber-200/50 border-amber-100' }`}>
            <div className="flex justify-between items-center">
              <label className="text-[10px] font-black text-black uppercase ml-1 flex items-center gap-1"><Paperclip className="h-3 w-3" /> Excuse Notes</label>
              <input 
                type="file" 
                className="hidden" 
                ref={fileInputRef} 
                multiple 
                onChange={handleFileChange} 
              />
              
              <button 
                type="button" 
                onClick={() => fileInputRef.current?.click()} 
                className="text-[9px] font-black bg-indigo-600 text-white px-3 py-1.5 rounded-full uppercase hover:bg-indigo-700 flex items-center gap-1"
              >
                <Plus className="h-3 w-3" /> Add File
              </button>
            </div>

            <div className="flex gap-2">
              {['pending', 'submitted', 'not_provided'].map((status) => {
                const isSelected = tempData.excuse_note_submitted === status;
                const fieldChanged = isChanged('excuse_note_submitted');
                // Logic: Disable "Not Provided" if any files exist (staged or database)
                const hasFiles = (tempData.excuse_note?.length || 0) > 0 || newFiles.length > 0;
                const notProvidedDisabled = status === 'not_provided' && hasFiles;

                return (
                  <button
                    key={status}
                    type="button"
                    disabled={notProvidedDisabled}
                    onClick={() => setTempData({ ...tempData, excuse_note_submitted: status as any })}
                    className={`flex-1 py-2.5 rounded-xl text-[10px] font-black uppercase border-2 transition-all ${
                      isSelected
                        ? 'bg-emerald-600 text-white border-emerald-600 shadow-md'
                        : fieldChanged
                          ? 'bg-white text-rose-400 border-rose-100 hover:border-rose-200'
                          : 'bg-white text-slate-400 border-slate-100 hover:border-slate-200'
                    }`}
                  >
                    {status.replace('_', ' ')}
                  </button>
                );
              })}
            </div>

            <div className="flex flex-wrap gap-2 pt-2 border-slate-200/60">
              {/* Existing Files from Database */}
              {tempData.excuse_note?.map((noteUrl, index) => (
                <div key={`existing-${index}`} className="group relative flex items-center">
                  <a href={noteUrl} target="_blank" rel="noopener noreferrer" className="flex items-center gap-2 text-[11px] font-bold text-emerald-700 bg-white border border-emerald-100 pl-3 pr-8 py-2 rounded-lg hover:bg-emerald-50 shadow-sm">
                    <FileText className="h-3.5 w-3.5" /> Note {index + 1}
                  </a>
                  <button 
                    type="button" 
                    onClick={() => {
                      setDeleteFiles(prev => [...prev, noteUrl]);
                      setTempData({ ...tempData, excuse_note: tempData.excuse_note?.filter((_, i) => i !== index) });
                    }} 
                    className="absolute right-1 p-1 text-rose-400 hover:text-rose-600"
                  >
                    <X className="h-3.5 w-3.5" />
                  </button>
                </div>
              ))}

              {/* Staged Files (Not yet uploaded) */}
              {newFiles.map((file, index) => (
                <div key={`new-${index}`} className="group relative flex items-center">
                  <div className="flex items-center gap-2 text-[11px] font-bold text-indigo-700 bg-indigo-50 border border-indigo-200 pl-3 pr-8 py-2 rounded-lg shadow-sm">
                    <Paperclip className="h-3.5 w-3.5" /> {file.name}
                    <span className="text-[8px] uppercase px-1 bg-indigo-200 rounded">New</span>
                  </div>
                  <button 
                    type="button" 
                    onClick={() => removeNewFile(index)} 
                    className="absolute right-1 p-1 text-indigo-400 hover:text-indigo-600"
                  >
                    <X className="h-3.5 w-3.5" />
                  </button>
                </div>
              ))}
            </div>
          </div>

          {/* Section 6: Approvals */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6 border-slate-100">
            {/* Manager Approval Box */}
            <div className={`space-y-3 p-3 col-span-2 md:col-span-1 rounded-2xl border transition-all shadow-sm ${isChanged('manager_approval') ? 'bg-rose-50 border-rose-200' : 'bg-emerald-200/50 border-emerald-200'}`}>
              <div className="flex justify-between items-center ml-1">
                <label className="text-[10px] font-black text-black uppercase">Manager Status</label>
                {tempData.manager_approval_name && (
                  <span className={`text-[9px] font-bold italic transition-colors ${isChanged('manager_approval') ? 'text-rose-600' : 'text-emerald-800'}`}>
                    Submitted: {tempData.manager_approval_name}
                  </span>
                )}
              </div>
              <div className="flex gap-1.5">
                {['pending', 'denied', 'approved', 'not_required'].map((status) => (
                  <button 
                    key={status} 
                    type="button" 
                    onClick={() => {
                      const isOriginalStatus = status === absence.manager_approval;
                      const isFinalizing = status === 'approved' || status === 'denied';
                      
                      setTempData({ 
                        ...tempData, 
                        manager_approval: status as any,
                        // 1. If back to original: use DB name. 
                        // 2. If new approval/denial: use current user. 
                        // 3. Otherwise (pending/skip): clear it.
                        manager_approval_name: isOriginalStatus 
                          ? absence.manager_approval_name 
                          : isFinalizing ? userName : ""
                      });
                    }} 
                    className={`flex-1 py-3 rounded-xl text-[9px] font-black uppercase border-2 transition-all ${tempData.manager_approval === status ? 'bg-slate-900 text-white border-slate-900 shadow-md' : 'bg-white text-slate-400 border-slate-100'}`}
                  > 
                    {status === 'not_required' ? 'Skip' : status} 
                  </button>
                ))}
              </div>
            </div>
            {/* Final Approval Box */}
            <div className={`space-y-3 p-3 col-span-2 md:col-span-1 rounded-2xl border transition-all shadow-sm ${isChanged('final_approval') ? 'bg-rose-50 border-rose-200' : 'bg-sky-200/50 border-sky-200'}`}>
              <div className="flex justify-between items-center ml-1">
                <div className="flex items-center gap-2">
                  <label className="text-[10px] font-black text-black uppercase ml-1">Final Approval</label>
                  {/* 💡 📝 Dynamic Notes Trigger Button inside Box */}
                  {!showFinalNotes && (
                    <button 
                      type="button" 
                      onClick={() => setIsNotesForcedOpen(true)}
                      className="text-[9px] font-black text-indigo-600 bg-white hover:bg-indigo-50 border border-indigo-200 px-2 py-0.5 rounded-full uppercase flex items-center gap-1 transition-all"
                    >
                      <FileEdit className="h-2.5 w-2.5" /> Add Notes
                    </button>
                  )}
                </div>
                {tempData.final_approval_name && (
                  <span className={`text-[9px] font-bold italic transition-colors ${isChanged('final_approval') ? 'text-rose-600' : 'text-emerald-800'}`}>
                    Submitted {tempData.final_approval_name}
                  </span>
                )}
              </div>
              <div className="flex gap-1.5">
                {['pending', 'denied', 'approved', 'approved_with_note'].map((status) => (
                  <button 
                    key={status} 
                    type="button" 
                    onClick={() => {
                      const isOriginalStatus = status === absence.final_approval;
                      const isFinalizing = status === 'approved' || status === 'denied';

                      setTempData({ 
                        ...tempData, 
                        final_approval: status as any,
                        // Same logic: Restore DB name if user returns to the original state
                        final_approval_name: isOriginalStatus 
                          ? absence.final_approval_name 
                          : isFinalizing ? userName : ""
                      });
                    }} 
                    className={`flex-1 py-3 rounded-xl text-[9px] font-black uppercase border-2 transition-all ${tempData.final_approval === status ? 'bg-slate-900 text-white border-slate-900 shadow-md' : 'bg-white text-slate-400 border-slate-100'}`}
                  > 
                    {status === 'approved_with_note' ? 'Approved w/ Note' : status} 
                  </button>
                ))}
              </div>
            </div>
          </div>

          {/* Final Notes for Employee to see, needs to press button to see this part of the form */}
          {showFinalNotes && (
            <div className={`space-y-1 p-3 rounded-2xl border transition-all shadow-sm animate-in fade-in slide-in-from-top-2 duration-200 ${ isChanged('final_notes') ? 'bg-rose-50 border-rose-200' : 'bg-sky-200/50 border-sky-200' }`}>
              <label className="text-[10px] font-black text-black uppercase ml-1 flex items-center gap-1 ">Final Notes</label>
              <textarea 
                className="w-full bg-transparent border-none text-sm font-bold focus:ring-0 resize-none overflow-hidden text-slate-800" 
                placeholder="(Final Notes for Employee Viewing)" 
                value={tempData.final_notes || ''} 
                onChange={e => {
                    setTempData({...tempData, final_notes: e.target.value});
                    e.target.style.height = 'auto';
                    e.target.style.height = e.target.scrollHeight + 'px';
                }}
                onFocus={(e) => {
                    e.target.style.height = 'auto';
                    e.target.style.height = e.target.scrollHeight + 'px';
                }}
              />
            </div>
          )}

          {/* Section 7: Manager Exemption and DOA Points*/}
          <div className = "grid grid-cols-2 gap-6 border-slate-100">
            {/* Pending DOA Points Box */}
            <div className={`p-4 rounded-2xl col-span-2 md:col-span-1 border transition-all flex items-center justify-between shadow-sm ${ isChanged('pendingDOAPoints') ? 'bg-rose-50 border-rose-200' : 'bg-sky-200/50 border-sky-200' }`}>
              <div className="flex items-center gap-3">
                <div className={`p-2 rounded-lg ${(tempData.pendingDOAPoints ?? 0) > 0 ? 'bg-amber-100 text-amber-600' : 'bg-blue-200 text-slate-500'}`}>
                  <AlertCircle className="h-4 w-4" />
                </div>
                <div>
                  <p className="text-[11px] font-black text-black uppercase leading-none">Pending DOA Points</p>
                  <p className="text-[9px] font-bold text-slate-400 uppercase tracking-tighter italic">Staged Penalty Points</p>
                </div>
              </div>
  
              <input 
                type="number" 
                min="0" 
                step="any"
                value={tempData.pendingDOAPoints ?? ''}
                onChange={e => {
                  const val = e.target.value; 
                  setTempData({...tempData, pendingDOAPoints: val === '' ? null : parseFloat(val)});}}
                className="w-20 h-10 text-center font-bold text-lg bg-white rounded-xl border-none shadow-inner focus:ring-2 focus:ring-blue-400"
                placeholder="-"
              />
            </div>

            {/* DOA Points Box */}
            <div className={`p-4 rounded-2xl col-span-2 md:col-span-1 border transition-all flex items-center justify-between shadow-sm ${ isChanged('DOAPoints') ? 'bg-rose-50 border-rose-200' : 'bg-sky-200/50 border-sky-200' }`}>
              <div className="flex items-center gap-3">
                <div className={`p-2 rounded-lg ${(tempData.DOAPoints ?? 0) > 0 ? 'bg-amber-100 text-amber-600' : 'bg-blue-200 text-slate-500'}`}>
                  <AlertCircle className="h-4 w-4" />
                </div>
                <div>
                  <p className="text-[11px] font-black text-black uppercase leading-none">DOA Points</p>
                  <p className="text-[9px] font-bold text-slate-400 uppercase tracking-tighter italic">Assign Penalty Points</p>
                </div>
              </div>
  
              <input 
                type="number" 
                min="0" 
                step="any"
                value={tempData.DOAPoints ?? ''} 
                onChange={e => {
                  const val = e.target.value; 
                  setTempData({...tempData, DOAPoints: val === '' ? null : parseFloat(val)});}}
                className="w-20 h-10 text-center font-bold text-lg bg-white rounded-xl border-none shadow-inner focus:ring-2 focus:ring-blue-400"
                placeholder="-"
              />
            </div>

            {/* Manager Exemption Box */}
            <div className={`p-4 rounded-2xl col-span-2 md:col-span-1 border transition-all flex items-center justify-between shadow-sm ${ isChanged('skipManagerApproval') ? 'bg-rose-50 border-rose-200' : 'bg-slate-200/50 border-slate-200' }`}>
              <div className="flex items-center gap-3">
                <div className={`p-2 rounded-lg ${tempData.skipManagerApproval ? 'bg-emerald-100 text-emerald-600' : 'bg-blue-200 text-slate-500'}`}><User className="h-4 w-4" /></div>
                <div><p className="text-[11px] font-black text-black uppercase leading-none">Manager Exemption</p><p className="text-[9px] font-bold text-slate-400 uppercase tracking-tighter italic">Skip Manager Approval</p></div>
              </div>
              <button type="button" onClick={() => setTempData({ ...tempData, skipManagerApproval: !tempData.skipManagerApproval, manager_approval: !tempData.skipManagerApproval ? 'not_required' : 'pending'})} className={`relative inline-flex h-6 w-11 items-center rounded-full transition-colors ${tempData.skipManagerApproval ? 'bg-emerald-500' : 'bg-slate-700'}`}><span className={`inline-block h-4 w-4 transform rounded-full bg-white transition-transform ${tempData.skipManagerApproval ? 'translate-x-6' : 'translate-x-1'}`} /></button>
            </div>

            {/* DAP Points Box */}
            <div className={`p-4 rounded-2xl col-span-2 md:col-span-1 border transition-all flex items-center justify-between shadow-sm ${ isChanged('DAP') ? 'bg-rose-50 border-rose-200' : 'bg-sky-200/50 border-sky-200' }`}>
              <div className="flex items-center gap-3">
                <div className={`p-2 rounded-lg ${(tempData.DAP ?? 0) > 0 ? 'bg-amber-100 text-amber-600' : 'bg-blue-200 text-slate-500'}`}>
                  <AlertCircle className="h-4 w-4" />
                </div>
                <div>
                  <p className="text-[11px] font-black text-black uppercase leading-none">DAP</p>
                  <p className="text-[9px] font-bold text-slate-400 uppercase tracking-tighter italic">DAP Points</p>
                </div>
              </div>
  
              <input 
                type="number" 
                min="0" 
                step="any"
                value={tempData.DAP ?? ''} 
                onChange={e => {
                  const val = e.target.value; 
                  setTempData({...tempData, DAP: val === '' ? null : parseFloat(val)});}}
                className="w-20 h-10 text-center font-bold text-lg bg-white rounded-xl border-none shadow-inner focus:ring-2 focus:ring-blue-400"
                placeholder="-"
              />
            </div>

            {/* Notes Box */}
            <div className={`col-span-2 space-y-1 p-3 rounded-2xl border transition-all shadow-sm ${ isChanged('manager_notes') ? 'bg-rose-50 border-rose-200' : 'bg-sky-200/50 border-sky-200' }`}>
              <label className="text-[10px] font-black text-black uppercase ml-1 flex items-center gap-1">Admin Notes</label>
              <textarea 
                rows={1}
                style={{ fieldSizing: 'content' } as React.CSSProperties}
                className="w-full bg-transparent border-none text-sm font-bold focus:ring-0 overflow-hidden" 
                placeholder="(Manager/HR Notes)" 
                value={tempData.manager_notes || ''} 
                onChange={e => setTempData({...tempData, manager_notes: e.target.value})}       
              />
            </div>
          </div>

          {/* Footer Actions */}
          <div className="flex gap-3 pt-4">
            {absence.status !== 'archived' ? (
              <button type="button" onClick={() => onArchive(absence.id)} className="p-4 text-rose-500 bg-rose-50 rounded-2xl shadow-xl hover:bg-rose-100 transition-colors">
                <Archive className="h-5 w-5"/>
              </button>
            ) : (
              <button type="button" onClick={() => onArchive(absence.id)} className="p-4 text-emerald-500 bg-emerald-50 rounded-2xl shadow-xl hover:bg-emerald-100 transition-colors">
                <ArchiveRestore className="h-5 w-5"/>
              </button>
            )}
            <button 
              type="submit" 
              disabled={isSaving || !hasChanges || isDateInvalid} 
              className={`flex-1 py-4 font-black rounded-2xl shadow-xl transition-all flex items-center justify-center gap-2 ${
                (!hasChanges  || isDateInvalid)
                  ? 'bg-slate-200 text-slate-400 cursor-not-allowed' 
                  : 'bg-indigo-600 text-white hover:bg-indigo-700'
              }`}
            >
              {isSaving ? "Syncing Database..." : hasChanges ? "Commit Changes" : "No Changes Detected"}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};