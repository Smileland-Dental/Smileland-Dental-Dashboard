"use client";

import React, { useState, useMemo } from 'react';
import { X, Plus, User, Search, Loader2, AlertCircle, ArrowRight, CheckCircle2 } from 'lucide-react';
import { db } from '@/lib/firebase.config';
import { collection, query, where, getDocs, addDoc } from "firebase/firestore";
import { AbsenceRequest } from "@/lib/types";
import FeedbackModal from '@/components/ui/FeedbackModal';
import { OFFICES } from '@/lib/constants';

const incidentTypes = ["Late In", "Early Out", "Absent", "Leave and Come Back", "Long Lunch", "Switch Shift"];
//const officeLocations = ["Corporate", "Ming", "Bernard", "California", "Ortho", "Delano", "Tulare", "Visalia", "Fresno"];

interface HRCreateAbsenceModalProps {
  isOpen: boolean;
  onClose: () => void;
  onSave: () => void;
}

// For the Reset when closing
const initialState = {
  office: '',
  type_of_incident: '',
  type_of_request: 'HR Call In',
  employee_id: '',
  employee_name: '',
  employee_title: '',
  employeeFirestoreID: '',
  skipManagerApproval: false,
  employee_comments: '',
  eta: '',
  etd: '',
  excuse_note: [],
  excuse_note_submitted: 'pending',
  final_approval: 'pending',
  final_approval_name: '',
  incident_start: '',
  incident_end: '',
  manager_approval: 'pending',
  manager_approval_name: '',
  manager_notes: '',
  DOAPoints: 0,
  pendingDOAPoints: 0,
  final_notes: '',
  DAP: 0,
  status: 'active',
};


export const HRCreateAbsenceModal = ({ isOpen, onClose, onSave }: HRCreateAbsenceModalProps) => {
  const [isSearching, setIsSearching] = useState(false);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [employeeExistingRequests, setEmployeeExistingRequests] = useState<AbsenceRequest[]>([]);
  
  // Split state into Employee Data and Incident Data
  const [selectedEmployee, setSelectedEmployee] = useState<any>(null);
  const [tempID, setTempID] = useState("");
  
  /*const [formData, setFormData] = useState({
    employee_comments: '',

    eta: '',
    etd: '',
    excuse_note: [],
    excuse_note_submitted: 'pending',
    final_approval: 'pending',
    final_approval_name: '', 
    incident_start: '',
    incident_end: '',
    manager_approval: 'pending',
    manager_approval_name: '',
    manager_notes: '',

    office: '',
    type_of_incident: '',
    type_of_request: '',

    employee_id: '',
    employee_name: '',
    employee_title: '',
    employeeFirestoreID: '',
    skipManagerApproval: false,
  });*/

  const [formData, setFormData] = useState(initialState);

  const [feedback, setFeedback] = useState({ isOpen: false, type: 'success' as 'success' | 'error', message: '' });

  const isOverlapping = useMemo(() => {
    if (!formData.type_of_incident || !formData.incident_start || !formData.incident_end) {
      return false;
    }

    const formStart = new Date(formData.incident_start + "T00:00:00").getTime();
    const formEnd = new Date(formData.incident_end + "T00:00:00").getTime();

    // Look through array instances currently stored in parent state
    return employeeExistingRequests.some((absence) => {
      if (absence.type_of_incident !== formData.type_of_incident) {
        return false;
      }
      const isConflictingStatus = ['active', 'pending_action'].includes(absence.status);
      if (!isConflictingStatus) {
        return false; // Skips 'archived'
      }

      const existingStart = new Date(absence.incident_start + "T00:00:00").getTime();
      const existingEnd = new Date(absence.incident_end + "T00:00:00").getTime();

      // Check date interval overlap logic: (StartA <= EndB) AND (EndA >= StartB)
      return formStart <= existingEnd && formEnd >= existingStart;
    });
  }, [formData.type_of_incident, formData.incident_start, formData.incident_end, employeeExistingRequests]);

  // Date Validation Logic
  const isDateInvalid = Boolean(
    formData.incident_start && 
    formData.incident_end && 
    new Date(formData.incident_end) < new Date(formData.incident_start)
  );

  // Step 1: Verification Logic
  const handleLookupEmployee = async () => {
    if (!tempID) return;
    setIsSearching(true);
    setError(null);

    try {
      const q = query(collection(db, "employees"), where("employeeID", "==", tempID));
      const querySnapshot = await getDocs(q);

      if (!querySnapshot.empty) {
        const empDoc = querySnapshot.docs[0];
        const data = empDoc.data();
        setSelectedEmployee({ id: empDoc.id, ...data });
        // Pre-fill form with employee context
        setFormData(prev => ({
          ...prev,
          employee_id: tempID,
          employee_name: `${data.firstName} ${data.lastName}`,
          employee_title: data.jobTitle || '',
          employeeFirestoreID: empDoc.id,
          skipManagerApproval: data.skipManagerApproval || false,
        }));

        const employeeAbsences = query(
          collection(db, "absences"), 
          where("employee_id", "==", tempID),
          where("status", "in", ["active", "pending_action", "cancellation_requested"])
        );
        const absenceSnapshot = await getDocs(employeeAbsences);
        const exisitngAbsences = absenceSnapshot.docs.map(doc => ({
          id: doc.id,
          ...doc.data()
        })) as AbsenceRequest[];

        setEmployeeExistingRequests(exisitngAbsences);

      } else {
        setError("Invalid Employee ID. Please check and try again.");
      }
    } catch (err) {
      setError("Database connection error.");
    } finally {
      setIsSearching(false);
    }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.employeeFirestoreID) {
      setError("Please verify an employee first.");
      return;
    }

    if (isDateInvalid) {
      setError("End date cannot be before start date.");
      return;
    }
    if (isOverlapping) {
      setError("A record for this incident type already exists within this date range.");
      return;
    }

    setIsSubmitting(true);
    try {
      const isSpecialCallIn = ["HR Call In", "No Call", "Call In After Shift"].includes(formData.type_of_request);

      let finalManagerApproval = formData.manager_approval;
      if (isSpecialCallIn) {
        finalManagerApproval = "not_required";
      } 
      else {
        // For Incident Notice or Time Off Request:
        // Default to 'pending' UNLESS the employee profile explicitly skips manager approval
        finalManagerApproval = selectedEmployee?.skipManagerApproval ? "not_required" : "pending";
      }

      const finalStatus = isSpecialCallIn ? "pending_action" : "active";

      await addDoc(collection(db, "absences"), {
        ...formData,
        createdAt: new Date(),
        manager_approval: finalManagerApproval,
        status: finalStatus
      });
      setFeedback({ isOpen: true, type: 'success', message: "Record synced successfully." });
    } catch (err) {
      console.error("Submit error:", err);
      setFeedback({ isOpen: true, type: 'error', message: "Failed to save record." });
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>) => {
    const { name, value } = e.target;

    setFormData(prev => ({
      ...prev,
      [name]: value,
      // Logic from your 2nd form: Reset times if incident type changes
      ...(name === 'type_of_incident' ? { eta: '', etd: '' } : {})
    }));
  };

  const handleReset = () => {
    setFormData(initialState);
    setSelectedEmployee(null);
    setEmployeeExistingRequests([]);
    setTempID("");
    setError(null);
  };

  const handleClose = () => {
    handleReset();
    onClose();
  };

  if (!isOpen) return null;

  return (
    <>
      <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-md z-[60] flex items-center justify-center p-4">
        <div className="bg-white w-full max-w-xl max-h-[90vh] flex flex-col rounded-[2.5rem] shadow-2xl overflow-hidden transition-all">
          
          {/* Header */}
          <div className="p-6 bg-slate-900 text-white flex justify-between items-center shrink-0">
            <h2 className="text-xl font-black flex items-center gap-2">
              HR Absence Entry
            </h2>
            <button onClick={handleClose} className="p-2 hover:bg-white/10 rounded-full transition-all"><X /></button>
          </div>

          <div className="p-8 space-y-6 overflow-y-auto">
            {/* STEP 1: EMPLOYEE SEARCH */}
            {!selectedEmployee ? (
              <div className="space-y-4 animate-in fade-in slide-in-from-bottom-4">
                <div className="text-center space-y-1">
                  <h3 className="text-lg font-bold text-slate-800">Employee Search</h3>
                  <p className="text-xs text-slate-500">Enter the Employee ID for employee receiving absence record.</p>
                </div>
                <div className="flex gap-2">
                  <input 
                    autoFocus
                    placeholder="Employee ID #"
                    className="flex-1 p-4 bg-slate-100 rounded-2xl border-2 border-transparent focus:border-indigo-500 focus:bg-white outline-none font-bold text-center text-xl transition-all"
                    value={tempID}
                    onChange={(e) => setTempID(e.target.value)}
                    onKeyDown={(e) => e.key === 'Enter' && handleLookupEmployee()}
                  />
                  <button 
                    onClick={handleLookupEmployee}
                    disabled={isSearching || !tempID}
                    className="p-4 bg-indigo-600 text-white rounded-2xl font-bold hover:bg-indigo-700 disabled:bg-slate-200 transition-all"
                  >
                    {isSearching ? <Loader2 className="animate-spin" /> : <ArrowRight />}
                  </button>
                </div>
                {error && <p className="text-center text-xs font-bold text-rose-500">{error}</p>}
              </div>
            ) : (
              /* STEP 2: INCIDENT DETAILS */
              <form onSubmit={handleSubmit} className="space-y-5 animate-in zoom-in-95 duration-300">
                {/* Profile Card */}
                <div className="flex items-center justify-between p-4 bg-indigo-50 rounded-2xl border border-indigo-100">
                  <div className="flex items-center gap-3">
                    <div className="h-10 w-10 bg-indigo-600 rounded-full flex items-center justify-center text-white font-bold">
                      {selectedEmployee.firstName[0]}
                    </div>
                    <div>
                      <p className="text-sm font-black text-slate-900">{selectedEmployee.firstName} {selectedEmployee.lastName}</p>
                      <p className="text-[10px] font-bold text-indigo-600 uppercase">ID: {selectedEmployee.employeeID} • {selectedEmployee.jobTitle}</p>
                    </div>
                  </div>
                  <button 
                    type="button" 
                    onClick={handleReset}
                    className="text-[10px] font-bold text-slate-400 hover:text-rose-500 uppercase underline"
                  >
                    Change
                  </button>
                </div>

                <div className="grid grid-cols-1 gap-4">
                  <div className="space-y-1">
                    <label className="text-[10px] font-black text-slate-400 uppercase ml-1">Type of Request</label>
                    <div className="w-full p-4 bg-slate-50 rounded-xl border-none grid grid-cols-3 gap-4 text-xs font-bold">
                    {["HR Call In", "Incident Notice", "Time Off Request", "No Call", "Call In After Shift"].map((type) => (
                      <label key={type} className="flex items-center gap-1.5 cursor-pointer select-none">
                        <input 
                        type="radio" 
                        name="type_of_request" 
                        value={type} 
                        checked={formData.type_of_request === type} 
                        onChange={handleChange}
                        className="accent-indigo-600 h-3.5 w-3.5 cursor-pointer" 
                      />
                      <span>{type}</span>
                    </label>
                    ))}
                    </div>
                  </div>
                </div>

                {/* Overlap / Duplicate Warning */}
                {isOverlapping && (
                  <div className="flex items-center gap-2 p-3 bg-amber-50 rounded-xl text-amber-700 text-xs font-bold border border-amber-200 animate-in fade-in">
                    <AlertCircle className="w-4 h-4 shrink-0 text-amber-600" />
                    <span>An active or pending record for "{formData.type_of_incident}" already exists in this date range.</span>
                  </div>
                )}

                {/* Form Fields */}
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <label className="text-[10px] font-black text-slate-400 uppercase ml-1">Incident Type</label>
                    <select required name="type_of_incident" className="w-full p-3 bg-slate-50 rounded-xl border-none font-bold text-sm" onChange={handleChange}>
                      <option value="">Select Type</option>
                      {incidentTypes.map(t => <option key={t} value={t}>{t}</option>)}
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] font-black text-slate-400 uppercase ml-1">Location</label>
                    <select required name="office" className="w-full p-3 bg-slate-50 rounded-xl border-none font-bold text-sm" onChange={handleChange}>
                      <option value="">Select Location</option>
                      {OFFICES.map(o => <option key={o} value={o}>{o}</option>)}
                    </select>
                  </div>
                </div>

                {(formData.type_of_incident === "Late In") && (
                  <div className="grid grid-cols-2 gap-4 animate-in fade-in slide-in-from-top-2">
                    <div className="space-y-1">
                      <label className="text-[10px] font-black text-slate-400 uppercase ml-1">Estimated Arrival (ETA)</label>
                      <input 
                        name="eta"
                        type="time" 
                        required 
                        className="w-full p-3 bg-slate-50 rounded-xl border-none font-bold text-sm focus:ring-2 focus:ring-indigo-500" 
                        value={formData.eta}
                        onChange={handleChange} 
                      />
                    </div>
                  </div>
                )}

                {(formData.type_of_incident === "Early Out") && (
                  <div className="grid grid-cols-2 gap-4 animate-in fade-in slide-in-from-top-2">
                    <div className="space-y-1">
                      <label className="text-[10px] font-black text-slate-400 uppercase ml-1">Estimated Departure (ETD)</label>
                      <input 
                        name="etd"
                        type="time" 
                        required 
                        className="w-full p-3 bg-slate-50 rounded-xl border-none font-bold text-sm focus:ring-2 focus:ring-indigo-500" 
                        value={formData.etd}
                        onChange={handleChange} 
                      />
                    </div>
                  </div>
                )}

                {(formData.type_of_incident === "Leave and Come Back") && (
                  <div className="grid grid-cols-2 gap-4 animate-in fade-in slide-in-from-top-2">
                    <div className="space-y-1">
                      <label className="text-[10px] font-black text-slate-400 uppercase ml-1">Estimated Departure (ETD)</label>
                      <input 
                        name="etd"
                        type="time" 
                        required 
                        className="w-full p-3 bg-slate-50 rounded-xl border-none font-bold text-sm focus:ring-2 focus:ring-indigo-500" 
                        value={formData.etd}
                        onChange={handleChange} 
                      />
                    </div>
                    <div className="space-y-1">
                      <label className="text-[10px] font-black text-slate-400 uppercase ml-1">Estimated Arrival (ETA)</label>
                      <input 
                        name="eta"
                        type="time" 
                        required 
                        className="w-full p-3 bg-slate-50 rounded-xl border-none font-bold text-sm focus:ring-2 focus:ring-indigo-500" 
                        value={formData.eta}
                        onChange={handleChange} 
                      />
                    </div>
                  </div>
                )}

                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <label className="text-[10px] font-black text-slate-400 uppercase ml-1">Start Date</label>
                    <input name="incident_start" type="date" required value={formData.incident_start} className="w-full p-3 bg-slate-50 rounded-xl border-none font-bold text-sm" onChange={handleChange} />
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] font-black text-slate-400 uppercase ml-1">End Date</label>
                    <input name="incident_end" type="date" required min={formData.incident_start || ""} value={formData.incident_end} className={`w-full p-3 rounded-xl border-none font-bold text-sm ${
                        isDateInvalid ? 'bg-rose-50 text-rose-600 ring-2 ring-rose-500' : 'bg-slate-50'
                      }`} onChange={handleChange} />
                  </div>
                </div>

                {/* Date Validation Error Message */}
                {isDateInvalid && (
                  <div className="flex items-center gap-2 p-3 bg-rose-50 rounded-xl text-rose-600 text-xs font-bold border border-rose-200 animate-in fade-in">
                    <AlertCircle className="w-4 h-4 shrink-0" />
                    <span>End date cannot be earlier than the start date.</span>
                  </div>
                )}

                <div className="space-y-1">
                  <label className="text-[10px] font-black text-slate-400 uppercase ml-1">HR/Manager Comments</label>
                  <textarea rows={3} name="manager_notes" className="w-full p-3 bg-slate-50 rounded-xl border-none font-bold text-sm" onChange={handleChange} />
                </div>

                <button 
                  type="submit" 
                  disabled={isSubmitting || isDateInvalid || isOverlapping}
                  className="w-full py-4 bg-slate-900 text-white font-black rounded-2xl shadow-xl hover:bg-slate-800 disabled:bg-slate-200 transition-all flex items-center justify-center gap-2"
                >
                  {isSubmitting ? <Loader2 className="h-5 w-5 animate-spin" /> : <CheckCircle2 className="h-5 w-5" />}
                  Submit Absence Record
                </button>
              </form>
            )}
          </div>
        </div>
      </div>

      <FeedbackModal 
        isOpen={feedback.isOpen} 
        type={feedback.type} 
        message={feedback.message} 
        onClose={() => {
          setFeedback(f => ({ ...f, isOpen: false }));
          if (feedback.type === 'success') {
            onSave();
            handleClose();
          }
        }} 
      />
    </>
  );
};