'use client';

import React, { useState } from 'react';
import { db } from '@/lib/firebase.config';
import { collection, addDoc, Timestamp, query, where, getDocs, serverTimestamp } from 'firebase/firestore';
import { X, Loader2, CheckCircle2, Clock, ArrowRight } from 'lucide-react';
import { OFFICES } from '@/lib/constants';
import FeedbackModal from '@/components/ui/FeedbackModal';
import { VolunteerEarlyOutRequest } from '@/lib/types';

interface EarlyOutFormProps {
  onClose: () => void;
}

export const EarlyOutForm = ({ onClose }: EarlyOutFormProps) => {
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [isSearching, setIsSearching] = useState(false);
  const [tempID, setTempID] = useState("");
  const [selectedEmployee, setSelectedEmployee] = useState<any | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [feedback, setFeedback] = useState({ isOpen: false, type: 'success' as 'success' | 'error', message: '' });

  const [isVolunteerChecked1, setIsVolunteerChecked1] = useState(false);
  const [isVolunteerChecked2, setIsVolunteerChecked2] = useState(false);

  const [supervisorID, setSupervisorID] = useState('');
  const [supervisorEmployeeName, setSupervisorEmployeeName] = useState<string | null>(null);
    const [supervisorEmployeeID, setSupervisorEmployeeID] = useState<string | null>(null);
  const [isVerifying, setIsVerifying] = useState(false);
  const [supervisorError, setSupervisorError] = useState<string | null>(null);

  const [formData, setFormData] = useState<Partial<VolunteerEarlyOutRequest>>({
    employeeFirestoreID: '',
    employee_id: '',
    employee_name: '',
    employee_title: '',
    office: '',
    type_of_request: 'Volunteer Early Out',
    supervisor_id: '',
    supervisor_name: '',
    incident_date: '',
    incident_time: '',
    //createdAt: Timestamp,
    //id: '',
  });

  // Drop this at the bottom of your file to drive the ticking clock efficiently
  const LiveTime = () => {
    const [time, setTime] = React.useState('');

    React.useEffect(() => {
      // Set immediate initial time
      setTime(new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' }));

      const timer = setInterval(() => {
        setTime(new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' }));
      }, 1000);

      return () => clearInterval(timer);
    }, []);

    return <>{time}</>;
  };

  const handleLookupEmployee = async () => {
    const cleanTempID = tempID.trim();
    if (!cleanTempID) return;
    setIsSearching(true);
    setError(null);

    try {
      const q = query(collection(db, "employees"), where("employeeID", "==", cleanTempID));
      const querySnapshot = await getDocs(q);

      if (!querySnapshot.empty) {
        const empDoc = querySnapshot.docs[0];
        const data = empDoc.data();
        setSelectedEmployee({ id: empDoc.id, ...data });
      
        // Pre-fill form with employee context
        setFormData(prev => ({
          ...prev,
          employee_id: cleanTempID,
          employee_name: `${data.firstName} ${data.lastName}`,
          employee_title: data.jobTitle || '',
          employeeFirestoreID: empDoc.id,
          // 💡 Default the office selection to the employee's office record
          office: data.office || '', 
        }));
      } 
      else {
        setError("Invalid Employee ID. Please check and try again.");
        setSelectedEmployee(null);
      }
    } catch (err) {
      setError("Database connection error.");
    } finally {
      setIsSearching(false);
    }
  };

  const handleLookupSupervisor = async () => {
    const cleanSupervisorID = supervisorID.trim();
    if (!cleanSupervisorID) return;

    // 💡 Check if the verifier ID matches the requesting employee's ID
    if (cleanSupervisorID === formData.employee_id) {
      setSupervisorError("The requesting employee cannot verify their own departure.");
      setSupervisorEmployeeName(null);
      return;
    }
    setIsVerifying(true);
    setSupervisorError(null);

    try {
      const q = query(collection(db, "employees"), where("employeeID", "==", cleanSupervisorID));
      const querySnapshot = await getDocs(q);

      if (!querySnapshot.empty) {
        const empDoc = querySnapshot.docs[0];
        const data = empDoc.data();
        setSupervisorEmployeeID(data.employeeID);
        setSupervisorEmployeeName(`${data.firstName} ${data.lastName}`);
      } 
      else {
        setSupervisorError("Invalid Supervisor Employee ID. Please check and try again.");
        setSupervisorEmployeeName(null);
      }
    } catch (err) {
      setSupervisorError("Database connection error.");
    } finally {
      setIsVerifying(false);
    }
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>) => {
    const { name, value } = e.target;
    setFormData(prev => ({ ...prev, [name]: value }));
  };

  const handleResetEmployee = () => {
    setTempID('');
    setSelectedEmployee(null);
    setError(null);

    // Clear checkboxes
    setIsVolunteerChecked1(false);
    setIsVolunteerChecked2(false);

    // Clear witness/verifier state
    setSupervisorID('');
    setSupervisorEmployeeName(null);
    setSupervisorError(null);
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.employeeFirestoreID) {
      setError("Please verify an employee first.");
      return;
    }

    if (!supervisorEmployeeName) {
      setSupervisorError("Supervisor verification is required.");
      return;
    }

    setError(null);
    setIsSubmitting(true);

    try {
      const now = new Date();

      const currentDateString = now.toLocaleDateString('en-CA');

      // Format current time matching typical 'HH:MM' input shapes
      const currentTimeString = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: false });

      await addDoc(collection(db, "volunteer-early-outs"), {
        ...formData,
        supervisor_id: supervisorEmployeeID,
        supervisor_name: supervisorEmployeeName, 
        incident_date: currentDateString,       
        incident_time: currentTimeString,
        createdAt: serverTimestamp(),
      });
      setFeedback({ isOpen: true, type: 'success', message: "Volunteer early departure recorded successfully." });
    } catch (err) {
      console.error("Submit error:", err);
      setFeedback({ isOpen: true, type: 'error', message: "Failed to save early departure record." });
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <>
      <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-md z-[60] flex items-center justify-center p-4">
        <div className="bg-white w-full max-w-xl rounded-[2.5rem] shadow-2xl overflow-hidden transition-all">
          {/* Header */}
          <div className="p-6 bg-slate-900 text-white flex justify-between items-center">
            <h2 className="text-xl font-black flex items-center gap-2">
              <Clock className="h-5 w-5 text-indigo-400" />
              Volunteer Early Departure Form
            </h2>
            <button onClick={onClose} className="p-2 hover:bg-white/10 rounded-full transition-all"><X /></button>
          </div>

          <div className="p-8 space-y-6">  
            {/* STAGE 1: Search Employee Field */}
            {!selectedEmployee ? (
              <div className="space-y-4 animate-in fade-in slide-in-from-bottom-4">
                <div className="text-center space-y-1">
                  <h3 className="text-lg font-bold text-slate-800">Employee Search</h3>
                  <p className="text-xs text-slate-500">Enter the Employee ID for employee leaving early.</p>
                </div>
                <div className="flex gap-2">
                  <input 
                    autoFocus
                    type="password"
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
              /* STAGE 2: Employee Verified -> Display Form Fields */
              <form onSubmit={handleSubmit} className="space-y-5 animate-in zoom-in-95 duration-300">
                {/* Contextual Employee Profile Card */}
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
                    onClick={handleResetEmployee}
                    className="text-[10px] font-bold text-slate-400 hover:text-rose-500 uppercase underline"
                  >
                    Change
                  </button>
                </div>

                {/* Form Input Grid */}
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-1 col-span-1">
                    <label className="text-[10px] font-black text-slate-400 uppercase ml-1">Location / Office</label>
                    <select 
                      required 
                      name="office" 
                      className="w-full p-3 bg-slate-50 rounded-xl border-none font-bold text-sm focus:ring-2 focus:ring-indigo-500" 
                      value={formData.office}
                      onChange={handleChange}
                    >
                      <option value="">Select Location</option>
                      {OFFICES.map(o => <option key={o} value={o}>{o}</option>)}
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] font-black text-slate-400 uppercase ml-1">Current Timestamp</label>
                    <div className="w-full p-3 bg-slate-50 text-slate-700 rounded-xl font-mono font-bold border border-slate-100 flex items-center justify-between">
                      <span>{new Date().toLocaleDateString()}</span>
                      {/*<span className="text-indigo-600 animate-pulse bg-indigo-50 px-2 py-0.5 rounded-md">*/}
                        <LiveTime />
                      {/*</span>*/}
                    </div>
                  </div>
                </div>

                <div className="space-y-1">
                  <p className="font-bold">I voluntarily request to leave work early on {<span className="text-rose-600">{new Date().toLocaleDateString()}</span>} because my assigned work tasks have been completed.</p>
                  <input 
                    type="checkbox" 
                    id="confirm" 
                    checked={isVolunteerChecked1} 
                    onChange={(e) => setIsVolunteerChecked1(e.target.checked)} 
                    className=" text-blue-600 rounded border-gray-300"
                    required 
                  />
                  <label htmlFor="confirm" className="pl-1 text-s text-gray-600 leading-tight">
                    I acknowledge and agree to the above statement.
                  </label>
                  <br/>
                  <input 
                    type="checkbox" 
                    id="confirm2" 
                    checked={isVolunteerChecked2} 
                    onChange={(e) => setIsVolunteerChecked2(e.target.checked)} 
                    className=" text-blue-600 rounded border-gray-300"
                    required 
                  />
                  <label htmlFor="confirm2" className="pl-1 text-s text-gray-600 leading-tight">
                    I acknowledge that I am volunarily requesting to leave work early because my assigned tasks have been completed. I understand this request is voluntary and subject to management approval.
                  </label>
                </div>

                <div className="space-y-1">
                  {/* 💡 STEP 3: Conditional Verifier Input Box */}
                  {isVolunteerChecked1 && isVolunteerChecked2 && (
                    <div className="p-4 bg-slate-50 border border-slate-200 rounded-2xl space-y-3 animate-in fade-in slide-in-from-bottom-2">
                      <div>
                        <label className="text-[10px] font-black text-slate-500 uppercase">Supervisor ID</label>
                        <p className="text-xs text-slate-400">Supervisor must verify this request.</p>
                      </div>
                      
                      <div className="flex gap-2">
                        <input 
                          placeholder="Supervisor Employee ID #"
                          className="flex-1 p-3 bg-white rounded-xl border border-slate-200 focus:border-indigo-500 outline-none font-bold text-sm transition-all"
                          value={supervisorID}
                          type="password"
                          disabled={Boolean(supervisorEmployeeName)}
                          onChange={(e) => setSupervisorID(e.target.value)}
                          onKeyDown={(e) => e.key === 'Enter' && (e.preventDefault(), handleLookupSupervisor())}
                        />
                        {!supervisorEmployeeName ? (
                          <button 
                            type="button"
                            onClick={handleLookupSupervisor}
                            disabled={isVerifying || !supervisorID.trim()}
                            className="px-4 bg-indigo-600 text-white rounded-xl font-bold text-xs hover:bg-indigo-700 disabled:bg-slate-200 transition-all"
                          >
                            {isVerifying ? <Loader2 className="animate-spin h-4 w-4" /> : 'Verify'}
                          </button>
                        ) : (
                          <button 
                            type="button"
                            onClick={() => { setSupervisorID(''); setSupervisorEmployeeName(null); }}
                            className="px-4 bg-rose-50 text-rose-600 rounded-xl font-bold text-xs hover:bg-rose-100 transition-all"
                          >
                            Clear
                          </button>
                        )}
                      </div>

                      {supervisorEmployeeName && (
                        <p className="text-xs font-bold text-emerald-600 flex items-center gap-1">
                          ✓ Confirmed: {supervisorEmployeeName}
                        </p>
                      )}
                      {supervisorError && <p className="text-xs font-bold text-rose-500">{supervisorError}</p>}
                    </div>
                  )}
                </div>

                <button 
                  type="submit" 
                  disabled={isSubmitting || !formData.office || !isVolunteerChecked1 || !isVolunteerChecked2 || !supervisorEmployeeName}
                  className="w-full py-4 bg-slate-900 text-white font-black rounded-2xl shadow-xl hover:bg-slate-800 disabled:bg-slate-200 transition-all flex items-center justify-center gap-2"
                >
                  {isSubmitting ? <Loader2 className="h-5 w-5 animate-spin" /> : <CheckCircle2 className="h-5 w-5" />}
                  Log Volunteer Departure
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
            onClose();
          }
        }} 
      />
    </>
  );
};