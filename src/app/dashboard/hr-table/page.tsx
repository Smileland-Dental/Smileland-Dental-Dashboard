'use client';

import React, { useState, useEffect, useMemo } from "react";
import { collection, getDocs, doc, writeBatch, addDoc, deleteDoc, Timestamp } from 'firebase/firestore';
import { db } from '@/lib/firebase.config';
import { AbsenceRequest } from "@/lib/types";
import * as XLSX from 'xlsx';
import { 
  Search, 
  Filter, 
  Download, 
  Plus, 
  Save, 
  Trash2, 
  Check, 
  X, 
  Calendar,
  FileText
} from 'lucide-react';

// --- Initial State for New Entries matching your new Type ---
const initialNewAbsenceState: Omit<AbsenceRequest, 'id'> = {
  employeeFirestoreID: '',
  employee_id: '',
  employee_name: '',
  employee_title: '',
  office: '',
  type_of_incident: '',
  incident_start: '',
  incident_end: '',
  manager_approval: 'pending',
  final_approval: 'pending',
  excuse_note_submitted: 'not_provided',
  eta: "",
  etd: "",
  excuse_note: [],
  createdAt: Timestamp.now(),
  type_of_request: "",
  employee_comments: "",
  manager_notes: "",
  manager_approval_name: "",
  final_approval_name: ""
};

const incidentTypes = ["Late In", "Early Out", "Absent", "Leave and Come Back", "Long Lunch", "Switch Shift", "Cancel Cell"];

export default function ManagementPage() {
  const [absences, setAbsences] = useState<AbsenceRequest[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [modifiedDocs, setModifiedDocs] = useState<Set<string>>(new Set());
  const [isSaving, setIsSaving] = useState(false);
  const [isAdding, setIsAdding] = useState(false);
  const [newAbsence, setNewAbsence] = useState<Omit<AbsenceRequest, 'id'>>(initialNewAbsenceState);
  const [isCreating, setIsCreating] = useState(false);
  const [feedbackMessage, setFeedbackMessage] = useState<{type: 'success' | 'error', text: string} | null>(null);
  
  const [filterStartDate, setFilterStartDate] = useState('');
  const [filterEndDate, setFilterEndDate] = useState('');
  const [isExportModalOpen, setIsExportModalOpen] = useState(false);
  const [selectedColumns, setSelectedColumns] = useState<Record<string, boolean>>({
    employee_name: true, employee_id: true, office: true, type_of_incident: true, 
    incident_start: true, incident_end: true, manager_approval: true, final_approval: true,
    manager_notes: true, employee_comments: false, employee_title: true
  });

  useEffect(() => {
    if (feedbackMessage) {
      const timer = setTimeout(() => setFeedbackMessage(null), 4000);
      return () => clearTimeout(timer);
    }
  }, [feedbackMessage]);

  const handleFetch = async () => {
    setLoading(true);
    setError(null);
    try {
      const querySnapshot = await getDocs(collection(db, "absences"));
      const fetched = querySnapshot.docs.map(doc => ({
        id: doc.id,
        ...(doc.data() as Omit<AbsenceRequest, 'id'>),
      }));
      setAbsences(fetched);
    } catch (err) {
      setError("Failed to fetch data. Check permissions.");
    } finally {
      setLoading(false);
    }
  };

  const filteredData = useMemo(() => {
    if (!filterStartDate || !filterEndDate) return absences;
    return absences.filter(absence => {
      const absenceStart = new Date(absence.incident_start);
      const absenceEnd = new Date(absence.incident_end);
      const filterStart = new Date(filterStartDate);
      const filterEnd = new Date(filterEndDate);
      return filterStart <= absenceEnd && absenceStart <= filterEnd;
    });
  }, [absences, filterStartDate, filterEndDate]);

  const handleInputChange = (id: string, field: keyof AbsenceRequest, value: any) => {
    setAbsences(curr => curr.map(a => a.id === id ? { ...a, [field]: value, updatedAt: Timestamp.now() } : a));
    setModifiedDocs(prev => new Set(prev).add(id));
  };

  const handleSaveChanges = async () => {
    setIsSaving(true);
    const batch = writeBatch(db);
    modifiedDocs.forEach(docId => {
      const data = absences.find(a => a.id === docId);
      if (data) {
        const { id, ...dataToSave } = data;
        batch.update(doc(db, "absences", docId), dataToSave);
      }
    });
    try {
      await batch.commit();
      setFeedbackMessage({ type: 'success', text: 'Changes saved!' });
      setModifiedDocs(new Set());
    } catch (err) {
      setFeedbackMessage({ type: 'error', text: 'Save failed.' });
    } finally { setIsSaving(false); }
  };

  const handleCreateAbsence = async () => {
    if (!newAbsence.employee_id || !newAbsence.employee_name) return alert("Required fields missing");
    setIsCreating(true);
    try {
      await addDoc(collection(db, "absences"), { ...newAbsence, createdAt: Timestamp.now() });
      setFeedbackMessage({ type: 'success', text: 'Added!' });
      setIsAdding(false);
      setNewAbsence(initialNewAbsenceState);
      handleFetch();
    } catch (err) { setFeedbackMessage({ type: 'error', text: 'Add failed.' }); }
    finally { setIsCreating(false); }
  };

  const handleDeleteAbsence = async (id: string) => {
    if (!window.confirm("Delete record?")) return;
    try {
      await deleteDoc(doc(db, "absences", id));
      setAbsences(curr => curr.filter(a => a.id !== id));
      setFeedbackMessage({ type: 'success', text: 'Deleted!' });
    } catch (err) { setFeedbackMessage({ type: 'error', text: 'Delete failed.' }); }
  };

  const handleExport = () => {
    const activeColumns = Object.entries(selectedColumns).filter(([_, v]) => v).map(([k]) => k);
    const dataToExport = filteredData.map(a => {
      const row: any = {};
      activeColumns.forEach(k => row[k] = (a as any)[k]);
      return row;
    });
    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Absences");
    XLSX.writeFile(wb, "EmployeeAbsences.xlsx");
    setIsExportModalOpen(false);
  };

  return (
    <main className="max-w-full overflow-hidden p-4 md:p-8 bg-gray-50 min-h-screen">
      <div className="max-w-7xl mx-auto space-y-6">
        <div>
          <h1 className="text-3xl font-extrabold text-gray-900">Absence Management</h1>
          <p className="text-gray-500 mt-1">Review and synchronize employee absence records.</p>
        </div>

        {/* --- Top Control Bar --- */}
        <div className="flex flex-col md:flex-row gap-4 items-center justify-between bg-white p-4 rounded-xl shadow-sm border border-gray-200">
          <div className="flex flex-wrap gap-2 w-full md:w-auto">
            <button onClick={handleFetch} disabled={loading} className="flex items-center gap-2 px-4 py-2 bg-white border border-gray-300 rounded-lg text-sm font-semibold hover:bg-gray-50 transition-all">
              <Search className={`h-4 w-4 ${loading ? 'animate-spin' : ''}`} />
              Sync Data
            </button>
            <button onClick={() => setIsAdding(true)} className="flex items-center gap-2 px-4 py-2 bg-indigo-600 text-white rounded-lg text-sm font-semibold hover:bg-indigo-700 transition-all shadow-md shadow-indigo-100">
              <Plus className="h-4 w-4" />
              Add Record
            </button>
            <button onClick={() => setIsExportModalOpen(true)} className="flex items-center gap-2 px-4 py-2 bg-teal-600 text-white rounded-lg text-sm font-semibold hover:bg-teal-700 transition-all shadow-md shadow-teal-100">
              <Download className="h-4 w-4" />
              Export
            </button>
          </div>

          {modifiedDocs.size > 0 && (
            <button onClick={handleSaveChanges} disabled={isSaving} className="flex items-center gap-2 px-6 py-2 bg-green-600 text-white rounded-lg text-sm font-bold animate-pulse">
              <Save className="h-4 w-4" />
              Save {modifiedDocs.size} Changes
            </button>
          )}
        </div>

        {/* --- Filter Bar --- */}
        <div className="flex flex-col md:flex-row items-center gap-4 p-4 bg-white rounded-xl border border-gray-200 shadow-sm">
          <div className="flex items-center gap-2 text-sm font-medium text-gray-700">
            <Filter className="h-4 w-4 text-indigo-500" />
            <span>Date Filter:</span>
          </div>
          <div className="flex items-center gap-2">
            <input type="date" value={filterStartDate} onChange={e => setFilterStartDate(e.target.value)} className="text-sm border-gray-300 rounded-md focus:ring-indigo-500" />
            <span className="text-gray-400">—</span>
            <input type="date" value={filterEndDate} onChange={e => setFilterEndDate(e.target.value)} className="text-sm border-gray-300 rounded-md focus:ring-indigo-500" />
            <button onClick={() => { setFilterStartDate(''); setFilterEndDate(''); }} className="text-xs text-gray-500 hover:text-indigo-600 font-medium">Clear</button>
          </div>
          <div className="md:ml-auto text-sm text-gray-500 bg-gray-50 px-3 py-1 rounded-full border">
            {filteredData.length} records found
          </div>
        </div>

        {feedbackMessage && (
          <div className={`flex items-center gap-2 p-3 rounded-lg text-white font-medium animate-in fade-in slide-in-from-top-2 ${feedbackMessage.type === 'success' ? 'bg-green-500 shadow-green-100' : 'bg-red-500'}`}>
            {feedbackMessage.type === 'success' ? <Check className="h-4 w-4" /> : <X className="h-4 w-4" />}
            {feedbackMessage.text}
          </div>
        )}

        {/* --- Main Table Container --- */}
        <div className="bg-white rounded-2xl border border-gray-200 shadow-xl overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="bg-gray-50/50 border-b border-gray-200">
                  <th className="px-4 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider sticky left-0 bg-gray-50 z-10">Employee</th>
                  <th className="px-4 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Office & Type</th>
                  <th className="px-4 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Incident Interval</th>
                  <th className="px-4 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Manager Approval</th>
                  <th className="px-4 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Final Approval</th>
                  <th className="px-4 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Manager Notes</th>
                  <th className="px-4 py-4 text-center text-xs font-bold text-gray-500 uppercase tracking-wider">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {/* --- Add New Row --- */}
                {isAdding && (
                  <tr className="bg-indigo-50/50">
                    <td className="px-4 py-3 sticky left-0 bg-indigo-50 z-10">
                      <div className="space-y-1">
                        <input type="text" placeholder="Name" value={newAbsence.employee_name} onChange={(e) => setNewAbsence({...newAbsence, employee_name: e.target.value})} className="w-full text-sm p-2 border border-indigo-200 rounded focus:ring-2 focus:ring-indigo-500 outline-none" />
                        <input type="text" placeholder="ID" value={newAbsence.employee_id} onChange={(e) => setNewAbsence({...newAbsence, employee_id: e.target.value})} className="w-full text-xs p-1 border border-indigo-200 rounded bg-indigo-50/50" />
                      </div>
                    </td>
                    <td className="px-4 py-3 space-y-1">
                      <input type="text" placeholder="Office" value={newAbsence.office} onChange={(e) => setNewAbsence({...newAbsence, office: e.target.value})} className="w-full text-sm p-2 border border-indigo-200 rounded" />
                      <select value={newAbsence.type_of_incident} onChange={(e) => setNewAbsence({...newAbsence, type_of_incident: e.target.value})} className="w-full text-sm p-2 border border-indigo-200 rounded">
                        <option value="">Incident Type</option>
                        {incidentTypes.map(t => <option key={t} value={t}>{t}</option>)}
                      </select>
                    </td>
                    <td className="px-4 py-3 space-y-1">
                       <div className="flex flex-col gap-1">
                          <input type="date" value={newAbsence.incident_start} onChange={(e) => setNewAbsence({...newAbsence, incident_start: e.target.value})} className="text-xs p-1 border border-indigo-200 rounded" />
                          <input type="date" value={newAbsence.incident_end} onChange={(e) => setNewAbsence({...newAbsence, incident_end: e.target.value})} className="text-xs p-1 border border-indigo-200 rounded" />
                       </div>
                    </td>
                    <td colSpan={3} className="px-4 py-3 text-center text-gray-400 text-xs italic">Set approvals after creation</td>
                    <td className="px-4 py-3">
                      <div className="flex justify-center gap-2">
                        <button onClick={handleCreateAbsence} disabled={isCreating} className="p-2 bg-green-500 text-white rounded-lg hover:bg-green-600 transition-colors"><Check className="h-4 w-4" /></button>
                        <button onClick={() => setIsAdding(false)} className="p-2 bg-gray-400 text-white rounded-lg hover:bg-gray-500 transition-colors"><X className="h-4 w-4" /></button>
                      </div>
                    </td>
                  </tr>
                )}

                {/* --- Existing Rows --- */}
                {filteredData.map((absence) => (
                  <tr key={absence.id} className="hover:bg-gray-50/50 transition-colors">
                    <td className="px-4 py-4 sticky left-0 bg-white z-10 border-r border-gray-100">
                      <div className="flex flex-col">
                        <span className="text-sm font-bold text-gray-900">{absence.employee_name}</span>
                        <span className="text-[10px] font-mono text-gray-500">#{absence.employee_id}</span>
                      </div>
                    </td>
                    <td className="px-4 py-4 min-w-[200px]">
                      <div className="space-y-1">
                        <input type="text" value={absence.office} onChange={(e) => handleInputChange(absence.id, 'office', e.target.value)} className="w-full text-xs p-1 bg-gray-50 border-none rounded focus:ring-1 focus:ring-indigo-500" />
                        <select value={absence.type_of_incident} onChange={(e) => handleInputChange(absence.id, 'type_of_incident', e.target.value)} className="w-full text-xs p-1 bg-gray-100 border-none rounded font-medium">
                          {incidentTypes.map(t => <option key={t} value={t}>{t}</option>)}
                        </select>
                      </div>
                    </td>
                    <td className="px-4 py-4">
                       <div className="flex flex-col gap-1 min-w-[130px]">
                          <div className="flex items-center gap-1">
                            <Calendar className="h-3 w-3 text-gray-400" />
                            <input type="date" value={absence.incident_start} onChange={(e) => handleInputChange(absence.id, 'incident_start', e.target.value)} className="text-[10px] border-none p-0 bg-transparent focus:ring-0" />
                          </div>
                          <div className="flex items-center gap-1">
                            <Calendar className="h-3 w-3 text-gray-400" />
                            <input type="date" value={absence.incident_end} onChange={(e) => handleInputChange(absence.id, 'incident_end', e.target.value)} className="text-[10px] border-none p-0 bg-transparent focus:ring-0" />
                          </div>
                       </div>
                    </td>
                    <td className="px-4 py-4">
                      <select value={absence.manager_approval || 'pending'} onChange={(e) => handleInputChange(absence.id, 'manager_approval', e.target.value)} className={`text-xs p-1 border-none rounded font-bold ${absence.manager_approval === 'approved' ? 'bg-green-100 text-green-700' : absence.manager_approval === 'denied' ? 'bg-red-100 text-red-700' : 'bg-yellow-100 text-yellow-700'}`}>
                        <option value="pending">Pending</option>
                        <option value="approved">Approved</option>
                        <option value="denied">Denied</option>
                      </select>
                    </td>
                    <td className="px-4 py-4">
                      <select value={absence.final_approval || 'pending'} onChange={(e) => handleInputChange(absence.id, 'final_approval', e.target.value)} className={`text-xs p-1 border-none rounded font-bold ${absence.final_approval === 'approved' ? 'bg-indigo-100 text-indigo-700' : absence.final_approval === 'denied' ? 'bg-red-100 text-red-700' : 'bg-gray-100 text-gray-700'}`}>
                        <option value="pending">Pending</option>
                        <option value="approved">Approved</option>
                        <option value="denied">Denied</option>
                      </select>
                    </td>
                    <td className="px-4 py-4 min-w-[180px]">
                      <textarea rows={1} value={absence.manager_notes || ''} onChange={(e) => handleInputChange(absence.id, 'manager_notes', e.target.value)} className="w-full text-xs p-1 bg-gray-50 border border-transparent rounded focus:border-gray-200 focus:bg-white transition-all resize-y" placeholder="Add notes..." />
                    </td>
                    <td className="px-4 py-4">
                      <button onClick={() => handleDeleteAbsence(absence.id)} className="mx-auto block p-2 text-gray-400 hover:text-red-600 transition-colors">
                        <Trash2 className="h-4 w-4" />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* --- Export Modal --- */}
      {isExportModalOpen && (
        <div className="fixed inset-0 bg-gray-900/60 backdrop-blur-sm flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md p-6">
            <div className="flex items-center gap-3 mb-4">
              <div className="p-2 bg-teal-100 rounded-lg text-teal-600"><FileText className="h-5 w-5" /></div>
              <h2 className="text-xl font-bold">Configure XLSX Export</h2>
            </div>
            <div className="grid grid-cols-2 gap-3 mb-6">
              {Object.keys(selectedColumns).map(key => (
                <label key={key} className="flex items-center gap-2 p-2 rounded-lg border border-gray-100 hover:bg-gray-50 cursor-pointer transition-colors">
                  <input type="checkbox" className="rounded text-teal-600 focus:ring-teal-500" checked={selectedColumns[key]} onChange={(e) => setSelectedColumns({...selectedColumns, [key]: e.target.checked})} />
                  <span className="text-xs font-medium text-gray-700 capitalize">{key.replace(/_/g, ' ')}</span>
                </label>
              ))}
            </div>
            <div className="flex gap-3">
              <button onClick={() => setIsExportModalOpen(false)} className="flex-1 px-4 py-2 bg-gray-100 text-gray-600 font-bold rounded-xl hover:bg-gray-200 transition-all">Cancel</button>
              <button onClick={handleExport} className="flex-1 px-4 py-2 bg-teal-600 text-white font-bold rounded-xl hover:bg-teal-700 transition-all shadow-lg shadow-teal-100">Generate File</button>
            </div>
          </div>
        </div>
      )}
    </main>
  );
}