'use client';

import React, { useState, useEffect, useMemo } from "react";
import { collection, getDocs, doc, writeBatch, addDoc, deleteDoc, Timestamp } from 'firebase/firestore';
import { db } from '@/lib/firebase.config';

import { AbsenceRequest } from "@/lib/types";
import * as XLSX from 'xlsx'; // Import the xlsx library

// --- New Type Definition ---
/* export type AbsenceRequest = {
  id: string; // The Firestore document ID
  employee_id: string;
  employee_name: string;
  employee_title: string;
  eta_etd?: string;
  excuse_note?: string[];
  excuse_note_submitted?: 'pending' | 'submitted' | 'not_provided';
  date_submitted?: string;
  type_of_incident: string;
  type_of_request?: string;
  office: string;
  incident_start: string;
  incident_end: string;
  employee_comments?: string;
  manager_notes?: string;
  manager_approval?: 'pending' | 'approved' | 'denied';
  manager_name?: string;
  final_approval?: 'pending' | 'approved' | 'denied';
  final_name?: string;
};
*/

// --- Initial State for New Entries ---
const initialNewAbsenceState: Omit<AbsenceRequest, 'id'> = {
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
  // Add other default fields as needed
};

const incidentTypes = ["Late In", "Early Out", "Absent", "Leave and Come Back", "Long Lunch", "Switch Shift", "Cancel Cell"];

export default function Page() {
  // --- Core State ---
  const [absences, setAbsences] = useState<AbsenceRequest[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // --- Editing and Saving State ---
  const [modifiedDocs, setModifiedDocs] = useState<Set<string>>(new Set());
  const [isSaving, setIsSaving] = useState(false);

  // --- Adding New Entry State ---
  const [isAdding, setIsAdding] = useState(false);
  const [newAbsence, setNewAbsence] = useState<Omit<AbsenceRequest, 'id'>>(initialNewAbsenceState);
  const [isCreating, setIsCreating] = useState(false);
  
  // --- UI Feedback State ---
  const [feedbackMessage, setFeedbackMessage] = useState<{type: 'success' | 'error', text: string} | null>(null);

  // --- 📅 Date Filtering State ---
  const [filterStartDate, setFilterStartDate] = useState('');
  const [filterEndDate, setFilterEndDate] = useState('');

  // --- 📄 XLSX Export State ---
  const [isExportModalOpen, setIsExportModalOpen] = useState(false);
  const [selectedColumns, setSelectedColumns] = useState<Record<keyof AbsenceRequest, boolean>>(
    // Initially select a few common columns by default
    {
      id: false, employee_id: true, employee_name: true, office: true, type_of_incident: true, 
      incident_start: true, incident_end: true, manager_approval: true, final_approval: false,
      employee_title: false, eta_etd: false, excuse_note: false, excuse_note_submitted: false,
      date_submitted: false, type_of_request: false, employee_comments: false, manager_notes: false,
      manager_name: false, final_name: false,
    }
  );

  // --- Display Feedback and Clear After a Few Seconds ---
  useEffect(() => {
    if (feedbackMessage) {
      const timer = setTimeout(() => setFeedbackMessage(null), 4000);
      return () => clearTimeout(timer);
    }
  }, [feedbackMessage]);
  
  // --- Fetch Logic (Simplified for the new data model) ---
  const handleFetch = async () => {
    setLoading(true);
    setError(null);
    setAbsences([]);
    setModifiedDocs(new Set());
    setIsAdding(false);
    try {
      const querySnapshot = await getDocs(collection(db, "absences"));
      const fetchedAbsences = querySnapshot.docs.map(doc => ({
        id: doc.id,
        ...(doc.data() as Omit<AbsenceRequest, 'id'>),
      }));
      setAbsences(fetchedAbsences);
    } catch (err) {
      console.error("Error fetching documents: ", err);
      setError("Failed to fetch data. Check collection names and permissions.");
    } finally {
      setLoading(false);
    }
  };

  // --- 📅 Memoized Filtering Logic ---
  const filteredData = useMemo(() => {
    if (!filterStartDate || !filterEndDate) {
      return absences; // No filter applied
    }
    return absences.filter(absence => {
      // An overlap occurs if (FilterStart <= AbsenceEnd) and (AbsenceStart <= FilterEnd)
      const absenceStart = new Date(absence.incident_start);
      const absenceEnd = new Date(absence.incident_end);
      const filterStart = new Date(filterStartDate);
      const filterEnd = new Date(filterEndDate);
      return filterStart <= absenceEnd && absenceStart <= filterEnd;
    });
  }, [absences, filterStartDate, filterEndDate]);


  // --- Handlers for Editing, Adding, Saving, Deleting ---
  const handleInputChange = (id: string, field: keyof AbsenceRequest, value: any) => {
    setAbsences(currentData =>
      currentData.map(absence => absence.id === id ? { ...absence, [field]: value } : absence)
    );
    setModifiedDocs(prev => new Set(prev).add(id));
  };

  const handleNewAbsenceChange = (field: keyof Omit<AbsenceRequest, 'id'>, value: any) => {
    setNewAbsence(prev => ({ ...prev, [field]: value }));
  };

  const handleSaveChanges = async () => {
    setIsSaving(true);
    setFeedbackMessage(null);
    const batch = writeBatch(db);
    modifiedDocs.forEach(docId => {
      const absenceData = absences.find(absence => absence.id === docId);
      if (absenceData) {
        const docRef = doc(db, "absences", docId);
        const { id, ...dataToSave } = absenceData; // Exclude the client-side 'id' field
        batch.update(docRef, dataToSave);
      }
    });
    try {
      await batch.commit();
      setFeedbackMessage({ type: 'success', text: 'All changes saved successfully!' });
      setModifiedDocs(new Set());
    } catch (err) {
      console.error("Error saving documents: ", err);
      setFeedbackMessage({ type: 'error', text: 'Failed to save changes.' });
    } finally {
      setIsSaving(false);
    }
  };

  const handleCreateAbsence = async () => {
    if (!newAbsence.employee_id || !newAbsence.employee_name) {
      alert("Employee ID and Name are required.");
      return;
    }
    setIsCreating(true);
    setFeedbackMessage(null);
    try {
      await addDoc(collection(db, "absences"), newAbsence);
      setFeedbackMessage({ type: 'success', text: 'New absence added successfully!' });
      setIsAdding(false);
      setNewAbsence(initialNewAbsenceState);
      await handleFetch(); // Refresh data
    } catch (err) {
      console.error("Error adding document: ", err);
      setFeedbackMessage({ type: 'error', text: 'Failed to add new absence.' });
    } finally {
      setIsCreating(false);
    }
  };

  const handleDeleteAbsence = async (idToDelete: string) => {
    if (window.confirm("Are you sure you want to permanently delete this record?")) {
      try {
        await deleteDoc(doc(db, "absences", idToDelete));
        setAbsences(currentData => currentData.filter(absence => absence.id !== idToDelete));
        setFeedbackMessage({ type: 'success', text: 'Record deleted successfully!' });
      } catch (err) {
        console.error("Error deleting document: ", err);
        setFeedbackMessage({ type: 'error', text: 'Failed to delete record.' });
      }
    }
  };
  
  // --- 📄 XLSX Export Handler ---
  const handleExport = () => {
    const activeColumns = Object.entries(selectedColumns)
      .filter(([, isSelected]) => isSelected)
      .map(([key]) => key as keyof AbsenceRequest);

    if (activeColumns.length === 0) {
      alert("Please select at least one column to export.");
      return;
    }

    // Prepare data for export by picking only the selected columns
    const dataToExport = filteredData.map(absence => {
      const row: Partial<AbsenceRequest> = {};
      for (const key of activeColumns) {
        const value = absence[key];
        // Handle array values for XLSX export, which expects strings
        if (key === 'excuse_note' && Array.isArray(value)) {
          (row as any)[key] = value.join(', ');
        } else {
          (row as any)[key] = value;
        }
      }
      return row;
    });

    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Absences");
    XLSX.writeFile(workbook, "EmployeeAbsences.xlsx");
    setIsExportModalOpen(false); // Close modal after export
  };


  return (
    <main className="p-8">
        <h1 className="text-2xl font-bold">Employee Absence Management</h1>
        <p className="mb-6 text-gray-600">
            View, edit, add, and filter absence records. Export your filtered data to XLSX.
        </p>

        {/* --- Controls: Fetch, Add, Save --- */}
        <div className="flex flex-wrap items-center gap-4 mb-4">
            <button onClick={handleFetch} disabled={loading} className="px-4 py-2 bg-blue-600 text-white font-semibold rounded-lg shadow-md hover:bg-blue-700 disabled:bg-gray-400">
                {loading ? 'Fetching...' : 'Fetch Absences'}
            </button>
            <button onClick={() => setIsAdding(true)} disabled={isAdding} className="px-4 py-2 bg-indigo-600 text-white font-semibold rounded-lg shadow-md hover:bg-indigo-700 disabled:bg-gray-400">
                Add New Record
            </button>
             <button onClick={() => setIsExportModalOpen(true)} disabled={filteredData.length === 0} className="px-4 py-2 bg-teal-600 text-white font-semibold rounded-lg shadow-md hover:bg-teal-700 disabled:bg-gray-400">
                Export to XLSX
            </button>
            {modifiedDocs.size > 0 && (
                <button onClick={handleSaveChanges} disabled={isSaving} className="px-4 py-2 bg-green-600 text-white font-semibold rounded-lg shadow-md hover:bg-green-700 disabled:bg-gray-400">
                    {isSaving ? 'Saving...' : `Save ${modifiedDocs.size} Changes`}
                </button>
            )}
        </div>

        {/* --- 📅 Date Filtering Controls --- */}
        <div className="flex items-center gap-4 p-4 bg-gray-100 rounded-lg mb-6">
            <label className="font-semibold">Filter by Incident Date:</label>
            <input type="date" value={filterStartDate} onChange={e => setFilterStartDate(e.target.value)} className="p-2 border rounded-md" />
            <span className="text-gray-600">to</span>
            <input type="date" value={filterEndDate} onChange={e => setFilterEndDate(e.target.value)} className="p-2 border rounded-md" />
            <button onClick={() => { setFilterStartDate(''); setFilterEndDate(''); }} className="px-4 py-2 text-sm bg-gray-300 rounded-md hover:bg-gray-400">Clear</button>
            <p className="ml-auto text-gray-700 font-medium">Showing {filteredData.length} of {absences.length} records</p>
        </div>

        {/* --- Feedback Messages --- */}
        {feedbackMessage && (
            <p className={`p-3 rounded-lg mb-4 text-white ${feedbackMessage.type === 'success' ? 'bg-green-500' : 'bg-red-500'}`}>
                {feedbackMessage.text}
            </p>
        )}
        
        {/* --- Table and Data Display --- */}
        <div className="mt-8">
            {loading && <p className="text-center text-gray-500">Loading data...</p>}
            {error && <p className="text-center text-red-500 bg-red-100 p-4 rounded-lg">{error}</p>}
            
            {(filteredData.length > 0 || isAdding) && (
                <div className="relative overflow-x-auto shadow-md sm:rounded-lg">
                    <table className="w-full text-sm text-left text-gray-500">
                      {/* ... Table Head (`<thead>`) ... */}
                      <thead className="text-xs text-gray-700 uppercase bg-gray-50">
                        <tr>
                          {/* Key columns first */}
                          <th scope="col" className="px-4 py-3">Employee Name</th>
                          <th scope="col" className="px-4 py-3">Employee ID</th>
                          <th scope="col" className="px-4 py-3">Office</th>
                          <th scope="col" className="px-4 py-3">Incident Type</th>
                          <th scope="col" className="px-4 py-3">Incident Start</th>
                          <th scope="col" className="px-4 py-3">Incident End</th>
                          <th scope="col" className="px-4 py-3">Manager Approval</th>
                          <th scope="col" className="px-4 py-3">Final Approval</th>
                          <th scope="col" className="px-4 py-3">Manager Notes</th>
                          <th scope="col" className="px-4 py-3 text-center">Actions</th>
                        </tr>
                      </thead>
                      <tbody>
                        {/* --- ADD NEW ROW --- */}
                        {isAdding && (
                            <tr className="bg-indigo-50 border-b border-indigo-200">
                              <td className="px-2 py-2"><input type="text" placeholder="Full Name" value={newAbsence.employee_name} onChange={(e) => handleNewAbsenceChange('employee_name', e.target.value)} className="w-full bg-white p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500" required/></td>
                              <td className="px-2 py-2"><input type="text" placeholder="Employee ID" value={newAbsence.employee_id} onChange={(e) => handleNewAbsenceChange('employee_id', e.target.value)} className="w-full bg-white p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500" required/></td>
                              <td className="px-2 py-2"><input type="text" placeholder="Office" value={newAbsence.office} onChange={(e) => handleNewAbsenceChange('office', e.target.value)} className="w-full bg-white p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500" /></td>
                              <td className="px-2 py-2">
                                <select value={newAbsence.type_of_incident} onChange={(e) => handleNewAbsenceChange('type_of_incident', e.target.value)} className="w-full bg-white p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500">
                                    <option value="" disabled>Select Type</option>
                                    {incidentTypes.map(type => <option key={type} value={type}>{type}</option>)}
                                </select>
                              </td>
                              <td className="px-2 py-2"><input type="date" value={newAbsence.incident_start} onChange={(e) => handleNewAbsenceChange('incident_start', e.target.value)} className="w-full bg-white p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500" /></td>
                              <td className="px-2 py-2"><input type="date" value={newAbsence.incident_end} onChange={(e) => handleNewAbsenceChange('incident_end', e.target.value)} className="w-full bg-white p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500" /></td>
                              <td colSpan={3} className="text-center italic text-gray-500">Set after creation</td>
                              <td className="px-2 py-2 whitespace-nowrap text-center">
                                  <div className="flex items-center justify-center space-x-2">
                                    <button onClick={handleCreateAbsence} disabled={isCreating} className="p-2 bg-green-500 text-white rounded hover:bg-green-600 disabled:bg-gray-300">✓</button>
                                    <button onClick={() => setIsAdding(false)} className="p-2 bg-red-500 text-white rounded hover:bg-red-600">✗</button>
                                  </div>
                              </td>
                            </tr>
                        )}
                        {/* --- EXISTING DATA ROWS (now uses filteredData) --- */}
                        {filteredData.map((absence) => (
                           <tr key={absence.id} className="bg-white border-b hover:bg-gray-50">
                                <td className="px-4 py-2 font-medium text-gray-900 bg-gray-50 whitespace-nowrap">{absence.employee_name}</td>
                                <td className="px-4 py-2 bg-gray-50">{absence.employee_id}</td>
                                <td className="px-2 py-2"><input type="text" value={absence.office} onChange={(e) => handleInputChange(absence.id, 'office', e.target.value)} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" /></td>
                                <td className="px-2 py-2">
                                    <select value={absence.type_of_incident} onChange={(e) => handleInputChange(absence.id, 'type_of_incident', e.target.value)} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500">
                                        {incidentTypes.map(type => <option key={type} value={type}>{type}</option>)}
                                    </select>
                                </td>
                                <td className="px-2 py-2"><input type="date" value={absence.incident_start} onChange={(e) => handleInputChange(absence.id, 'incident_start', e.target.value)} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" /></td>
                                <td className="px-2 py-2"><input type="date" value={absence.incident_end} onChange={(e) => handleInputChange(absence.id, 'incident_end', e.target.value)} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" /></td>
                                <td className="px-2 py-2">
                                    <select value={absence.manager_approval || 'pending'} onChange={(e) => handleInputChange(absence.id, 'manager_approval', e.target.value)} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500">
                                        <option value="pending">Pending</option>
                                        <option value="approved">Approved</option>
                                        <option value="denied">Denied</option>
                                    </select>
                                </td>
                                <td className="px-2 py-2">
                                    <select value={absence.final_approval || 'pending'} onChange={(e) => handleInputChange(absence.id, 'final_approval', e.target.value)} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500">
                                        <option value="pending">Pending</option>
                                        <option value="approved">Approved</option>
                                        <option value="denied">Denied</option>
                                    </select>
                                </td>
                                <td className="px-2 py-2"><textarea rows={1} title={absence.manager_notes || ''} value={absence.manager_notes || ''} onChange={(e) => handleInputChange(absence.id, 'manager_notes', e.target.value)} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 resize-y min-h-[40px]" /></td>
                                <td className="px-2 py-2 text-center">
                                    <button onClick={() => handleDeleteAbsence(absence.id)} className="p-2 bg-red-500 text-white rounded hover:bg-red-600" title="Delete Record">✗</button>
                                </td>
                           </tr>
                        ))}
                      </tbody>
                    </table>
                </div>
            )}
        </div>

        {/* --- 📄 XLSX Export Modal --- */}
        {isExportModalOpen && (
            <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
                <div className="bg-white p-8 rounded-lg shadow-2xl w-full max-w-lg">
                    <h2 className="text-xl font-bold mb-4">Select Columns to Export</h2>
                    <div className="grid grid-cols-2 gap-4 max-h-80 overflow-y-auto pr-2">
                        {Object.keys(selectedColumns).map(key => (
                            <label key={key} className="flex items-center space-x-3 p-2 rounded-md hover:bg-gray-100 cursor-pointer">
                                <input
                                    type="checkbox"
                                    className="h-5 w-5 rounded text-blue-600 focus:ring-blue-500"
                                    checked={selectedColumns[key as keyof AbsenceRequest]}
                                    onChange={(e) => setSelectedColumns(prev => ({...prev, [key]: e.target.checked}))}
                                />
                                <span className="text-gray-800 capitalize">{key.replace(/_/g, ' ')}</span>
                            </label>
                        ))}
                    </div>
                    <div className="mt-6 flex justify-end space-x-4">
                        <button onClick={() => setIsExportModalOpen(false)} className="px-4 py-2 bg-gray-300 text-gray-800 font-semibold rounded-lg hover:bg-gray-400">Cancel</button>
                        <button onClick={handleExport} className="px-4 py-2 bg-teal-600 text-white font-semibold rounded-lg hover:bg-teal-700">Generate XLSX</button>
                    </div>
                </div>
            </div>
        )}
    </main>
  );
}