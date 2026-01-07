'use client';

import React, { useState, useEffect } from "react";
import { collection, getDocs, doc, writeBatch, addDoc, deleteDoc } from 'firebase/firestore';
import { db } from '@/lib/firebase.config';

// --- Interfaces ---

// Original Absence interface from Firestore
interface Absence {
  absence_id: string; 
  employee_id: string;
  office?: string;
  type_of_incident?: string;
  incident_start?: string;
  incident_end?: string;
  notes?: string;
  incident_submitted?: boolean;
  date_submitted?: string;
  comments?: string;
  excuse_note?: string;
}

// Interface for the combined data we'll use in the table state
interface UpdatedAbsence extends Absence {
    name?: string;  // Added from 'employees' collection
    email?: string; // Added from 'employees' collection
}

// Employee data structure
interface Employee {
    active: boolean;
    email: string;
    name: string;
    year: number;
}

type NewAbsence = Omit<Absence, 'absence_id'>;

const initialNewAbsenceState: NewAbsence = {
  employee_id: '',
  office: '',
  type_of_incident: '',
  incident_start: '',
  incident_end: '',
  notes: '',
  incident_submitted: false, 
  date_submitted: '',      
  comments: '',            
  excuse_note: '',         
};

const incidentTypes = [
  "Late In",
  "Early Out",
  "Absent",
  "Leave and Come Back",
  "Long Lunch",
  "Switch Shift",
  "Cancel Cell"
];


export default function Page() {
  // --- State now uses the UpdatedAbsence interface ---
  const [data, setData] = useState<UpdatedAbsence[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // --- State for Editing and Updating ---
  const [modifiedDocs, setModifiedDocs] = useState<Set<string>>(new Set());
  const [isSaving, setIsSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);
  const [saveSuccess, setSaveSuccess] = useState<string | null>(null);

  // --- State for Creating New Entries ---
  const [isAdding, setIsAdding] = useState(false);
  const [newAbsence, setNewAbsence] = useState<NewAbsence>(initialNewAbsenceState);
  const [isCreating, setIsCreating] = useState(false);


  const handleFetch = async () => {
    setLoading(true);
    setError(null);
    setData([]);
    setModifiedDocs(new Set());
    setIsAdding(false); 

    try {
      // Step 1: Fetch both collections concurrently for efficiency
      const [absencesSnapshot, employeesSnapshot] = await Promise.all([
          getDocs(collection(db, "absences")),
          getDocs(collection(db, "employees"))
      ]);

      // Step 2: Create a lookup map for employees for quick access
      const employeesMap = new Map<string, Employee>();
      employeesSnapshot.forEach(doc => {
          employeesMap.set(doc.id, doc.data() as Employee);
      });

      // Step 3: Map over absences and enrich them with employee data
      const enrichedAbsencesData = absencesSnapshot.docs.map(doc => {
          const absenceData = doc.data() as Omit<Absence, 'absence_id'>;
          const employeeDetails = employeesMap.get(absenceData.employee_id);

          return {
              ...absenceData,
              absence_id: doc.id,
              name: employeeDetails?.name || 'Unknown',
              email: employeeDetails?.email || 'N/A',
          };
      }) as UpdatedAbsence[];

      setData(enrichedAbsencesData);

    } catch (err) {
      console.error("Error fetching documents: ", err);
      setError("Failed to fetch data. Check collection names and permissions.");
    } finally {
      setLoading(false);
    }
  };

  // Handler for existing row input changes
  const handleInputChange = (id: string, field: keyof NewAbsence, value: string | boolean | number) => {
    setData(currentData =>
      currentData.map(absence => absence.absence_id === id ? { ...absence, [field]: value } : absence)
    );
    setModifiedDocs(prev => new Set(prev).add(id));
    setSaveSuccess(null);
  };

  // Handler for NEW row input changes
  const handleNewAbsenceChange = (field: keyof NewAbsence, value: string | boolean | number) => {
    setNewAbsence(prev => ({ ...prev, [field]: value }));
  };

  // SAVE CHANGES to existing documents
  const handleSaveChanges = async () => {
    setIsSaving(true);
    setSaveError(null);
    setSaveSuccess(null);
    const batch = writeBatch(db);
    modifiedDocs.forEach(docId => {
      const absenceData = data.find(absence => absence.absence_id === docId);
      if (absenceData) {
        const docRef = doc(db, "absences", docId);
        // IMPORTANT: Destructure out the added fields (name, email) so they aren't saved to the 'absences' collection
        const { absence_id, name, email, ...dataToSave } = absenceData;
        batch.update(docRef, dataToSave as { [key: string]: any });
      }
    });

    try {
      await batch.commit();
      setSaveSuccess("All changes saved successfully!");
      setModifiedDocs(new Set());
    } catch (err) {
      console.error("Error saving documents: ", err);
      setSaveError("Failed to save changes. Please try again.");
    } finally {
      setIsSaving(false);
    }
  };

  // ADD NEW absence to the database
  const handleCreateAbsence = async () => {
    if (!newAbsence.employee_id) {
      alert("Employee ID is required to add a new entry.");
      return;
    }
    setIsCreating(true);
    setSaveError(null);

    try {
      await addDoc(collection(db, "absences"), newAbsence);
      setSaveSuccess(`Absence for employee "${newAbsence.employee_id}" added successfully!`);
      setIsAdding(false); 
      setNewAbsence(initialNewAbsenceState); 
      await handleFetch(); 
    } catch (err) {
      console.error("Error adding document: ", err);
      setSaveError("Failed to add new absence.");
    } finally {
      setIsCreating(false);
    }
  };
  
  // DELETE an absence from the database
  const handleDeleteAbsence = async (idToDelete: string) => {
    if (window.confirm("Are you sure you want to permanently delete this record?")) {
        try {
            await deleteDoc(doc(db, "absences", idToDelete));
            setData(currentData => currentData.filter(absence => absence.absence_id !== idToDelete));
            setSaveSuccess("Record deleted successfully!");
        } catch (err) {
            console.error("Error deleting document: ", err);
            setSaveError("Failed to delete record. Please try again.");
        }
    }
  };

  const handleCancelAdd = () => {
    setIsAdding(false);
    setNewAbsence(initialNewAbsenceState);
  };

  useEffect(() => {
    if (saveSuccess || saveError) {
      const timer = setTimeout(() => {
        setSaveSuccess(null);
        setSaveError(null);
      }, 4000);
      return () => clearTimeout(timer);
    }
  }, [saveSuccess, saveError]);


  return (
    <main className="pt-8">
      <div className="flex justify-between items-center mb-4">
        <h1 className="text-2xl font-bold">Employee Absences</h1>
      </div>
      
      <p className="mb-6 text-gray-600">
        Fetch data, edit in the table, or add a new absence record. A "Save" button will appear when changes are made.
      </p>
      
      <div className="flex space-x-4">
        <button onClick={handleFetch} disabled={loading} className="cursor-pointer px-4 py-2 bg-blue-600 text-white font-semibold rounded-lg shadow-md hover:bg-blue-700 disabled:bg-gray-400">
          {loading ? 'Fetching...' : 'Fetch Absences'}
        </button>
        <button onClick={() => setIsAdding(true)} disabled={isAdding} className="cursor-pointer px-4 py-2 bg-indigo-600 text-white font-semibold rounded-lg shadow-md hover:bg-indigo-700 disabled:bg-gray-400">
          Add New Absence
        </button>

                  {saveSuccess && <p className="text-green-600 font-semibold">{saveSuccess}</p>}
          {saveError && <p className="text-red-600 font-semibold">{saveError}</p>}
          {modifiedDocs.size > 0 && (
            <button
              onClick={handleSaveChanges}
              disabled={isSaving}
              className="cursor-pointer px-4 py-2 bg-green-600 text-white font-semibold rounded-lg shadow-md hover:bg-green-700 disabled:bg-gray-400"
            >
              {isSaving ? 'Saving...' : `Save ${modifiedDocs.size} Changes`}
            </button>
          )}
      </div>

      <div className="mt-8">
        {loading && <p className="text-center text-gray-500">Loading data...</p>}
        {error && <p className="text-center text-red-500 bg-red-100 p-4 rounded-lg">{error}</p>}
        
        {data.length > 0 || isAdding ? (
          <div className="relative overflow-x-auto shadow-md sm:rounded-lg">
            <table className="w-full text-sm text-left text-gray-500">
              <thead className="text-xs text-gray-700 uppercase bg-gray-50">
                <tr>
                  <th scope="col" className="px-6 py-3">Name</th>
                  <th scope="col" className="px-6 py-3">Email</th>
                  <th scope="col" className="px-6 py-3">ID</th>
                  <th scope="col" className="px-6 py-3">Office</th>
                  <th scope="col" className="px-6 py-3">Incident Type</th>
                  <th scope="col" className="px-6 py-3">Incident Start</th>
                  <th scope="col" className="px-6 py-3">Incident End</th>
                  <th scope="col" className="px-6 py-3">Notes</th>
                  <th scope="col" className="px-6 py-3">Submitted</th>
                  <th scope="col" className="px-6 py-3">Date Submitted</th>
                  <th scope="col" className="px-6 py-3">Comments</th>
                  <th scope="col" className="px-6 py-3">Excuse Note</th>
                  <th scope="col" className="px-6 py-3 text-center">Actions</th>
                </tr>
              </thead>
              <tbody>
                {/* --- ADD NEW ROW --- */}
                {isAdding && (
                  <tr className="bg-indigo-50 border-b border-indigo-200">
                    <td className="px-6 py-4 bg-indigo-100 text-gray-500 italic">Info appears after adding</td>
                    <td className="px-6 py-4 bg-indigo-100 text-gray-500 italic">Info appears after adding</td>
                    <td className="px-2 py-2"><input type="text" placeholder="Employee ID" value={newAbsence.employee_id} onChange={(e) => handleNewAbsenceChange('employee_id', e.target.value)} className="w-full bg-white p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500" required/></td>
                    <td className="px-2 py-2"><input type="text" placeholder="Office" value={newAbsence.office} onChange={(e) => handleNewAbsenceChange('office', e.target.value)} className="w-full bg-white p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500" /></td>
                    <td className="px-2 py-2">
  <select
    value={newAbsence.type_of_incident}
    onChange={(e) => handleNewAbsenceChange('type_of_incident', e.target.value)}
    className="w-full bg-white p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500"
  >
    <option value="" disabled>Select Type</option>
    {incidentTypes.map(type => (
      <option key={type} value={type}>{type}</option>
    ))}
  </select>
</td>
                    <td className="px-2 py-2"><input type="date" value={newAbsence.incident_start} onChange={(e) => handleNewAbsenceChange('incident_start', e.target.value)} className="w-full bg-white p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500" /></td>
                    <td className="px-2 py-2"><input type="date" value={newAbsence.incident_end} onChange={(e) => handleNewAbsenceChange('incident_end', e.target.value)} className="w-full bg-white p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500" /></td>
                    <td className="px-2 py-2"><textarea placeholder="Initial notes..." value={newAbsence.notes} onChange={(e) => handleNewAbsenceChange('notes', e.target.value)} className="w-full bg-white p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500" rows={1} /></td>
                    <td colSpan={4}></td>
                    <td className="px-2 py-2 whitespace-nowrap text-center">
                       <div className="flex items-center justify-center space-x-2">
                          <button onClick={handleCreateAbsence} disabled={isCreating} className="p-2 bg-green-500 text-white rounded hover:bg-green-600 disabled:bg-gray-300">✓</button>
                          <button onClick={handleCancelAdd} className="p-2 bg-red-500 text-white rounded hover:bg-red-600">✗</button>
                       </div>
                    </td>
                  </tr>
                )}
                {/* --- EXISTING DATA ROWS --- */}
                {data.map((absence) => (
                  <tr key={absence.absence_id} className="bg-white border-b hover:bg-gray-50">
                    <td className="px-6 py-4 font-medium text-gray-900 bg-gray-50">{absence.name}</td>
                    <td className="px-6 py-4 bg-gray-50">{absence.email}</td>
                    <td className="px-2 py-2"><input type="text" value={absence.employee_id || ''} onChange={(e) => handleInputChange(absence.absence_id, 'employee_id', e.target.value)} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" /></td>
                    <td className="px-2 py-2"><input type="text" value={absence.office || ''} onChange={(e) => handleInputChange(absence.absence_id, 'office', e.target.value)} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" /></td>
                      <td className="px-2 py-2"><select
    value={absence.type_of_incident || ''}
    onChange={(e) => handleInputChange(absence.absence_id, 'type_of_incident', e.target.value)}
    className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
  >
    <option value="" disabled>Select Type</option>
    {incidentTypes.map(type => (
      <option key={type} value={type}>{type}</option>
    ))}
  </select>
</td>
                    <td className="px-2 py-2"><input type="date" value={absence.incident_start || ''} onChange={(e) => handleInputChange(absence.absence_id, 'incident_start', e.target.value)} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" /></td>
                    <td className="px-2 py-2"><input type="date" value={absence.incident_end || ''} onChange={(e) => handleInputChange(absence.absence_id, 'incident_end', e.target.value)} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" /></td>
                    <td className="px-2 py-2"><textarea rows={1} title={absence.notes || ''} value={absence.notes || ''} onChange={(e) => handleInputChange(absence.absence_id, 'notes', e.target.value)} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 resize-y min-h-[40px]" /></td>
                    <td className="px-6 py-4 text-center"><input type="checkbox" checked={!!absence.incident_submitted} onChange={(e) => handleInputChange(absence.absence_id, 'incident_submitted', e.target.checked)} className="h-5 w-5 rounded text-blue-600 focus:ring-blue-500" /></td>
                    <td className="px-2 py-2"><input type="date" value={absence.date_submitted || ''} onChange={(e) => handleInputChange(absence.absence_id, 'date_submitted', e.target.value)} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" /></td>
                    <td className="px-2 py-2"><textarea rows={1} title={absence.comments || ''} value={absence.comments || ''} onChange={(e) => handleInputChange(absence.absence_id, 'comments', e.target.value)} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 resize-y min-h-[40px]" /></td>
                    <td className="px-2 py-2"><input type="text" value={absence.excuse_note || ''} onChange={(e) => handleInputChange(absence.absence_id, 'excuse_note', e.target.value)} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" /></td>
                    <td className="px-2 py-2 text-center">
                        <button onClick={() => handleDeleteAbsence(absence.absence_id)} className="p-2 bg-red-500 text-white rounded hover:bg-red-600" title="Delete Record">
                            ✗
                        </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : null}
      </div>
    </main>
  );
}