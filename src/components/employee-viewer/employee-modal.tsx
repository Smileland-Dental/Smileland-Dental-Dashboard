'use client';

import React, { useState, useEffect } from 'react';
import { collection, query, where, getDocs } from 'firebase/firestore';
import { db } from '@/lib/firebase.config';
import { Employee, AbsenceRequest } from '@/lib/types';
// Import icons for the UI
import { User, CalendarDays, BarChart3, GraduationCap, X } from 'lucide-react';

// Define the types for our tabs for type safety
type TabCategory = 'contact' | 'absences' | 'performance' | 'training' | 'notes' | 'documents' | 'history'; // Added more tabs to demonstrate scrolling

interface EmployeeModalProps {
  employee: Employee;
  onClose: () => void;
}

const EmployeeModal: React.FC<EmployeeModalProps> = ({ employee, onClose }) => {
  // State for managing the active tab
  const [activeTab, setActiveTab] = useState<TabCategory>('contact');
  const [absences, setAbsences] = useState<AbsenceRequest[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);

  // Effect to lock/unlock background scroll
  useEffect(() => {
    document.body.style.overflow = 'hidden';
    return () => {
      document.body.style.overflow = 'unset';
    };
  }, []);

  // Effect to fetch absences when the modal opens or employee changes
  useEffect(() => {
    if (!employee?.id) return;

    const fetchAbsences = async () => {
      setLoading(true);
      setError(null);
      try {
        const absencesCol = collection(db, 'absences');
        const q = query(absencesCol, where('employee_id', '==', employee.id));
        const querySnapshot = await getDocs(q);
        const absencesData = querySnapshot.docs.map(doc => ({
          id: doc.id,
          ...doc.data(),
        })) as AbsenceRequest[];
        
        const sorted = absencesData.sort((a, b) => new Date(b.date_submitted).getTime() - new Date(a.date_submitted).getTime());
        setAbsences(sorted);
      } catch (err) {
        console.error("Error fetching absences:", err);
        setError("Failed to load absence history.");
      } finally {
        setLoading(false);
      }
    };

    fetchAbsences();
  }, [employee.id]);

  // Data for our navigation tabs
  const tabs = [
    { id: 'contact', label: 'Contact', icon: <User size={16} /> },
    { id: 'absences', label: 'Absences', icon: <CalendarDays size={16} /> },
    { id: 'performance', label: 'Performance', icon: <BarChart3 size={16} /> },
    { id: 'training', label: 'Training', icon: <GraduationCap size={16} /> },
    { id: 'notes', label: 'Manager Notes', icon: <User size={16} /> },
    { id: 'documents', label: 'Documents', icon: <User size={16} /> },
    { id: 'history', label: 'Work History', icon: <User size={16} /> },
    { id: 'label1', label: 'Label 1', icon: <User size={16} /> },
    { id: 'label2', label: 'Label 2', icon: <User size={16} /> },
    { id: 'label3', label: 'Label 3', icon: <User size={16} /> },
    { id: 'label4', label: 'Label 4', icon: <User size={16} /> },
    { id: 'label5', label: 'Label 5', icon: <User size={16} /> },
    { id: 'label6', label: 'Label 6', icon: <User size={16} /> },
  ];

  // Helper for status pills (unchanged)
  const getStatusPill = (status: 'pending' | 'approved' | 'denied') => {
    /* ... same as before ... */
    const baseClasses = 'px-2 py-1 text-xs font-bold rounded-full';

    switch (status) {

      case 'approved':

        return <span className={`${baseClasses} bg-green-100 text-green-700`}>Approved</span>;

      case 'denied':

        return <span className={`${baseClasses} bg-red-100 text-red-700`}>Denied</span>;

      default:

        return <span className={`${baseClasses} bg-yellow-100 text-yellow-700`}>Pending</span>;

    } 
  };

  // Renders the content for the right-side panel based on the active tab
  const renderContent = () => {
    switch (activeTab) {
      case 'contact':
        return (
          <div>
            <h3 className="text-xl font-bold text-gray-800 mb-4">Employee Information</h3>
            <div className="space-y-3 text-gray-700">
              <p><span className="font-semibold">ID:</span> {employee.id || 'N/A'}</p>
              <p><span className="font-semibold">Email:</span> {employee.email || 'N/A'}</p>
              <p><span className="font-semibold">Phone:</span> {employee.phone || 'N/A'}</p>
              <p><span className="font-semibold">Department:</span> {employee.department || 'N/A'}</p>
              <p><span className="font-semibold">Office:</span> {employee.office || 'N/A'}</p>
              <p><span className="font-semibold">Start Date:</span> {employee.startDate || 'N/A'}</p>
            </div>
          </div>
        );
      case 'absences':
        return (
          <div>
            <h3 className="text-xl font-bold text-gray-800 mb-4">Absence History</h3>
            {loading ? <p>Loading history...</p> : error ? <p className="text-red-500">{error}</p> : absences.length > 0 ? (
              <ul className="space-y-3 max-h overflow-y-auto pr-2">
                {absences.map((absence) => (
                  <li key={absence.id} className="p-3 bg-gray-50 rounded-lg border">
                    <div className="flex justify-between items-center">
                      <p className="font-bold text-gray-800">{absence.type_of_incident}</p>
                      {getStatusPill(absence.manager_approval)}
                    </div>
                    <p className="text-sm text-gray-600 mt-1">
                      {absence.incident_start} to {absence.incident_end}
                    </p>
                  </li>
                ))}
              </ul>
            ) : (
              <p className="text-gray-500 mt-4">No absence records found.</p>
            )}
          </div>
        );
      case 'performance':
        return (
            <div>
              <h3 className="text-xl font-bold text-gray-800 mb-4">Performance Reviews</h3>
              <p className="text-gray-500 mt-4">Performance review data is not yet available.</p>
            </div>
        );
      case 'training':
        return (
            <div>
              <h3 className="text-xl font-bold text-gray-800 mb-4">Training Records</h3>
              <p className="text-gray-500 mt-4">No completed training records found.</p>
            </div>
        );
      default:
        return null;
    }
  };

  return (
   <div
      className="fixed inset-0 bg-black/60 flex justify-center items-center z-50 p-4"
      onClick={onClose}
    >
      {/* CHANGE 1: Main modal container is now a vertical flexbox.
        This allows us to have a fixed header and a flexible (growing) body.
        max-h-[90vh] sets the maximum height for the entire modal.
      */}
      <div
        className="bg-white rounded-lg shadow-2xl w-full max-w-4xl max-h-[90vh] flex flex-col"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Modal Header: It does NOT scroll. 'flex-shrink-0' prevents it from shrinking. */}
        <div className="flex justify-between items-center p-6 border-b flex-shrink-0">
          <h2 className="text-2xl font-bold text-gray-900">Employee Profile</h2>
          <button onClick={onClose} className="text-gray-500 hover:text-gray-800">
            <X size={24} />
          </button>
        </div>

        {/* Modal Body: The grid now fills the remaining space and handles internal overflow.
          'flex-grow' makes this section expand. 'overflow-hidden' prevents this container
          from showing a scrollbar, forcing its children to manage their own scrolling.
        */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 p-6 flex-grow overflow-hidden">
          {/* CHANGE 2: The left column is now a vertical flexbox to manage its children.
          */}
          <div className="lg:col-span-1 lg:border-r lg:pr-6 flex flex-col overflow-hidden">
            {/* Employee Info Flexbox: This part stays fixed at the top of the column. */}
            <div className="flex flex-col items-center text-center sm:flex-row sm:text-left sm:items-start gap-4 mb-6 flex-shrink-0">
              <img
                src={employee.imageURL || '/defaultProfile.jpg'}
                alt={employee.name}
                className="w-24 h-24 rounded-full object-cover border-4 border-gray-200 flex-shrink-0"
              />
              <div className="flex-grow pt-2">
                <h3 className="text-xl font-bold text-gray-900">{employee.name}</h3>
                <p className="text-md text-blue-600 font-semibold">{employee.role}</p>
              </div>
            </div>
            
            {/* Tab Navigation: This is now the scrollable part of the left column.
              'flex-grow' tells it to take all available vertical space.
              'overflow-y-auto' adds a scrollbar ONLY when the tabs don't fit.
            */}
            <nav className="flex flex-col space-y-2 flex-grow overflow-y-auto pr-2">
              {tabs.map((tab) => (
                <button
                  key={tab.id}
                  onClick={() => setActiveTab(tab.id as TabCategory)}
                  className={`flex items-center gap-3 p-3 rounded-md text-sm font-medium transition-colors duration-200 w-full ${
                    activeTab === tab.id
                      ? 'bg-blue-100 text-blue-700'
                      : 'text-gray-600 hover:bg-gray-100 hover:text-gray-900'
                  }`}
                >
                  {tab.icon}
                  <span>{tab.label}</span>
                </button>
              ))}
            </nav>
          </div>

          {/* CHANGE 3: The right column now has its own independent scrollbar.
            'overflow-y-auto' will show a scrollbar if the content inside is taller
            than the available space in the column.
          */}
          <div className="lg:col-span-2 flex flex-col overflow-hidden">
            {renderContent()}
          </div>
        </div>
      </div>
    </div>
  );
};

export default EmployeeModal;