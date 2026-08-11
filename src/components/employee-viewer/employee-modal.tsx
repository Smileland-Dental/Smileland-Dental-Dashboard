'use client';

import React, { useState, useEffect } from 'react';
import { collection, query, where, getDocs, doc, updateDoc, deleteDoc } from 'firebase/firestore';
import { ref, uploadBytes, getDownloadURL, deleteObject, listAll } from 'firebase/storage';
import { db, storage } from '@/lib/firebase.config';
import { Employee, AbsenceRequest } from '@/lib/types';
// 1. ADDED: More icons for the additional tabs
// import { User, CalendarDays, BarChart3, GraduationCap, X, Pencil, Save, Upload, FileText, Folder, History, Tag } from 'lucide-react';
import { 
  Pencil,
  Save,
  Upload,
  X,
  LayoutDashboard, 
  UserCircle, 
  Award, 
  AlertTriangle, 
  ClipboardCheck, 
  CalendarClock, 
  Stethoscope, 
  Banknote,
  Tag
} from 'lucide-react';

import { GeneralTab } from '@/components/employee-viewer/tabs/general-tab';
import { EmployeeInfoTab } from '@/components/employee-viewer/tabs/employee-information-tab';
import { IncidentNoticesTab } from '@/components/employee-viewer/tabs/incident-notices-tab';
import { LicenseCertificatesTab } from '@/components/employee-viewer/tabs/license-certificates-tab';

// 2. UPDATED: Restored the full list of types
type TabCategory = 'general' | 'employee_information' | 'license_certificates' | 'incident_notices' | 'review' | 'leave_of_absence' | 'workers_comp' | 'edd_claims' | 'temp'; // Added more tabs to demonstrate scrolling

interface EmployeeModalProps {
  employee: Employee;
  userRole: string;
  onClose: () => void;
  onUpdate?: () => void;
}

const EmployeeModal: React.FC<EmployeeModalProps> = ({ employee, userRole, onClose, onUpdate }) => {
  const [activeTab, setActiveTab] = useState<TabCategory>('general');
  
  // EDITING STATE
  const [isEditing, setIsEditing] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [formData, setFormData] = useState<Employee>(employee);
  const [newImageFile, setNewImageFile] = useState<File | null>(null);

  const hasUnsavedChanges = JSON.stringify(formData) !== JSON.stringify(employee) || newImageFile !== null;

  const defaultProfileURL = "https://firebasestorage.googleapis.com/v0/b/smileland-dental-dashboard.firebasestorage.app/o/employee-pictures%2F.DefaultProfile%2Fprofile.png?alt=media&token=70b1a79f-1a33-4b6c-9b9b-c8e3debd209d"

  useEffect(() => {
    setFormData(employee);
    setNewImageFile(null);
    setIsEditing(false);
  }, [employee]);

  useEffect(() => {
    // 1. Calculate scrollbar width
    const scrollbarWidth = window.innerWidth - document.documentElement.clientWidth;
    
    // 2. Hide scrollbar AND add padding to "fill the gap"
    document.body.style.overflow = 'hidden';
    document.body.style.paddingRight = `${scrollbarWidth}px`;

    return () => {
      // 3. Reset everything when closing
      document.body.style.overflow = 'unset';
      document.body.style.paddingRight = '0px';
    };
  }, []);

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>) => {
    const { name, value } = e.target;
    setFormData(prev => ({ ...prev, [name]: value }));
  };

  const handleImageChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setNewImageFile(e.target.files[0]);
    }
  };

  const handleSave = async () => {
    setIsSaving(true);
    try {
      // 1. Reference the document using the permanent Firestore ID, not the employeeID which can change
      const docRef = doc(db, 'employees', employee.id); 
      const oldID = employee.employeeID;
      const newID = formData.employeeID;
      let finalImageURL = formData.imageURL;

      // 2. Check for numeric ID duplicated, if employeeID is being changed
      if (newID !== oldID) {
        const q = query(collection(db, 'employees'), where('employeeID', '==', newID));
        const querySnapshot = await getDocs(q);
        if (!querySnapshot.empty) {
          alert(`Error: Employee Number "${newID}" is already assigned to someone else.`);
          setIsSaving(false);
          return;
        }
      }

      // 3. IMAGE UPLOAD
      // Using employee.id ensures the photo stays with the record even if names/numbers change
      if (newImageFile) {
        const storageRef = ref(storage, `employee-pictures/${employee.id}/profile-picture`);
        const snapshot = await uploadBytes(storageRef, newImageFile);
        finalImageURL = await getDownloadURL(snapshot.ref);
      }

      // 4. PERFORM UPDATE
      const updatePayload = {
        ...formData,
        imageURL: finalImageURL,
        pay: Number(formData.pay) || 0,
        updatedAt: new Date(),
      };

      await updateDoc(docRef, updatePayload);
      
      setIsEditing(false);
      if (onUpdate) onUpdate();

    } catch (error) {
      console.error("Error updating employee:", error);
      alert("Failed to update employee information.");
    } finally {
      setIsSaving(false);
    }
  };

  const handleArchiveEmployee = async () => {
    const idToArchive = employee?.id;

    if (!idToArchive) {
      alert("Error: Could not find the database record ID.");
      return;
    }

    const confirmArchive = window.confirm(
      `Are you sure you want to ARCHIVE ${formData.firstName} ${formData.lastName}? They will be hidden from active lists.`
    );

    if (!confirmArchive) return;

    setIsSaving(true);
    try {
      const employeeRef = doc(db, 'employees', idToArchive);
      
      // Update status to 'Archived' (or add a dedicated boolean field 'isArchived: true')
      await updateDoc(employeeRef, {
        status: 'Terminated', 
        archivedAt: new Date(),
      });

      alert("Employee archived successfully.");
      
      if (onUpdate) onUpdate();
      onClose();

    } catch (error) {
      console.error("Error archiving employee:", error);
      alert("Failed to archive employee record.");
    } finally {
      setIsSaving(false);
    }
  };

  const handleCancelEdit = () => {
    setFormData(employee);
    setNewImageFile(null);
    setIsEditing(false);
  };

  const handleBeforeClose = () => {
    if (isEditing && hasUnsavedChanges) {
      const confirmClose = window.confirm("You have unsaved changes. If you close now, your changes will be lost.");
      if (!confirmClose) return;
    }
    onClose();
  };

  // --- RENDER HELPERS ---

  const renderContent = () => {
    const commonProps = { formData, handleInputChange, isEditing, setIsEditing, handleCancelEdit, handleSave, isSaving, hasUnsavedChanges };

    switch (activeTab) {
      case 'general':
        return <GeneralTab {...commonProps} />;
      case 'employee_information':
        return (
        <EmployeeInfoTab 
          {...commonProps}
          setFormData={setFormData}
          userRole={userRole}
        />
      );
      
      case 'incident_notices':
        return <IncidentNoticesTab employee={employee} />;
      // 3. UPDATED: Grouped all other tabs here so they don't crash
      case 'license_certificates':
        return (
          <LicenseCertificatesTab 
            isEditing={isEditing} 
            setIsEditing={setIsEditing} 
            handleCancelEdit={handleCancelEdit}
          />
        );
      case 'review':
      case 'leave_of_absence':
      case 'workers_comp':
      case 'edd_claims':
      case 'temp':
        return (
           <div>
             <h3 className="text-xl font-bold text-gray-800 mb-4 capitalize">{activeTab.replace('label', 'Label ')}</h3>
             <p className="text-gray-500 mt-4">This module is under development.</p>
           </div>
        );
      default:
        return null;
    }
  };

  // 4. UPDATED: Restored the full tabs array with new icons
  const tabs = [
    { id: 'general', label: 'General', icon: <LayoutDashboard size={16} /> },
    { id: 'employee_information', label: 'Employee Information', icon: <UserCircle size={16} /> },
    { id: 'license_certificates', label: 'License Certificates', icon: <Award size={16} /> },
    { id: 'incident_notices', label: 'Incident Notices', icon: <AlertTriangle size={16} /> },
    { id: 'review', label: 'Review', icon: <ClipboardCheck size={16} /> },
    { id: 'leave_of_absence', label: 'Leave of Absence', icon: <CalendarClock size={16} /> },
    { id: 'workers_comp', label: 'Workers Comp', icon: <Stethoscope size={16} /> },
    { id: 'edd_claims', label: 'EDD Claims', icon: <Banknote size={16} /> },
    { id: 'temp', label: 'Temp', icon: <Tag size={16} /> }
  ];

  return (
    <div className="fixed inset-0 bg-black/60 flex justify-center items-center z-50 p-4">
      <div className="bg-white rounded-lg shadow-2xl w-full max-w-6xl max-h-[90vh] flex flex-col">
        
        {/* Header */}
        <div className="flex justify-between items-center p-6 border-b flex-shrink-0">
          <h2 className="text-2xl font-bold text-gray-900">Employee Information</h2>
          {/*<h3 className="text-xs font-semibold text-blue-400/50">{formData.id}</h3>*/}
          <button onClick={handleBeforeClose} className="text-gray-500 hover:text-gray-800"><X size={24} /></button>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 p-6 flex-grow lg:overflow-hidden min-h-0">
          
          {/* LEFT SIDEBAR */}
          <div className="lg:col-span-1 lg:border-r lg:pr-6 flex flex-col lg:overflow-hidden">
            <div className="flex flex-col items-center text-center gap-4 mb-6 flex-shrink-0 relative">
                {isEditing && (
                  <button
                    onClick={handleArchiveEmployee}
                    className="mt-4 flex items-center justify-center gap-2 px-4 py-2 text-red-600 hover:bg-red-50 rounded-md border border-red-200 transition-colors"
                  >
                    <X size={16} /> Archive Employee Record
                  </button>
                )}

                <div className="relative group">
                    <img
                        src={newImageFile ? URL.createObjectURL(newImageFile) : (formData.imageURL || defaultProfileURL)}
                        alt={formData.firstName + ' ' + formData.lastName}
                        className="w-40 h-40 md:w-48 md:h-48 rounded-full object-cover border-4 border-gray-200 flex-shrink-0"
                    />
                    {isEditing && (
                        <label className="absolute inset-0 bg-black/50 rounded-full flex items-center justify-center cursor-pointer opacity-0 group-hover:opacity-100 transition-opacity">
                            <Upload className="text-white" size={24} />
                            <input type="file" className="hidden" accept="image/*" onChange={handleImageChange} />
                        </label>
                    )}
                </div>

                <div className="flex-grow pt-2">
                    <h3 className="text-xl font-bold text-gray-900">{formData.firstName + ' ' + formData.lastName}</h3>
                    <p className="text-md text-blue-600 font-semibold">{formData.jobTitle}</p>
                </div>
            </div>

            <nav className="flex flex-row lg:flex-col space-x-2 lg:space-x-0 lg:space-y-2 overflow-x-auto lg:overflow-y-auto pb-4 lg:pb-0 no-scrollbar">
              {tabs.map((tab) => (
                <button
                  key={tab.id}
                  onClick={() => {
                  if (isEditing && hasUnsavedChanges) {
                    const confirmLeave = window.confirm("You have unsaved changes. Are you sure you want to switch tabs? Your progress will be kept, but not saved to the database.");
                    if (!confirmLeave) return;

                  }

                  handleCancelEdit();
                  
                  setActiveTab(tab.id as TabCategory);}}
                  className={`flex items-center gap-3 p-3 rounded-md text-sm font-medium transition-colors duration-200 flex-shrink-0 lg:w-full ${
                    activeTab === tab.id ? 'bg-blue-100 text-blue-700' : 'text-gray-600 hover:bg-gray-100'
                  }`}
                >
                  {tab.icon} <span className="whitespace-nowrap">{tab.label}</span>
                </button>
              ))}
            </nav>
          </div>

          {/* RIGHT CONTENT AREA */}
          <div className="lg:col-span-2 flex flex-col min-h-0 overflow-hidden px-2">
            {/* By adding 'h-full' and 'min-h-0' to the parent 
                and 'overflow-y-auto' here, the scrollbar will finally appear.
            */}
            <div className="flex-grow overflow-y-auto pr-2">
                {renderContent()}
            </div>
          </div>

        </div>
      </div>
    </div>
  );
};

export default EmployeeModal;