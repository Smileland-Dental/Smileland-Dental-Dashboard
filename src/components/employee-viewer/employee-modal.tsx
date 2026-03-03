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
  const [absences, setAbsences] = useState<AbsenceRequest[]>([]);
  const [loadingHistory, setLoadingHistory] = useState<boolean>(true);
  
  // EDITING STATE
  const [isEditing, setIsEditing] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [formData, setFormData] = useState<Employee>(employee);
  const [newImageFile, setNewImageFile] = useState<File | null>(null);

  const hasUnsavedChanges = JSON.stringify(formData) !== JSON.stringify(employee) || newImageFile !== null;

  const offices = ["Corporate", "Ortho", "California", "Ming", "Bernard", "Delano", "Tulare", "Visalia", "Fresno"];
  const departments = ["Back Office", "Front Office", "Support Services", "Management", "Call Center", "Accounts Receivable", "Dentist"];
  const jobStatuses = ["Full-Time", "Part-Time"];
  const employmentStatuses = ["Exempt", "Non-Exempt", "Contracted"];
  const statusOptions = ["Current", "Terminated", "On Leave"];
  const defaultProfileURL = "https://firebasestorage.googleapis.com/v0/b/smileland-dental-dashboard.firebasestorage.app/o/employee-pictures%2F.DefaultProfile%2Fprofile.png?alt=media&token=70b1a79f-1a33-4b6c-9b9b-c8e3debd209d"

  useEffect(() => {
    setFormData(employee);
    setNewImageFile(null);
    setIsEditing(false);
  }, [employee]);

  useEffect(() => {
    document.body.style.overflow = 'hidden';
    return () => { document.body.style.overflow = 'unset'; };
  }, []);

  /* Edit this when dealing with Absences
  useEffect(() => {
    if (!employee?.employeeID) return;
    const fetchAbsences = async () => {
      setLoadingHistory(true);
      try {
        const absencesCol = collection(db, 'absences');
        const q = query(absencesCol, where('employee_id', '==', employee.employeeID));
        const querySnapshot = await getDocs(q);
        const absencesData = querySnapshot.docs.map(doc => ({ id: doc.id, ...doc.data() })) as AbsenceRequest[];
        const sorted = absencesData.sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());
        setAbsences(sorted);
      } catch (err) {
        console.error("Error fetching absences:", err);
      } finally {
        setLoadingHistory(false);
      }
    };
    fetchAbsences();
  }, [employee.employeeID]);*/

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

  const handleDeleteEmployee = async () => {
    // 1. IMPORTANT: Use the Firestore Auto-ID (employee.id)
    // This is the permanent ID we saved in AddEmployeeForm
    const idToDelete = employee?.id;

    // 2. Safety check: If ID is missing, stop immediately
    if (!idToDelete) {
      console.error("No Firestore Document ID found to delete");
      alert("Error: Could not find the database record ID.");
      return;
    }

    const confirmDelete = window.confirm(
      `Are you sure you want to delete ${formData.firstName} ${formData.lastName}? This action cannot be undone.`
    );

    if (!confirmDelete) return;

    setIsSaving(true); // Reusing isSaving to show a loading state
    try {
      //const employeeId = employee.employeeID;

      // 1. Delete the Firestore Document
      const employeeRef = doc(db, 'employees', idToDelete);
      await deleteDoc(employeeRef);

      // 2. Clean up Firebase Storage
      // We list all files in the employee's folder and delete them
      const folderRef = ref(storage, `employee-pictures/${idToDelete}`);

      try {
        const fileList = await listAll(folderRef);
        const deletePromises = fileList.items.map((fileItem) => deleteObject(fileItem));
        await Promise.all(deletePromises);
      } catch (storageErr) {
        console.warn("Storage folder was empty or already deleted:", storageErr);
      }

      alert("Employee record deleted successfully.");
      
      // 3. Refresh Parent and Close
      if (onUpdate) onUpdate();
      onClose();

    } catch (error) {
      console.error("Error deleting employee:", error);
      alert("Failed to delete employee record.");
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
    switch (activeTab) {
      case 'general':
        return (
          <div className="relative">
            <div className="sticky top-0 z-10 bg-white pb-4 border-b mb-4 flex justify-between items-center">
               <h3 className="text-2xl font-bold text-gray-800">General</h3>
               {!isEditing ? (
                 <button 
                   onClick={() => setIsEditing(true)} 
                   className="flex items-center gap-2 text-sm text-blue-600 hover:text-blue-800 font-semibold"
                 >
                   <Pencil size={16} /> Edit Details
                 </button>
               ) : (
                <div className="col-span-2 flex justify-end gap-3 mt-4">
                    <button onClick={handleCancelEdit} className="px-4 py-2 text-gray-600 hover:bg-gray-100 rounded">Cancel</button>
                    <button 
                      onClick={handleSave} 
                      disabled={isSaving}
                      className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700 disabled:bg-blue-400"
                    >
                      {isSaving ? 'Saving...' : <><Save size={16}/> Save Changes</>}
                    </button>
                 </div>
                )
                }
            </div>

            {isEditing ? (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                 <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">First Name</label>
                    <input name="firstName" value={formData.firstName || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
                  <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Last Name</label>
                    <input name="lastName" value={formData.lastName || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
                 <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Email</label>
                    <input name="email" value={formData.email || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
                 <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Employee ID</label>
                    <input name="employeeID" value={formData.employeeID || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
                 <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Address</label>
                    <input name="address" value={formData.address || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
                 <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">City</label>
                    <input name="city" value={formData.city || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
                 <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">State</label>
                    <input name="state" value={formData.state || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
                 <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Zip Code</label>
                    <input name="zipCode" value={formData.zipCode || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
                <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Primary Phone</label>
                    <input name="phone" value={formData.phone || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
                  <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Secondary Phone</label>
                    <input name="secondaryPhone" value={formData.secondaryPhone || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
                  <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Emergency Contact 1</label>
                    <input name="emergencyContact1Name" value={formData.emergencyContact1Name || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
                  <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Emergency Contact 2</label>
                    <input name="emergencyContact2Name" value={formData.emergencyContact2Name || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
                  <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Emergency Contact 1 Phone</label>
                    <input name="emergencyContact1Phone" value={formData.emergencyContact1Phone || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
                  <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Emergency Contact 2 Phone</label>
                    <input name="emergencyContact2Phone" value={formData.emergencyContact2Phone || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
                  <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Emergency Contact 1 Relationship</label>
                    <input name="emergencyContact1Relation" value={formData.emergencyContact1Relation || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
                  <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Emergency Contact 2 Relationship</label>
                    <input name="emergencyContact2Relation" value={formData.emergencyContact2Relation || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
                 
                <div className="space-y-1 md:col-span-2"> {/* Added col-span-2 to make it full width */}
                  <label className="text-xs font-semibold text-gray-500">Comments</label>
                  <textarea 
                    name="comments" 
                    value={formData.comments || ''} 
                    onChange={handleInputChange} 
                    rows={4} // Sets the initial height
                    placeholder="Enter any additional notes here..."
                    className="w-full border p-2 rounded resize-y focus:ring-2 focus:ring-blue-500 outline-none" 
                  />
                </div>

                <div className="space-y-1 md:col-span-2">
                  <label className="text-xs font-semibold text-gray-500">Permission To Release Check</label>
                  <input name="permissionToReleaseCheck" value={formData.permissionToReleaseCheck || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                </div>
                
              </div>
            ) : (
              <div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-y-4 text-gray-700">
             
             {/* Column 1: Identity */}
             <div className="space-y-3">
               <p className="flex">
                 {/* Changed w-32 to w-42 and added whitespace-nowrap */}
                 <span className="font-semibold w-42 shrink-0 text-gray-900 whitespace-nowrap">First Name:</span> 
                 <span>{formData.firstName || 'N/A'}</span>
               </p>
               <p className="flex">
                 <span className="font-semibold w-42 shrink-0 text-gray-900 whitespace-nowrap">Last Name:</span> 
                 <span>{formData.lastName || 'N/A'}</span>
               </p>
               <p className="flex">
                 <span className="font-semibold w-42 shrink-0 text-gray-900 whitespace-nowrap">Email:</span> 
                 {/* Added truncate to prevent long emails from breaking layout */}
                 <span className="truncate hover:text-clip hover:whitespace-normal">{formData.email || 'N/A'}</span>
               </p>
               <p className="flex">
                 <span className="font-semibold w-42 shrink-0 text-gray-900 whitespace-nowrap">Employee ID:</span> 
                 <span>{formData.employeeID || 'N/A'}</span>
               </p>
                <p className="flex">
                 <span className="font-semibold w-42 shrink-0 text-gray-900 whitespace-nowrap">Address:</span> 
                 <span>{formData.address || 'N/A'}</span>
               </p>
               <p className="flex">
                 <span className="font-semibold w-42 shrink-0 text-gray-900 whitespace-nowrap">City:</span> 
                 <span>{formData.city || 'N/A'}</span>
               </p>
               <p className="flex">
                 <span className="font-semibold w-42 shrink-0 text-gray-900 whitespace-nowrap">State:</span> 
                 <span>{formData.state || 'N/A'}</span>
               </p>
               <p className="flex">
                 <span className="font-semibold w-42 shrink-0 text-gray-900 whitespace-nowrap">Zip Code:</span> 
                 <span>{formData.zipCode || 'N/A'}</span>
               </p>
             </div>

             {/* Column 2: Contact & Location */}
             <div className="space-y-3">
               <p className="flex">
                 <span className="font-semibold w-42 shrink-0 text-gray-900 whitespace-nowrap">Primary Phone:</span> 
                 <span>{formData.phone || 'N/A'}</span>
               </p>
               <p className="flex">
                  <span className="font-semibold w-42 shrink-0 text-gray-900 whitespace-nowrap">Secondary Phone:</span> 
                  <span>{formData.secondaryPhone || 'N/A'}</span>
               </p>
               <p className="flex">
                  <span className="font-semibold w-42 shrink-0 text-gray-900 whitespace-nowrap">Emergency Contact 1:</span> 
                  <span>{formData.emergencyContact1Name || 'N/A'}</span>
               </p>
                <p className="flex">
                 <span className="font-semibold w-42 shrink-0 text-gray-900 whitespace-nowrap">Phone:</span> 
                 <span>{formData.emergencyContact1Phone || 'N/A'}</span>
               </p>
                <p className="flex">
                 <span className="font-semibold w-42 shrink-0 text-gray-900 whitespace-nowrap">Relationship:</span> 
                 <span>{formData.emergencyContact1Relation || 'N/A'}</span>
               </p>
                <p className="flex">
                 <span className="font-semibold w-42 shrink-0 text-gray-900 whitespace-nowrap">Emergency Contact 2:</span> 
                 <span>{formData.emergencyContact2Name || 'N/A'}</span>
               </p>
                <p className="flex">
                 <span className="font-semibold w-42 shrink-0 text-gray-900 whitespace-nowrap">Phone:</span> 
                 <span>{formData.emergencyContact2Phone || 'N/A'}</span>
               </p>
               <p className="flex">
                 <span className="font-semibold w-42 shrink-0 text-gray-900 whitespace-nowrap">Relationship:</span> 
                 <span>{formData.emergencyContact2Relation || 'N/A'}</span>
               </p>
             </div>
           </div>
            <p className="flex padding-top-4 text-gray-700 border-t border-gray-200 mt-4 pt-2">
                 <span className="font-semibold w-42 shrink-0 text-gray-900 whitespace-nowrap">Comments:</span> 
                 <span className="truncate">{formData.comments || 'N/A'}</span>
            </p>
            <p className="flex padding-top-4 text-gray-700 border-t border-gray-200 mt-4 pt-2">
                 <span className="font-semibold w-42 shrink-0 text-gray-900">Permission To Release Check:</span> 
                 <span className="truncate">{formData.permissionToReleaseCheck || 'N/A'}</span>
            </p>
            
            </div>
            )}
          </div>
        );
      case 'employee_information':
        return (
          <div className="relative">
            <div className="sticky top-0 z-10 bg-white pb-4 border-b mb-4 flex justify-between items-center">
               <h3 className="text-2xl font-bold text-gray-800">Employee Information</h3>
               {!isEditing ? (
                 <button 
                   onClick={() => setIsEditing(true)} 
                   className="flex items-center gap-2 text-sm text-blue-600 hover:text-blue-800 font-semibold"
                 >
                   <Pencil size={16} /> Edit Details
                 </button>
               ) : (
                <div className="col-span-2 flex justify-end gap-3 mt-4">
                    <button onClick={handleCancelEdit} className="px-4 py-2 text-gray-600 hover:bg-gray-100 rounded">Cancel</button>
                    <button 
                      onClick={handleSave} 
                      disabled={isSaving}
                      className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700 disabled:bg-blue-400"
                    >
                      {isSaving ? 'Saving...' : <><Save size={16}/> Save Changes</>}
                    </button>
                 </div>
                )
                }
            </div>

            {isEditing ? (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                 <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Status</label>
                    <select name="status" value={formData.status || ''} onChange={handleInputChange} className="w-full border p-2 rounded">
                      <option value="" disabled hidden>Select Status</option>
                      {statusOptions.map(status => (
                        <option key={status} value={status}>{status}</option>
                      ))}
                    </select>
                 </div>
                 <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Date of Hire</label>
                    <input type="date" name="dateOfHire" value={formData.dateOfHire || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
                 <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Date of Termination</label>
                    <input type="date" name="dateOfTermination" value={formData.dateOfTermination || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
                 <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Department</label>
                    <select name="department" value={formData.department || ''} onChange={handleInputChange} className="w-full border p-2 rounded">
                      <option value="" disabled hidden>Select Department</option>
                      {departments.map(dept => (
                        <option key={dept} value={dept}>{dept}</option>
                      ))}
                    </select>
                 </div>
                 <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Job Title</label>
                    <input name="jobTitle" value={formData.jobTitle || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
                 <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Job Status</label>
                    <select name="jobStatus" value={formData.jobStatus || ''} onChange={handleInputChange} className="w-full border p-2 rounded">
                      <option value="" disabled hidden>Select Job Status</option>
                      {jobStatuses.map(status => (
                        <option key={status} value={status}>{status}</option>
                      ))}
                    </select>
                 </div>
                 <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Employee Status</label>
                    <select name="employmentStatus" value={formData.employmentStatus || ''} onChange={handleInputChange} className="w-full border p-2 rounded">
                      <option value="" disabled hidden>Select Employment Status</option>
                      {employmentStatuses.map(status => (
                        <option key={status} value={status}>{status}</option>
                      ))}
                    </select>
                 </div>
                 {userRole === 'Director' && (
                    <div className="space-y-1">
                      <label className="text-xs font-semibold text-gray-500">Pay</label>
                        <input name="pay" value={formData.pay || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                    </div>
                  )}
                 <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Office</label>
                    <select name="office" value={formData.office || ''} onChange={handleInputChange} className="w-full border p-2 rounded">
                      <option value="" disabled hidden>Select Office</option>
                      {offices.map(office => (
                        <option key={office} value={office}>{office}</option>
                      ))}
                    </select>
                 </div>
                  <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Employee SSN</label>
                    <input name="employeeSSN" value={formData.employeeSSN || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
                  <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Date of Birth</label>
                    <input name="dateOfBirth" value={formData.dateOfBirth || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
                  <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Other Job Titles</label>
                    <input name="otherJobTitles" value={formData.otherJobTitles || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
              </div>
            ) : (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-y-4 text-gray-700">
                <div className="space-y-3 text-gray-700">
                  <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">Status</span> {formData.status || 'N/A'}</p>
                  <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">Date of Hire:</span> {formData.dateOfHire || 'N/A'}</p>
                  <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">Date of Terminiation:</span> {formData.dateOfTermination || 'N/A'}</p>
                  <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">Department:</span> {formData.department || 'N/A'}</p>
                  <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">Job Title:</span> {formData.jobTitle || 'N/A'}</p>
                  <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">Job Status:</span> {formData.jobStatus || 'N/A'}</p>
                </div>
                <div className="space-y-3">
                  <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">Employee Status:</span> {formData.employmentStatus || 'N/A'}</p>
                    {userRole === 'Director' && (
                      <p className="flex">
                        <span className="font-semibold w-42 whitespace-nowrap shrink-0 text-gray-900">Pay:</span> 
                        <span className="text-blue-600 font-medium">{formData.pay || 'N/A'}</span>
                      </p>
                    )}
                  <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">Office:</span> {formData.office || 'N/A'}</p>
                  <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">Employee SSN:</span> {formData.employeeSSN || 'N/A'}</p>
                  <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">Date of Birth:</span> {formData.dateOfBirth || 'N/A'}</p>
                  <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">Other Job Titles:</span> {formData.otherJobTitles || 'N/A'}</p>
                </div>
              </div>
            )}
            <div className="flex justify-between items-center mb-4 pt-5">
               <h3 className="text-xl font-bold text-gray-800">Employee Training</h3>
            </div>
            {isEditing ? (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                 <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">Sexual Harassment Training</label>
                    <input name="sexualHarassmentTraining" type="date" value={formData.sexualHarassmentTraining || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
                 <div className="space-y-1">
                    <label className="text-xs font-semibold text-gray-500">OSHA Training</label>
                    <input name="oshaTraining" type="date" value={formData.oshaTraining || ''} onChange={handleInputChange} className="w-full border p-2 rounded" />
                 </div>
              </div>
            ) : (
              <div className="space-y-3 text-gray-700">
                <p className="flex"><span className="font-semibold w-55 whitespace-nowrap inline-block text-gray-900">Sexual Harassment Training:</span> {formData.sexualHarassmentTraining || 'N/A'}</p>
                <p className="flex"><span className="font-semibold w-55 whitespace-nowrap inline-block text-gray-900">OSHA Training:</span> {formData.oshaTraining || 'N/A'}</p>
              </div>
            )}
          </div>
        );
      
      // 3. UPDATED: Grouped all other tabs here so they don't crash
      case 'license_certificates':
      case 'incident_notices':
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
    <div className="fixed inset-0 bg-black/60 flex justify-center items-center z-50 p-4" onClick={handleBeforeClose}>
      <div className="bg-white rounded-lg shadow-2xl w-full max-w-6xl max-h-[90vh] flex flex-col" onClick={(e) => e.stopPropagation()}>
        
        {/* Header */}
        <div className="flex justify-between items-center p-6 border-b flex-shrink-0">
          <h2 className="text-2xl font-bold text-gray-900">Employee Information</h2>
          <button onClick={handleBeforeClose} className="text-gray-500 hover:text-gray-800"><X size={24} /></button>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 p-6 flex-grow lg:overflow-hidden min-h-0">
          
          {/* LEFT SIDEBAR */}
          <div className="lg:col-span-1 lg:border-r lg:pr-6 flex flex-col lg:overflow-hidden">
            <div className="flex flex-col items-center text-center gap-4 mb-6 flex-shrink-0 relative">
                {isEditing && (
                  <button
                    onClick={handleDeleteEmployee}
                    className="mt-4 flex items-center justify-center gap-2 px-4 py-2 text-red-600 hover:bg-red-50 rounded-md border border-red-200 transition-colors"
                  >
                    <X size={16} /> Delete Employee Record
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