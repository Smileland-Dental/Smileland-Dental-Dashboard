'use client';

import React, { useState } from 'react';
// 1. CHANGE: Import setDoc and doc instead of addDoc
import { doc, getDocs, setDoc, addDoc, collection, query, where } from 'firebase/firestore';
import { ref, uploadBytes, getDownloadURL } from 'firebase/storage';
import { db, storage } from '@/lib/firebase.config';
import FeedbackModal from '@/components/ui/FeedbackModal';

interface AddEmployeeFormProps {
  onSuccess: () => void;
  onCancel: () => void;
}

export default function AddEmployeeForm({ onSuccess, onCancel }: AddEmployeeFormProps) {
  const [formData, setFormData] = useState({
    jobTitle: '',
    office: '',
    dateOfHire: '',
    employeeID: '', 
    jobStatus: '',
    employmentStatus: '',
    department: '',
    firstName: '',
    lastName: '',
    dateOfBirth: '',
    skipManagerApproval: false,
  });
  
  const [imageFile, setImageFile] = useState<File | null>(null);
  const [isSubmitting, setIsSubmitting] = useState(false);

  // Feedback Modal State
  const [feedback, setFeedback] = useState<{ isOpen: boolean; type: 'success' | 'error'; message: string }>({
    isOpen: false,
    type: 'success',
    message: '',
  });

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) => {
    setFormData({ ...formData, [e.target.name]: e.target.value });
  };

  const handleImageChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
        if (e.target.files[0].size > 2 * 1024 * 1024) {
            alert("File is too large. Max 2MB.");
            e.target.value = "";
            setImageFile(null);
            return;
        }
        setImageFile(e.target.files[0]);
    }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setIsSubmitting(true);

    try {
      // 1. Safety check: ensure we have an ID to use
      const q = query(
        collection(db, 'employees'), 
        where('employeeID', '==', String(formData.employeeID))
      );
      
      const querySnapshot = await getDocs(q);

      if (!querySnapshot.empty) {
        setFeedback({
          isOpen: true,
          type: 'error',
          message: `Employee ID "${formData.employeeID}" is already in use by another staff member.`,
        });
        setIsSubmitting(false);
        return; // Stop the submission here
      }

      // 2. GENERATE THE FIRESTORE ID FIRST
      // This creates the reference (and the ID) without writing to the database yet.
      const newEmployeeRef = doc(collection(db, 'employees'));
      const firestoreId = newEmployeeRef.id;

      // 3. Image Upload: If there's an image, upload it using the firestore id and get the URL
      let imageURL = '';
      if (imageFile) {
        // We use the firestoreId here so the photo is permanently linked to this record
        const storageRef = ref(storage, `employee-pictures/${firestoreId}/profile-picture`);
        const snapshot = await uploadBytes(storageRef, imageFile);
        imageURL = await getDownloadURL(snapshot.ref);
      }

      // 4. CHANGE: Use addDoc to create a new document with an auto-generated ID, instead of setDoc with a specific ID
      // doc(db, 'collectionName', 'specificID')
      await setDoc(newEmployeeRef, {
        ...formData,
        id: firestoreId,
        employeeID: String(formData.employeeID),
        imageURL: imageURL || null,
        status: 'Current', // Default status for new employees
        skipManagerApproval: formData.skipManagerApproval, // Explicitly included
        createdAt: new Date(),
      });

      setFeedback({
        isOpen: true,
        type: 'success',
        message: `${formData.firstName} ${formData.lastName} has been added successfully.`,
      });
    } catch (error) {
      console.error(error);
      setFeedback({
        isOpen: true,
        type: 'error',
        message: 'Could not save employee data. This ID might already exist or connection failed.',
      });
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleFeedbackClose = () => {
    setFeedback({ ...feedback, isOpen: false });
    if (feedback.type === 'success') {
      onSuccess(); 
    }
  };

  const offices = ["Corporate", "Ortho", "California", "Ming", "Bernard", "Delano", "Tulare", "Visalia", "Fresno"];
  const departments = ["Back Office", "Front Office", "Support Services", "Management", "Call Center", "Accounts Receivable", "Dentist"];
  const jobStatuses = ["Full-Time", "Part-Time"];
  const employmentStatuses = ["Exempt", "Non-Exempt", "Contracted"];

  return (
    <>
      <div className="fixed inset-0 bg-black/50 flex justify-center items-center z-50" onClick={onCancel}>
        <div className="bg-white p-6 rounded-lg shadow-xl w-full max-w-lg max-h-[90vh] overflow-y-auto" onClick={(e) => e.stopPropagation()}>
          <h2 className="text-2xl font-bold mb-4">Add New Employee</h2>
          
          <form onSubmit={handleSubmit} className="space-y-4">
            
            <div className="space-y-1">
                <label className="block text-sm font-medium text-gray-700 underline">Employee Picture</label>
                <input 
                    type="file" 
                    accept="image/*"
                    onChange={handleImageChange}
                    className="w-full text-sm text-gray-500
                      file:mr-4 file:py-2 file:px-4
                      file:rounded-full file:border-0
                      file:text-sm file:font-semibold
                      file:bg-blue-50 file:text-blue-700
                      hover:file:bg-blue-100"
                />
            </div>

            <div className="grid grid-cols-2 gap-4">
              <input required name="firstName" placeholder="First Name" onChange={handleChange} className="w-full border p-2 rounded" />
              <input required name="lastName" placeholder="Last Name" onChange={handleChange} className="w-full border p-2 rounded" />
            </div>

            <div className="grid grid-cols-2 gap-4">
              <input 
                  required 
                  inputMode="numeric"
                  name="employeeID" 
                  type="number" 
                  placeholder="Employee ID" 
                  onChange={handleChange} 
                  className="w-full border p-2 rounded" 
              />
              <input required name="jobTitle" placeholder="Job Title" onChange={handleChange} className="w-full border p-2 rounded" />
            </div>
            
            <div className="grid grid-cols-2 gap-4">
              {/* Office Dropdown */}
              <div className="space-y-1">
                <label className="text-xs text-gray-500">Office</label>
                <select 
                  required 
                  name="office" 
                  value={formData.office}
                  onChange={handleChange} 
                  className="w-full border p-2 rounded bg-white text-gray-700 h-[42px]" // h-[42px] matches standard input height
                >
                  <option value="" disabled hidden>Select Office</option>
                  {offices.map(loc => (
                      <option key={loc} value={loc}>{loc}</option>
                    ))}
                </select>
              </div>

              {/* Department Input */}
              <div className="space-y-1">
                <label className="text-xs text-gray-500">Department</label>
                <select 
                  required 
                  name="department" 
                  value={formData.department}
                  onChange={handleChange} 
                  className="w-full border p-2 rounded bg-white text-gray-700 h-[42px]"
                >
                  <option value="" disabled hidden>Select Department</option>
                  {departments.map(dept => (
                    <option key={dept} value={dept}>{dept}</option>
                  ))}
                </select>
              </div>
            </div>

            <div className="grid grid-cols-2 gap-4">
                <div className="space-y-1">
                    <label className="text-xs text-gray-500">Date of Hire</label>
                    <input required name="dateOfHire" type="date" onChange={handleChange} className="w-full border p-2 rounded" />
                </div>
                  <div className="space-y-1">
                    <label className="text-xs text-gray-500">Date of Birth</label>
                    <input required name="dateOfBirth" type="date" onChange={handleChange} className="w-full border p-2 rounded" />
                </div>
            </div>

            <div className="grid grid-cols-2 gap-4">
              {/* Job Status Input */}
              <div className="space-y-1">
                <label className="text-xs text-gray-500">Job Status</label>
                <select 
                  required 
                  name="jobStatus" 
                  value={formData.jobStatus}
                  onChange={handleChange} 
                  className="w-full border p-2 rounded bg-white text-gray-700 h-[42px]"
                >
                  <option value="" disabled hidden>Select Job Status</option>
                  {jobStatuses.map(status => (
                    <option key={status} value={status}>{status}</option>
                  ))}
                </select>
              </div>
              {/* Employment Status Input */}
              <div className="space-y-1">
                <label className="text-xs text-gray-500">Employment Status</label>
                <select 
                  required
                  name="employmentStatus" 
                  value={formData.employmentStatus}
                  onChange={handleChange} 
                  className="w-full border p-2 rounded bg-white text-gray-700 h-[42px]"
                >
                  <option value="" disabled hidden>Select Employment Status</option>
                  {employmentStatuses.map(status => (
                    <option key={status} value={status}>{status}</option>
                  ))}
                </select>
              </div>  
            </div>

            {/* Manager Approval Override Toggle */}
            <div className="p-3 bg-gray-50 rounded-lg border border-gray-200 flex items-center justify-between">
              <div className="space-y-0.5">
                <label className="text-sm font-semibold text-gray-700">Skip Manager Approval</label>
                <p className="text-xs text-gray-500">Enable to Bypass Manager Review for Absence Requests.</p>
              </div>
              <button
                type="button"
                onClick={() => setFormData(prev => ({ ...prev, skipManagerApproval: !prev.skipManagerApproval }))}
                className={`relative inline-flex h-6 w-11 items-center rounded-full transition-colors focus:outline-none ${
                  formData.skipManagerApproval ? 'bg-blue-600' : 'bg-gray-300'
                }`}
              >
                <span
                  className={`inline-block h-4 w-4 transform rounded-full bg-white transition-transform ${
                    formData.skipManagerApproval ? 'translate-x-6' : 'translate-x-1'
                  }`}
                />
              </button>
            </div>

            <div className="flex justify-end gap-2 mt-6">
              <button type="button" onClick={onCancel} className="px-4 py-2 text-gray-600 hover:bg-gray-100 rounded">Cancel</button>
              <button type="submit" disabled={isSubmitting} className="px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700 disabled:bg-blue-300">
                {isSubmitting ? 'Saving...' : 'Add Employee'}
              </button>
            </div>
          </form>
        </div>
      </div>

      <FeedbackModal 
        isOpen={feedback.isOpen} 
        type={feedback.type} 
        message={feedback.message} 
        onClose={handleFeedbackClose} 
      />
    </>
  );
}