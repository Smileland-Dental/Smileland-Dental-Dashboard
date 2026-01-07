'use client';

import React, { useState } from 'react';
import { collection, addDoc } from 'firebase/firestore';
import { ref, uploadBytes, getDownloadURL } from 'firebase/storage';
import { db, storage } from '@/lib/firebase.config';
import FeedbackModal from '@/components/ui/FeedbackModal'; // Import the new modal

interface AddEmployeeFormProps {
  onSuccess: () => void;
  onCancel: () => void;
}

export default function AddEmployeeForm({ onSuccess, onCancel }: AddEmployeeFormProps) {
  // ... existing form state (formData, imageFile, isSubmitting) ...
  const [formData, setFormData] = useState({
    name: '', role: '', office: '', dateOfHire: '', pay: '', tenure: '',
  });
  const [imageFile, setImageFile] = useState<File | null>(null);
  const [isSubmitting, setIsSubmitting] = useState(false);

  // NEW: State for the feedback modal
  const [feedback, setFeedback] = useState<{ isOpen: boolean; type: 'success' | 'error'; message: string }>({
    isOpen: false,
    type: 'success',
    message: '',
  });

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) => {
    setFormData({ ...formData, [e.target.name]: e.target.value });
  };

  const handleImageChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) setImageFile(e.target.files[0]);
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setIsSubmitting(true);

    try {
      let imageUrl = '';
      if (imageFile) {
        const storageRef = ref(storage, `profile-pictures/${Date.now()}-${imageFile.name}`);
        const snapshot = await uploadBytes(storageRef, imageFile);
        imageUrl = await getDownloadURL(snapshot.ref);
      }

      await addDoc(collection(db, 'employees'), {
        ...formData,
        pay: Number(formData.pay),
        imageUrl: imageUrl || null,
        createdAt: new Date(),
      });

      // Show Success Modal
      setFeedback({
        isOpen: true,
        type: 'success',
        message: `${formData.name} has been added to the directory.`,
      });
    } catch (error) {
      console.error(error);
      // Show Error Modal
      setFeedback({
        isOpen: true,
        type: 'error',
        message: 'Could not save employee data. Please check your connection.',
      });
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleFeedbackClose = () => {
    setFeedback({ ...feedback, isOpen: false });
    if (feedback.type === 'success') {
      onSuccess(); // Only trigger parent update if successful
    }
  };

  return (
    <>
      <div className="fixed inset-0 bg-black bg-opacity-50 flex justify-center items-center z-50">
        <div className="bg-white p-6 rounded-lg shadow-xl w-full max-w-md max-h-[90vh] overflow-y-auto">
          <h2 className="text-2xl font-bold mb-4">Add New Employee</h2>
          
          <form onSubmit={handleSubmit} className="space-y-4">
             {/* ... existing input fields ... */}
             <input required name="name" placeholder="Full Name" onChange={handleChange} className="w-full border p-2 rounded" />
             <input required name="role" placeholder="Role / Job Title" onChange={handleChange} className="w-full border p-2 rounded" />
             {/* ... rest of your form inputs ... */}

            <div className="flex justify-end gap-2 mt-6">
              <button type="button" onClick={onCancel} className="px-4 py-2 text-gray-600 hover:bg-gray-100 rounded">Cancel</button>
              <button type="submit" disabled={isSubmitting} className="px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700">
                {isSubmitting ? 'Saving...' : 'Add Employee'}
              </button>
            </div>
          </form>
        </div>
      </div>

      {/* The Feedback Modal */}
      <FeedbackModal 
        isOpen={feedback.isOpen} 
        type={feedback.type} 
        message={feedback.message} 
        onClose={handleFeedbackClose} 
      />
    </>
  );
}