import React from 'react';
import { Pencil, Save } from 'lucide-react';
import { Employee } from '@/lib/types';

interface TabProps {
  formData: Employee;
  handleInputChange: (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>) => void;
  setFormData: React.Dispatch<React.SetStateAction<Employee>>;
  isEditing: boolean;
  setIsEditing: (val: boolean) => void;
  handleCancelEdit: () => void;
  handleSave: () => void;
  isSaving: boolean;
  hasUnsavedChanges: boolean;
  userRole: string;
  // Options for selects
  statusOptions: string[];
  departments: string[];
  jobStatuses: string[];
  employmentStatuses: string[];
  offices: string[];
}

export const EmployeeInfoTab = ({
  formData, handleInputChange, setFormData, isEditing, setIsEditing, 
  handleCancelEdit, handleSave, isSaving, hasUnsavedChanges, userRole,
  statusOptions, departments, jobStatuses, employmentStatuses, offices
}: TabProps) => {
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
                className={`flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700 disabled:bg-blue-400 transition-all ${
                  hasUnsavedChanges && !isSaving ? 'animate-pulse-save shadow-lg' : ''
                }`}
              >
                {isSaving ? 'Saving...' : <><Save size={16}/> Save Changes</>}
              </button>
            </div>
          )}
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

          <div className="space-y-1 md:col-span-2 mt-4 p-4 bg-gray-50 rounded-lg border border-gray-200 flex items-center justify-between">
            <div className="space-y-0.5">
              <label className="text-sm font-bold text-gray-700">Skip Manager Approval</label>
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
            <p className="flex items-center min-h-[32px]">
              <span className="font-semibold w-42 shrink-0 text-gray-900">Manager Approval:</span> 
              <span className="flex-grow">
                {formData.skipManagerApproval ? (
                  <span className="text-xs font-bold text-amber-700 bg-amber-50 px-2 py-1 rounded border border-amber-200 inline-block leading-none">
                    Skipped (Auto-Approve)
                  </span>
                ) : (
                  <span className="text-xs font-bold text-green-700 bg-green-50 px-2 py-1 rounded border border-green-200 inline-block leading-none">
                    Required
                  </span>
                )}
              </span>
            </p>
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
}