import React from 'react';
import { Pencil, Save } from 'lucide-react';

interface TabProps {
  formData: any;
  handleInputChange: (e: any) => void;
  isEditing: boolean;
  setIsEditing: (val: boolean) => void;
  handleCancelEdit: () => void;
  handleSave: () => void;
  isSaving: boolean;
  hasUnsavedChanges: boolean;
}

export const GeneralTab = ({ formData, handleInputChange, isEditing, setIsEditing, handleCancelEdit, handleSave, isSaving, hasUnsavedChanges }: TabProps) => {
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
                className={`flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700 disabled:bg-blue-400 transition-all ${
                hasUnsavedChanges && !isSaving ? 'animate-pulse-save shadow-lg' : ''
              }`}
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
}