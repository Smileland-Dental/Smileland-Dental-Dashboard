import React from 'react';
import { Pencil, Save } from 'lucide-react';

interface LicenseCertificatesProps {
  isEditing: boolean;
  setIsEditing: (val: boolean) => void;
  handleCancelEdit: () => void;
}

const handleLicenseCertificatesSave = () => {
  // Implement save logic here, e.g., update Firestore with new license/certificate info
  console.log("Saving license/certificate information...");
}

export const LicenseCertificatesTab = ({ isEditing, setIsEditing, handleCancelEdit }: LicenseCertificatesProps) => {
  return (
    <div className="relative">
      <div className="sticky top-0 z-10 bg-white pb-4 border-b mb-4 flex justify-between items-center">
        <h3 className="text-2xl font-bold text-gray-800">License/Certificates</h3>
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
              <button onClick={handleLicenseCertificatesSave} className="px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700 flex items-center gap-2">
                Save Changes
              </button>
            </div>
          )
          }
      </div>
      
      {isEditing ? (
        <div className="p-4 bg-yellow-50 border-l-4 border-yellow-400 text-yellow-700">
          <p className="mb-4">You are in edit mode. Make your changes and click "Save" to update the employee's license and certificate information.</p>
        </div>
        ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 gap-y-4 text-gray-700">
          <div className="space-y-1 md:col-span-2">
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">CPR Renewal:</span> 'N/A'</p>
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">Hep. B Status:</span> 'N/A'</p>
          g</div>
          <div className="space-y-3 text-gray-700">

            <h3 className="text-xl font-bold text-gray-800">Doctor</h3>
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">Title:</span> 'N/A'</p>
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">License #:</span> 'N/A'</p>
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">Expiration Date:</span> 'N/A'</p>
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">OCS #:</span> 'N/A'</p>
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">OCS Expiration Date:</span> 'N/A'</p>
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">DEA #:</span> 'N/A'</p>
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">DEA Expiration Date:</span> 'N/A'</p>
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">Malpractice #:</span> 'N/A'</p>
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">Expiration Date:</span> 'N/A'</p>
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">NPI #:</span> 'N/A'</p>
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">ADA/CDA #:</span> 'N/A'</p>
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">ADA/CDA Renewal:</span> 'N/A'</p>
          </div>
          <div className="space-y-3">
            <h3 className="text-xl font-bold text-gray-800">Registered Dental Assistants</h3>
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">License #:</span> 'N/A'</p>
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">Expiration Date:</span> 'N/A'</p>

            <h3 className="text-xl font-bold text-gray-800">Dental Assistants</h3>
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">DA Certificate:</span> 'N/A'</p>
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">California Dental Practice Act (CDPA):</span> 'N/A'</p>
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">Infection Control:</span> 'N/A'</p>
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">Radiation Certificate:</span> 'N/A'</p>
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">Pit and Fissure:</span> 'N/A'</p>
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">Coronal Polishing:</span> 'N/A'</p>
            <p className="flex"><span className="font-semibold w-42 whitespace-nowrap inline-block text-gray-900">HIPAA Training:</span> 'N/A'</p>
          </div>
        </div>
      )}
    </div>
  );
};