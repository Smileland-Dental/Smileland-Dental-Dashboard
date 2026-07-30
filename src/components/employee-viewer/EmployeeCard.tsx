import React from 'react';
import Image from 'next/image';
import { Employee } from '@/lib/types';

interface EmployeeCardProps {
  employee: Employee;
  onClick: () => void;
}

const EmployeeCard: React.FC<EmployeeCardProps> = ({ employee, onClick }) => {
  const defaultProfileURL = 'https://firebasestorage.googleapis.com/v0/b/smileland-dental-dashboard.firebasestorage.app/o/employee-pictures%2F.DefaultProfile%2Fprofile.png?alt=media&token=70b1a79f-1a33-4b6c-9b9b-c8e3debd209d';
  const imageSrc = employee.imageURL || defaultProfileURL;
  return (
    <div
      onClick={onClick}
      className="bg-white rounded-lg shadow-md p-4 flex flex-col items-center text-center cursor-pointer hover:shadow-xl hover:scale-105 transition-transform duration-200"
    >
      <Image
        src={imageSrc}
        alt={employee.firstName + " " + employee.lastName || 'Employee Image'}
        width={128}
        height={128}
        className="w-32 h-32 rounded-full object-cover mb-4 border-4 border-gray-200"
      />
      <h3 className="text-xl font-semibold text-gray-800">{`${employee.firstName || 'No Name'} ${employee.lastName || ''}`.trim() || 'No Name'}</h3>
      <p className="text-gray-500">{employee.jobTitle}</p>
    </div>
  );
};

export default EmployeeCard;