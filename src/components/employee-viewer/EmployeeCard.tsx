import React from 'react';
import Image from 'next/image';
import { Employee } from '@/lib/types';
//import defaultProfile from '/public/defaultProfile.jpg';

interface EmployeeCardProps {
  employee: Employee;
  onClick: () => void;
}

const EmployeeCard: React.FC<EmployeeCardProps> = ({ employee, onClick }) => {
  const imageSrc = employee.imageURL || '/public/defaultProfile.jpg'
  return (
    <div
      onClick={onClick}
      className="bg-white rounded-lg shadow-md p-4 flex flex-col items-center text-center cursor-pointer hover:shadow-xl hover:scale-105 transition-transform duration-200"
    >
      <Image
        src={imageSrc}
        alt={employee.name}
        width={128}
        height={128}
        className="w-32 h-32 rounded-full object-cover mb-4 border-4 border-gray-200"
      />
      <h3 className="text-xl font-semibold text-gray-800">{employee.name}</h3>
      <p className="text-gray-500">{employee.role}</p>
    </div>
  );
};

export default EmployeeCard;