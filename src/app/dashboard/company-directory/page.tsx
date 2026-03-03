'use client'; 

import React, { useState, useEffect, useMemo } from 'react';
import { collection, getDocs, getDoc, doc } from 'firebase/firestore';
import { db, auth } from '@/lib/firebase.config';
import { Employee, GroupedEmployees } from '@/lib/types';
import EmployeeCard from '@/components/employee-viewer/EmployeeCard';
import EmployeeModal from '@/components/employee-viewer/employee-modal';
import AddEmployeeForm from '@/components/employee-viewer/AddEmployeeForm'; // Import the new form
import Image from 'next/image';
import { useAuth } from '@/contexts/AuthContext';

export default function DirectoryPage() {
  const { user, loading: authLoading } = useAuth();

  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);

  const [employees, setEmployees] = useState<Employee[]>([]);

  // Filter States
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedOffice, setSelectedOffice] = useState('All');
  const [selectedTitle, setSelectedTitle] = useState('All');
  //Need to Add filter by status of employee (active/terminated)

  // Modal States
  const [selectedEmployee, setSelectedEmployee] = useState<Employee | null>(null);
  const [isAddModalOpen, setIsAddModalOpen] = useState(false); // State for Add Form

  // 1. Refactored fetch logic into a function so we can call it after adding a user
  const fetchEmployees = async () => {
    setLoading(true);
    try {
      const querySnapshot = await getDocs(collection(db, 'employees'));
      const employeesData = querySnapshot.docs.map((doc) => ({
        id: doc.id,
        ...doc.data(),
      })) as Employee[];
      setEmployees(employeesData);
    } catch (err) {
      setError('Failed to fetch employee data.');
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchEmployees();
  }, []);

  // Memoize filters (Kept your existing logic)
  const groupedAndSortedEmployees = useMemo(() => {
    let filteredEmployees = employees;

    if (searchQuery) {
      filteredEmployees = filteredEmployees.filter(
        (emp) =>
          emp.firstName.toLowerCase().includes(searchQuery.toLowerCase()) ||
          emp.lastName.toLowerCase().includes(searchQuery.toLowerCase()) ||
          emp.employeeID.toLowerCase().includes(searchQuery.toLowerCase()) 
      );
    }
    if (selectedOffice !== 'All') {
      filteredEmployees = filteredEmployees.filter((emp) => emp.office === selectedOffice);
    }
    if (selectedTitle !== 'All') {
      filteredEmployees = filteredEmployees.filter((emp) => emp.jobTitle === selectedTitle);
    }

    const grouped = filteredEmployees.reduce((acc, emp) => {
      const office = emp.office || 'Unassigned';
      if (!acc[office]) acc[office] = [];
      acc[office].push(emp);
      return acc;
    }, {} as GroupedEmployees);

    for (const office in grouped) {
      grouped[office].sort((a, b) => a.jobTitle.localeCompare(b.jobTitle));
    }

    return grouped;
  }, [employees, searchQuery, selectedOffice, selectedTitle]);
  
  const uniqueOffices = useMemo(() => ['All', ...new Set(employees.map(e => e.office))], [employees]);
  const uniqueTitles = useMemo(() => ['All',...new Set(employees.map(e => e.jobTitle))], [employees]);

  if (authLoading) return <div className="text-center py-20">Loading</div>;

  return (
    <div className="container mx-auto px-4 py-8">
      <div className="flex justify-between items-center mb-8">
         <h1 className="text-4xl font-bold">Company Directory</h1>
         {user?.role}
         {/* Show Add Button ONLY if user is HR */}
         {user?.role === 'Manager' && (
           <button 
             onClick={() => setIsAddModalOpen(true)}
             className="bg-green-600 text-white px-4 py-2 rounded hover:bg-green-700 transition"
           >
             + Add Employee
           </button>
         )}
      </div>

      {/* Filter Controls (Kept your existing code) */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-8 p-4 bg-gray-100 rounded-lg">
        <input
          type="text"
          placeholder="Search by name or ID..."
          value={searchQuery}
          onChange={(e) => setSearchQuery(e.target.value)}
          className="p-2 border border-gray-300 rounded-md w-full"
        />
        <select
          value={selectedOffice}
          onChange={(e) => setSelectedOffice(e.target.value)}
          className="p-2 border border-gray-300 rounded-md w-full"
        >
          {uniqueOffices.map(office => <option key={office} value={office}>{office}</option>)}
        </select>
        <select
          value={selectedTitle}
          onChange={(e) => setSelectedTitle(e.target.value)}
          className="p-2 border border-gray-300 rounded-md w-full"
        >
          {uniqueTitles.map(title => <option key={title} value={title}>{title}</option>)}
        </select>
      </div>

      {/* Loading & Error States */}
      {loading && <div className="text-center py-10">Loading directory...</div>}
      {error && <div className="text-center py-10 text-red-500">{error}</div>}

      {/* Employee List */}
      {!loading && !error && (
          <div className="space-y-12">
            {Object.keys(groupedAndSortedEmployees).length > 0 ? (
              Object.keys(groupedAndSortedEmployees).sort().map((office) => (
                <section key={office}>
                  <h2 className="text-3xl font-semibold border-b-2 border-blue-500 pb-2 mb-6">
                    {office}
                  </h2>
                  <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-6">
                    {groupedAndSortedEmployees[office].map((employee) => (
                      <EmployeeCard
                        key={employee.id}
                        employee={employee}
                        onClick={() => setSelectedEmployee(employee)}
                      />
                    ))}
                  </div>
                </section>
              ))
            ) : (
              <p className="text-center text-gray-500 text-xl">No employees found.</p>
            )}
          </div>
      )}

      {/* Employee Detail Modal */}
      {selectedEmployee && (
        <EmployeeModal employee={selectedEmployee} userRole={user?.role || "User"} onClose={() => setSelectedEmployee(null)} onUpdate={fetchEmployees}/>
      )}

      {/* Add Employee Form Modal */}
      {isAddModalOpen && (
        <AddEmployeeForm 
          onSuccess={() => {
            setIsAddModalOpen(false);
            fetchEmployees(); // Refresh list after adding
          }} 
          onCancel={() => setIsAddModalOpen(false)} 
        />
      )}
    </div>
  );
}

