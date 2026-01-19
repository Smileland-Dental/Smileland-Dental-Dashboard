/* 'use client'; // This directive is necessary for using hooks

import React, { useState, useEffect, useMemo } from 'react';
import { collection, getDocs } from 'firebase/firestore';
import { db } from '@/lib/firebase.config';
import { Employee, GroupedEmployees } from '@/lib/types';
import EmployeeCard from '@/components/employee-viewer/EmployeeCard';
import EmployeeModal from '@/components/employee-viewer/employee-modal';
//import defaultProfile from '/defaultProfile.jpg';
import Image from 'next/image';

export default function DirectoryPage() {
  const [employees, setEmployees] = useState<Employee[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);

  // State for filters and search
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedOffice, setSelectedOffice] = useState('All');
  const [selectedRole, setSelectedRole] = useState('All');

  // State for the modal
  const [selectedEmployee, setSelectedEmployee] = useState<Employee | null>(null);

  useEffect(() => {
    const fetchEmployees = async () => {
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
    fetchEmployees();
  }, []);

  // Memoize the filtering and grouping logic to avoid re-computation on every render
  const groupedAndSortedEmployees = useMemo(() => {
    let filteredEmployees = employees;

    // 1. Filter by search query (name or id)
    if (searchQuery) {
      filteredEmployees = filteredEmployees.filter(
        (emp) =>
          emp.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
          emp.id.toLowerCase().includes(searchQuery.toLowerCase())
      );
    }

    // 2. Filter by selected office
    if (selectedOffice !== 'All') {
      filteredEmployees = filteredEmployees.filter((emp) => emp.office === selectedOffice);
    }

    // 3. Filter by selected role
    if (selectedRole !== 'All') {
      filteredEmployees = filteredEmployees.filter((emp) => emp.role === selectedRole);
    }

    // 4. Group by office
    const grouped = filteredEmployees.reduce((acc, emp) => {
      const office = emp.office || 'Unassigned';
      if (!acc[office]) {
        acc[office] = [];
      }
      acc[office].push(emp);
      return acc;
    }, {} as GroupedEmployees);

    // 5. Sort employees within each office group by role
    for (const office in grouped) {
      grouped[office].sort((a, b) => a.role.localeCompare(b.role));
    }

    return grouped;
  }, [employees, searchQuery, selectedOffice, selectedRole]);
  
  // Get unique offices and roles for the filter dropdowns
  const uniqueOffices = useMemo(() => ['All', ...new Set(employees.map(e => e.office))], [employees]);
  const uniqueRoles = useMemo(() => ['All',...new Set(employees.map(e => e.role))], [employees]);

  const handleOpenModal = (employee: Employee) => {
    setSelectedEmployee(employee);
  };

  const handleCloseModal = () => {
    setSelectedEmployee(null);
  };

  if (loading) {
    return <div className="text-center py-10">Loading directory...</div>;
  }

  if (error) {
    return <div className="text-center py-10 text-red-500">{error}</div>;
  }

  return (
    <div className="container mx-auto px-4 py-8">
      <h1 className="text-4xl font-bold text-center mb-8">Company Directory</h1>
      <Image src="/defaultProfile.jpg" alt="Default Profile" width={100} height={100} />

      {/* Filter and Search Controls }
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
          value={selectedRole}
          onChange={(e) => setSelectedRole(e.target.value)}
          className="p-2 border border-gray-300 rounded-md w-full"
        >
          {uniqueRoles.map(role => <option key={role} value={role}>{role}</option>)}
        </select>
      </div>

      {/* Employee List }
      <div className="space-y-12">
        {Object.keys(groupedAndSortedEmployees).length > 0 ? (
          Object.keys(groupedAndSortedEmployees)
            .sort() // Sort office names alphabetically
            .map((office) => (
              <section key={office}>
                <h2 className="text-3xl font-semibold border-b-2 border-blue-500 pb-2 mb-6">
                  {office}
                </h2>
                <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-6">
                  {groupedAndSortedEmployees[office].map((employee) => (
                    <EmployeeCard
                      key={employee.id}
                      employee={employee}
                      onClick={() => handleOpenModal(employee)}
                    />
                    
                  ))}
                </div>
              </section>
            ))
        ) : (
          <p className="text-center text-gray-500 text-xl">No employees found matching your criteria.</p>
        )}
      </div>

      {/* Modal }
      {selectedEmployee && (
        <EmployeeModal employee={selectedEmployee} onClose={handleCloseModal} />
      )}
    </div>
  );
}
*/

'use client'; 

import React, { useState, useEffect, useMemo } from 'react';
import { collection, getDocs, doc } from 'firebase/firestore';
import { db, auth } from '@/lib/firebase.config';
import { Employee, GroupedEmployees } from '@/lib/types';
import EmployeeCard from '@/components/employee-viewer/EmployeeCard';
import EmployeeModal from '@/components/employee-viewer/employee-modal';
import AddEmployeeForm from '@/components/employee-viewer/AddEmployeeForm'; // Import the new form
import Image from 'next/image';
import { onAuthStateChanged, User } from "firebase/auth";


// MOCK AUTH STATE: In a real app, get this from your AuthProvider
const currentUser = {
  role: 'hr', // Change this to 'user' to test hiding the button
};

export default function DirectoryPage() {
  const [user, setUser] = useState<User | null>(null);
  const [userRole, setUserRole] = useState<string | "">("");
  const [employees, setEmployees] = useState<Employee[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    // This listener checks for changes in the user's login state.
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      setLoading(true);
      setError(null);
      if (currentUser) {
        // User is signed in
        setUser(currentUser);
      } else {
        // User is signed out
        setUser(null);
      }
    });
    // Cleanup the subscription when the component unmounts
    return () => unsubscribe();
  }, []); 

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      setLoading(true);
      setError(null);
      if (currentUser) {
        setUser(currentUser);

        const userDocRef = doc(db, 'users', currentUser.uid);

        try{
          const userDocSnap = await getDoc(userDocRef);
          if (userDocSnap.exists()) {
            const userData = userDocSnap.data();

            const name = currentUser.displayName || "";
            const role = userData.role || null;
            const offices = userData.offices || [];

            setUserName(name);
            setUserRole(role);
            setUserOffices(offices);
            // Fetch absence requests based on the user's offices and role
            if ((userData.offices).length > 0) {
              const requests = await getAbsenceRequestsByOffices(offices, userName, role);
              setAllRequests(requests as AbsenceRequest[]);
            }
            else{
              setAllRequests([]);
            }
          }
          else{
            setUserRole("");
          }
        }
        catch (err:any){
          console.error("Error fetching user data or requests:", err);
          setError("Failed to load data. Please try again later.");
        }
      }
      else {
        setUser(null);
        setUserRole("");
        setAllRequests([]);
      }
      setLoading(false);
    });
    return () => unsubscribe();
  }, []);

  // Filter States
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedOffice, setSelectedOffice] = useState('All');
  const [selectedRole, setSelectedRole] = useState('All');

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
          emp.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
          emp.id.toLowerCase().includes(searchQuery.toLowerCase())
      );
    }
    if (selectedOffice !== 'All') {
      filteredEmployees = filteredEmployees.filter((emp) => emp.office === selectedOffice);
    }
    if (selectedRole !== 'All') {
      filteredEmployees = filteredEmployees.filter((emp) => emp.role === selectedRole);
    }

    const grouped = filteredEmployees.reduce((acc, emp) => {
      const office = emp.office || 'Unassigned';
      if (!acc[office]) acc[office] = [];
      acc[office].push(emp);
      return acc;
    }, {} as GroupedEmployees);

    for (const office in grouped) {
      grouped[office].sort((a, b) => a.role.localeCompare(b.role));
    }

    return grouped;
  }, [employees, searchQuery, selectedOffice, selectedRole]);
  
  const uniqueOffices = useMemo(() => ['All', ...new Set(employees.map(e => e.office))], [employees]);
  const uniqueRoles = useMemo(() => ['All',...new Set(employees.map(e => e.role))], [employees]);

  return (
    <div className="container mx-auto px-4 py-8">
      <div className="flex justify-between items-center mb-8">
         <h1 className="text-4xl font-bold">Company Directory</h1>
         
         {/* Show Add Button ONLY if user is HR */}
         {currentUser.role === 'hr' && (
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
          value={selectedRole}
          onChange={(e) => setSelectedRole(e.target.value)}
          className="p-2 border border-gray-300 rounded-md w-full"
        >
          {uniqueRoles.map(role => <option key={role} value={role}>{role}</option>)}
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
        <EmployeeModal employee={selectedEmployee} onClose={() => setSelectedEmployee(null)} />
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