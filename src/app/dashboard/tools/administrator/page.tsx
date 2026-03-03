'use client';

import { collection, onSnapshot, query } from 'firebase/firestore';
import { db } from '@/lib/firebase.config';
import { Search, Loader2 } from "lucide-react";
import { useState, useEffect } from "react";
import UserCard from '@/components/administrator/UserCard';
import UserModal from '@/components/administrator/user-modal';
import { useAuth } from '@/contexts/AuthContext';
import { Employee, User } from '@/lib/types';

export default function AdministratorUsersPage() {
  const { user: loggedInUser } = useAuth();
  const [allUsers, setAllUsers] = useState<User[]>([]);
  const [allEmployees, setAllEmployees] = useState<Employee[]>([]);
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedUser, setSelectedUser] = useState<User | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    // 1. Set up Real-time Listener for Users
    // This replaces fetchUsers() entirely.
    const qUsers = query(collection(db, "users"));
    const unsubscribeUsers = onSnapshot(qUsers, (snapshot) => {
      const usersData = snapshot.docs.map(doc => ({ 
        id: doc.id, 
        ...doc.data() 
      })) as User[];
      
      setAllUsers(usersData);
      setLoading(false);

      // CRITICAL: If a user is currently open in a modal, 
      // this ensures the modal updates live if someone else edits them!
      if (selectedUser) {
        const updated = usersData.find(u => u.id === selectedUser.id);
        if (updated) setSelectedUser(updated);
      }
    });

    // 2. Set up Real-time Listener for Employees
    const qEmps = query(collection(db, "employees"));
    const unsubscribeEmps = onSnapshot(qEmps, (snapshot) => {
      setAllEmployees(snapshot.docs.map(doc => ({ 
        id: doc.id, 
        ...doc.data() 
      })) as Employee[]);
    });

    // 3. Cleanup: Stop listening when the admin leaves the page
    return () => {
      unsubscribeUsers();
      unsubscribeEmps();
    };
  }, [selectedUser?.id]); // Re-sync modal if selectedUser changes

  const filteredUsers = allUsers.filter(u =>
    u
    //u.username?.toLowerCase().includes(searchTerm.toLowerCase())
  );

  return (
    <div className="p-4 min-h-screen bg-gray-50">
      <div className="max-w-6xl mx-auto">
        <header className="flex justify-between items-center mb-8">
          <div>
            <h1 className="text-3xl font-bold text-gray-900">Dashboard User Setup</h1>
            <p className="text-gray-500">Adjust the user roles and managed staff.</p>
          </div>
          <div className="relative w-64">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400 h-4 w-4" />
            <input 
              placeholder="Search users..."
              className="pl-10 pr-4 py-2 border rounded-xl w-full outline-none focus:ring-2 focus:ring-blue-500"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
        </header>

        {loading ? (
          <div className="flex flex-col items-center justify-center py-20">
            <Loader2 className="animate-spin text-blue-500 mb-2" size={32} />
            <p className="text-gray-500">Syncing with database...</p>
          </div>
        ) : (
          <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-6">
            {filteredUsers.map(userItem => (
              <UserCard
                key={userItem.id}
                user={userItem}
                onClick={() => setSelectedUser(userItem)}
              />
            ))}
          </div>
        )}
      </div>
      
      {selectedUser && (
        <UserModal
          user={selectedUser} 
          allEmployees={allEmployees}
          loggedInUserRole={loggedInUser?.role || 'User'} 
          onClose={() => setSelectedUser(null)}
          // NOTICE: No need to call fetchUsers here anymore!
          // Firestore pushes the update automatically.
          onUpdate={() => {}} 
        />
      )}
    </div>
  );
}