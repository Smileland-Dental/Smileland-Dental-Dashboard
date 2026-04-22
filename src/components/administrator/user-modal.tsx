import React, { useState, useMemo } from 'react';
import { db } from '@/lib/firebase.config';
import { User, Employee } from '@/lib/types';
import { doc, updateDoc, arrayUnion, arrayRemove } from 'firebase/firestore';
import { X, Users, Info, Trash2, Loader2, CheckCircle2, Search, UserPlus, MapPin, Building2, RotateCcw } from 'lucide-react';

interface UserModalProps {
  user: User;
  allEmployees: Employee[];
  allUsers: User[]; // Added to check for linked employees
  loggedInUserRole: string;
  onClose: () => void;
  onUpdate?: () => void;
}

const OFFICE_LIST = ["Fresno", "Visalia", "Tulare", "Delano", "Bernard", "Ming", "California", "Ortho"];

const UserModal: React.FC<UserModalProps> = ({ user, allEmployees, allUsers,loggedInUserRole, onClose, onUpdate }) => {
  const [activeTab, setActiveTab] = useState<'info' | 'managed'>('info');
  const [loading, setLoading] = useState(false);
  const [empSearch, setEmpSearch] = useState('');
  const [selectedForBulk, setSelectedForBulk] = useState<string[]>([]);
  const [linkSearch, setLinkSearch] = useState('');

const filteredEmployeesForLink = useMemo(() => {
  // 1. Safety check: default to empty array if allUsers is undefined
  const usersList = allUsers || [];
  const term = linkSearch.toLowerCase();

  // 2. Create a set of IDs already taken by OTHER users
  // Added optional chaining (?.) for extra safety
  const takenEmployeeIds = new Set(
    usersList
      .filter(u => u.id !== user.id) 
      .map(u => u.linkedEmployeeId)
      .filter(Boolean)
  );

  return allEmployees.filter(emp => {
    // 3. Is this employee record already claimed?
    const isAlreadyLinked = takenEmployeeIds.has(emp.id); 
    
    // 4. Search logic (with optional chaining on employee fields)
    const matches = (emp.firstName?.toLowerCase().includes(term) || 
                     emp.lastName?.toLowerCase().includes(term) || 
                     emp.employeeID?.toString().includes(term));
                     
    return !isAlreadyLinked && matches;
  });
}, [allEmployees, allUsers, linkSearch, user.id]);

// Inside UserModal
const handleLinkEmployee = async (employeeId: string | null) => {
  setLoading(true);
  try {
    const newId = (employeeId && employeeId.trim() !== "") ? employeeId : null;

    // 1. Prepare the update object for the User
    const userUpdates: any = { 
      linkedEmployeeId: newId 
    };

    // 2. CRITICAL FIX: If we are linking an employee, 
    // remove them from the managed list immediately
    if (newId && user.managedEmployeeIds?.includes(newId)) {
      userUpdates.managedEmployeeIds = arrayRemove(newId);
    }

    // 3. Update the User Document
    await updateDoc(doc(db, 'users', user.id), userUpdates);

    // 4. Update the NEWLY linked employee record
    if (newId) {
      await updateDoc(doc(db, 'employees', newId), { 
        linkedUserId: user.id 
      });
    } 

    // 5. Clean up the OLD employee record (if switching)
    if (user.linkedEmployeeId && user.linkedEmployeeId !== newId) {
       await updateDoc(doc(db, 'employees', user.linkedEmployeeId), { 
         linkedUserId: null 
       });
    }

    setLinkSearch('');
    if (onUpdate) onUpdate();
  } catch (err) {
    console.error("Linking failed:", err);
    alert("Database Error: Could not update the link.");
  } finally {
    setLoading(false);
  }
};

  // 1. Update the core permission check
  // Only Directors can edit HR/Directors. 
  // HR can only edit Managers/Employees.
  const isEditingHigherOrEqualPower = (loggedInUserRole === 'HR' && ['HR', 'Director', 'Administrator'].includes(user.role)) || (loggedInUserRole === 'Director' && ['Director', 'Administrator'].includes(user.role));
  const canEdit = (loggedInUserRole === 'HR' || loggedInUserRole === 'Director') && !isEditingHigherOrEqualPower;

  // 1. Role Logic
  // 2. Update the availableRoles Memo
  const availableRoles = useMemo(() => {
    const roles = ["Employee", "Manager", "HR", "Director", "Administrator"];
    
    // If HR is looking at another HR/Director, they get NO roles (can't change them)
    if (isEditingHigherOrEqualPower) return [];

    if (loggedInUserRole === 'HR') {
      return roles.filter(r => ["Employee", "Manager"].includes(r));
    }
    if (loggedInUserRole === 'Director') {
      return roles.filter(r => ["Employee", "Manager", "HR"].includes(r));
    }
    return [];
  }, [loggedInUserRole, user.role, isEditingHigherOrEqualPower]);

  const managedStaff = useMemo(() => 
    allEmployees.filter(emp => user.managedEmployeeIds?.includes(emp.id)),
    [allEmployees, user.managedEmployeeIds]
  );

  // 2. Filter available employees for the dropdown/search
  const filteredAvailableToAssign = useMemo(() => {
    const unassigned = allEmployees.filter(emp => 
      // 1. Not already in the managed list
      !user.managedEmployeeIds?.includes(emp.id) && 
      // 2. Not currently sitting in the bulk staging area
      !selectedForBulk.includes(emp.id) &&
      // 3. ADDED: Not the user's own linked profile
      emp.id !== user.linkedEmployeeId
    );

    if (!empSearch) return unassigned;
    const term = empSearch.toLowerCase();
    return unassigned.filter(emp => 
      emp.firstName?.toLowerCase().includes(term) || 
      emp.lastName?.toLowerCase().includes(term) || 
      emp.employeeID?.toString().toLowerCase().includes(term)
    );
  }, [allEmployees, user.managedEmployeeIds, user.linkedEmployeeId, empSearch, selectedForBulk]);

  // 3. Office-based "Select All" Logic
const officeAvailability = useMemo(() => {
  const stats: Record<string, string[]> = {};
  OFFICE_LIST.forEach(office => {
    const ids = allEmployees
      .filter(emp => 
        emp.office === office && 
        !user.managedEmployeeIds?.includes(emp.id) && 
        !selectedForBulk.includes(emp.id) &&
        emp.id !== user.linkedEmployeeId // ADDED THIS LINE
      )
      .map(e => e.id);
    if (ids.length > 0) stats[office] = ids;
  });
  return stats;
}, [allEmployees, user.managedEmployeeIds, user.linkedEmployeeId, selectedForBulk]);

  const bulkSelectionList = allEmployees.filter(emp => selectedForBulk.includes(emp.id));

  // Firebase Operations
  const updateUserData = async (field: string, value: any) => {
    setLoading(true);
    try {
      await updateDoc(doc(db, 'users', user.id), { [field]: value });
      if (onUpdate) onUpdate();
    } catch (err) { console.error(err); } finally { setLoading(false); }
  };

const handleBulkAdd = async () => {
  setLoading(true);
  try {
    // Filter out the linkedEmployeeId from the bulk selection before saving
    const safeSelection = selectedForBulk.filter(id => id !== user.linkedEmployeeId);
    
    await updateDoc(doc(db, 'users', user.id), {
      managedEmployeeIds: arrayUnion(...safeSelection)
    });
    setSelectedForBulk([]);
    if (onUpdate) onUpdate();
  } catch (err) { 
    console.error(err); 
  } finally { 
    setLoading(false); 
  }
};

  return (
    <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-50 p-4" onClick={onClose}>
      <div className="bg-white rounded-xl shadow-2xl w-full max-w-lg overflow-hidden relative" onClick={(e) => e.stopPropagation()}>
        
        {/* Header */}
        <div className="p-6 border-b bg-gray-50 flex justify-between items-start">
          <div>
            <h2 className="text-2xl font-bold text-gray-800">{user.username}</h2>
            <p className="text-sm text-gray-500">{user.email}</p>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 transition-colors"><X size={24} /></button>
        </div>

        {/* Tabs */}
        <div className="flex border-b">
          <button onClick={() => setActiveTab('info')} className={`flex-1 py-3 text-sm font-medium ${activeTab === 'info' ? 'border-b-2 border-blue-600 text-blue-600' : 'text-gray-500'}`}>Details</button>
          <button onClick={() => setActiveTab('managed')} className={`flex-1 py-3 text-sm font-medium ${activeTab === 'managed' ? 'border-b-2 border-blue-600 text-blue-600' : 'text-gray-500'}`}>Managed Staff ({user.managedEmployeeIds?.length || 0})</button>
        </div>

        <div className="p-6 max-h-[600px] overflow-y-auto">
          {activeTab === 'info' ? (
             <div className="space-y-6">
                {/* User Role Selection */}
                <div>
                  <label className="text-xs font-bold text-gray-400 uppercase block mb-2">User Role</label>
                  {canEdit ? (
                    <select 
                      value={user.role} 
                      onChange={(e) => updateUserData('role', e.target.value)}
                      disabled={loading}
                      className="w-full p-2 border rounded-lg bg-white outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      {!availableRoles.includes(user.role) && <option value={user.role}>{user.role}</option>}
                      {availableRoles.map(r => <option key={r} value={r}>{r}</option>)}
                    </select>
                  ) : <p className="font-semibold text-blue-600">{user.role}</p>}
                </div>

                {/* Tie to Employee Database */}
                {/* Employee Linking Section */}
{/* Employee Linking Section */}
<div className="relative">
  <label className="text-xs font-bold text-gray-400 uppercase block mb-2">
    Link Employee Profile
  </label>
  
  {user.linkedEmployeeId ? (
    /* VIEW: Profile is already linked */
    <div className="flex items-center justify-between p-3 bg-green-50 border border-green-200 rounded-lg shadow-sm">
      <div className="flex items-center gap-3">
        <div className="bg-green-100 p-2 rounded-full text-green-600">
          <CheckCircle2 size={18} />
        </div>
        <div>
          <p className="text-sm font-bold text-green-800">
            {allEmployees.find(e => e.id === user.linkedEmployeeId)?.firstName}{' '}
            {allEmployees.find(e => e.id === user.linkedEmployeeId)?.lastName}
          </p>
          <p className="text-[10px] text-green-600 font-medium uppercase tracking-wider">
            Successfully Linked to HR Record
          </p>
        </div>
      </div>
      <button 
        onClick={() => handleLinkEmployee(null)}
        disabled={loading}
        className="text-gray-400 hover:text-red-500 transition-colors p-1"
        title="Unlink Profile"
      >
        <RotateCcw size={16} className={loading ? 'animate-spin' : ''} />
      </button>
    </div>
  ) : (
    /* SEARCH: Profile is NOT linked */
    <div className="space-y-2">
      <div className="relative">
        <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
        <input
          type="text"
          placeholder="Type employee name or ID..."
          className="w-full pl-10 pr-4 py-2 border rounded-lg text-sm outline-none focus:ring-2 focus:ring-blue-500 bg-white"
          value={linkSearch}
          onChange={(e) => setLinkSearch(e.target.value)}
        />
      </div>

      {/* Search Results Dropdown */}
      {linkSearch && (
        <div className="absolute z-30 w-full mt-1 bg-white border border-gray-200 rounded-lg shadow-xl max-h-60 overflow-y-auto">
          {filteredEmployeesForLink.length > 0 ? (
            filteredEmployeesForLink.map(emp => (
              <button
                key={emp.id}
                onClick={() => handleLinkEmployee(emp.id)}
                disabled={loading}
                className="w-full text-left px-4 py-3 hover:bg-blue-50 transition-colors border-b last:border-b-0 flex flex-col"
              >
                <span className="text-sm font-semibold text-gray-800">
                  {emp.firstName} {emp.lastName}
                </span>
                <span className="text-xs text-gray-500">
                  ID: {emp.employeeID} • {emp.office}
                </span>
              </button>
            ))
          ) : (
            <div className="px-4 py-3 text-xs text-gray-400 italic bg-gray-50">
              No matching unlinked employees found.
            </div>
          )}
        </div>
      )}
      
      {!linkSearch && (
        <p className="text-[10px] text-gray-400 italic ml-1">
          Search for an HR record to associate with this login account.
        </p>
      )}
    </div>
  )}
</div>

                {/* Office Toggles */}
                <div>
                  <label className="text-xs font-bold text-gray-400 uppercase block mb-2">Assigned Offices</label>
                  <div className="grid grid-cols-2 gap-2">
                    {OFFICE_LIST.map(office => (
                      <button
                        key={office}
                        disabled={!canEdit || loading}
                        onClick={() => {
                          const current = user.offices || [];
                          const next = current.includes(office) ? current.filter(o => o !== office) : [...current, office];
                          updateUserData('offices', next);
                        }}
                        className={`flex items-center justify-between p-2 rounded-lg border text-sm transition-all ${user.offices?.includes(office) ? 'bg-blue-50 border-blue-200 text-blue-700 font-medium' : 'bg-white text-gray-500 hover:bg-gray-50'}`}
                      >
                        {office} {user.offices?.includes(office) && <CheckCircle2 size={14} />}
                      </button>
                    ))}
                  </div>
                </div>
             </div>
          ) : (
            <div className="space-y-4">
              {/* STAGING AREA */}
              <div className="bg-blue-50/50 p-4 rounded-xl border border-blue-100 space-y-3">
                <div className="flex justify-between items-center">
                  <label className="text-xs font-bold text-blue-600 uppercase flex items-center gap-2"><UserPlus size={14}/> Bulk Selection</label>
                  {selectedForBulk.length > 0 && (
                    <button 
                      onClick={() => setSelectedForBulk([])}
                      className="text-[10px] font-bold text-red-500 flex items-center gap-1 hover:underline"
                    >
                      <RotateCcw size={10} /> Clear Selection
                    </button>
                  )}
                </div>

                {/* Office Shortcuts */}
                <div className="flex flex-wrap gap-2">
                  {Object.entries(officeAvailability).map(([office, ids]) => (
                    <button
                      key={office}
                      onClick={() => setSelectedForBulk(prev => [...new Set([...prev, ...ids])])}
                      className="px-2 py-1 bg-white border border-blue-200 text-blue-600 rounded text-[10px] font-bold hover:bg-blue-100 transition-colors"
                    >
                      +{office} ({ids.length})
                    </button>
                  ))}
                </div>
                
                <div className="relative">
                  <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
                  <input 
                    type="text"
                    placeholder="Search Name or ID..."
                    className="w-full pl-10 pr-4 py-2 border rounded-lg text-sm outline-none"
                    value={empSearch}
                    onChange={(e) => setEmpSearch(e.target.value)}
                  />
                </div>

                <select 
                  className="w-full p-2 border rounded-lg bg-white text-sm"
                  onChange={(e) => {
                    if (e.target.value) {
                      setSelectedForBulk(prev => [...new Set([...prev, e.target.value])]);
                      setEmpSearch('');
                    }
                  }}
                  value=""
                >
                  <option value="">{filteredAvailableToAssign.length} available to add...</option>
                  {filteredAvailableToAssign.map(emp => (
                    <option key={emp.id} value={emp.id}>{emp.firstName} {emp.lastName} (ID: {emp.employeeID})</option>
                  ))}
                </select>

                {/* STAGING AREA (Items picked but not yet saved to Firestore) */}
                {selectedForBulk.length > 0 && (
                  <div className="pt-2 border-t border-blue-100">
                    <div className="flex flex-wrap gap-2 mb-3">
                      {bulkSelectionList.map(emp => (
                        <span 
                          key={emp.id} 
                          className="inline-flex items-center gap-1 px-2 py-1 bg-white border border-blue-200 text-blue-700 rounded-md text-[10px] font-bold shadow-sm"
                        >
                          <span>{emp.firstName} {emp.lastName}</span>
                          <span className="text-blue-400 font-mono">({emp.employeeID})</span>
                          <button 
                            onClick={() => setSelectedForBulk(p => p.filter(id => id !== emp.id))}
                            className="hover:text-red-500 transition-colors ml-1 border-l pl-1 border-blue-100"
                          >
                            <X size={12}/>
                          </button>
                        </span>
                      ))}
                    </div>
                    
                    <button 
                      onClick={handleBulkAdd} 
                      disabled={loading} 
                      className="w-full py-2 bg-blue-600 text-white rounded-lg text-xs font-bold hover:bg-blue-700 shadow-md flex items-center justify-center gap-2 transition-all active:scale-[0.98]"
                    >
                      {loading ? (
                        <>
                          <Loader2 className="animate-spin" size={14} />
                          Processing...
                        </>
                      ) : (
                        `Add ${selectedForBulk.length} Selected Staff`
                      )}
                    </button>
                  </div>
                )}
              </div>

              {/* CURRENTLY MANAGED LIST */}
{/* 1. The Wrapper: Force a height and prevent it from growing */}
<div className="mt-4 border border-gray-100 rounded-xl p-3 bg-gray-50/30">
  <div className="sticky top-0 pb-2 flex justify-between items-center bg-transparent z-10">
    <label className="text-xs font-bold text-gray-400 uppercase">
      Currently Managing ({managedStaff.length})
    </label>
  </div>
  
  {/* 2. The Scroller: This is where the magic happens */}
  <div 
    className="overflow-y-auto p-3" 
    style={{ height: '300px', maxHeight: '300px' }} // Inline styles often override parent flex behaviors better
  >
    <div className="space-y-2 pb-8">
      {managedStaff.length > 0 ? (
        managedStaff.map(emp => (
          <div key={emp.id} className="flex items-center justify-between p-3 bg-white rounded-lg border border-gray-100 shadow-sm">
            <div>
              <p className="text-sm font-medium text-gray-700">{emp.firstName} {emp.lastName}</p>
              <p className="text-[10px] text-gray-400 flex items-center gap-1">
                <MapPin size={10}/> {emp.office} • ID: {emp.employeeID}
              </p>
            </div>
            <button 
              onClick={async () => {
                setLoading(true);
                await updateDoc(doc(db, 'users', user.id), { managedEmployeeIds: arrayRemove(emp.id) });
                if (onUpdate) onUpdate();
                setLoading(false);
              }} 
              disabled={loading} 
              className="text-gray-300 hover:text-red-500 transition-colors"
            >
              <Trash2 size={18} />
            </button>
          </div>
        ))
      ) : (
        <div className="text-center py-10 text-gray-400 text-xs italic">
          No staff assigned.
        </div>
      )}
    </div>
  </div>
</div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default UserModal;