/*'use client';

import React, { useEffect, useState } from "react";
import { collection, doc, getDocs, getFirestore } from 'firebase/firestore';
import { db } from '@/lib/firebase.config';



export default function Test2Page() {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);


  const handleClick = async () => {
    const querySnapshot = await getDocs(collection(db, "cities"));
    querySnapshot.forEach((doc) => {
      // doc.data() is never undefined for query doc snapshots
      console.log(doc.id, " => ", doc.data());
    });
  };

  
  return (
    <main>
      <h1>Test2 Page</h1>
      
      <p>Welcome to the Test2 page of the dashboard.</p>
      <button onClick={handleClick}>
        Fetch Cities
      </button>
    </main>
  );
} */

/*

'use client';

import React, { useState, useEffect } from "react";
import { collection, getDocs, doc, writeBatch } from 'firebase/firestore';
import { db } from '@/lib/firebase.config';

// The interface remains the same
interface City {
  id: string;
  capital?: boolean;
  country?: string;
  name?: string;
  population?: number;
  regions?: string[];
  state?: string;
}

export default function Test2Page() {
  // --- Core Data State ---
  const [data, setData] = useState<City[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // --- State for Editing and Saving ---
  const [modifiedDocs, setModifiedDocs] = useState<Set<string>>(new Set());
  const [isSaving, setIsSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);
  const [saveSuccess, setSaveSuccess] = useState<string | null>(null);


  const handleFetch = async () => {
    setLoading(true);
    setError(null);
    setData([]);
    setModifiedDocs(new Set()); // Clear modifications on new fetch

    try {
      const querySnapshot = await getDocs(collection(db, "cities"));
      const citiesData = querySnapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      })) as City[];
      setData(citiesData);
    } catch (err) {
      console.error("Error fetching documents: ", err);
      setError("Failed to fetch city data.");
    } finally {
      setLoading(false);
    }
  };

  // Handler to update local state when an input changes
  const handleInputChange = (id: string, field: keyof Omit<City, 'id'>, value: string | boolean | number) => {
    // Update the main data array
    setData(currentData =>
      currentData.map(city => {
        if (city.id === id) {
          // For regions, convert comma-separated string back to array
          const finalValue = field === 'regions' && typeof value === 'string'
            ? value.split(',').map(s => s.trim())
            : value;

          return { ...city, [field]: finalValue };
        }
        return city;
      })
    );

    // Add the document id to our set of modified docs
    setModifiedDocs(prev => new Set(prev).add(id));
    setSaveSuccess(null); // Clear success message on new edit
  };

  // Handler to save all modified documents to Firestore
  const handleSaveChanges = async () => {
    setIsSaving(true);
    setSaveError(null);
    setSaveSuccess(null);

    // Use a Firestore Batch Write for efficiency
    const batch = writeBatch(db);

    modifiedDocs.forEach(docId => {
      const cityData = data.find(city => city.id === docId);
      if (cityData) {
        const docRef = doc(db, "cities", docId);
        // Destructure to remove 'id' before sending to Firestore
        const { id, ...dataToSave } = cityData;
        batch.update(docRef, dataToSave);
      }
    });

    try {
      await batch.commit();
      setSaveSuccess("All changes saved successfully!");
      setModifiedDocs(new Set()); // Clear modifications after successful save
    } catch (err) {
      console.error("Error saving documents: ", err);
      setSaveError("Failed to save changes. Please try again.");
    } finally {
      setIsSaving(false);
    }
  };
  
  // Clear success message after 3 seconds
  useEffect(() => {
    if (saveSuccess) {
      const timer = setTimeout(() => setSaveSuccess(null), 3000);
      return () => clearTimeout(timer);
    }
  }, [saveSuccess]);


  return (
    <main className="p-8">
      <div className="flex justify-between items-center mb-4">
        <h1 className="text-2xl font-bold">Editable Cities from Firestore</h1>
        {/* Save button and feedback messages change this->}*/  /*
        <div className="flex items-center space-x-4">
          {saveSuccess && <p className="text-green-600 font-semibold">{saveSuccess}</p>}
          {saveError && <p className="text-red-600 font-semibold">{saveError}</p>}
          {modifiedDocs.size > 0 && (
            <button
              onClick={handleSaveChanges}
              disabled={isSaving}
              className="px-4 py-2 bg-green-600 text-white font-semibold rounded-lg shadow-md hover:bg-green-700 focus:outline-none focus:ring-2 focus:ring-green-400 disabled:bg-gray-400"
            >
              {isSaving ? 'Saving...' : `Save ${modifiedDocs.size} Changes`}
            </button>
          )}
        </div>
      </div>
      
      <p className="mb-6 text-gray-600">
        Click the button to fetch data. You can edit the values directly in the table. A "Save" button will appear when changes are made.
      </p>
      
      <button 
        onClick={handleFetch}
        disabled={loading}
        className="px-4 py-2 bg-blue-600 text-white font-semibold rounded-lg shadow-md hover:bg-blue-700 focus:outline-none disabled:bg-gray-400"
      >
        {loading ? 'Fetching...' : 'Fetch Cities'}
      </button>

      <div className="mt-8">
        {loading && <p className="text-center text-gray-500">Loading data...</p>}
        {error && <p className="text-center text-red-500 bg-red-100 p-4 rounded-lg">{error}</p>}
        
        {data.length > 0 && (
          <div className="relative overflow-x-auto shadow-md sm:rounded-lg">
            <table className="w-full text-sm text-left text-gray-500 dark:text-gray-400">
              <thead className="text-xs text-gray-700 uppercase bg-gray-50 dark:bg-gray-700 dark:text-gray-400">
                <tr>
                  <th scope="col" className="px-6 py-3">City Name</th>
                  <th scope="col" className="px-6 py-3">State</th>
                  <th scope="col" className="px-6 py-3">Country</th>
                  <th scope="col" className="px-6 py-3">Population</th>
                  <th scope="col" className="px-6 py-3">Is Capital</th>
                  <th scope="col" className="px-6 py-3">Regions</th>
                </tr>
              </thead>
              <tbody>
                {data.map((city) => (
                  <tr key={city.id} className="bg-white border-b dark:bg-gray-800 dark:border-gray-700">
                    {/* Each cell is now an input changethis->}*/ /*
                    <td className="px-2 py-2">
                      <input type="text" value={city.name || ''} onChange={(e) => handleInputChange(city.id, 'name', e.target.value)} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" />
                    </td>
                    <td className="px-2 py-2">
                      <input type="text" value={city.state || ''} onChange={(e) => handleInputChange(city.id, 'state', e.target.value)} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" />
                    </td>
                    <td className="px-2 py-2">
                       <input type="text" value={city.country || ''} onChange={(e) => handleInputChange(city.id, 'country', e.target.value)} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" />
                    </td>
                    <td className="px-2 py-2">
                      <input type="number" value={city.population || 0} onChange={(e) => handleInputChange(city.id, 'population', parseInt(e.target.value, 10))} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" />
                    </td>
                    <td className="px-6 py-4 text-center">
                      <input type="checkbox" checked={!!city.capital} onChange={(e) => handleInputChange(city.id, 'capital', e.target.checked)} className="h-5 w-5 rounded text-blue-600 focus:ring-blue-500" />
                    </td>
                    <td className="px-2 py-2">
                      <input type="text" value={city.regions?.join(', ') || ''} onChange={(e) => handleInputChange(city.id, 'regions', e.target.value)} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </main>
  );
} */

'use client';

import React, { useState, useEffect } from "react";
import { collection, getDocs, doc, writeBatch, addDoc } from 'firebase/firestore';
import { db } from '@/lib/firebase.config';

interface City {
  id: string;
  capital?: boolean;
  country?: string;
  name?: string;
  population?: number;
  regions?: string[];
  state?: string;
}

// Define the shape of a new city, which won't have an ID yet
type NewCity = Omit<City, 'id'>;

// Define the initial state for the 'add new' form
const initialNewCityState: NewCity = {
  name: '',
  state: '',
  country: '',
  population: 0,
  capital: false,
  regions: [],
};


export default function Page() {
  // --- Core Data State ---
  const [data, setData] = useState<City[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // --- State for Editing and Updating ---
  const [modifiedDocs, setModifiedDocs] = useState<Set<string>>(new Set());
  const [isSaving, setIsSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);
  const [saveSuccess, setSaveSuccess] = useState<string | null>(null);

  // --- State for Creating New Entries ---
  const [isAdding, setIsAdding] = useState(false);
  const [newCity, setNewCity] = useState<NewCity>(initialNewCityState);
  const [isCreating, setIsCreating] = useState(false);


  const handleFetch = async () => {
    setLoading(true);
    setError(null);
    setData([]);
    setModifiedDocs(new Set()); 
    setIsAdding(false); // Close 'add' form on fetch

    try {
      const querySnapshot = await getDocs(collection(db, "cities"));
      const citiesData = querySnapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      })) as City[];
      setData(citiesData);
    } catch (err) {
      console.error("Error fetching documents: ", err);
      setError("Failed to fetch city data.");
    } finally {
      setLoading(false);
    }
  };

  // Handler for existing row input changes
  const handleInputChange = (id: string, field: keyof NewCity, value: string | boolean | number) => {
    setData(currentData =>
      currentData.map(city => city.id === id ? { ...city, [field]: value } : city)
    );
    setModifiedDocs(prev => new Set(prev).add(id));
    setSaveSuccess(null);
  };

  // Handler for NEW row input changes
  const handleNewCityChange = (field: keyof NewCity, value: string | boolean | number) => {
    // Handle region string-to-array conversion directly here
    const finalValue = field === 'regions' && typeof value === 'string'
      ? value.split(',').map(s => s.trim())
      : value;
    setNewCity(prev => ({ ...prev, [field]: finalValue }));
  };

  // SAVE CHANGES to existing documents
  const handleSaveChanges = async () => {
    setIsSaving(true);
    setSaveError(null);
    setSaveSuccess(null);
    const batch = writeBatch(db);
    modifiedDocs.forEach(docId => {
      const cityData = data.find(city => city.id === docId);
      if (cityData) {
        const docRef = doc(db, "cities", docId);
        const { id, ...dataToSave } = cityData;
        batch.update(docRef, dataToSave as { [key: string]: any });
      }
    });

    try {
      await batch.commit();
      setSaveSuccess("All changes saved successfully!");
      setModifiedDocs(new Set());
    } catch (err) {
      console.error("Error saving documents: ", err);
      setSaveError("Failed to save changes. Please try again.");
    } finally {
      setIsSaving(false);
    }
  };

  // ADD NEW city to the database
  const handleCreateCity = async () => {
    if (!newCity.name) {
      alert("City name is required to add a new entry.");
      return;
    }
    setIsCreating(true);
    setSaveError(null);

    try {
      await addDoc(collection(db, "cities"), newCity);
      setSaveSuccess(`City "${newCity.name}" added successfully!`);
      setIsAdding(false); // Hide the form
      setNewCity(initialNewCityState); // Reset the form
      await handleFetch(); // Refresh the table with the new data
    } catch (err) {
      console.error("Error adding document: ", err);
      setSaveError("Failed to add new city.");
    } finally {
      setIsCreating(false);
    }
  };

  // Helper to cancel adding
  const handleCancelAdd = () => {
    setIsAdding(false);
    setNewCity(initialNewCityState);
  };

  useEffect(() => {
    if (saveSuccess) {
      const timer = setTimeout(() => setSaveSuccess(null), 4000);
      return () => clearTimeout(timer);
    }
  }, [saveSuccess]);


  return (
    <main className="p-8">
      <div className="flex justify-between items-center mb-4">
        <h1 className="text-2xl font-bold">Editable Cities from Firestore</h1>
        <div className="flex items-center space-x-4">
          {saveSuccess && <p className="text-green-600 font-semibold">{saveSuccess}</p>}
          {saveError && <p className="text-red-600 font-semibold">{saveError}</p>}
          {modifiedDocs.size > 0 && (
            <button
              onClick={handleSaveChanges}
              disabled={isSaving}
              className="px-4 py-2 bg-green-600 text-white font-semibold rounded-lg shadow-md hover:bg-green-700 disabled:bg-gray-400"
            >
              {isSaving ? 'Saving...' : `Save ${modifiedDocs.size} Changes`}
            </button>
          )}
        </div>
      </div>
      
      <p className="mb-6 text-gray-600">
        Fetch data, edit in the table, or add a new city. A "Save" button will appear when changes are made.
      </p>
      
      <div className="flex space-x-4">
        <button onClick={handleFetch} disabled={loading} className="px-4 py-2 bg-blue-600 text-white font-semibold rounded-lg shadow-md hover:bg-blue-700 disabled:bg-gray-400">
          {loading ? 'Fetching...' : 'Fetch Cities'}
        </button>
        <button onClick={() => setIsAdding(true)} disabled={isAdding} className="px-4 py-2 bg-indigo-600 text-white font-semibold rounded-lg shadow-md hover:bg-indigo-700 disabled:bg-gray-400">
          Add New City
        </button>
      </div>

      <div className="mt-8">
        {loading && <p className="text-center text-gray-500">Loading data...</p>}
        {error && <p className="text-center text-red-500 bg-red-100 p-4 rounded-lg">{error}</p>}
        
        {data.length > 0 || isAdding ? (
          <div className="relative overflow-x-auto shadow-md sm:rounded-lg">
            <table className="w-full text-sm text-left text-gray-500">
              <thead className="text-xs text-gray-700 uppercase bg-gray-50">
                <tr>
                  <th scope="col" className="px-6 py-3">City Name</th>
                  <th scope="col" className="px-6 py-3">State</th>
                  <th scope="col" className="px-6 py-3">Country</th>
                  <th scope="col" className="px-6 py-3">Population</th>
                  <th scope="col" className="px-6 py-3">Is Capital</th>
                  <th scope="col" className="px-6 py-3">Regions / Actions</th>
                </tr>
              </thead>
              <tbody>
                {/* --- ADD NEW ROW --- */}
                {isAdding && (
                  <tr className="bg-indigo-50 border-b border-indigo-200">
                    <td className="px-2 py-2"><input type="text" placeholder="Name" value={newCity.name} onChange={(e) => handleNewCityChange('name', e.target.value)} className="w-full bg-white p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500" /></td>
                    <td className="px-2 py-2"><input type="text" placeholder="State" value={newCity.state} onChange={(e) => handleNewCityChange('state', e.target.value)} className="w-full bg-white p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500" /></td>
                    <td className="px-2 py-2"><input type="text" placeholder="Country" value={newCity.country} onChange={(e) => handleNewCityChange('country', e.target.value)} className="w-full bg-white p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500" /></td>
                    <td className="px-2 py-2"><input type="number" placeholder="Population" value={newCity.population} onChange={(e) => handleNewCityChange('population', parseInt(e.target.value, 10))} className="w-full bg-white p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500" /></td>
                    <td className="px-6 py-4 text-center"><input type="checkbox" checked={newCity.capital} onChange={(e) => handleNewCityChange('capital', e.target.checked)} className="h-5 w-5 rounded text-indigo-600 focus:ring-indigo-500" /></td>
                    <td className="px-2 py-2 whitespace-nowrap">
                       <div className="flex items-center space-x-2">
                          <input type="text" placeholder="region1,region2" value={newCity.regions?.join(', ')} onChange={(e) => handleNewCityChange('regions', e.target.value)} className="w-full bg-white p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500" />
                          <button onClick={handleCreateCity} disabled={isCreating} className="p-2 bg-green-500 text-white rounded hover:bg-green-600 disabled:bg-gray-300">✓</button>
                          <button onClick={handleCancelAdd} className="p-2 bg-red-500 text-white rounded hover:bg-red-600">✗</button>
                       </div>
                    </td>
                  </tr>
                )}
                {/* --- EXISTING DATA ROWS --- */}
                {data.map((city) => (
                  <tr key={city.id} className="bg-white border-b">
                    <td className="px-2 py-2"><input type="text" value={city.name || ''} onChange={(e) => handleInputChange(city.id, 'name', e.target.value)} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" /></td>
                    <td className="px-2 py-2"><input type="text" value={city.state || ''} onChange={(e) => handleInputChange(city.id, 'state', e.target.value)} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" /></td>
                    <td className="px-2 py-2"><input type="text" value={city.country || ''} onChange={(e) => handleInputChange(city.id, 'country', e.target.value)} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" /></td>
                    <td className="px-2 py-2"><input type="number" value={city.population || 0} onChange={(e) => handleInputChange(city.id, 'population', parseInt(e.target.value, 10))} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" /></td>
                    <td className="px-6 py-4 text-center"><input type="checkbox" checked={!!city.capital} onChange={(e) => handleInputChange(city.id, 'capital', e.target.checked)} className="h-5 w-5 rounded text-blue-600 focus:ring-blue-500" /></td>
                   {/*} <td className="px-2 py-2"><input type="text" value={city.regions?.join(', ') || ''} onChange={(e) => handleInputChange(city.id, 'regions', e.target.value.split(',').map(s=>s.trim()))} className="w-full bg-transparent p-2 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" /></td>*/}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : null}
      </div>
    </main>
  );
}