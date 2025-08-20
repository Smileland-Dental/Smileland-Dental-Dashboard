'use client';

import Link from 'next/link';
import { useState, useEffect } from 'react';
import { onAuthStateChanged, User } from "firebase/auth";
import { auth } from '@/lib/firebase.config';
import { doc, getDoc, getFirestore } from 'firebase/firestore';
import { Button } from '@/components/ui/button';

// Type for a single link
type OfficeLink = {
  title: string;
  url: string;
};

// New type for a group of links belonging to one office
type OfficeLinkGroup = {
  officeName: string;
  links: OfficeLink[];
};

function getColor(){ 
  return "hsl(" + 360 * Math.random() + ',' +
             (25 + 70 * Math.random()) + '%,' + 
             (85 + 10 * Math.random()) + '%)'
};

export default function Page() {
  const [user, setUser] = useState<User | null>(null);
  const [userRole, setUserRole] = useState<string | null>(null);
  // State is now an array of OfficeLinkGroup objects
  const [groupedOfficeLinks, setGroupedOfficeLinks] = useState<OfficeLinkGroup[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      if (currentUser) {
        setUser(currentUser);
        const db = getFirestore();

        const userDocRef = doc(db, 'users', currentUser.uid);
        const userDocSnap = await getDoc(userDocRef);

        if (userDocSnap.exists()) {
          const userData = userDocSnap.data();
          setUserRole(userData.role);

          const officeIds: string[] = userData.offices || [];

          if (officeIds.length > 0) {
            const officePromises = officeIds.map(id => getDoc(doc(db, 'offices', id)));
            const officeDocsSnapshots = await Promise.all(officePromises);

            // Process the snapshots into our new grouped structure
            const allGroupedLinks = officeDocsSnapshots.map(snapshot => {
              if (snapshot.exists()) {
                const officeData = snapshot.data();
                // Create an object with the office name and its links
                return {
                  officeName: officeData.name || 'Unnamed Office', // Fallback name
                  links: officeData.links || []
                };
              }
              return null;
            }).filter(group => group !== null) as OfficeLinkGroup[]; // Filter out any nulls

            setGroupedOfficeLinks(allGroupedLinks);
          }
        } else {
          setUserRole('user');
        }
      } else {
        setUser(null);
        setUserRole(null);
        setGroupedOfficeLinks([]);
      }
      setLoading(false);
    });

    return () => unsubscribe();
  }, []);

  if (loading) {
    return <div>Loading...</div>;
  }

  return (
    <div className='mx-2'>
      {user ? (
        <div>
          <p>This is the forms page.</p>
          <p>User ID: {user.uid}</p>
          <p>Role: {userRole}</p>

          {/* New rendering logic for grouped links */}
          {userRole === 'manager' && groupedOfficeLinks.length > 0 && (
            <div className="flex flex-wrap gap-8 w-full">
              {/* Outer loop: iterates through each office group */}
              {groupedOfficeLinks.map((officeGroup, groupIndex) => (
                <div 
                  key={groupIndex} 
                  className="p-8 border rounded-lg shadow-sm w-full md:w-[calc(50%-1rem)]" 
                  style={{ backgroundColor: getColor() }}
                >
                  <h3 className="text-xl font-semibold mb-4">{officeGroup.officeName} Office</h3>
                  <div className="flex flex-col space-y-4">
                    {/* Inner loop: iterates through links for the current office */}
                    {officeGroup.links.map((link, linkIndex) => (
                      <Button asChild key={linkIndex} variant="secondary" className="py-2.5 px-5 me-2 mb-2 text-lg font-medium text-gray-900 focus:outline-none bg-white rounded-lg border border-gray-200 hover:bg-gray-100 hover:text-blue-700 focus:z-10 focus:ring-4 focus:ring-gray-100 dark:focus:ring-gray-700 dark:bg-gray-800 dark:text-gray-400 dark:border-gray-600 dark:hover:text-white dark:hover:bg-gray-700">
                        <Link href={link.url} target="_blank" rel="noopener noreferrer">
                          {link.title}
                        </Link>
                      </Button>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          )}

          {/* ... other role-based components */}
          {userRole === 'employee' && groupedOfficeLinks.length > 0 && (
            <div className="space-y-4">
              {/* Outer loop: iterates through each office group */}
              {groupedOfficeLinks.map((officeGroup, groupIndex) => (
                <div key={groupIndex} className="p-4 border rounded-lg shadow-sm">
                  <h3 className="text-lg font-semibold mb-3">Employee Forms</h3>
                  <div className="flex flex-col space-y-2">
                    {/* Inner loop: iterates through links for the current office */}
                    {officeGroup.links.map((link, linkIndex) => (
                      <Button asChild key={linkIndex} variant="secondary">
                        <Link href={link.url} target="_blank" rel="noopener noreferrer">
                          {link.title}
                        </Link>
                      </Button>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      ) : (
        <p>Please log in to access this page.</p>
      )}
    </div>
  );
}