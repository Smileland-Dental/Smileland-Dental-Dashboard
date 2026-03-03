'use client';

import Link from 'next/link';
import { useState, useEffect } from 'react';
import { onAuthStateChanged, User } from "firebase/auth";
import { auth } from '@/lib/firebase.config';
import { doc, getDoc, getFirestore } from 'firebase/firestore';
import { Button } from '@/components/ui/button';

// Type for the objects in the office links array
type OfficeLink = {
  title: string;
  url: string;
};

// The array for the office links
type OfficeLinkGroup = {
  officeName: string;
  links: OfficeLink[];
};

// Function to generate a random pastel color for the background of the links boxes
function getColor(){ 
  return "hsl(" + 360 * Math.random() + ',' +
             (25 + 70 * Math.random()) + '%,' + 
             (85 + 10 * Math.random()) + '%)'
};

// Main /dashboard/forms page component 
export default function Page() {
  // State to hold the current user and their role
  const [user, setUser] = useState<User | null>(null);
  const [userRole, setUserRole] = useState<string | null>(null);
  // State to hold the links per office
  const [groupedOfficeLinks, setGroupedOfficeLinks] = useState<OfficeLinkGroup[]>([]);
  // State to keep track of loading status of the page
  const [loading, setLoading] = useState(true);
  //console.log("Rendering /dashboard/forms page. Current user:", user, "Role:", userRole, "Grouped Links:", groupedOfficeLinks);

  // Effect to set up the auth state listener and fetch user role and links
  useEffect(() => {

    // This listener checks for changes in the user's login state.
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      // If a user is logged in, fetch their role and office links from Firebase Firestore
      if (currentUser) {
        setUser(currentUser);
        const db = getFirestore();

        // Fetch user document to get role and associated offices
        const userDocRef = doc(db, 'users', currentUser.uid);
        const userDocSnap = await getDoc(userDocRef);

        // If user document exists, extract role and offices
        if (userDocSnap.exists()) {
          const userData = userDocSnap.data();
          setUserRole(userData.role);


          const officeIds: string[] = userData.forms || [];
          console.log("User's office IDs for forms:", officeIds);

          // If office IDs exist, fetch their documents to get links
          if (officeIds.length > 0) {
            const officePromises = officeIds.map(id => getDoc(doc(db, 'forms', id)));
            const officeDocsSnapshots = await Promise.all(officePromises);
            console.log(officeDocsSnapshots, "office documents fetched for user.");

            // Process the snapshots into an array of links for each office.
            const allGroupedLinks = officeDocsSnapshots.map(snapshot => {
              if (snapshot.exists()) {
                const officeData = snapshot.data();
                console.log(officeData);
                // Create an object with the office name and its links
                return {
                  officeName: officeData.name || 'Unnamed Office', // If there isn't a name, default to 'Unnamed Office'
                  links: officeData.links || []
                };
              }
              // If the document doesn't exist, skip it
              return null;
              // Filtering out any nulls from non-existent documents
            }).filter(group => group !== null) as OfficeLinkGroup[]; // Filter out any nulls

            // Set the grouped links state with the fetched data
            setGroupedOfficeLinks(allGroupedLinks);
          }
        } 
        // Default to 'user' role if none found
        else {
          setUserRole('User');
        }
    } 
    // If no user is logged in, clear the user and role state
    else {
        setUser(null);
        setUserRole(null);
        setGroupedOfficeLinks([]);
      }
      setLoading(false);
    });

    // When the component unmounts, unsubscribe from the listener
    return () => unsubscribe();
  // Only run this effect once on component mount
  }, []);

  // If the page is still loading, show a loading message  
  if (loading) {
    return <div>Loading...</div>;
  }

  // Main render of the /dashboard/forms page
  return (
    <div className="mx-2">
      {/* If a user is logged in, show content */}
      {user ? (
      
        <div>
          <h1 className="text-2xl mb-5 font-extrabold leading-none tracking-tight text-gray-900 md:text-3xl lg:text-4xl dark:text-white">Employee Forms</h1>
          {/*<p className="capitalize mb-5">Role: {userRole}</p>*/}

          {/* Rendering logic for grouped links for managers */}
          {userRole === 'Manager' || userRole === 'HR' && groupedOfficeLinks.length > 0 && (
            <div className="flex flex-wrap gap-8 w-screen sm:w-[100vw] md:w-[70vw]">
              {/* Outer loop: iterates through each office group */}
              {groupedOfficeLinks.map((officeGroup, groupIndex) => (
                <div 
                  key={groupIndex} 
                  className="p-8 border rounded-lg shadow-sm w-[96vw] md:w-[calc(50%-1rem)]" 
                  style={{ backgroundColor: getColor() }}
                >
                  <h3 className="text-xl font-semibold mb-4">{officeGroup.officeName}</h3>
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

          {/* Other role-based components (discontinuted for now)*/}
          {userRole === 'Employee' && groupedOfficeLinks.length > 0 && (
            <div className="flex flex-wrap gap-8 w-screen">
              {/* Outer loop: iterates through each office group */}
              {groupedOfficeLinks.map((officeGroup, groupIndex) => (
                <div 
                  key={groupIndex} 
                  className="p-8 border rounded-lg shadow-sm w-[35vw]" 
                  style={{ backgroundColor: getColor() }}
                >
                  <h3 className="text-xl font-semibold mb-4">Employee Forms</h3>
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

          {/* If there are no links available.*/}
          {groupedOfficeLinks.length === 0 && (
            <div className="space-y-4">
              You do not have any office links available.
              <p>Please contact IT to set up your office links.</p>
            </div>
          )}

          </div>
      ) : (
        // If the user is not logged in, display nothing
        <p>Please log in to access this page.</p>
      )}
    </div>
  );
}