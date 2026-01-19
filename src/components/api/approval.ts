import { db } from '@/lib/firebase.config'; 
import { collection, query, where, getDocs, doc, updateDoc, serverTimestamp } from "firebase/firestore";

const peopleList = ["Justin Lee"] 

export async function getAbsenceRequestsByOffices(offices: string[], userName: string,  userRole: string) {
  if (!offices || offices.length === 0) {
    return []; // Return empty if the manager has no offices assigned
  }

  // Make it so that when users don't have the specified role, they see no requests
  if (!["HR", "Director", "Manager"].includes(userRole)) {
    return [];
  }
  const absencesRef = collection(db, "absences");
  let absenceQuery;

  // Base query for the user's offices
  const queryConstraints = [where("office", "in", offices)];

  // Add role-based filtering directly into the single query
  if (true) {
    queryConstraints.push(where("employee_title", "in", ["DA"]));
  } 
  else if (userRole === "HR") {
    queryConstraints.push(where("employee_title", "in", ["DA", "Dentist"]));
  }
  // For 'Director', we add no extra constraints, so they see all titles.

  absenceQuery = query(absencesRef, ...queryConstraints);

  const querySnapshot = await getDocs(absenceQuery);
  const requests: Array<{ id: string; [key: string]: any }> = [];

  querySnapshot.forEach((doc) => {
    const data = doc.data();
    // Convert Firestore Timestamps to JS Date objects for easier handling on the client
    const request = {
      id: doc.id,
      ...data,
    };
    requests.push(request);
  });

  // Sort by submission date, newest first
  requests.sort((a, b) => b.date_submitted - a.date_submitted);
  return requests;
}