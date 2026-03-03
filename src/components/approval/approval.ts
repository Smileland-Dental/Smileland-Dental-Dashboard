import { db } from '@/lib/firebase.config'; 
import { collection, query, where, getDocs, DocumentData, QuerySnapshot } from "firebase/firestore";
import { AbsenceRequest, User } from '@/lib/types';

/**
 * Splits an array into smaller chunks to satisfy Firestore's 30-item 'in' limit.
 */
const chunkArray = (arr: string[], size: number): string[][] => {
  return Array.from({ length: Math.ceil(arr.length / size) }, (v, i) =>
    arr.slice(i * size, i * size + size)
  );
};

/**
 * Fetches requests based on role permissions:
 * - HR/Director: See all requests in their assigned offices.
 * - Manager: See ONLY requests from employees in their managedEmployeeIds list.
 */
export async function getAbsenceRequestsByUser(user: User): Promise<AbsenceRequest[]> {

  const userRole = user.role || "";
  const managedEmployeeIds: string[] = user.managedEmployeeIds || [];
  const offices = user.offices || [];
  const absencesRef = collection(db, "absences");
  let queryConstraints = [];
  //console.log("User Role:", userRole);
  //console.log("Managed Employee IDs:", managedEmployeeIds);
  //console.log("Offices:", offices);

// 1. Basic Permission Guard
  if (!["HR", "Director", "Manager"].includes(userRole)) return [];

  try {
    let allRequests: AbsenceRequest[] = [];

    // 2. Logic: If no specific employees are managed, fetch EVERYTHING (Global Access)
    // This applies to HR/Directors who don't have a specific team assigned.
    if (managedEmployeeIds.length === 0 && ["HR", "Director"].includes(userRole)) {
      const globalQuery = query(absencesRef);
      const snap = await getDocs(globalQuery);
      allRequests = snap.docs.map(doc => ({ 
        id: doc.id, 
        ...doc.data() 
      } as AbsenceRequest));
    }

    // 3. Logic: Filtered access for Managers (or HR/Directors with a specific team)
    else {
      // Chunk the IDs to stay under the Firestore 30-item limit
      const chunks = chunkArray(managedEmployeeIds, 30);
      
      // Execute all queries in parallel
      const queryPromises = chunks.map(chunk => {
        const q = query(absencesRef, where("employeeFirestoreID", "in", chunk));
        return getDocs(q); 
      });

      const snapshots: QuerySnapshot<DocumentData>[] = await Promise.all(queryPromises);
      
      // Merge results and cast to AbsenceRequest
      allRequests = snapshots.flatMap(snap => 
        snap.docs.map(doc => ({ 
            id: doc.id, 
            ...doc.data() 
        } as AbsenceRequest))
      );
    }

    // 4. Sort (Oldest first)
    // Using .toMillis() for Firestore Timestamps and new Date() for strings
    return allRequests.sort((a, b) => {
      const dateA = a.createdAt?.toMillis() || 0;
      const dateB = b.createdAt?.toMillis() || 0;
      return dateA - dateB;
    });

  } catch (error) {
    console.error("Error in getAbsenceRequestsByUsers:", error);
    return [];
  }
}