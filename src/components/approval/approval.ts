import { db } from '@/lib/firebase.config'; 
import { collection, query, where, getDocs, DocumentData, QuerySnapshot } from "firebase/firestore";
import { AbsenceRequest, User } from '@/lib/types';
import { Timestamp } from 'firebase/firestore';

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
export async function getAbsenceRequestsByUser(user: User, startDateString: string, endDateString: string): Promise<AbsenceRequest[]> {

  const userRole = user.role || "";

  const managedEmployeeIds: string[] = (user.managedEmployeeIds || []).filter(
    (id) => id !== user.linkedEmployeeId
  );

  const absencesRef = collection(db, "absences");
  //console.log("User Role:", userRole);
  //console.log("Managed Employee IDs:", managedEmployeeIds);
  //console.log("Offices:", offices);
  //const start = Timestamp.fromDate(new Date(`${startDateString}T00:00:00`));
  //const end = Timestamp.fromDate(new Date(`${endDateString}T23:59:59`));

// 1. Basic Permission Guard
  if (!["HR", "Director", "Manager"].includes(userRole)) return [];

  try {
    let allRequests: AbsenceRequest[] = [];

    // 2. Logic: If no specific employees are managed, fetch EVERYTHING (Global Access)
    // This applies to HR/Directors who don't have a specific team assigned.
    // CHANGED: HR/Directors now see all requests across the system, not just their offices, to align with the new requirement.
    if (["HR", "Director"].includes(userRole)) {
      const globalQuery = query(absencesRef, where("incident_start", "<=", endDateString), where("incident_end", ">=", startDateString));
      const snap = await getDocs(globalQuery);
      allRequests = snap.docs.map(doc => ({ 
        id: doc.id, 
        ...doc.data() 
      } as AbsenceRequest));
    }

    // 3. Logic: Filtered access for Managers as long as they have employees assigned. They see requests ONLY from their managed employees, regardless of office.
    else if (userRole === "Manager" && managedEmployeeIds.length > 0) {
      // Chunk the IDs to stay under the Firestore 30-item limit
      const chunks = chunkArray(managedEmployeeIds, 30);
      
      // Execute all queries in parallel
      const queryPromises = chunks.map(chunk => {
        const q = query(absencesRef, where("employeeFirestoreID", "in", chunk), where("incident_start", "<=", endDateString), where("incident_end", ">=", startDateString), where("skipManagerApproval", "==", false));
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
      const dateA = a.incident_start || "";
      const dateB = b.incident_end || "";
      return dateA.localeCompare(dateB);
    });

  } catch (error) {
    console.error("Error in getAbsenceRequestsByUsers:", error);
    return [];
  }
}

// Helper to derive the visible text status from an object
export const getRequestStatusText = (req: AbsenceRequest): string => {
  if (req.status === 'pending_action') return "Pending Employee";
  if (req.manager_approval === 'denied' && req.final_approval === 'pending') return "Manager Denied";
  if (req.manager_approval === 'denied' || req.final_approval === 'denied') return "Denied";
  if (req.final_approval === 'approved') return "Approved";
  if (req.final_approval === 'approved_with_note') return "Approved With Note"
  if (req.manager_approval === 'approved') return "Manager Approved";
  if (req.manager_approval === 'not_required' && req.final_approval === 'pending') return "Pending Final Approval";
  return "Pending Manager Approval";
};