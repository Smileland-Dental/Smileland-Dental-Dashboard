import { Timestamp } from "firebase/firestore";

export type AbsenceRequest = {
  id: string; // Firestore document ID for the absence request
  employeeFirestoreID: string; // Firestore document ID for the employee
  createdAt: Timestamp;
  employee_comments?: string;
  employee_id:string;
  employee_name: string;
  employee_title: string;
  eta?: string;
  etd?: string;
  excuse_note?: string[];
  excuse_note_submitted: 'pending' | 'submitted' | 'not_provided' | 'na';
  final_approval: 'pending' | 'approved' | 'denied' | 'approved_with_note';
  final_approval_name: string;
  incident_start: string;
  incident_end: string;
  manager_approval: 'pending' | 'approved' | 'denied' | 'not_required';
  manager_approval_name: string;
  manager_notes?: string;
  office: string; // Office location associated with the employee making request
  type_of_incident: string; // "Late In", "Early Out", "Absent", "Leave and Come Back", "Long Lunch", "Switch Shift"
  type_of_request: string; // Either "Time Off Request" or "Incident Notice" for the User // "HR Call In" for HR created requests // "No Call", "Call In After Shift", "Previously Not Approved"
  updatedAt?: Timestamp;
  skipManagerApproval: boolean; // Field to indicate if manager approval is skipped (true/false)
  DOAPoints: number; // Points assigned for the DOA system based on the hours missed
  pendingDOAPoints: number; // Points that are pending and will be added to DOAPoints if the request is approved
  DAP: number;
  final_notes?: string;
  status: 'active' | 'archived' | 'pending_action'; // New field to indicate if the request is active or archived, including a way to check if it is an initial 'HR Call In' Request
};

export type VolunteerEarlyOutRequest = {
  id?: string;
  employeeFirestoreID: string;
  employee_id: string;
  employee_name: string;
  employee_title: string;
  office: string;
  type_of_request: string;
  supervisor_id: string;
  supervisor_name: string;
  created_at?: Timestamp;
  incident_date: string;
  incident_time: string;
}

export interface User {
  id: string; // Firestore document ID for the user
  createdAt?: Timestamp;
  email: string | null;
  role: string | " User";
  username: string | null;
  managedEmployeeIds: string[]; // Array of Firestore document IDs for employees managed by this user
  offices: string[] | null;
  linkedEmployeeId?: string; // Firestore document ID for the linked employee, if applicable
}

export interface Employee {
  //EmployeeID is number in database
  // All variables here are from using the add employee form
  id: string; // Firestore document ID
  employeeID: string;
  firstName: string;
  lastName: string;
  jobTitle: string;
  office: string;
  department: string;
  dateOfHire: string;
  dateOfBirth: string;
  jobStatus: 'Full-Time' | 'Part-Time';
  employmentStatus: 'Exempt' | 'Non-Exempt' | 'Contracted';
  imageURL: string;
  skipManagerApproval: boolean; 

  // Input boxes for edit details after adding the employee
  // General Tab
  email: string;
  address: string;
  city: string;
  state: string;
  zipCode: string;
  phone: string;
  secondaryPhone: string;
  emergencyContact1Name: string;
  emergencyContact1Phone: string;
  emergencyContact1Relation: string;
  emergencyContact2Name: string;
  emergencyContact2Phone: string;
  emergencyContact2Relation: string;

  // General Different Section
  comments: string;
  permissionToReleaseCheck: string;

  // Input boxes for edit details after adding the employee
  // Employee Information Tab
  status: string;
  dateOfTermination: string;
  employeeSSN: string;
  otherJobTitles: string;

  // Keep this variable a secret Pay
  pay: number;

  // Input boxes for edit details after adding the employee
  // Employee Training Portion of Employee Information Tab
  sexualHarassmentTraining: string;
  oshaTraining: string;
}

export type GroupedEmployees = {
  [office: string]: Employee[];
};