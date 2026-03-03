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
  excuse_note_submitted: 'pending' | 'submitted' | 'not_provided';
  final_approval: 'pending' | 'approved' | 'denied';
  final_approval_name: string;
  incident_end: string;
  incident_start: string;
  manager_approval: 'pending' | 'approved' | 'denied';
  manager_approval_name: string;
  manager_notes?: string;
  office: string;
  type_of_incident: string;
  type_of_request: string;
  updatedAt?: Timestamp;
};

export interface User {
  id: string; // Firestore document ID for the user
  createdAt?: Timestamp;
  email: string | null;
  role: string | " User";
  username: string | null;
  managedEmployeeIds: string[]; // Array of Firestore document IDs for employees managed by this user
  offices: string[] | null;
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