
export type AbsenceRequest = {
  id: string;
  employee_id: string;
  employee_name: string; // Recommended: Denormalize name for easy display
  employee_title: string;
  eta_etd: string; // Hour:Minutes | For Late In/Early Out, when they expect to arrive/leave
  excuse_note: string[]; // File URL(s) for uploaded excuse notes
  excuse_note_submitted: 'pending' | 'submitted' | 'not_provided';
  date_submitted: string;
  type_of_incident: string; // e.g., "Late In", "Early Out", "Full Day"
  type_of_request: string;
  office: string;
  incident_start: string;
  incident_end: string;
  employee_comments: string;
  manager_notes: string;
  manager_approval: 'pending' | 'approved' | 'denied';
  manager_name: string;
  final_approval: 'pending' | 'approved' | 'denied';
  final_name: string;
};