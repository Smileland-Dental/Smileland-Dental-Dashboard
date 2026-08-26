export const OFFICES = [
  //"Corporate" Split into Departments: Corporate, AR, Call Center, Marcom 
  "Corporate",
  "AR",
  "Call Center",
  "Marcom",
  "Dentist",
  "Ortho", 
  "California", 
  "Ming", 
  "Bernard", 
  "Delano", 
  "Tulare", 
  "Visalia", 
  "Fresno"
];

export const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
export const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

// Helper function to return active background based on selection
export const getRequestColor = (value: string) => {
  switch (value) {
    case 'HR Call In': return 'bg-sky-100 border-sky-300 text-sky-900';
    case 'Incident Notice': return 'bg-orange-100 border-orange-300 text-orange-900';
    case 'Time Off Request': return 'bg-emerald-100 border-emerald-300 text-emerald-900';
    case 'No Call': return 'bg-violet-100 border-violet-400 text-violet-900';
    case 'Call In After Shift': return 'bg-pink-100 border-pink-300 text-pink-900';
    case 'Previously Not Approved': return 'bg-lime-100 border-lime-300 text-lime-900';
    default: return 'bg-slate border-slate-200 text-slate-700';
  }
};