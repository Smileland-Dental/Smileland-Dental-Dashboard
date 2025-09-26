
function formatDate(dateString: string): string {
  const parts = dateString.split('/'); // Splits "09/22/2025" into ["09", "22", "2025"]
  const month = parts[0];
  const day = parts[1];
  const year = parts[2];

  return `${year}-${month}-${day}`; // Reassembles as "2025-09-22"
}

const options: Intl.DateTimeFormatOptions = {
  year: 'numeric',
  month: '2-digit',
  day: '2-digit',
  timeZone: 'America/Los_Angeles'
};

export const getPSTDate = () => {
  const today = new Date();
  return formatDate(today.toLocaleString('en-US', options));
};