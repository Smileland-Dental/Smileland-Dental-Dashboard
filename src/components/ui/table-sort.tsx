import { ArrowUpDown, ArrowUp, ArrowDown } from 'lucide-react';

interface SortIconProps {
  columnKey: string;
  currentSortKey: string;
  direction: 'asc' | 'desc' | string;
}

export function SortIcon({ columnKey, currentSortKey, direction }: SortIconProps) {
  // If this column isn't the active sorting column, show the neutral up/down icon
  if (currentSortKey !== columnKey || direction === 'none') {
    return (
      <ArrowUpDown className="ml-1 h-3 w-3 text-slate-300 group-hover:text-slate-400 inline-block transition-colors" />
    );
  }

  // Show active sorting states
  if (direction === 'asc') {
    return <ArrowUp className="ml-1 h-3 w-3 text-indigo-600 inline-block animate-fade-in" />;
  }
  
  return <ArrowDown className="ml-1 h-3 w-3 text-indigo-600 inline-block animate-fade-in" />;
}