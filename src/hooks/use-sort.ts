import { useState } from 'react';

// Define the structure of our sorting state
export interface SortConfig {
  key: string;
  direction: 'asc' | 'desc' | 'none';
}

export function useSort(initialKey = '') {
  const [sortConfig, setSortConfig] = useState<SortConfig>({
    key: initialKey,
    direction: initialKey ? 'asc' : 'none',
  });

  const handleSort = (columnKey: string) => {
    setSortConfig((prev) => {
      // Clicking a new column initializes it to ascending
      if (prev.key !== columnKey) {
        return { key: columnKey, direction: 'asc' };
      }
      // Cycle through states if it's the same column
      if (prev.direction === 'asc') {
        return { key: columnKey, direction: 'desc' };
      }
      if (prev.direction === 'desc') {
        return { key: '', direction: 'none' }; // Reset back to normal
      }
      return { key: columnKey, direction: 'asc' };
    });
  };

  return { sortConfig, handleSort };
}