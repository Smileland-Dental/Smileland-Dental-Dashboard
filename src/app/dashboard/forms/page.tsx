'use client';

import Link from 'next/link';

import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select"

import { Combobox } from "@/components/combobox"
import { Button } from '@/components/ui/button'; 

export default function Page() {
  return (
    <div>
        <p>This is the forms page.</p>
        <Select>
          <SelectTrigger className="w-[180px]">
            <SelectValue placeholder="Theme" />
          </SelectTrigger>
          <SelectContent>
            <SelectItem value="light">Light</SelectItem>
            <SelectItem value="dark">Dark</SelectItem>
            <SelectItem value="system">System</SelectItem>
          </SelectContent>
        </Select>
    </div>
  )
}
