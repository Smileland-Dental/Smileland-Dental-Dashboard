'use client';

import React from "react";
import { AbsenceRequest } from "@/lib/types";
import { getRequestStatusText } from '@/components/approval/approval';

const styles: Record<string, string> = {
  green: 'bg-emerald-100 text-emerald-700 ring-emerald-600/20',
  red: 'bg-rose-100 text-rose-700 ring-rose-600/20',
  yellow: 'bg-amber-100 text-amber-700 ring-amber-600/20',
  gray: 'bg-slate-100 text-slate-600 ring-slate-600/20',
  blue: 'bg-cyan-100 text-cyan-600 ring-cyan-600/20' 
};

export const Badge = ({ color, text }: { color: string; text: string }) => {
  return (
    <span className={`inline-flex items-center rounded-full px-2 py-1 text-xs font-black text-[10px] uppercase ring-1 ring-inset shadow-sm ${styles[color]}`}>
      {text}
    </span>
  );
};

export const StatusBadge = ({ req }: { req: AbsenceRequest }) => {
  const statusText = getRequestStatusText(req);
  if (statusText === "Pending Employee") return <Badge color="blue" text={statusText} />;
  if (statusText === "Manager Denied" || statusText === "Denied") return <Badge color="red" text={statusText} />;
  if (statusText === "Approved" || statusText === "Approved With Note") return <Badge color="green" text={statusText} />;
  if (statusText === "Manager Approved" || statusText === "Pending Final Approval") return <Badge color="yellow" text={statusText} />;
  return <Badge color="gray" text={statusText} />;
};