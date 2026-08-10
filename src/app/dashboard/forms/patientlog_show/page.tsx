'use client'

import React, { useState, useEffect, useRef, useMemo } from "react";
import { doc, collection, getDocs, getDoc, updateDoc, deleteDoc, addDoc } from "firebase/firestore";
import { db } from "@/lib/firebase.config";

function sanitizeFirebaseDataClient(data: any, depth: number = 0): any {
  if (depth > 20) return null;
  
  if (data === null || data === undefined) return null;
  
  if (typeof data !== 'object') {
    if (typeof data === 'string') {
      if (data.length > 900 * 1024) {
        return data.slice(0, 900 * 1024);
      }
      return data.replace(/[\x00-\x1F\x7F-\x9F]/g, '').trim();
    }
    if (typeof data === 'number' && (isNaN(data) || !isFinite(data))) {
      return 0;
    }
    return data;
  }
  
  if (Array.isArray(data)) {
    if (data.length > 10000) {
      return data.slice(0, 10000).map(item => sanitizeFirebaseDataClient(item, depth + 1));
    }
    return data.map(item => sanitizeFirebaseDataClient(item, depth + 1));
  }
  
  const sanitized: any = {};
  let keyCount = 0;
  const maxKeys = 1000;
  
  for (const key in data) {
    if (Object.prototype.hasOwnProperty.call(data, key)) {
      if (keyCount >= maxKeys) break;
      
      if (key.length > 1500 || key.length === 0) continue;
      
      const safeKey = key.replace(/[.$[\]#\/]/g, '_').slice(0, 1500);
      
      sanitized[safeKey] = sanitizeFirebaseDataClient(data[key], depth + 1);
      keyCount++;
    }
  }
  
  return sanitized;
}

function getCurrentDate(): string {
  return new Date().toISOString().split('T')[0];
}

function getCurrentMonth(): string {
  return getCurrentDate().slice(0, 7);
}

function formatMonthLabel(monthStr: string): string {
  return monthStr.replace('-', '/');
}

function formatShortDate(dateStr: string): string {
  return String(Number(dateStr.split('-')[2]));
}

function normalizePatientField(value: unknown, maxLen: number): string {
  return typeof value === 'string' ? value.replace(/[\x00-\x1F\x7F-\x9F]/g, '').slice(0, maxLen) : '';
}

function isSamePatient(patient: any, appointment: any): boolean {
  return (
    normalizePatientField(patient?.name, 100) === appointment.name &&
    normalizePatientField(patient?.office, 50) === appointment.office
  );
}

function isSameAppointment(a: any, b: any): boolean {
  return a.docId === b.docId && a.name === b.name && a.office === b.office;
}

function removePatientsFromList(patients: any[], toRemove: any[]): any[] {
  const remaining = [...patients];
  for (const apt of toRemove) {
    const idx = remaining.findIndex(p => isSamePatient(p, apt));
    if (idx >= 0) remaining.splice(idx, 1);
  }
  return remaining;
}

export default function ShowCheckSystem() {
  const validateInput = React.useCallback((field: string, value: any): any => {
    if (typeof value === 'string') {
      const maxLengths: { [key: string]: number } = {
        selectedDate: 20,
        selectedMonth: 7,
        selectedOffice: 50
      };
      
      const maxLength = maxLengths[field] || 500;
      if (value.length > maxLength) {
        return value.slice(0, maxLength);
      }
      
      if (field === 'selectedMonth' && value) {
        const monthRegex = /^\d{4}-\d{2}$/;
        if (!monthRegex.test(value)) {
          return '';
        }
      }

      if (field === 'selectedDate' && value) {
        const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
        if (!dateRegex.test(value)) {
          return '';
        }
      }
      
      if (field === 'selectedOffice' && value && value !== 'All') {
        const validOffices = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];
        if (!validOffices.includes(value)) {
          return '';
        }
      }
      
      return value.replace(/[\x00-\x1F\x7F-\x9F]/g, '');
    }
    
    return value;
  }, []);

  const [loading, setLoading] = useState(false);
  const [appointments, setAppointments] = useState<any[]>([]);
  const appointmentsRef = useRef<any[]>([]); // 최신 appointments 추적
  const [filteredAppointments, setFilteredAppointments] = useState<any[]>([]);
  const [selectedMonth, setSelectedMonth] = useState(getCurrentMonth);
  const [selectedDate, setSelectedDate] = useState(getCurrentDate);
  const [selectedOffice, setSelectedOffice] = useState('All');
  const [submitting, setSubmitting] = useState(false);
  const [submitStatus, setSubmitStatus] = useState('');
  const [progress, setProgress] = useState(0);
  const [editingAppointment, setEditingAppointment] = useState<any>(null);
  const [editOffice, setEditOffice] = useState('');
  const [editDate, setEditDate] = useState('');
  const [editSaving, setEditSaving] = useState(false);

  const officeOptions = ['All', 'Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];
  const editableOfficeOptions = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];

  useEffect(() => {
    loadAppointments();

    if (process.env.NODE_ENV === 'production' && 
        typeof window !== 'undefined' && 
        window.location.protocol !== 'https:') {
      window.location.href = window.location.href.replace('http:', 'https:');
    }
  }, []);

  useEffect(() => {
    filterAppointments();
  }, [appointments, selectedDate, selectedOffice]);

  const availableMonths = useMemo(() => {
    const months = new Set<string>();
    appointments.forEach(apt => {
      if (selectedOffice && selectedOffice !== 'All' && apt.office !== selectedOffice) return;
      if (typeof apt.appt_date === 'string' && apt.appt_date.length >= 7) {
        months.add(apt.appt_date.slice(0, 7));
      }
    });
    return Array.from(months).sort();
  }, [appointments, selectedOffice]);

  const datesForSelectedMonth = useMemo(() => {
    const byDate = new Map<string, { total: number; pending: number }>();
    appointments.forEach(apt => {
      if (selectedOffice && selectedOffice !== 'All' && apt.office !== selectedOffice) return;
      if (!apt.appt_date?.startsWith(selectedMonth)) return;
      const entry = byDate.get(apt.appt_date) || { total: 0, pending: 0 };
      entry.total++;
      if (apt.showStatus === 'pending' || !apt.showStatus) entry.pending++;
      byDate.set(apt.appt_date, entry);
    });
    return Array.from(byDate.entries())
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([date, stats]) => ({ date, ...stats }));
  }, [appointments, selectedMonth, selectedOffice]);

  useEffect(() => {
    if (availableMonths.length === 0) return;
    if (!availableMonths.includes(selectedMonth)) {
      setSelectedMonth(availableMonths[availableMonths.length - 1]);
    }
  }, [availableMonths, selectedOffice]);

  useEffect(() => {
    if (datesForSelectedMonth.length === 0) return;
    const dates = datesForSelectedMonth.map(d => d.date);
    if (!dates.includes(selectedDate)) {
      const withPending = datesForSelectedMonth.find(d => d.pending > 0);
      setSelectedDate(withPending?.date ?? dates[dates.length - 1]);
    }
  }, [datesForSelectedMonth, selectedMonth, selectedOffice]);

  useEffect(() => {
    const handleBeforeUnload = (e: any) => {
      if (submitting) {
        e.preventDefault();
        e.returnValue = '';
      }
    };

    const handlePopState = (e: any) => {
      if (submitting) {
        e.preventDefault();
        window.history.pushState(null, '', window.location.href);
      }
    };

    if (submitting) {
      window.addEventListener('beforeunload', handleBeforeUnload);
      window.addEventListener('popstate', handlePopState);
      window.history.pushState(null, '', window.location.href);
    }

    return () => {
      window.removeEventListener('beforeunload', handleBeforeUnload);
      window.removeEventListener('popstate', handlePopState);
    };
  }, [submitting]);

  const loadAppointments = async () => {
    try {
      if (loading) return;

      setLoading(true);
      const showSnapshot = await getDocs(collection(db, "show-noshow"));

      const allAppointments: any[] = [];

      showSnapshot.forEach((docSnap) => {
        const data = docSnap.data();
        const appt_date = typeof data.appt_date === 'string' ? data.appt_date.trim() : '';
        if (!appt_date) return;

        if (Array.isArray(data.patients)) {
          data.patients.forEach((p: any, index: number) => {
            const name = typeof p.name === 'string' ? p.name.replace(/[\x00-\x1F\x7F-\x9F]/g, '').slice(0, 100) : '';
            const office = typeof p.office === 'string' ? p.office.replace(/[\x00-\x1F\x7F-\x9F]/g, '').slice(0, 50) : '';
            const showStatus = (p.showStatus === 'show' || p.showStatus === 'no-show' || p.showStatus === 'pending') ? p.showStatus : 'pending';
            allAppointments.push({
              name,
              office,
              appt_date,
              showStatus,
              docId: docSnap.id,
              rowIndex: index,
            });
          });
        }
      });

      setAppointments(allAppointments);
      appointmentsRef.current = allAppointments;
    } catch (error) {
    } finally {
      setLoading(false);
    }
  };

  const filterAppointmentList = (list: any[]) => {
    let filtered = list;

    if (selectedDate) {
      filtered = filtered.filter(apt =>
        typeof apt.appt_date === 'string' && apt.appt_date === selectedDate
      );
    }

    if (selectedOffice && selectedOffice !== 'All') {
      filtered = filtered.filter(apt => apt.office === selectedOffice);
    }

    return filtered;
  };

  const filterAppointments = () => {
    setFilteredAppointments(filterAppointmentList(appointments));
  };

  const updateShowStatus = async (appointment: any, newStatus: string) => {
    try {
      const validStatuses = ['show', 'no-show', 'pending'];
      if (!validStatuses.includes(newStatus)) {
        alert('⚠️ Invalid value.');
        return;
      }
      
      const docRef = doc(db, "show-noshow", appointment.docId);
      const docSnap = await getDoc(docRef);
      const currentData = docSnap.exists() ? docSnap.data() : null;

      if (currentData && Array.isArray(currentData.patients)) {
        const patients = [...currentData.patients];
        const patientIndex = patients.findIndex(p => isSamePatient(p, appointment));
        if (patientIndex >= 0) {
          patients[patientIndex] = { ...patients[patientIndex], showStatus: newStatus };
          await updateDoc(docRef, sanitizeFirebaseDataClient({
            patients,
            lastUpdated: new Date().toISOString(),
            updatedAt: new Date().toISOString()
          }));
        }
      }

      if (currentData && Array.isArray(currentData.patients) && currentData.patients.some(p => isSamePatient(p, appointment))) {
        setAppointments(prev => {
          const updated = prev.map((apt: any) =>
            isSameAppointment(apt, appointment)
              ? { ...apt, showStatus: newStatus }
              : apt
          );
          
          appointmentsRef.current = updated;
          
          const filtered = filterAppointmentList(updated);
          
          setFilteredAppointments(filtered);
          
          return updated;
        });
        
      }
    } catch (error) {
    }
  };

  const handleSubmit = async () => {
    if (submitting) {
      return;
    }

    if (filteredAppointments.length === 0) {
      alert('⚠️ No appointments to submit.');
      return;
    }

    const hasPendingStatus = filteredAppointments.some(apt => {
      const s = apt.showStatus ?? apt.actions;
      return s !== 'show' && s !== 'no-show';
    });
    
    if (hasPendingStatus) {
      alert('⚠️ Please select Show or No Show for all appointments before submitting.');
      return;
    }

    try {
      setSubmitting(true);
      setSubmitStatus('Processing...');
      setProgress(50);

      const latestAppointments = appointmentsRef.current.length > 0 ? appointmentsRef.current : appointments;
      const submitFilteredAppointments = filterAppointmentList(latestAppointments);

      const markedAppointments = submitFilteredAppointments.filter(apt =>
        apt.showStatus === 'show' || apt.showStatus === 'no-show'
      );

      if (markedAppointments.length === 0) {
        alert('⚠️ Please select show or no show.');
        setSubmitting(false);
        setSubmitStatus('');
        setProgress(0);
        return;
      }

      if (!selectedOffice || selectedOffice === 'All') {
        setSubmitStatus('Deleting...');
        setProgress(70);

        const snapshot = await getDocs(collection(db, "show-noshow"));
        const deleteTasks: Promise<void>[] = [];

        snapshot.forEach((docSnap) => {
          const data = docSnap.data();
          const appt_date = typeof data.appt_date === 'string' ? data.appt_date.trim() : '';
          if (appt_date === selectedDate) {
            deleteTasks.push(deleteDoc(doc(db, "show-noshow", docSnap.id)));
          }
        });

        await Promise.all(deleteTasks);
        await loadAppointments();
      }

      setSubmitStatus('Complete!');
      setProgress(100);

      setTimeout(() => {
        setSubmitting(false);
        setSubmitStatus('');
        setProgress(0);
      }, 2000);
    } catch (error) {
      setSubmitStatus('error');
      setProgress(0);
      setTimeout(() => {
        setSubmitting(false);
        setSubmitStatus('');
      }, 3000);
    }
  };

  const deleteProcessedAppointments = async (appointmentsToDelete?: any[]) => {
    try {
      const appointmentsToProcess = appointmentsToDelete || filteredAppointments;
      const byDoc = new Map<string, any[]>();
      appointmentsToProcess.forEach(apt => {
        if (!byDoc.has(apt.docId)) byDoc.set(apt.docId, []);
        byDoc.get(apt.docId)!.push(apt);
      });

      const successfullyRemoved: any[] = [];

      for (const [docId, apts] of byDoc) {
        const docRef = doc(db, "show-noshow", docId);
        const docSnap = await getDoc(docRef);
        const data = docSnap.exists() ? docSnap.data() : null;
        if (!data || !Array.isArray(data.patients)) continue;

        const patients = removePatientsFromList(data.patients, apts);
        if (patients.length === data.patients.length) continue;

        if (patients.length === 0) await deleteDoc(docRef);
        else await updateDoc(docRef, sanitizeFirebaseDataClient({ patients, lastUpdated: new Date().toISOString() }));

        for (const apt of apts) {
          const wasInDoc = data.patients.some(p => isSamePatient(p, apt));
          const stillInDoc = patients.some(p => isSamePatient(p, apt));
          if (wasInDoc && !stillInDoc) successfullyRemoved.push(apt);
        }
      }

      if (successfullyRemoved.length === 0) {
        throw new Error('No matching patients found in database');
      }

      const updated = appointmentsRef.current.filter(apt =>
        !successfullyRemoved.some(removed => isSameAppointment(apt, removed))
      );
      setAppointments(updated);
      appointmentsRef.current = updated;
      setFilteredAppointments(prev =>
        prev.filter(apt => !successfullyRemoved.some(removed => isSameAppointment(apt, removed)))
      );
    } catch (error) {
      throw error;
    }
  };

  const handleDelete = async (appointment: any) => {
    try {
      await deleteProcessedAppointments([appointment]);
    } catch (error) {
      alert('⚠️ Failed to remove appointment. Please try again.');
    }
  };

  const openEditModal = (appointment: any) => {
    setEditingAppointment(appointment);
    setEditOffice(appointment.office || '');
    setEditDate(appointment.appt_date || '');
  };

  const closeEditModal = () => {
    if (editSaving) return;
    setEditingAppointment(null);
    setEditOffice('');
    setEditDate('');
  };

  const handleSaveEdit = async () => {
    if (!editingAppointment || editSaving) return;

    const newOffice = validateInput('selectedOffice', editOffice.trim());
    const newDate = validateInput('selectedDate', editDate.trim());

    if (!newOffice || !editableOfficeOptions.includes(newOffice)) {
      alert('⚠️ Please enter a valid Office.');
      return;
    }

    if (!newDate || !/^\d{4}-\d{2}-\d{2}$/.test(newDate)) {
      alert('⚠️ Date (YYYY-MM-DD).');
      return;
    }

    if (newOffice === editingAppointment.office && newDate === editingAppointment.appt_date) {
      closeEditModal();
      return;
    }

    try {
      setEditSaving(true);

      const sourceRef = doc(db, "show-noshow", editingAppointment.docId);
      const sourceSnap = await getDoc(sourceRef);
      const sourceData = sourceSnap.exists() ? sourceSnap.data() : null;

      if (!sourceData || !Array.isArray(sourceData.patients)) {
        alert('⚠️ Appointment not found.');
        return;
      }

      const patients = [...sourceData.patients];
      const patientIndex = patients.findIndex(p => isSamePatient(p, editingAppointment));
      if (patientIndex < 0) {
        alert('⚠️ Appointment not found.');
        return;
      }

      const updatedPatient = {
        ...patients[patientIndex],
        office: newOffice,
      };

      if (newDate === editingAppointment.appt_date) {
        patients[patientIndex] = updatedPatient;
        await updateDoc(sourceRef, sanitizeFirebaseDataClient({
          patients,
          lastUpdated: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
        }));
      } else {
        const snapshot = await getDocs(collection(db, "show-noshow"));
        let targetDocId: string | null = null;
        let targetPatients: any[] | null = null;

        snapshot.forEach((docSnap) => {
          if (targetDocId || docSnap.id === editingAppointment.docId) return;
          const data = docSnap.data();
          if (typeof data.appt_date === 'string' && data.appt_date.trim() === newDate) {
            targetDocId = docSnap.id;
            targetPatients = Array.isArray(data.patients) ? data.patients : [];
          }
        });

        if (!targetDocId && patients.length === 1) {
          patients[patientIndex] = updatedPatient;
          await updateDoc(sourceRef, sanitizeFirebaseDataClient({
            patients,
            appt_date: newDate,
            lastUpdated: new Date().toISOString(),
            updatedAt: new Date().toISOString(),
          }));
        } else {
          patients.splice(patientIndex, 1);
          if (patients.length === 0) {
            await deleteDoc(sourceRef);
          } else {
            await updateDoc(sourceRef, sanitizeFirebaseDataClient({
              patients,
              lastUpdated: new Date().toISOString(),
              updatedAt: new Date().toISOString(),
            }));
          }

          if (targetDocId) {
            await updateDoc(doc(db, "show-noshow", targetDocId), sanitizeFirebaseDataClient({
              patients: [...(targetPatients || []), updatedPatient],
              lastUpdated: new Date().toISOString(),
              updatedAt: new Date().toISOString(),
            }));
          } else {
            await addDoc(collection(db, "show-noshow"), sanitizeFirebaseDataClient({
              appt_date: newDate,
              patients: [updatedPatient],
              lastUpdated: new Date().toISOString(),
              updatedAt: new Date().toISOString(),
              createdAt: new Date().toISOString(),
            }));
          }
        }
      }

      await loadAppointments();

      setEditingAppointment(null);
      setEditOffice('');
      setEditDate('');
    } catch (error) {
      alert('⚠️ Failed to update appointment. Please try again.');
    } finally {
      setEditSaving(false);
    }
  };

  const containerStyle = {
    maxWidth: '1400px',
    margin: '40px auto',
    padding: '30px',
    backgroundColor: 'rgba(255, 255, 255, 0.95)',
    backdropFilter: 'blur(5px)',
    borderRadius: '12px',
    boxShadow: '0 8px 16px rgba(0, 0, 0, 0.15)',
    fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
    color: '#023047',
    lineHeight: '1.6'
  };

  const bodyStyle = {
    padding: '20px',
    background: 'linear-gradient(to bottom, #a2d2ff, #f0f8ff)',
    minHeight: '100vh'
  };

  const sectionStyle = {
    marginBottom: '30px',
    padding: '20px',
    backgroundColor: '#f0f8ff',
    borderRadius: '8px',
    border: '1px solid #BDE0FE'
  };

  const inputStyle = {
    padding: '8px 12px',
    border: '1px solid #BDE0FE',
    borderRadius: '4px',
    fontSize: '1em',
    backgroundColor: 'white',
    color: '#023047',
    width: '100%'
  };

  const buttonStyle = {
    padding: '8px 16px',
    borderRadius: '6px',
    fontSize: '14px',
    fontWeight: 'bold',
    cursor: 'pointer',
    margin: '2px',
    border: 'none',
    transition: 'all 0.3s ease'
  };

  const showButtonStyle = {
    ...buttonStyle,
    backgroundColor: '#28a745',
    color: 'white'
  };

  const noShowButtonStyle = {
    ...buttonStyle,
    backgroundColor: '#dc3545',
    color: 'white'
  };

  const deleteButtonStyle = {
    ...buttonStyle,
    backgroundColor: '#6c757d',
    color: 'white'
  };

  const pendingButtonStyle = {
    ...buttonStyle,
    backgroundColor: '#ffc107',
    color: '#212529'
  };

  const editButtonStyle = {
    ...buttonStyle,
    backgroundColor: '#0077B6',
    color: 'white'
  };

  const getStatusColor = (status: string, index: number): string => {
    switch (status) {
      case 'show': return '#d4edda';
      case 'no-show': return '#f8d7da'; 
      case 'pending': return '#fff3cd';
      default: return index % 2 === 0 ? '#f9f9f9' : 'white'; 
    }
  };

  const getStatusText = (status: string): string => {
    switch (status) {
      case 'show': return 'Show';
      case 'no-show': return 'No Show';
      default: return 'Pending';
    }
  };

  return (
    <>
      <style>{`
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
      `}</style>
      <div style={bodyStyle}>
        {submitting && (
          <div style={{
            position: "fixed",
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: "rgba(0, 0, 0, 0.7)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            zIndex: 9999
          }}>
            <div style={{
              backgroundColor: "white",
              padding: "40px",
              borderRadius: "20px",
              textAlign: "center",
              boxShadow: "0 20px 40px rgba(0, 0, 0, 0.3)",
              maxWidth: "400px",
              width: "90%"
            }}>
              <div style={{
                fontSize: "18px",
                fontWeight: "600",
                color: "#4a6fa1",
                marginBottom: "20px"
              }}>
                {submitStatus || "Processing..."}
              </div>
              {progress > 0 && (
                <div style={{
                  width: "100%",
                  height: "8px",
                  backgroundColor: "#f0f0f0",
                  borderRadius: "4px",
                  overflow: "hidden",
                  marginBottom: "10px"
                }}>
                  <div style={{
                    width: `${progress}%`,
                    height: "100%",
                    background: "linear-gradient(90deg, #4a90e2, #51cf66)",
                    transition: "width 0.3s ease",
                    borderRadius: "4px"
                  }} />
                </div>
              )}
              <div style={{
                fontSize: "14px",
                color: "#666",
                marginTop: "10px"
              }}>
                {progress}%
              </div>
            </div>
          </div>
        )}

        <div style={containerStyle}>
        <h1 style={{ 
          color: '#0077B6', 
          textAlign: 'center', 
          marginBottom: '20px', 
          fontSize: '2.5rem', 
          fontWeight: 'bold',
          textShadow: '2px 2px 4px rgba(0,0,0,0.1)'
        }}>Appointment Show/No Show Check</h1>

        <div style={sectionStyle}>
          <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap', marginBottom: '20px' }}>
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Office:
              </label>
              <select
                value={selectedOffice}
                onChange={(e) => {
                  const validatedValue = validateInput('selectedOffice', e.target.value);
                  setSelectedOffice(validatedValue);
                }}
                style={inputStyle}
              >
                {officeOptions.map((office: string) => (
                  <option key={office} value={office}>{office}</option>
                ))}
              </select>
            </div>

            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Month:
              </label>
              <select
                value={availableMonths.includes(selectedMonth) ? selectedMonth : ''}
                onChange={(e) => {
                  const validatedValue = validateInput('selectedMonth', e.target.value);
                  if (validatedValue) setSelectedMonth(validatedValue);
                }}
                style={inputStyle}
                disabled={availableMonths.length === 0}
              >
                {availableMonths.length === 0 ? (
                  <option value="">No data</option>
                ) : (
                  availableMonths.map(month => (
                    <option key={month} value={month}>{formatMonthLabel(month)}</option>
                  ))
                )}
              </select>
            </div>

            <div style={{ flex: '1', minWidth: '240px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Date:
              </label>
              <select
                value={datesForSelectedMonth.some(d => d.date === selectedDate) ? selectedDate : ''}
                onChange={(e) => {
                  const validatedValue = validateInput('selectedDate', e.target.value);
                  if (validatedValue) setSelectedDate(validatedValue);
                }}
                style={inputStyle}
                disabled={datesForSelectedMonth.length === 0}
              >
                {datesForSelectedMonth.length === 0 ? (
                  <option value="">No appointments this month</option>
                ) : (
                  datesForSelectedMonth.map(({ date }) => (
                    <option key={date} value={date}>
                      {formatShortDate(date)}
                    </option>
                  ))
                )}
              </select>
            </div>
            

            <div style={{ flex: '1', minWidth: '200px', display: 'flex', alignItems: 'end' }}>
              <button 
                onClick={loadAppointments}
                disabled={loading}
                style={{
                  ...buttonStyle,
                  backgroundColor: '#0077B6',
                  color: 'white',
                  width: '100%'
                }}
              >
                {loading ? 'Loading...' : '🔄 Refresh Data'}
              </button>
            </div>
          </div>

          <div style={{ 
            padding: '10px', 
            backgroundColor: '#e3f2fd',
            borderRadius: '5px', 
            fontSize: '14px',
            color: '#1565c0'
          }}>
            📊 <strong>Total Appointments:</strong> {filteredAppointments.length}
            {selectedOffice && selectedOffice !== 'All' && ` | Office: ${selectedOffice}`}
          </div>
            </div>

        <div style={sectionStyle}>
          {loading ? (
            <div style={{ textAlign: 'center', padding: '40px', color: '#666' }}>
              Loading appointments...
            </div>
          ) : filteredAppointments.length === 0 ? (
            <div style={{ textAlign: 'center', padding: '40px', color: '#666' }}>
              No appointments found for the selected criteria.
            </div>
          ) : (
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', margin: '10px 0' }}>
              <thead style={{ backgroundColor: '#0077B6', color: 'white' }}>
                <tr>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Patient Name</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Office</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Appt. Date</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Status</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                  {filteredAppointments.map((appointment: any, index: number) => (
                    <tr 
                      key={`${appointment.docId}-${appointment.rowIndex ?? index}`} 
                      style={{ 
                        backgroundColor: getStatusColor(appointment.showStatus, index),
                        opacity: appointment.showStatus === 'pending' ? 1 : 0.8
                      }}
                    >
                      <td style={{ padding: '12px 8px', fontWeight: 'bold' }}>
                        {appointment.name}
                    </td>
                      <td style={{ padding: '12px 8px', textAlign: 'center' }}>
                        {appointment.office}
                    </td>
                      <td style={{ padding: '12px 8px', textAlign: 'center' }}>
                        {appointment.appt_date}
                    </td>
                      <td style={{ padding: '12px 8px', textAlign: 'center', fontWeight: 'bold' }}>
                        {getStatusText(appointment.showStatus)}
                    </td>
                      <td style={{ padding: '12px 8px', textAlign: 'center' }}>
                        <div style={{ display: 'flex', gap: '4px', justifyContent: 'center', flexWrap: 'wrap' }}>
                          <button
                            onClick={() => updateShowStatus(appointment, 'show')}
                            style={showButtonStyle}
                            disabled={appointment.showStatus === 'show'}
                          >
                            Show
                          </button>
                          <button
                            onClick={() => updateShowStatus(appointment, 'no-show')}
                            style={noShowButtonStyle}
                            disabled={appointment.showStatus === 'no-show'}
                          >
                            No Show
                          </button>
                          <button
                            onClick={() => handleDelete(appointment)}
                            style={deleteButtonStyle}
                          >
                            Delete
                          </button>
                          <button
                            onClick={() => updateShowStatus(appointment, 'pending')}
                            style={pendingButtonStyle}
                            disabled={appointment.showStatus === 'pending'}
                          >
                            Pending
                          </button>
                          <button
                            onClick={() => openEditModal(appointment)}
                            style={editButtonStyle}
                          >
                            Edit
                          </button>
                        </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          )}
        </div>

        {editingAppointment && (
          <div style={{
            position: "fixed",
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: "rgba(0, 0, 0, 0.55)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            zIndex: 10000
          }}>
            <div style={{
              backgroundColor: "white",
              padding: "28px",
              borderRadius: "12px",
              width: "90%",
              maxWidth: "420px",
              boxShadow: "0 16px 32px rgba(0, 0, 0, 0.25)"
            }}>
              <h3 style={{ marginTop: 0, marginBottom: '8px', color: '#0077B6' }}>
                Edit Appointment
              </h3>
              <p style={{ marginTop: 0, marginBottom: '18px', color: '#555' }}>
                {editingAppointment.name}
              </p>

              <div style={{ marginBottom: '14px' }}>
                <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                  Office:
                </label>
                <input
                  type="text"
                  value={editOffice}
                  onChange={(e) => setEditOffice(e.target.value)}
                  list="edit-office-options"
                  style={inputStyle}
                  disabled={editSaving}
                />
                <datalist id="edit-office-options">
                  {editableOfficeOptions.map(office => (
                    <option key={office} value={office} />
                  ))}
                </datalist>
              </div>

              <div style={{ marginBottom: '22px' }}>
                <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                  Appt. Date:
                </label>
                <input
                  type="text"
                  value={editDate}
                  onChange={(e) => setEditDate(e.target.value)}
                  placeholder="YYYY-MM-DD"
                  style={inputStyle}
                  disabled={editSaving}
                />
              </div>

              <div style={{ display: 'flex', gap: '10px', justifyContent: 'flex-end' }}>
                <button
                  onClick={closeEditModal}
                  disabled={editSaving}
                  style={{
                    ...buttonStyle,
                    backgroundColor: '#6c757d',
                    color: 'white',
                    opacity: editSaving ? 0.6 : 1
                  }}
                >
                  Cancel
                </button>
                <button
                  onClick={handleSaveEdit}
                  disabled={editSaving}
                  style={{
                    ...buttonStyle,
                    backgroundColor: '#0077B6',
                    color: 'white',
                    opacity: editSaving ? 0.6 : 1
                  }}
                >
                  {editSaving ? 'Saving...' : 'Save'}
                </button>
              </div>
            </div>
          </div>
        )}

        {filteredAppointments.length > 0 && (
        <div style={sectionStyle}>
            <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap' }}>
              {(() => {
                const showCount = filteredAppointments.filter(apt => apt.showStatus === 'show').length;
                const noShowCount = filteredAppointments.filter(apt => apt.showStatus === 'no-show').length;
                const pendingCount = filteredAppointments.filter(apt => apt.showStatus === 'pending').length;
                const rateDenom = showCount + noShowCount;
                const showRate = rateDenom > 0 ? ((showCount / rateDenom) * 100).toFixed(1) : '0';
                
                return (
                  <>
                    <div style={{ 
                      flex: '1', 
                      minWidth: '150px', 
                      padding: '15px', 
                      backgroundColor: '#d4edda', 
                      borderRadius: '8px',
                      textAlign: 'center'
                    }}>
                      <div style={{ fontSize: '24px', fontWeight: 'bold', color: '#155724' }}>{showCount}</div>
                      <div style={{ color: '#155724' }}>Show</div>
                    </div>
                    <div style={{ 
                      flex: '1', 
                      minWidth: '150px', 
                      padding: '15px', 
                      backgroundColor: '#f8d7da', 
                      borderRadius: '8px',
                      textAlign: 'center'
                    }}>
                      <div style={{ fontSize: '24px', fontWeight: 'bold', color: '#721c24' }}>{noShowCount}</div>
                      <div style={{ color: '#721c24' }}>No Show</div>
                    </div>
                    <div style={{ 
                      flex: '1', 
                      minWidth: '150px', 
                      padding: '15px', 
                      backgroundColor: '#fff3cd', 
                      borderRadius: '8px',
                      textAlign: 'center'
                    }}>
                      <div style={{ fontSize: '24px', fontWeight: 'bold', color: '#856404' }}>{pendingCount}</div>
                      <div style={{ color: '#856404' }}>Pending</div>
                    </div>
                    <div style={{ 
                      flex: '1', 
                      minWidth: '150px', 
                      padding: '15px', 
                      backgroundColor: '#e3f2fd', 
                      borderRadius: '8px',
                      textAlign: 'center'
                    }}>
                      <div style={{ fontSize: '24px', fontWeight: 'bold', color: '#1565c0' }}>{showRate}%</div>
                      <div style={{ color: '#1565c0' }}>Show Rate</div>
                    </div>
                  </>
                );
              })()}
            </div>
        </div>
        )}

        {filteredAppointments.length > 0 && (!selectedOffice || selectedOffice === 'All') && (
          <div style={{textAlign: 'center', margin: '30px 0'}}>
          {(() => {
            const hasPendingStatus = filteredAppointments.some(apt => {
              const s = apt.showStatus ?? apt.actions;
              return s !== 'show' && s !== 'no-show';
            });
            const canSubmit = !submitting && !hasPendingStatus;
            const titleMsg = hasPendingStatus
              ? '⚠️ Please select Show or No Show for all appointments before submitting.'
              : '';
            
            return (
              <button 
                onClick={handleSubmit}
                disabled={!canSubmit}
                style={{
                  backgroundColor: canSubmit ? '#3b82f6' : '#9ca3af',
                  color: 'white',
                  border: 'none',
                  padding: '15px 30px',
                  borderRadius: '8px',
                  fontSize: '16px',
                  fontWeight: 'bold',
                  cursor: canSubmit ? 'pointer' : 'not-allowed',
                  boxShadow: '0 4px 6px rgba(0, 0, 0, 0.1)',
                  transition: 'all 0.2s ease'
                }}
                title={titleMsg}
              >
                {submitting ? 'Deleting...' : 'Delete All'}
              </button>
            );
          })()}
        </div>
        )}
      </div>
    </div>
    </>
  );
}
