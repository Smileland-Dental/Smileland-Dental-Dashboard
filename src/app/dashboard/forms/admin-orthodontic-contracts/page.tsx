'use client'

import React, { useState, useEffect, useMemo } from "react";
import { collection, getDocs, query, orderBy, deleteDoc, doc } from "firebase/firestore";
import { db } from "@/lib/firebase.config";
import { getStorage, ref as storageRef, deleteObject } from "firebase/storage";

interface OrthodonticContract {
  id: string;
  patientName: string;
  dob: string;
  responsibleParty: string;
  relationship: string;
  ssn: string;
  driversLicense: string;
  typeOfTreatment: string;
  contractDate: string;
  timestamp: string;
  language?: string;
  approved?: boolean;
  pdfUrl?: string;
  pdfFileName?: string;
  quotePresentedBy?: string;
  quotePresentedDate?: string;
  [key: string]: any;
}

export default function AdminOrthodonticContracts() {
  const [contracts, setContracts] = useState<OrthodonticContract[]>([]);
  const [loading, setLoading] = useState(true);
  const [selectedContract, setSelectedContract] = useState<OrthodonticContract | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [filterLanguage, setFilterLanguage] = useState<'all' | 'english' | 'spanish'>('all');
  const [downloadingExcel, setDownloadingExcel] = useState(false);
  const [selectedMonths, setSelectedMonths] = useState<Set<string>>(new Set());

  useEffect(() => {
    fetchContracts();
  }, []);

  const fetchContracts = async () => {
    try {
      setLoading(true);
      const contractsRef = collection(db, "orthodontic-contracts");
      const q = query(contractsRef, orderBy("timestamp", "desc"));
      const querySnapshot = await getDocs(q);
      
      const contractsData: OrthodonticContract[] = [];
      querySnapshot.forEach((doc) => {
        contractsData.push({
          id: doc.id,
          ...doc.data()
        } as OrthodonticContract);
      });
      
      setContracts(contractsData);
      setLoading(false);
    } catch (error) {
      console.error("Error fetching contracts:", error);
      alert("Failed to load contracts: " + (error instanceof Error ? error.message : 'Unknown error'));
      setLoading(false);
    }
  };

  const handleDelete = async (contract: OrthodonticContract) => {
    const confirmDelete = window.confirm(
      `Are you sure you want to delete the contract for ${contract.patientName}?\nThis action cannot be undone.`
    );
    
    if (!confirmDelete) return;

    try {
      // 1. Firebase Storage에서 PDF 삭제
      if (contract.pdfUrl && contract.pdfFileName) {
        const storage = getStorage();
        const pdfRef = storageRef(storage, `orthodontic-contracts/${contract.id}/${contract.pdfFileName}`);
        try {
          await deleteObject(pdfRef);
        } catch (error) {
          console.log("PDF file not found in storage or already deleted");
        }
      }

      // 2. Firestore에서 문서 삭제
      await deleteDoc(doc(db, "orthodontic-contracts", contract.id));

      // 3. 로컬 상태 업데이트
      setContracts(contracts.filter(c => c.id !== contract.id));
      setSelectedContract(null);
      
      alert("✅ Contract deleted successfully!");
    } catch (error) {
      console.error("Error deleting contract:", error);
      alert("❌ Failed to delete contract: " + (error instanceof Error ? error.message : 'Unknown error'));
    }
  };

  const handleDeleteAll = async () => {
    const confirmDelete = window.confirm(
      `⚠️ WARNING: This will delete ALL ${contracts.length} orthodontic contracts!\n\nThis includes:\n- All contract data from database\n- All PDF files from storage\n\nThis action CANNOT be undone!\n\nAre you absolutely sure?`
    );
    
    if (!confirmDelete) return;

    // 두 번째 확인
    const doubleConfirm = window.confirm(
      `🚨 FINAL CONFIRMATION\n\nYou are about to permanently delete ${contracts.length} contracts.\n\nType YES in your mind and click OK to proceed.`
    );
    
    if (!doubleConfirm) return;

    try {
      setLoading(true);
      
      let deletedCount = 0;
      let errorCount = 0;
      
      // 모든 계약서 삭제
      for (const contract of contracts) {
        try {
          // 1. Firebase Storage에서 PDF 삭제
          if (contract.pdfUrl && contract.pdfFileName) {
            const storage = getStorage();
            const pdfRef = storageRef(storage, `orthodontic-contracts/${contract.id}/${contract.pdfFileName}`);
            try {
              await deleteObject(pdfRef);
              console.log(`✅ Deleted PDF for ${contract.patientName}`);
            } catch (error) {
              console.log(`⚠️ PDF not found for ${contract.patientName}`);
            }
          }

          // 2. Firestore에서 문서 삭제
          await deleteDoc(doc(db, "orthodontic-contracts", contract.id));
          deletedCount++;
          console.log(`✅ Deleted contract ${deletedCount}/${contracts.length}: ${contract.patientName}`);
          
        } catch (error) {
          errorCount++;
          console.error(`❌ Failed to delete contract for ${contract.patientName}:`, error);
        }
      }

      // 3. 로컬 상태 초기화
      setContracts([]);
      setSelectedContract(null);
      setLoading(false);
      
      if (errorCount === 0) {
        alert(`✅ Successfully deleted all ${deletedCount} contracts!`);
      } else {
        alert(`⚠️ Deleted ${deletedCount} contracts, but ${errorCount} failed.\n\nPlease refresh and try again for remaining contracts.`);
      }
      
    } catch (error) {
      console.error("Error deleting all contracts:", error);
      alert("❌ Failed to delete contracts: " + (error instanceof Error ? error.message : 'Unknown error'));
      setLoading(false);
    }
  };

  const handleDownloadExcel = async () => {
    try {
      setDownloadingExcel(true);
      
      // Firestore에서 모든 계약서 데이터 가져오기
      const contractsRef = collection(db, "orthodontic-contracts");
      const q = query(contractsRef, orderBy("timestamp", "desc"));
      const querySnapshot = await getDocs(q);
      
      if (querySnapshot.empty) {
        alert('❌ No contracts found. Please submit at least one contract first.');
        return;
      }
      
      // CSV 헤더
      const csvData = [
        ['Patient Name', 'DOB', 'Responsible Party', 'Relationship', 'Type of Treatment', 'Contract Date', 'Language', 'Quote Presented By', 'Quote Date', 'First Option - Treatment', 'First Option - Appliance', 'First Option - Deposit', 'First Option - Subtotal', 'First Option - Est. Insurance', 'First Option - Net Balance', 'First Option - Est. Period (months)', 'First Option - Monthly Payment', 'Second Option - Treatment', 'Second Option - Appliance', 'Second Option - Deposit', 'Second Option - Subtotal', 'Second Option - Est. Insurance', 'Second Option - Net Balance', 'Second Option - Est. Period (months)', 'Second Option - Monthly Payment', 'Submission Date', 'Status', 'PDF Link']
      ];
      
      // 각 계약서를 CSV 행으로 변환
      querySnapshot.forEach((doc) => {
        const data = doc.data();
        const row = [
          data.patientName || 'N/A',
          data.dob || 'N/A',
          data.responsibleParty || 'N/A',
          data.relationship || 'N/A',
          data.typeOfTreatment || 'N/A',
          data.contractDate || 'N/A',
          data.language === 'spanish' ? 'Español' : 'English',
          data.quotePresentedBy || 'N/A',
          data.quotePresentedDate || 'N/A',
          data.firstOption?.treatment ? `$${data.firstOption.treatment}` : 'N/A',
          data.firstOption?.appliance ? `$${data.firstOption.appliance}` : 'N/A',
          data.firstOption?.deposit ? `$${data.firstOption.deposit}` : 'N/A',
          data.firstOption?.subtotal ? `$${data.firstOption.subtotal}` : 'N/A',
          data.firstOption?.estimatedInsurance ? `$${data.firstOption.estimatedInsurance}` : 'N/A',
          data.firstOption?.netBalance ? `$${data.firstOption.netBalance}` : 'N/A',
          data.firstOption?.estimatedTreatmentPeriod || 'N/A',
          data.firstOption?.monthlyPayment ? `$${data.firstOption.monthlyPayment}` : 'N/A',
          data.secondOption?.treatment ? `$${data.secondOption.treatment}` : 'N/A',
          data.secondOption?.appliance ? `$${data.secondOption.appliance}` : 'N/A',
          data.secondOption?.deposit ? `$${data.secondOption.deposit}` : 'N/A',
          data.secondOption?.subtotal ? `$${data.secondOption.subtotal}` : 'N/A',
          data.secondOption?.estimatedInsurance ? `$${data.secondOption.estimatedInsurance}` : 'N/A',
          data.secondOption?.netBalance ? `$${data.secondOption.netBalance}` : 'N/A',
          data.secondOption?.estimatedTreatmentPeriod || 'N/A',
          data.secondOption?.monthlyPayment ? `$${data.secondOption.monthlyPayment}` : 'N/A',
          data.timestamp ? new Date(data.timestamp).toLocaleDateString() : 'N/A',
          'Approved & PDF Generated',
          data.pdfUrl || 'N/A'
        ];
        csvData.push(row);
      });
      
      // CSV 문자열 생성
      const csvString = csvData.map(row => row.join(',')).join('\n');
      
      // CSV 파일 다운로드
      const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `orthodontic-contracts-${new Date().toISOString().split('T')[0]}.csv`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
      
      console.log(`✅ Excel downloaded with ${csvData.length - 1} contracts`);
      
    } catch (error) {
      console.error('Error downloading Excel:', error);
      alert('❌ Failed to download Excel file: ' + (error instanceof Error ? error.message : 'Unknown error'));
    } finally {
      setDownloadingExcel(false);
    }
  };

  // 사용 가능한 월 목록 추출
  const availableMonths = useMemo(() => {
    const monthsSet = new Set<string>();
    contracts.forEach(contract => {
      if (contract.timestamp) {
        const date = new Date(contract.timestamp);
        const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
        monthsSet.add(monthKey);
      }
    });
    return Array.from(monthsSet).sort().reverse(); // 최신순으로 정렬
  }, [contracts]);

  // 월 표시 형식 변환 (YYYY-MM -> "MMM YYYY")
  const formatMonth = (monthKey: string) => {
    const [year, month] = monthKey.split('-');
    const date = new Date(parseInt(year), parseInt(month) - 1);
    return date.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
  };

  // 월 선택 토글
  const toggleMonth = (monthKey: string) => {
    setSelectedMonths(prev => {
      const newSet = new Set(prev);
      if (newSet.has(monthKey)) {
        newSet.delete(monthKey);
      } else {
        newSet.add(monthKey);
      }
      return newSet;
    });
  };

  // 모든 월 선택/해제
  const toggleAllMonths = () => {
    if (selectedMonths.size === availableMonths.length) {
      setSelectedMonths(new Set());
    } else {
      setSelectedMonths(new Set(availableMonths));
    }
  };

  const filteredContracts = contracts.filter(contract => {
    const matchesSearch = contract.patientName.toLowerCase().includes(searchTerm.toLowerCase()) ||
                         contract.responsibleParty.toLowerCase().includes(searchTerm.toLowerCase()) ||
                         contract.id.toLowerCase().includes(searchTerm.toLowerCase());
    
    const matchesLanguage = filterLanguage === 'all' || 
                           (filterLanguage === 'spanish' && contract.language === 'spanish') ||
                           (filterLanguage === 'english' && (!contract.language || contract.language === 'english'));
    
    // 날짜 필터링: 선택된 월이 없으면 모든 계약서 표시, 있으면 선택된 월만 표시
    let matchesDate = true;
    if (selectedMonths.size > 0 && contract.timestamp) {
      const date = new Date(contract.timestamp);
      const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
      matchesDate = selectedMonths.has(monthKey);
    }
    
    return matchesSearch && matchesLanguage && matchesDate;
  });

  const styles = {
    container: {
      padding: '20px',
      maxWidth: '1400px',
      margin: '0 auto',
      fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
    },
    header: {
      marginBottom: '30px',
      borderBottom: '3px solid #1976d2',
      paddingBottom: '15px'
    },
    title: {
      fontSize: '2rem',
      fontWeight: 'bold',
      color: '#1976d2',
      margin: '0 0 10px 0'
    },
    filters: {
      display: 'flex',
      gap: '15px',
      marginBottom: '20px',
      flexWrap: 'wrap' as const,
      alignItems: 'center'
    },
    searchInput: {
      padding: '10px 15px',
      fontSize: '1rem',
      border: '2px solid #e0e0e0',
      borderRadius: '6px',
      minWidth: '300px',
      outline: 'none'
    },
    filterButton: (active: boolean) => ({
      padding: '10px 20px',
      fontSize: '0.95rem',
      fontWeight: active ? 'bold' : 'normal',
      backgroundColor: active ? '#1976d2' : 'white',
      color: active ? 'white' : '#666',
      border: `2px solid ${active ? '#1976d2' : '#e0e0e0'}`,
      borderRadius: '6px',
      cursor: 'pointer',
      transition: 'all 0.3s'
    }),
    tableContainer: {
      overflowX: 'auto' as const,
      backgroundColor: 'white',
      borderRadius: '8px',
      boxShadow: '0 2px 8px rgba(0,0,0,0.1)'
    },
    table: {
      width: '100%',
      borderCollapse: 'collapse' as const,
      minWidth: '800px'
    },
    th: {
      backgroundColor: '#f5f5f5',
      padding: '15px',
      textAlign: 'left' as const,
      fontWeight: 'bold',
      color: '#333',
      borderBottom: '2px solid #e0e0e0'
    },
    td: {
      padding: '12px 15px',
      borderBottom: '1px solid #f0f0f0'
    },
    button: {
      padding: '8px 16px',
      fontSize: '0.9rem',
      fontWeight: 'bold',
      border: 'none',
      borderRadius: '4px',
      cursor: 'pointer',
      marginRight: '5px',
      transition: 'all 0.3s'
    },
    viewButton: {
      backgroundColor: '#1976d2',
      color: 'white'
    },
    downloadButton: {
      backgroundColor: '#4caf50',
      color: 'white'
    },
    deleteButton: {
      backgroundColor: '#f44336',
      color: 'white'
    },
    badge: (language?: string) => ({
      display: 'inline-block',
      padding: '4px 10px',
      fontSize: '0.75rem',
      fontWeight: 'bold',
      borderRadius: '12px',
      backgroundColor: language === 'spanish' ? '#ff9800' : '#2196f3',
      color: 'white'
    })
  };

  return (
    <>
      <style>{`
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
      `}</style>
      
      <div style={styles.container}>
        <div style={styles.header}>
          <h1 style={styles.title}>📋 Orthodontic Contracts Management</h1>
          <p style={{ color: '#666', margin: 0 }}>
            Total Contracts: <strong>{filteredContracts.length}</strong> 
            {searchTerm && ` (filtered from ${contracts.length})`}
          </p>
        </div>

        {/* Filters */}
        <div style={styles.filters}>
          <input
            type="text"
            placeholder="🔍 Search by patient name, responsible party, or ID..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            style={styles.searchInput}
          />
          
          <button
            onClick={() => setFilterLanguage('all')}
            style={styles.filterButton(filterLanguage === 'all')}
          >
            All Languages
          </button>
          <button
            onClick={() => setFilterLanguage('english')}
            style={styles.filterButton(filterLanguage === 'english')}
          >
            English
          </button>
          <button
            onClick={() => setFilterLanguage('spanish')}
            style={styles.filterButton(filterLanguage === 'spanish')}
          >
            Español
          </button>
          
          <div style={{ marginLeft: 'auto', display: 'flex', gap: '10px' }}>
            <button
              onClick={handleDownloadExcel}
              disabled={downloadingExcel}
              style={{
                ...styles.button,
                backgroundColor: downloadingExcel ? '#bdc3c7' : '#2e7d32',
                color: 'white',
                cursor: downloadingExcel ? 'not-allowed' : 'pointer'
              }}
            >
              {downloadingExcel ? '⏳ Downloading...' : '📊 Download Excel'}
            </button>
            <button
              onClick={handleDeleteAll}
              disabled={loading || contracts.length === 0}
              style={{
                ...styles.button,
                backgroundColor: (loading || contracts.length === 0) ? '#bdc3c7' : '#d32f2f',
                color: 'white',
                cursor: (loading || contracts.length === 0) ? 'not-allowed' : 'pointer'
              }}
            >
              🗑️ Delete All
            </button>
          </div>
        </div>

        {/* Month Filter */}
        {availableMonths.length > 0 && (
          <div style={{
            marginBottom: '15px',
            padding: '8px 12px',
            backgroundColor: 'white',
            borderRadius: '6px',
            boxShadow: '0 1px 4px rgba(0,0,0,0.1)'
          }}>
            <div style={{ display: 'flex', alignItems: 'center', marginBottom: '6px', gap: '8px', flexWrap: 'wrap' }}>
              <strong style={{ color: '#333', fontSize: '0.85rem' }}>📅 Filter by Month:</strong>
              <button
                onClick={toggleAllMonths}
                style={{
                  padding: '4px 8px',
                  fontSize: '0.75rem',
                  backgroundColor: selectedMonths.size === availableMonths.length ? '#1976d2' : 'white',
                  color: selectedMonths.size === availableMonths.length ? 'white' : '#666',
                  border: '1px solid #e0e0e0',
                  borderRadius: '4px',
                  cursor: 'pointer',
                  fontWeight: '500'
                }}
              >
                {selectedMonths.size === availableMonths.length ? 'Clear All' : 'Select All'}
              </button>
              {selectedMonths.size > 0 && (
                <span style={{ color: '#666', fontSize: '0.8rem' }}>
                  ({selectedMonths.size} selected)
                </span>
              )}
            </div>
            <div style={{
              display: 'flex',
              flexWrap: 'wrap',
              gap: '6px'
            }}>
              {availableMonths.map(monthKey => (
                <label
                  key={monthKey}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    padding: '4px 10px',
                    backgroundColor: selectedMonths.has(monthKey) ? '#e3f2fd' : '#f5f5f5',
                    border: `1px solid ${selectedMonths.has(monthKey) ? '#1976d2' : '#e0e0e0'}`,
                    borderRadius: '4px',
                    cursor: 'pointer',
                    transition: 'all 0.2s',
                    fontWeight: selectedMonths.has(monthKey) ? '600' : 'normal',
                    fontSize: '0.8rem'
                  }}
                >
                  <input
                    type="checkbox"
                    checked={selectedMonths.has(monthKey)}
                    onChange={() => toggleMonth(monthKey)}
                    style={{
                      marginRight: '6px',
                      cursor: 'pointer',
                      width: '14px',
                      height: '14px'
                    }}
                  />
                  <span style={{ color: selectedMonths.has(monthKey) ? '#1976d2' : '#666' }}>
                    {formatMonth(monthKey)}
                  </span>
                </label>
              ))}
            </div>
          </div>
        )}

        {/* Loading */}
        {loading && (
          <div style={{ textAlign: 'center', padding: '50px' }}>
            <div style={{
              border: '4px solid #f3f3f3',
              borderTop: '4px solid #1976d2',
              borderRadius: '50%',
              width: '50px',
              height: '50px',
              animation: 'spin 1s linear infinite',
              margin: '0 auto 20px'
            }}></div>
            <p style={{ color: '#666' }}>Loading contracts...</p>
          </div>
        )}

        {/* Table */}
        {!loading && filteredContracts.length === 0 && (
          <div style={{
            textAlign: 'center',
            padding: '50px',
            backgroundColor: 'white',
            borderRadius: '8px',
            boxShadow: '0 2px 8px rgba(0,0,0,0.1)'
          }}>
            <p style={{ fontSize: '1.2rem', color: '#666' }}>
              {searchTerm ? '🔍 No contracts found matching your search.' : '📭 No contracts submitted yet.'}
            </p>
          </div>
        )}

        {!loading && filteredContracts.length > 0 && (
          <div style={styles.tableContainer}>
            <table style={styles.table}>
              <thead>
                <tr>
                  <th style={styles.th}>Patient Name</th>
                  <th style={styles.th}>DOB</th>
                  <th style={styles.th}>Responsible Party</th>
                  <th style={styles.th}>Treatment Type</th>
                  <th style={styles.th}>Date</th>
                  <th style={styles.th}>Language</th>
                  <th style={styles.th}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {filteredContracts.map((contract) => (
                  <tr key={contract.id} style={{ backgroundColor: 'white' }}>
                    <td style={styles.td}>
                      <strong>{contract.patientName}</strong>
                    </td>
                    <td style={styles.td}>{contract.dob}</td>
                    <td style={styles.td}>{contract.responsibleParty}</td>
                    <td style={styles.td}>{contract.typeOfTreatment || 'N/A'}</td>
                    <td style={styles.td}>
                      {new Date(contract.timestamp).toLocaleDateString('en-US', {
                        year: 'numeric',
                        month: 'short',
                        day: 'numeric',
                        hour: '2-digit',
                        minute: '2-digit'
                      })}
                    </td>
                    <td style={styles.td}>
                      <span style={styles.badge(contract.language)}>
                        {contract.language === 'spanish' ? 'Español' : 'English'}
                      </span>
                    </td>
                    <td style={styles.td}>
                      {contract.pdfUrl && (
                        <a
                          href={contract.pdfUrl}
                          target="_blank"
                          rel="noopener noreferrer"
                          style={{ textDecoration: 'none' }}
                        >
                          <button style={{ ...styles.button, ...styles.downloadButton }}>
                            📥 PDF
                          </button>
                        </a>
                      )}
                      <button
                        onClick={() => handleDelete(contract)}
                        style={{ ...styles.button, ...styles.deleteButton }}
                      >
                        🗑️ Delete
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {/* Detail Modal */}
        {selectedContract && (
          <div style={{
            position: 'fixed',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: 'rgba(0, 0, 0, 0.7)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 9999,
            padding: '20px'
          }}>
            <div style={{
              backgroundColor: 'white',
              borderRadius: '12px',
              maxWidth: '800px',
              width: '100%',
              maxHeight: '90vh',
              overflow: 'auto',
              boxShadow: '0 10px 40px rgba(0, 0, 0, 0.3)'
            }}>
              {/* Modal Header */}
              <div style={{
                padding: '20px',
                borderBottom: '2px solid #e0e0e0',
                backgroundColor: '#f8f9fa',
                position: 'sticky',
                top: 0,
                zIndex: 1
              }}>
                <h2 style={{ margin: '0 0 5px 0', color: '#1976d2' }}>
                  Contract Details
                </h2>
                <p style={{ margin: 0, color: '#666', fontSize: '0.9rem' }}>
                  ID: {selectedContract.id}
                </p>
              </div>

              {/* Modal Body */}
              <div style={{ padding: '25px' }}>
                <div style={{ marginBottom: '20px' }}>
                  <h3 style={{ color: '#333', marginBottom: '10px', borderBottom: '2px solid #e0e0e0', paddingBottom: '5px' }}>
                    Patient Information
                  </h3>
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '15px' }}>
                    <div>
                      <strong>Patient Name:</strong>
                      <div>{selectedContract.patientName}</div>
                    </div>
                    <div>
                      <strong>Date of Birth:</strong>
                      <div>{selectedContract.dob}</div>
                    </div>
                    <div>
                      <strong>Responsible Party:</strong>
                      <div>{selectedContract.responsibleParty}</div>
                    </div>
                    <div>
                      <strong>Relationship:</strong>
                      <div>{selectedContract.relationship}</div>
                    </div>
                    <div>
                      <strong>SS #:</strong>
                      <div>{selectedContract.ssn || 'N/A'}</div>
                    </div>
                    <div>
                      <strong>Driver's License:</strong>
                      <div>{selectedContract.driversLicense || 'N/A'}</div>
                    </div>
                  </div>
                </div>

                <div style={{ marginBottom: '20px' }}>
                  <h3 style={{ color: '#333', marginBottom: '10px', borderBottom: '2px solid #e0e0e0', paddingBottom: '5px' }}>
                    Treatment Information
                  </h3>
                  <div>
                    <strong>Type of Treatment:</strong>
                    <div>{selectedContract.typeOfTreatment || 'N/A'}</div>
                  </div>
                  <div style={{ marginTop: '10px' }}>
                    <strong>Services Required:</strong>
                    <div>{selectedContract.servicesRequired?.join(', ') || 'None'}</div>
                  </div>
                  <div style={{ marginTop: '10px' }}>
                    <strong>Additional Appliances:</strong>
                    <div>{selectedContract.additionalAppliances?.join(', ') || 'None'}</div>
                  </div>
                </div>

                {selectedContract.firstOption && (
                  <div style={{ marginBottom: '20px' }}>
                    <h3 style={{ color: '#333', marginBottom: '10px', borderBottom: '2px solid #e0e0e0', paddingBottom: '5px' }}>
                      Payment Options
                    </h3>
                    <div style={{ backgroundColor: '#f5f5f5', padding: '15px', borderRadius: '8px', marginBottom: '10px' }}>
                      <strong>First Option:</strong>
                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px', marginTop: '10px' }}>
                        <div>Treatment: ${selectedContract.firstOption.treatment}</div>
                        <div>Appliance: ${selectedContract.firstOption.appliance}</div>
                        <div>Deposit: ${selectedContract.firstOption.deposit}</div>
                        <div>Subtotal: ${selectedContract.firstOption.subtotal}</div>
                        <div>Est. Insurance: ${selectedContract.firstOption.estimatedInsurance}</div>
                        <div>Net Balance: ${selectedContract.firstOption.netBalance}</div>
                        <div>Est. Treatment Period: ${selectedContract.firstOption.estimatedTreatmentPeriod} months</div>
                        <div>Monthly Payment: ${selectedContract.firstOption.monthlyPayment}</div>
                      </div>
                    </div>
                    {selectedContract.secondOption && selectedContract.secondOption.treatment && (
                      <div style={{ backgroundColor: '#f5f5f5', padding: '15px', borderRadius: '8px' }}>
                        <strong>Second Option:</strong>
                        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px', marginTop: '10px' }}>
                          <div>Treatment: ${selectedContract.secondOption.treatment}</div>
                          <div>Appliance: ${selectedContract.secondOption.appliance}</div>
                          <div>Deposit: ${selectedContract.secondOption.deposit}</div>
                          <div>Subtotal: ${selectedContract.secondOption.subtotal}</div>
                          <div>Est. Insurance: ${selectedContract.secondOption.estimatedInsurance}</div>
                          <div>Net Balance: ${selectedContract.secondOption.netBalance}</div>
                          <div>Est. Treatment Period: ${selectedContract.secondOption.estimatedTreatmentPeriod} months</div>
                          <div>Monthly Payment: ${selectedContract.secondOption.monthlyPayment}</div>
                        </div>
                      </div>
                    )}
                  </div>
                )}

                <div style={{ marginBottom: '20px' }}>
                  <h3 style={{ color: '#333', marginBottom: '10px', borderBottom: '2px solid #e0e0e0', paddingBottom: '5px' }}>
                    Additional Information
                  </h3>
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '15px' }}>
                    <div>
                      <strong>Quote Presented By:</strong>
                      <div>{selectedContract.quotePresentedBy || 'N/A'}</div>
                    </div>
                    <div>
                      <strong>Quote Date:</strong>
                      <div>{selectedContract.quotePresentedDate || 'N/A'}</div>
                    </div>
                    <div>
                      <strong>Contract Date:</strong>
                      <div>{selectedContract.contractDate}</div>
                    </div>
                    <div>
                      <strong>Submitted:</strong>
                      <div>{new Date(selectedContract.timestamp).toLocaleString()}</div>
                    </div>
                  </div>
                </div>

                {selectedContract.pdfUrl && (
                  <div style={{ marginTop: '20px', padding: '15px', backgroundColor: '#e3f2fd', borderRadius: '8px' }}>
                    <strong>📄 PDF Document:</strong>
                    <div style={{ marginTop: '10px' }}>
                      <a
                        href={selectedContract.pdfUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                        style={{
                          display: 'inline-block',
                          padding: '10px 20px',
                          backgroundColor: '#1976d2',
                          color: 'white',
                          textDecoration: 'none',
                          borderRadius: '6px',
                          fontWeight: 'bold'
                        }}
                      >
                        📥 Download PDF
                      </a>
                    </div>
                  </div>
                )}
              </div>

              {/* Modal Footer */}
              <div style={{
                padding: '15px 20px',
                borderTop: '2px solid #e0e0e0',
                backgroundColor: '#f8f9fa',
                textAlign: 'right'
              }}>
                <button
                  onClick={() => setSelectedContract(null)}
                  style={{
                    padding: '10px 25px',
                    fontSize: '1rem',
                    fontWeight: 'bold',
                    backgroundColor: '#666',
                    color: 'white',
                    border: 'none',
                    borderRadius: '6px',
                    cursor: 'pointer'
                  }}
                >
                  Close
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    </>
  );
}

