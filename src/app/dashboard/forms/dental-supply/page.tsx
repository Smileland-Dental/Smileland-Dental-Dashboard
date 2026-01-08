'use client'

import React, { useState, useEffect, useCallback, useRef, Suspense } from "react";
import { useSearchParams } from "next/navigation";
import { collection, getDocs, addDoc, doc, setDoc, deleteDoc } from "firebase/firestore";
import { db } from "@/lib/firebase.config";
import { enableAllSecurityMeasures, sanitizeFirebaseDataClient } from "@/lib/security-client";

// 개별 아이템 행 컴포넌트
const ItemRow = React.memo(({ 
  item, 
  inputStyle,
  supplyType,
  orderQuantity,
  onQuantityChange,
  editingQuantity
}: {
  item: any;
  inputStyle: any;
  supplyType: string;
  orderQuantity: string | number;
  onQuantityChange: (itemId: string, value: string) => void;
  editingQuantity?: string | number;
}) => {
  // 편집 중인 값이 있으면 그것을 사용, 없으면 orderQuantity 사용
  const displayValue = editingQuantity !== undefined ? editingQuantity : (orderQuantity || '');
  
  return (
    <tr style={{ backgroundColor: item.displayId % 2 === 0 ? '#f9f9f9' : 'white' }}>
      <td style={{ padding: '8px', textAlign: 'center', fontWeight: 'bold' }}>
        {item.displayId}
      </td>
      {supplyType === 'processing-request' && (
        <td style={{ padding: '8px', textAlign: 'center' }}>
          <span style={{
            padding: '4px 8px',
            borderRadius: '4px',
            fontSize: '12px',
            fontWeight: 'bold',
            backgroundColor: item.supplyType === 'dental' ? '#e7f3ff' : '#f0f8ff',
            color: item.supplyType === 'dental' ? '#0077B6' : '#495057'
          }}>
            {item.supplyType === 'dental' ? '🦷 Dental' : '📋 Office'}
          </span>
        </td>
      )}
      {supplyType === 'dental' && (
        <td style={{ padding: '8px' }}>
          {item.category}
        </td>
      )}
      <td style={{ padding: '8px' }}>
        {supplyType === 'processing-request' ? (
          <div>
            <div style={{ fontWeight: 'bold', marginBottom: '4px' }}>
              {item.item}
            </div>
            <div style={{ fontSize: '12px', color: '#666' }}>
              📅 {item.orderDate ? new Date(item.orderDate).toLocaleDateString() : 'N/A'} | 
              🆔 Order #{item.requestDisplayId}
            </div>
          </div>
        ) : (
          item.item
        )}
      </td>
      <td style={{ padding: '8px' }}>
        {item.extraInfo}
      </td>
      <td style={{ padding: '8px' }}>
        {item.url && item.url.trim() !== '' ? (
          <a 
            href={item.url} 
            target="_blank" 
            rel="noopener noreferrer"
            style={{ color: '#0077B6', textDecoration: 'none' }}
          >
            {item.seller}
          </a>
        ) : (
          item.seller
        )}
      </td>
      <td style={{ padding: '8px' }}>
        {item.code}
      </td>
      <td style={{ padding: '8px', textAlign: 'center' }}>
        {item.seller === 'JB' ? (
          supplyType === 'processing-request' ? (
            // Processing Request에서는 읽기 전용으로 quantity 표시
            <span style={{
              display: 'inline-block',
              padding: '4px 8px',
              backgroundColor: '#e9ecef',
              borderRadius: '4px',
              fontWeight: 'bold',
              color: '#495057',
              minWidth: '60px',
              textAlign: 'center'
            }}>
              {supplyType === 'processing-request' ? (item.quantity || '0') : (displayValue || '0')}
            </span>
          ) : (
            // Dental/Office Supply에서는 입력 가능
            <input
              type="number"
              min="0"
              value={displayValue}
              onChange={(e) => onQuantityChange(item.id, e.target.value)}
              placeholder="0"
              style={{
                ...inputStyle,
                width: '80px',
                textAlign: 'center',
                padding: '4px 8px'
              }}
            />
          )
        ) : (
          ''
        )}
      </td>
    </tr>
  );
}, (prevProps, nextProps) => {
  // 커스텀 비교 함수로 불필요한 리렌더링 방지
  return (
    prevProps.item.id === nextProps.item.id &&
    prevProps.item.displayId === nextProps.item.displayId &&
    prevProps.item.category === nextProps.item.category &&
    prevProps.item.item === nextProps.item.item &&
    prevProps.item.extraInfo === nextProps.item.extraInfo &&
    prevProps.item.seller === nextProps.item.seller &&
    prevProps.item.code === nextProps.item.code &&
    prevProps.item.url === nextProps.item.url &&
    prevProps.supplyType === nextProps.supplyType &&
    prevProps.orderQuantity === nextProps.orderQuantity &&
    prevProps.editingQuantity === nextProps.editingQuantity
  );
});

function SupplyViewSystemContent() {
  const searchParams = useSearchParams();
  
  // Supply type 상태
  const [supplyType, setSupplyType] = useState('dental'); // 'dental', 'office', or 'processing-request'
  
  // 기본 상태
  const [loading, setLoading] = useState(false);
  const [items, setItems] = useState<any[]>([]);
  const [filteredItems, setFilteredItems] = useState<any[]>([]);
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 45;
  
  // Dental과 Office 아이템을 각각 저장
  const [dentalItems, setDentalItems] = useState<any[]>([]);
  const [officeItems, setOfficeItems] = useState<any[]>([]);
  const [processingRequests, setProcessingRequests] = useState<any[]>([]);

  // 필터 상태
  const [categoryFilter, setCategoryFilter] = useState('');
  const [sellerFilter, setSellerFilter] = useState('');
  const [searchInput, setSearchInput] = useState('');

  // Order Quantity 상태 (오피스별로 분리 저장)
  const [orderQuantitiesByOffice, setOrderQuantitiesByOffice] = useState<{ [office: string]: { [itemId: string]: string | number } }>({});
  
  // 편집 중인 Quantity 값 (오피스별로 분리 저장)
  const [editingQuantitiesByOffice, setEditingQuantitiesByOffice] = useState<{ [office: string]: { [itemId: string]: string | number } }>({});

  // Office 선택 상태
  const [selectedOffice, setSelectedOffice] = useState('');
  
  // 세션 ID (브라우저 세션마다 고유)
  const [sessionId] = useState(() => {
    // 브라우저 환경에서만 sessionStorage 사용
    if (typeof window !== 'undefined') {
      // sessionStorage에서 기존 ID 가져오거나 새로 생성
      let id = sessionStorage.getItem('dental-supply-session-id');
      if (!id) {
        id = `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
        sessionStorage.setItem('dental-supply-session-id', id);
      }
      return id;
    }
    // 서버 사이드에서는 임시 ID 생성
    return `temp-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
  });
  
  // 현재 선택된 오피스의 orderQuantities
  const orderQuantities = orderQuantitiesByOffice[selectedOffice] || {};
  
  // 현재 선택된 오피스의 editingQuantities
  const editingQuantities = editingQuantitiesByOffice[selectedOffice] || {};
  
  // Office 선택 시 임시 저장된 값들을 불러오는 함수
  const handleOfficeSelect = useCallback(async (office: string) => {
    if (!office) return;
    
    // 비밀번호 확인
    const password = prompt(`Enter password for ${office}:`);
    if (!password) return;
    
    const expectedPassword = office.charAt(0).toLowerCase();
    if (password !== expectedPassword) {
      alert('❌ Incorrect password!');
      return;
    }
    
    setSelectedOffice(office);
    
    // 임시 저장된 quantity 불러오기 (직접 호출)
    if (office) {
      try {
        const draftDoc = await getDocs(collection(db, 'office-draft-orders'));
        const currentDraft = draftDoc.docs.find(doc => {
          const data = doc.data();
          return data.office === office && data.sessionId === sessionId;
        });
        
        if (currentDraft) {
          const draftData = currentDraft.data();
          const sanitizedData = sanitizeFirebaseDataClient(draftData);
          setOrderQuantitiesByOffice(prev => {
            const newState = {
              ...prev,
              [office]: sanitizedData.quantities || {}
            };
            return newState;
          });
        } else {
        }
        
        // 같은 오피스의 다른 세션 draft들은 삭제 (오래된 draft 정리)
        const oldDrafts = draftDoc.docs.filter(doc => {
          const data = doc.data();
          return data.office === office && data.sessionId !== sessionId;
        });
        for (const oldDraft of oldDrafts) {
          try {
            await deleteDoc(doc(db, 'office-draft-orders', oldDraft.id));
          } catch (err) {
            console.error('Failed to delete old draft:', err);
          }
        }
      } catch (error) {
        console.error("Error loading draft quantities:", error);
      }
    }
  }, [sessionId]);
  
  // 디버깅용 로그
  React.useEffect(() => {
    if (selectedOffice) {
    }
  }, [selectedOffice, orderQuantities, editingQuantities]);
  
  // debounce 타이머 저장
  const quantityTimersRef = useRef<{ [key: string]: NodeJS.Timeout }>({});
  
  // Office 옵션 목록
  const officeOptions = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia', 'Corporate'];

  // Office 비밀번호 검증 (첫 알파벳 소문자)
  const getOfficePassword = (office: string) => {
    return office.charAt(0).toLowerCase();
  };



  const collectionName = supplyType === 'dental' ? 'dental-supplies' : 'office-supplies';
  const supplyTypeLabel = supplyType === 'dental' ? 'Dental' : (supplyType === 'office' ? 'Office' : 'Processing Request');
  const supplyTypeEmoji = supplyType === 'dental' ? '🦷' : (supplyType === 'office' ? '📋' : '📦');

  // 카테고리 옵션 (실제 데이터에서 동적으로 생성)
  const categoryOptions = [...new Set(items.map(item => item.category).filter(Boolean))].sort();

  // URL 파라미터에서 supply type 설정
  useEffect(() => {
    const type = searchParams.get('type');
    if (type === 'dental' || type === 'office') {
      setSupplyType(type);
    }
  }, [searchParams]);

  // supply type 변경 시 items 업데이트
  useEffect(() => {
    // supply type 변경 전에 편집 중인 값들을 먼저 저장
    if (selectedOffice && Object.keys(editingQuantities).length > 0) {
      const updatedQuantities = {
        ...(orderQuantitiesByOffice[selectedOffice] || {}),
        ...editingQuantities
      };
      setOrderQuantitiesByOffice(prev => ({
        ...prev,
        [selectedOffice]: updatedQuantities
      }));
      // Firebase에 즉시 저장
      saveDraftQuantities(selectedOffice, updatedQuantities);
    }

    if (supplyType === 'dental') {
      setItems(dentalItems);
    } else if (supplyType === 'office') {
      setItems(officeItems);
    } else if (supplyType === 'processing-request') {
      // Processing Request는 selectedOffice가 있을 때만 로드
      if (selectedOffice) {
        loadProcessingRequests();
      } else {
        setItems([]);
      }
    }
    setCategoryFilter('');
    setSellerFilter('');
    setSearchInput('');
    // supplyType 변경 시 모든 오피스의 editingQuantities 초기화
    setEditingQuantitiesByOffice({});
  }, [supplyType, dentalItems, officeItems]);

  // Processing Requests 로드 함수 (선택된 오피스의 주문 내역)
  const loadProcessingRequests = useCallback(async () => {
    
    if (!selectedOffice) {
      setItems([]);
      return;
    }

    try {
      setLoading(true);
      const requestsSnapshot = await getDocs(collection(db, 'order-requests'));
      
      const requestsList: any[] = [];
      
      requestsSnapshot.forEach((doc) => {
        const rawData = doc.data();
        const sanitizedData = sanitizeFirebaseDataClient(rawData);
        
        
        // 현재 선택된 오피스의 주문만 필터링
        if (sanitizedData.office === selectedOffice) {
          
          // 새로운 구조 (items 배열) 또는 기존 구조 (개별 문서) 처리
          if (sanitizedData.items && sanitizedData.items.length > 0) {
            // 새로운 구조: items 배열이 있는 경우
            const itemsWithQuantity = sanitizedData.items.map((item: any) => {
              const quantity = sanitizedData.quantities?.[item.id] || 0;
              return {
                ...item,
                quantity: quantity
              };
            });
            
            requestsList.push({ 
              id: doc.id, 
              ...sanitizedData,
              items: itemsWithQuantity,
              displayId: requestsList.length + 1
            });
          } else if (sanitizedData.item) {
            // 기존 구조: 개별 문서인 경우
            const quantity = sanitizedData.quantity || 0;
            const itemWithQuantity = {
              ...sanitizedData,
              quantity: quantity
            };
            
            requestsList.push({ 
              id: doc.id, 
              ...sanitizedData,
              items: [itemWithQuantity],
              displayId: requestsList.length + 1
            });
          } else {
          }
        }
      });

      
      // 주문 날짜 기준 최신순 정렬
      requestsList.sort((a, b) => {
        return new Date(b.orderDate || 0).getTime() - new Date(a.orderDate || 0).getTime();
      });

      // 주문별로 그룹화하여 표시
      setProcessingRequests(requestsList);
      setItems(requestsList);
    } catch (error) {
      console.error("Error loading processing requests:", error);
      alert('❌ 주문 내역 로드 중 오류가 발생했습니다.');
    } finally {
      setLoading(false);
    }
  }, [selectedOffice]);

  // Dental과 Office 데이터 모두 로드하는 함수
  const loadAllItems = useCallback(async () => {
    try {
      setLoading(true);
      
      // Dental items 로드 (sourceType 추가 + 보안 검증)
      const dentalSnapshot = await getDocs(collection(db, 'dental-supplies'));
      const dentalList: any[] = [];
      dentalSnapshot.forEach((doc) => {
        const rawData = doc.data();
        const sanitizedData = sanitizeFirebaseDataClient(rawData);
        dentalList.push({ id: doc.id, ...sanitizedData, sourceType: 'dental' });
      });
      dentalList.sort((a, b) => {
        if (a.order && b.order) return a.order - b.order;
        if (a.order && !b.order) return -1;
        if (!a.order && b.order) return 1;
        return new Date(a.createdAt || 0).getTime() - new Date(b.createdAt || 0).getTime();
      });
      setDentalItems(dentalList);
      
      // Office items 로드 (sourceType 추가 + 보안 검증)
      const officeSnapshot = await getDocs(collection(db, 'office-supplies'));
      const officeList: any[] = [];
      officeSnapshot.forEach((doc) => {
        const rawData = doc.data();
        const sanitizedData = sanitizeFirebaseDataClient(rawData);
        officeList.push({ id: doc.id, ...sanitizedData, sourceType: 'office' });
      });
      officeList.sort((a, b) => {
        if (a.order && b.order) return a.order - b.order;
        if (a.order && !b.order) return -1;
        if (!a.order && b.order) return 1;
        return new Date(a.createdAt || 0).getTime() - new Date(b.createdAt || 0).getTime();
      });
      setOfficeItems(officeList);
      
      // 현재 선택된 supply type에 맞는 items 설정
      if (supplyType === 'dental') {
        setItems(dentalList);
      } else {
        setItems(officeList);
      }
    } catch (error) {
      console.error("Error loading items:", error);
      alert('❌ 데이터 로드 중 오류가 발생했습니다.');
    } finally {
      setLoading(false);
    }
  }, [supplyType]);

  // 컴포넌트 마운트 시 데이터 로드
  useEffect(() => {
    loadAllItems();
  }, [loadAllItems]);

  // 보안 조치 활성화
  useEffect(() => {
    enableAllSecurityMeasures({
      disableConsole: true,        // console 비활성화
      disableRightClick: true,     // 우클릭 방지
      disableShortcuts: true,      // F12 등 단축키 방지
      disableCopy: false,          // 복사 허용 (사용자 편의)
      disableSelection: false,     // 텍스트 선택 허용 (사용자 편의)
      monitorDevTools: true        // 개발자 도구 실시간 모니터링 활성화
    });

    // 추가 보안 조치
    if (process.env.NODE_ENV === 'production') {
      // 1. 페이지 가시성 변경 감지 (탭 전환 등)
      document.addEventListener('visibilitychange', () => {
        if (document.hidden) {
        }
      });

      // 2. 페이지 포커스 감지
      window.addEventListener('blur', () => {
      });

      // 3. 키보드 이벤트 로깅 (의심스러운 패턴 감지)
      let keySequence: string[] = [];
      document.addEventListener('keydown', (e) => {
        keySequence.push(e.key);
        if (keySequence.length > 10) {
          keySequence.shift();
        }
        
        // 의심스러운 키 조합 감지
        const suspiciousPatterns = [
          ['F12'],
          ['Control', 'Shift', 'I'],
          ['Control', 'Shift', 'J'],
          ['Control', 'u']
        ];
        
        suspiciousPatterns.forEach(pattern => {
          if (keySequence.slice(-pattern.length).join(',') === pattern.join(',')) {
          }
        });
      });

      // 4. 마우스 이벤트 모니터링
      let mouseActivity = 0;
      document.addEventListener('mousemove', () => {
        mouseActivity++;
        if (mouseActivity > 10000) {
          mouseActivity = 0;
        }
      });

      // 5. 페이지 새로고침 방지 (Ctrl+R, F5)
      document.addEventListener('keydown', (e) => {
        if ((e.ctrlKey && e.key === 'r') || e.key === 'F5') {
          e.preventDefault();
          return false;
        }
      });

      // 6. 브라우저 뒤로가기 방지
      window.addEventListener('popstate', (e) => {
        e.preventDefault();
        window.history.pushState(null, '', window.location.href);
      });

      // 7. 페이지 언로드 시 경고
      window.addEventListener('beforeunload', (e) => {
        e.preventDefault();
        e.returnValue = 'Are you sure you want to leave?';
        return 'Are you sure you want to leave?';
      });
    }
  }, []);

  // 컴포넌트 언마운트 시 타이머 정리 + 편집 중인 값 저장
  useEffect(() => {
    return () => {
      // 언마운트 전에 편집 중인 값들을 먼저 저장 (모든 오피스)
      Object.keys(editingQuantitiesByOffice).forEach(office => {
        // office 미선택 상태로 생성된 임시 키('temp')는 저장 금지
        if (!office || office === 'temp') {
          return;
        }
        const editingValues = editingQuantitiesByOffice[office];
        if (Object.keys(editingValues).length > 0) {
          const updatedQuantities = {
            ...(orderQuantitiesByOffice[office] || {}),
            ...editingValues
          };
          // 동기적으로 Firebase에 저장 (async 불가능하므로 best effort)
          const draftId = `${office}-${sessionId}`;
          const draftRef = doc(db, 'office-draft-orders', draftId);
          const draftData = {
            office: office,
            sessionId: sessionId,
            quantities: updatedQuantities,
            lastUpdated: new Date().toISOString()
          };
          const sanitizedData = sanitizeFirebaseDataClient(draftData);
          setDoc(draftRef, sanitizedData).catch(err => console.error('Failed to save on unmount:', err));
        }
      });
      
      // 타이머 정리
      Object.keys(quantityTimersRef.current).forEach((key) => {
        clearTimeout(quantityTimersRef.current[key]);
      });
    };
  }, [selectedOffice, editingQuantitiesByOffice, orderQuantitiesByOffice, sessionId]);

  // supply type 변경 시 편집 중인 값 초기화 (office 변경은 유지)
  useEffect(() => {
    if (selectedOffice) {
      setEditingQuantitiesByOffice(prev => ({
        ...prev,
        [selectedOffice]: {}
      }));
    }
  }, [supplyType, selectedOffice]);

  // Office 선택 시 임시 값(temp)을 폐기하여 사전 입력이 따라오지 않도록 처리
  useEffect(() => {
    if (selectedOffice && editingQuantitiesByOffice['temp'] && Object.keys(editingQuantitiesByOffice['temp']).length > 0) {
      setEditingQuantitiesByOffice(prev => ({
        ...prev,
        temp: {}
      }));
    }
  }, [selectedOffice, editingQuantitiesByOffice]);

  // selectedOffice 변경 시 Processing Request 로드
  useEffect(() => {
    if (supplyType === 'processing-request' && selectedOffice) {
      loadProcessingRequests();
    }
  }, [selectedOffice, supplyType, loadProcessingRequests]);

  // 필터 변경 시 데이터 필터링
  useEffect(() => {
    filterItems();
  }, [items, categoryFilter, sellerFilter, searchInput]);

  // 페이지 변경 시 필터링된 데이터 업데이트
  useEffect(() => {
    setCurrentPage(1);
  }, [categoryFilter, sellerFilter, searchInput]);

  // 필터링 함수
  const filterItems = () => {
    let filtered = items;
    
    // 카테고리 필터
    if (categoryFilter) {
      filtered = filtered.filter(item => item.category === categoryFilter);
    }
    
    // 판매자 필터
    if (sellerFilter) {
      filtered = filtered.filter(item => item.seller === sellerFilter);
    }
    
    // 검색 필터
    if (searchInput) {
      filtered = filtered.filter(item => 
        item.item.toLowerCase().includes(searchInput.toLowerCase())
      );
    }
    
    setFilteredItems(filtered);
  };

  // 페이지네이션 계산
  const totalPages = Math.ceil(filteredItems.length / itemsPerPage);
  const startIndex = (currentPage - 1) * itemsPerPage;
  const endIndex = startIndex + itemsPerPage;
  const currentPageItems = filteredItems.slice(startIndex, endIndex);

  // 페이지 변경 함수
  const goToPage = (page: number) => {
    setCurrentPage(page);
  };

  // Firebase에서 오피스별 임시 저장된 quantity 불러오기
  const loadDraftQuantities = useCallback(async (office: string) => {
    if (!office) return;
    
    try {
      const draftDoc = await getDocs(collection(db, 'office-draft-orders'));
      const currentDraft = draftDoc.docs.find(doc => {
        const data = doc.data();
        return data.office === office && data.sessionId === sessionId;
      });
      
      if (currentDraft) {
        const draftData = currentDraft.data();
        const sanitizedData = sanitizeFirebaseDataClient(draftData);
        setOrderQuantitiesByOffice(prev => {
          const newState = {
            ...prev,
            [office]: sanitizedData.quantities || {}
          };
          return newState;
        });
      } else {
      }
      
      // 같은 오피스의 다른 세션 draft들은 삭제 (오래된 draft 정리)
      const oldDrafts = draftDoc.docs.filter(doc => {
        const data = doc.data();
        return data.office === office && data.sessionId !== sessionId;
      });
      for (const oldDraft of oldDrafts) {
        try {
          await deleteDoc(doc(db, 'office-draft-orders', oldDraft.id));
        } catch (err) {
          console.error('Failed to delete old draft:', err);
        }
      }
    } catch (error) {
      console.error("Error loading draft quantities:", error);
    }
  }, [sessionId]);

  // Firebase에 오피스별 임시 quantity 저장
  const saveDraftQuantities = useCallback(async (office: string, quantities: { [itemId: string]: string | number }) => {
    if (!office || office === 'temp') return;
    
    try {
      const draftId = `${office}-${sessionId}`;
      const draftRef = doc(db, 'office-draft-orders', draftId);
      const draftData = {
        office: office,
        sessionId: sessionId,
        quantities: quantities,
        lastUpdated: new Date().toISOString()
      };
      const sanitizedData = sanitizeFirebaseDataClient(draftData);
      await setDoc(draftRef, sanitizedData);
    } catch (error) {
      console.error("Error saving draft quantities:", error);
    }
  }, [sessionId]);

  // Order Quantity 변경 핸들러 (편집 중인 값을 별도 관리 + Firebase에 저장)
  const handleQuantityChange = useCallback((itemId: string, value: string) => {
    
    // 편집 중인 값을 항상 저장 (office 선택 여부와 관계없이)
    // 임시 키를 사용하여 office가 선택되기 전까지 저장
    const tempKey = selectedOffice || 'temp';
    setEditingQuantitiesByOffice(prev => ({
      ...prev,
      [tempKey]: {
        ...(prev[tempKey] || {}),
        [itemId]: value
      }
    }));
    
    // office가 선택되지 않았으면 여기서 종료 (편집 값은 저장됨)
    if (!selectedOffice) {
      return;
    }
    
    // 이전 타이머 취소
    const timerKey = `${selectedOffice}-${itemId}`;
    if (quantityTimersRef.current[timerKey]) {
      clearTimeout(quantityTimersRef.current[timerKey]);
    }
    
    // 새 타이머 설정 (300ms 후 실제 상태 업데이트 + Firebase 저장)
    quantityTimersRef.current[timerKey] = setTimeout(() => {
      const updatedQuantities = {
        ...(orderQuantitiesByOffice[selectedOffice] || {}),
        [itemId]: value
      };
      
      setOrderQuantitiesByOffice(prev => ({
        ...prev,
        [selectedOffice]: updatedQuantities
      }));
      
      // Firebase에 임시 저장
      saveDraftQuantities(selectedOffice, updatedQuantities);
      
      // 편집 완료된 값 제거 (오피스별로)
      setEditingQuantitiesByOffice(prev => {
        const officeQuantities = prev[selectedOffice] || {};
        const { [itemId]: _, ...restQuantities } = officeQuantities;
        return {
          ...prev,
          [selectedOffice]: restQuantities
        };
      });
      
      // 타이머 정리
      delete quantityTimersRef.current[timerKey];
    }, 300);
  }, [selectedOffice, orderQuantitiesByOffice, saveDraftQuantities]);

  // 주문 제출 핸들러
  const handleSubmitOrder = useCallback(async () => {
    // Office가 선택되었는지 확인
    if (!selectedOffice) {
      alert('⚠️ Please select an office first!');
      return;
    }

    // 편집 중인 값들을 포함한 최종 수량
    const currentOrderQuantities = {
      ...(orderQuantitiesByOffice[selectedOffice] || {}),
      ...(editingQuantitiesByOffice[selectedOffice] || {})  // 편집 중인 값들도 포함
    };

    // Dental과 Office 양쪽에서 주문할 아이템 확인
    const allItems = [...dentalItems, ...officeItems];
    const orderedItems = allItems.filter(item => 
      item.seller === 'JB' && currentOrderQuantities[item.id] && parseInt(String(currentOrderQuantities[item.id])) > 0
    );

    if (orderedItems.length === 0) {
      alert('⚠️ Please enter at least one order quantity!');
      return;
    }

    try {
      setLoading(true);

      // 주문 세션을 위한 고유 ID 생성
      const orderSessionId = `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
      const orderDate = new Date().toISOString();

      // 주문 데이터를 하나의 문서로 저장 (items 배열 포함)
      const orderData = {
        orderSessionId: orderSessionId,
        office: selectedOffice,
        orderDate: orderDate,
        status: 'pending',
        items: orderedItems.map(item => ({
          id: item.id,
          category: item.category || '',
          item: item.item,
          extraInfo: item.extraInfo || '',
          seller: item.seller,
          code: item.code || '',
          url: item.url || '',
          supplyType: item.sourceType || 'dental'
        })),
        quantities: orderedItems.reduce((acc, item) => {
          acc[item.id] = parseInt(String(currentOrderQuantities[item.id]));
          return acc;
        }, {} as { [itemId: string]: number })
      };

      // 데이터 검증 후 저장
      const sanitizedOrderData = sanitizeFirebaseDataClient(orderData);
      await addDoc(collection(db, 'order-requests'), sanitizedOrderData);
      

      // 주문 내역 요약
      const dentalCount = orderedItems.filter(item => item.sourceType === 'dental').length;
      const officeCount = orderedItems.filter(item => item.sourceType === 'office').length;
      
      let summary = `✅ Order submitted successfully!\n\n`;
      summary += `Total: ${orderedItems.length} items from ${selectedOffice}\n`;
      if (dentalCount > 0) summary += `🦷 Dental: ${dentalCount} items\n`;
      if (officeCount > 0) summary += `📋 Office: ${officeCount} items`;
      
      alert(summary);
      
      // 주문 후 완전한 페이지 리셋
      setOrderQuantitiesByOffice(prev => ({
        ...prev,
        [selectedOffice]: {}
      }));
      
      // 편집 중인 값들도 초기화
      setEditingQuantitiesByOffice(prev => ({
        ...prev,
        [selectedOffice]: {}
      }));
      
      // 필터 및 검색 초기화
      setCategoryFilter('');
      setSellerFilter('');
      setSearchInput('');
      
      // 페이지 초기화
      setCurrentPage(1);
      
      // Firebase에서 임시 저장 삭제 (현재 세션의 draft만)
      try {
        const draftId = `${selectedOffice}-${sessionId}`;
        const draftRef = doc(db, 'office-draft-orders', draftId);
        await deleteDoc(draftRef);
      } catch (error) {
        console.error('Error deleting draft:', error);
        // Draft 삭제 실패해도 주문은 완료됐으므로 에러 무시
      }
      
    } catch (error) {
      console.error('Error submitting order:', error);
      alert('❌ Failed to submit order. Please try again.');
    } finally {
      setLoading(false);
    }
  }, [selectedOffice, dentalItems, officeItems, orderQuantitiesByOffice, editingQuantities, sessionId]);

  // 스타일 정의
  const containerStyle = {
    maxWidth: '95%',
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

  const headerStyle: React.CSSProperties = {
    color: '#0077B6',
    textAlign: 'center',
    marginBottom: '30px',
    paddingBottom: '10px',
    borderBottom: '2px solid #BDE0FE',
    fontSize: '2.5em',
    fontWeight: 'bold'
  };

  const sectionStyle = {
    marginBottom: '30px',
    padding: '20px',
    backgroundColor: '#f0f8ff',
    borderRadius: '8px',
    borderWidth: '1px',
    borderStyle: 'solid',
    borderColor: '#BDE0FE'
  };

  const inputStyle = {
    padding: '8px 12px',
    borderWidth: '1px',
    borderStyle: 'solid',
    borderColor: '#BDE0FE',
    borderRadius: '4px',
    fontSize: '1em',
    backgroundColor: 'white',
    color: '#023047',
    width: '100%'
  };

  const buttonStyle = {
    backgroundColor: '#0077B6',
    color: 'white',
    borderWidth: '0',
    borderStyle: 'none',
    borderColor: 'transparent',
    padding: '12px 24px',
    borderRadius: '6px',
    fontSize: '1em',
    fontWeight: 'bold',
    cursor: 'pointer',
    margin: '5px',
    transition: 'all 0.3s ease'
  };

  const tableStyle: React.CSSProperties = {
    width: '100%',
    borderCollapse: 'collapse',
    marginTop: '20px',
    backgroundColor: 'white',
    borderRadius: '8px',
    overflow: 'hidden',
    boxShadow: '0 4px 8px rgba(0, 0, 0, 0.1)'
  };

  const paginationStyle = {
    display: 'flex',
    justifyContent: 'center',
    gap: '5px',
    marginTop: '20px'
  };

  const pageButtonStyle = {
    padding: '8px 16px',
    backgroundColor: 'white',
    borderWidth: '1px',
    borderStyle: 'solid',
    borderColor: '#BDE0FE',
    borderRadius: '4px',
    color: '#023047',
    cursor: 'pointer',
    transition: 'all 0.2s'
  };

  const activePageButtonStyle = {
    ...pageButtonStyle,
    backgroundColor: '#0077B6',
    color: 'white',
    borderColor: '#0077B6'
  };

  return (
    <div style={bodyStyle}>
      <div style={containerStyle}>
        {/* 헤더 + Office 선택 */}
        <div style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          marginBottom: '30px',
          paddingBottom: '20px',
          borderBottom: '2px solid #BDE0FE',
          flexWrap: 'wrap',
          gap: '20px'
        }}>
          <h1 style={{ 
            ...headerStyle, 
            marginBottom: 0, 
            paddingBottom: 0, 
            borderBottom: 'none',
            flex: '0 0 auto'
          }}>📋 Supply View</h1>
          
          {/* Office 선택 - 헤더 옆에 배치 */}
          <div style={{ flex: '0 0 auto', minWidth: '300px' }}>
            <label style={{ display: 'block', marginBottom: '8px', fontWeight: 'bold', color: '#495057', fontSize: '14px' }}>
              🏢 Select Office:
            </label>
            <select
              value={selectedOffice}
              onChange={(e) => handleOfficeSelect(e.target.value)}
              style={{
                ...inputStyle,
                fontSize: '16px',
                padding: '10px 12px',
                cursor: 'pointer',
                width: '100%'
              }}
            >
              <option value="">-- Select an Office --</option>
              {officeOptions.map(office => (
                <option key={office} value={office}>{office}</option>
              ))}
            </select>
          </div>
        </div>



        {/* Supply Type 선택 */}
        <div style={sectionStyle}>
          <h2 style={{ color: '#0077B6', marginBottom: '15px' }}>📋 Supply Type Selection</h2>
          <div style={{ display: 'flex', gap: '20px', alignItems: 'center', flexWrap: 'wrap' }}>
            <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer' }}>
              <input
                type="radio"
                name="supplyType"
                value="dental"
                checked={supplyType === 'dental'}
                onChange={(e) => setSupplyType(e.target.value)}
                style={{ margin: 0 }}
              />
              <span style={{ fontSize: '16px', fontWeight: 'bold' }}>🦷 Dental Supply</span>
            </label>
            <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer' }}>
              <input
                type="radio"
                name="supplyType"
                value="office"
                checked={supplyType === 'office'}
                onChange={(e) => setSupplyType(e.target.value)}
                style={{ margin: 0 }}
              />
              <span style={{ fontSize: '16px', fontWeight: 'bold' }}>📋 Office Supply</span>
            </label>
            <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer' }}>
              <input
                type="radio"
                name="supplyType"
                value="processing-request"
                checked={supplyType === 'processing-request'}
                onChange={(e) => setSupplyType(e.target.value)}
                disabled={!selectedOffice}
                style={{ margin: 0 }}
              />
              <span style={{ 
                fontSize: '16px', 
                fontWeight: 'bold',
                color: !selectedOffice ? '#ccc' : 'inherit'
              }}>
                📦 Processing Request
                {!selectedOffice && ' (Select Office First)'}
              </span>
            </label>
          </div>
        </div>



        {/* Processing Request가 아닐 때만 필터 적용 */}
        {supplyType !== 'processing-request' && false && (
        <div style={sectionStyle} hidden>
          <h2 style={{ color: '#0077B6', marginBottom: '15px' }}>Old Section</h2>
          <div style={{ maxWidth: '500px' }}>
            <label style={{ display: 'block', marginBottom: '8px', fontWeight: 'bold', color: '#495057' }}>
              Select Office:
            </label>
            <select
              value={selectedOffice}
              onChange={(e) => handleOfficeSelect(e.target.value)}
              style={{
                ...inputStyle,
                fontSize: '16px',
                padding: '10px 12px',
                cursor: 'pointer'
              }}
            >
              <option value="">-- Select an Office --</option>
              {officeOptions.map(office => (
                <option key={office} value={office}>{office}</option>
              ))}
            </select>
            
            
            {/* Office가 확인된 경우 표시 */}
            {selectedOffice && (
              <div style={{
                marginTop: '10px',
                padding: '10px 12px',
                backgroundColor: '#d4edda',
                borderRadius: '4px',
                color: '#155724',
                fontSize: '15px',
                fontWeight: 'bold',
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center'
              }}>
                <span>✅ Confirmed: {selectedOffice}</span>
                <button
                  onClick={() => {
                    setSelectedOffice('');
                  }}
                  style={{
                    backgroundColor: '#6c757d',
                    color: 'white',
                    border: 'none',
                    padding: '5px 10px',
                    borderRadius: '3px',
                    fontSize: '13px',
                    cursor: 'pointer'
                  }}
                >
                  Change
                </button>
              </div>
            )}
          </div>
          
        </div>
        )}

        {/* 통합 아이템 관리 섹션 */}
        <div style={sectionStyle}>
          {/* 헤더와 통계 정보 */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px' }}>
            <h2 style={{ color: '#0077B6', margin: 0 }}>📋 {supplyTypeLabel} Supply Items</h2>
            <div style={{ 
              fontSize: '14px',
              color: '#666'
            }}>
              📊 {filteredItems.length} of {items.length} {supplyType === 'processing-request' ? 'orders' : 'items'}
              {categoryFilter && ` • ${categoryFilter}`}
              {sellerFilter && ` • ${sellerFilter}`}
              {searchInput && ` • "${searchInput}"`}
            </div>
          </div>
          
          {/* 필터 컨트롤 */}
          <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap', marginBottom: '20px' }}>
            {supplyType === 'dental' && (
              <div style={{ flex: '1', minWidth: '200px' }}>
                <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                  Category:
                </label>
                <select
                  value={categoryFilter}
                  onChange={(e) => setCategoryFilter(e.target.value)}
                  style={inputStyle}
                >
                  <option value="">All Categories</option>
                  {categoryOptions.map(category => (
                    <option key={category} value={category}>{category}</option>
                  ))}
                </select>
              </div>
            )}
            
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Seller:
              </label>
              <select
                value={sellerFilter}
                onChange={(e) => setSellerFilter(e.target.value)}
                style={inputStyle}
              >
                <option value="">All Sellers</option>
                {[...new Set(items.map(item => item.seller).filter(Boolean))].sort().map(seller => (
                  <option key={seller} value={seller}>{seller}</option>
                ))}
              </select>
            </div>

            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Search:
              </label>
              <input
                type="text"
                value={searchInput}
                onChange={(e) => setSearchInput(e.target.value)}
                placeholder="Search item name..."
                style={inputStyle}
              />
            </div>

            <div style={{ flex: '1', minWidth: '200px', display: 'flex', alignItems: 'end' }}>
              <button 
                onClick={loadAllItems}
                disabled={loading}
                style={{
                  ...buttonStyle,
                  backgroundColor: '#28a745',
                  width: '100%'
                }}
              >
                {loading ? 'Loading...' : '🔄 Refresh Data'}
              </button>
            </div>
          </div>
          
          {loading ? (
            <div style={{ textAlign: 'center', padding: '40px', color: '#666' }}>
              Loading items...
            </div>
          ) : filteredItems.length === 0 ? (
            <div style={{ textAlign: 'center', padding: '40px', color: '#666' }}>
              No items found for the selected criteria.
            </div>
          ) : supplyType === 'processing-request' ? (
            // Processing Request: 주문별로 그룹화하여 표시
            <div style={{ overflowX: 'auto' }}>
              {currentPageItems.map((order, orderIndex) => (
                <div key={order.id} style={{ marginBottom: '30px', border: '1px solid #e0e0e0', borderRadius: '8px', overflow: 'hidden' }}>
                  {/* 주문 헤더 */}
                  <div style={{ 
                    backgroundColor: '#0077B6', 
                    color: 'white', 
                    padding: '15px 20px',
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center'
                  }}>
                    <div>
                      <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 'bold' }}>
                        📦 Order #{startIndex + orderIndex + 1}
                      </h3>
                      <div style={{ fontSize: '14px', opacity: 0.9, marginTop: '5px' }}>
                        📅 {order.orderDate ? new Date(order.orderDate).toLocaleDateString() : 'N/A'} | 
                        📊 {order.items?.length || 0} items
                      </div>
                    </div>
                  </div>

                  {/* 주문 아이템 테이블 */}
                  {order.items && order.items.length > 0 && (
                    <table style={{ ...tableStyle, margin: 0, border: 'none' }}>
                      <thead style={{ backgroundColor: '#f8f9fa', color: '#495057' }}>
                        <tr>
                          <th style={{ padding: '10px 8px', textAlign: 'center', minWidth: '60px' }}>#</th>
                          <th style={{ padding: '10px 8px', textAlign: 'center', minWidth: '120px' }}>Supply Type</th>
                          <th style={{ padding: '10px 8px', textAlign: 'center', minWidth: '400px' }}>Item</th>
                          <th style={{ padding: '10px 8px', textAlign: 'center', minWidth: '250px' }}>Extra Info</th>
                          <th style={{ padding: '10px 8px', textAlign: 'center', minWidth: '200px' }}>Seller</th>
                          <th style={{ padding: '10px 8px', textAlign: 'center', minWidth: '150px' }}>Code</th>
                          <th style={{ padding: '10px 8px', textAlign: 'center', minWidth: '120px' }}>Quantity</th>
                        </tr>
                      </thead>
                      <tbody>
                        {order.items.map((item: any, itemIndex: number) => (
                          <tr key={`${order.id}-${item.id}-${itemIndex}`} style={{ backgroundColor: itemIndex % 2 === 0 ? '#f9f9f9' : 'white' }}>
                            <td style={{ padding: '8px', textAlign: 'center', fontWeight: 'bold' }}>
                              {itemIndex + 1}
                            </td>
                            <td style={{ padding: '8px', textAlign: 'center' }}>
                              <span style={{
                                padding: '4px 8px',
                                borderRadius: '4px',
                                fontSize: '12px',
                                fontWeight: 'bold',
                                backgroundColor: item.supplyType === 'dental' ? '#e7f3ff' : '#f0f8ff',
                                color: item.supplyType === 'dental' ? '#0077B6' : '#495057'
                              }}>
                                {item.supplyType === 'dental' ? '🦷 Dental' : '📋 Office'}
                              </span>
                            </td>
                            <td style={{ padding: '8px' }}>
                              <div style={{ fontWeight: 'bold', marginBottom: '4px' }}>
                                {item.item}
                              </div>
                            </td>
                            <td style={{ padding: '8px' }}>
                              {item.extraInfo}
                            </td>
                            <td style={{ padding: '8px' }}>
                              {item.url && item.url.trim() !== '' ? (
                                <a 
                                  href={item.url} 
                                  target="_blank" 
                                  rel="noopener noreferrer"
                                  style={{ color: '#0077B6', textDecoration: 'none' }}
                                >
                                  {item.seller}
                                </a>
                              ) : (
                                item.seller
                              )}
                            </td>
                            <td style={{ padding: '8px' }}>
                              {item.code}
                            </td>
                            <td style={{ padding: '8px', textAlign: 'center' }}>
                              <span style={{
                                display: 'inline-block',
                                padding: '4px 8px',
                                backgroundColor: '#e9ecef',
                                borderRadius: '4px',
                                fontWeight: 'bold',
                                color: '#495057',
                                minWidth: '60px',
                                textAlign: 'center'
                              }}>
                                {item.quantity || '0'}
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>
              ))}
            </div>
          ) : (
            // Dental/Office Supply: 기존 테이블 구조
            <div style={{ overflowX: 'auto' }}>
              <table style={tableStyle}>
                <thead style={{ backgroundColor: '#0077B6', color: 'white' }}>
                  <tr>
                    <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '60px' }}>#</th>
                    {supplyType === 'dental' && (
                      <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '250px' }}>Category</th>
                    )}
                    <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '400px' }}>Item</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '250px' }}>Extra Info</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '200px' }}>Seller</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '150px' }}>Code</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '180px' }}>Order Quantity</th>
                  </tr>
                </thead>
                <tbody>
                  {currentPageItems.map((item, index) => (
                    <ItemRow
                      key={`${item.id}-${item.requestId || 'no-request'}-${index}`}
                      item={{...item, displayId: startIndex + index + 1}}
                      inputStyle={inputStyle}
                      supplyType={supplyType}
                      orderQuantity={orderQuantities[item.id]}
                      onQuantityChange={handleQuantityChange}
                      editingQuantity={editingQuantities[item.id]}
                    />
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* 페이지네이션 */}
          {totalPages > 1 && (
            <div style={paginationStyle}>
              {/* Previous 버튼 */}
              {currentPage > 1 && (
                <button
                  onClick={() => goToPage(currentPage - 1)}
                  style={pageButtonStyle}
                >
                  Previous
                </button>
              )}

              {/* 페이지 번호들 */}
              {Array.from({ length: totalPages }, (_, i) => i + 1).map(page => (
                <button
                  key={page}
                  onClick={() => goToPage(page)}
                  style={currentPage === page ? activePageButtonStyle : pageButtonStyle}
                >
                  {page}
                </button>
              ))}

              {/* Next 버튼 */}
              {currentPage < totalPages && (
                <button
                  onClick={() => goToPage(currentPage + 1)}
                  style={pageButtonStyle}
                >
                  Next
                </button>
              )}
            </div>
          )}

          {/* 페이지 정보 */}
          <div style={{ 
            textAlign: 'center', 
            marginTop: '10px', 
            color: '#666', 
            fontSize: '14px' 
          }}>
            {supplyType === 'processing-request' 
              ? `Showing ${startIndex + 1} to ${Math.min(endIndex, filteredItems.length)} of ${filteredItems.length} orders`
              : `Showing ${startIndex + 1} to ${Math.min(endIndex, filteredItems.length)} of ${filteredItems.length} items`
            }
          </div>
        </div>

        {/* Submit Order Button - 맨 아래 중앙에 위치 */}
        {selectedOffice && supplyType !== 'processing-request' && (
          <div style={{
            textAlign: 'center',
            padding: '30px'
          }}>
            <button
              onClick={handleSubmitOrder}
              disabled={loading}
              style={{
                ...buttonStyle,
                backgroundColor: '#28a745',
                fontSize: '18px',
                padding: '16px 32px',
                fontWeight: 'bold',
                boxShadow: '0 4px 8px rgba(0,0,0,0.1)',
                transition: 'all 0.3s ease'
              }}
            >
              {loading ? '⏳ Submitting...' : '📦 Submit Order'}
            </button>
          </div>
        )}

      </div>
    </div>
  );
}

export default function SupplyViewPage() {
  return (
    <>
      <style>{`
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
      `}</style>
      <Suspense fallback={
        <div style={{ 
          display: 'flex', 
          justifyContent: 'center', 
          alignItems: 'center', 
          minHeight: '100vh',
          fontFamily: 'Segoe UI, Tahoma, Geneva, Verdana, sans-serif',
          backgroundColor: '#f8f9fa'
        }}>
          <div style={{ textAlign: 'center' }}>
            <div style={{
              border: '4px solid #f3f3f3',
              borderTop: '4px solid #0077B6',
              borderRadius: '50%',
              width: '50px',
              height: '50px',
              animation: 'spin 1s linear infinite',
              margin: '0 auto 20px'
            }}></div>
            <p style={{ color: '#666' }}>Loading...</p>
          </div>
        </div>
      }>
        <SupplyViewSystemContent />
      </Suspense>
    </>
  );
}
