'use client'

import React, { useState, useEffect, useCallback, useRef, Suspense } from "react";
import { useSearchParams, useRouter } from "next/navigation";
import { doc, setDoc, collection, getDocs, getDoc, updateDoc, deleteDoc, writeBatch } from "firebase/firestore";
import { db } from "@/lib/firebase.config";
import { enableAllSecurityMeasures, sanitizeFirebaseDataClient } from "@/lib/security-client";

// 개별 아이템 행 컴포넌트 (엑셀 스타일)
interface ItemRowProps {
  item: any;
  updateItem: (itemId: string, field: string, value: string) => void;
  deleteItem: (itemId: string) => void;
  isSelected: boolean;
  onSelect: (itemId: string) => void;
  excelInputStyle: React.CSSProperties;
  supplyType: string;
  onInsertBelow: (displayId: number) => void;
  editingValues: { [key: string]: any };
}

const ItemRow = React.memo(({ 
  item, 
  updateItem, 
  deleteItem, 
  isSelected,
  onSelect,
  excelInputStyle,
  supplyType,
  onInsertBelow,
  editingValues
}: ItemRowProps) => {
  // 편집 중인 값이 있으면 그것을 사용, 없으면 item 값 사용
  const getValue = (field: string) => {
    return editingValues[item.id]?.[field] ?? item[field];
  };

  return (
    <tr 
      style={{ 
        backgroundColor: isSelected ? '#e3f2fd' : (item.displayId % 2 === 0 ? '#f9f9f9' : 'white'),
        borderBottom: '1px solid #e0e0e0'
      }}
      onClick={() => onSelect(item.id)}
    >
      <td style={{ padding: '6px', textAlign: 'center', fontWeight: 'bold', minWidth: '50px' }}>
        <input
          type="checkbox"
          checked={isSelected}
          onChange={() => onSelect(item.id)}
          style={{ margin: 0, width: '16px', height: '16px', cursor: 'pointer' }}
        />
      </td>
      <td style={{ padding: '6px', textAlign: 'center', fontWeight: 'bold', minWidth: '50px' }}>
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '4px' }}>
          <span style={{ fontSize: '15px' }}>{item.displayId}</span>
          <button
            onClick={(e) => {
              e.stopPropagation();
              onInsertBelow(item.displayId);
            }}
            style={{
              fontSize: '11px',
              padding: '3px 7px',
              backgroundColor: '#17a2b8',
              color: 'white',
              border: 'none',
              borderRadius: '3px',
              cursor: 'pointer'
            }}
            title={`Insert new item after this row (position ${item.displayId + 1})`}
          >
            ➕
          </button>
        </div>
      </td>
      {supplyType === 'dental' && (
        <td style={{ padding: '0' }}>
          <input
            type="text"
            value={getValue('category')}
            onChange={(e) => updateItem(item.id, 'category', e.target.value)}
            style={excelInputStyle}
          />
        </td>
      )}
      <td style={{ padding: '0' }}>
        <input
          type="text"
          value={getValue('item')}
          onChange={(e) => updateItem(item.id, 'item', e.target.value)}
          style={excelInputStyle}
        />
      </td>
      <td style={{ padding: '0' }}>
        <input
          type="text"
          value={getValue('extraInfo')}
          onChange={(e) => updateItem(item.id, 'extraInfo', e.target.value)}
          style={excelInputStyle}
        />
      </td>
      <td style={{ padding: '0' }}>
        <input
          type="text"
          value={getValue('price')}
          onChange={(e) => updateItem(item.id, 'price', e.target.value)}
          style={excelInputStyle}
        />
      </td>
      <td style={{ padding: '0' }}>
        <input
          type="text"
          value={getValue('seller')}
          onChange={(e) => updateItem(item.id, 'seller', e.target.value)}
          style={excelInputStyle}
        />
      </td>
      <td style={{ padding: '0' }}>
        <input
          type="text"
          value={getValue('code')}
          onChange={(e) => updateItem(item.id, 'code', e.target.value)}
          style={excelInputStyle}
        />
      </td>
      <td style={{ padding: '0' }}>
        <input
          type="text"
          value={getValue('url')}
          onChange={(e) => updateItem(item.id, 'url', e.target.value)}
          style={excelInputStyle}
        />
      </td>
      <td style={{ padding: '6px', textAlign: 'center' }}>
        <button
          onClick={(e) => {
            e.stopPropagation();
            deleteItem(item.id);
          }}
          style={{
            backgroundColor: '#dc3545',
            color: 'white',
            borderWidth: '0',
            borderStyle: 'none',
            borderColor: 'transparent',
            padding: '6px 10px',
            borderRadius: '3px',
            fontSize: '14px',
            cursor: 'pointer'
          }}
        >
          Delete
        </button>
      </td>
    </tr>
  );
}, (prevProps, nextProps) => {
  // editingValues의 해당 item에 대한 값만 비교
  const prevEditing = prevProps.editingValues[prevProps.item.id] || {};
  const nextEditing = nextProps.editingValues[nextProps.item.id] || {};
  
  return (
    prevProps.item.id === nextProps.item.id &&
    prevProps.item.displayId === nextProps.item.displayId &&
    prevProps.item.category === nextProps.item.category &&
    prevProps.item.item === nextProps.item.item &&
    prevProps.item.extraInfo === nextProps.item.extraInfo &&
    prevProps.item.price === nextProps.item.price &&
    prevProps.item.seller === nextProps.item.seller &&
    prevProps.item.code === nextProps.item.code &&
    prevProps.item.url === nextProps.item.url &&
    prevProps.isSelected === nextProps.isSelected &&
    prevProps.supplyType === nextProps.supplyType &&
    JSON.stringify(prevEditing) === JSON.stringify(nextEditing)
  );
});

ItemRow.displayName = 'ItemRow';

function SupplyManagerSystemContent() {
  const searchParams = useSearchParams();
  const router = useRouter();
  const [supplyType, setSupplyType] = useState('order-request'); // 'dental', 'office', or 'order-request'
  const [items, setItems] = useState<any[]>([]);
  const [filteredItems, setFilteredItems] = useState<any[]>([]);
  const [selectedItems, setSelectedItems] = useState<Set<string>>(new Set());
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [categoryFilter, setCategoryFilter] = useState('');
  const [sellerFilter, setSellerFilter] = useState('');
  const [searchInput, setSearchInput] = useState('');
  const [newItem, setNewItem] = useState({
    category: '',
    item: '',
    extraInfo: '',
    price: '',
    seller: '',
    code: '',
    url: ''
  });
  const [insertAfterRow, setInsertAfterRow] = useState('');
  const [expandedGroups, setExpandedGroups] = useState<Set<string>>(new Set());
  const [editingValues, setEditingValues] = useState<{ [key: string]: any }>({});
  
  

  const collectionName = supplyType === 'order-request' ? 'order-requests' : (supplyType === 'dental' ? 'dental-supplies' : 'office-supplies');
  const supplyTypeLabel = supplyType === 'order-request' ? 'Order Request' : (supplyType === 'dental' ? 'Dental' : 'Office');
  const supplyTypeEmoji = supplyType === 'order-request' ? '📦' : (supplyType === 'dental' ? '🦷' : '📋');

  // 카테고리 옵션 동적 생성
  const categoryOptions = [...new Set(items.map(item => item.category).filter(Boolean))].sort();

  // 데이터 로드
  const loadItems = useCallback(async () => {
    try {
      setLoading(true);
      setError(null);
      
      const querySnapshot = await getDocs(collection(db, collectionName));
      const itemsList: any[] = [];
      
      querySnapshot.forEach((doc: any) => {
        const rawData = doc.data();
        const sanitizedData = sanitizeFirebaseDataClient(rawData);
        // 임시 저장된 주문(office가 비어있거나 'temp')은 목록에서 제외
        if (supplyType === 'order-request') {
          const officeValue = (sanitizedData.office || '').toString().trim().toLowerCase();
          if (!officeValue || officeValue === 'temp') {
            return; // skip this document
          }
        }
        
        // 새로운 구조 (items 배열) 처리
        if (sanitizedData.items && sanitizedData.items.length > 0) {
          // items 배열의 각 아이템을 개별 문서로 변환
          sanitizedData.items.forEach((item: any, index: number) => {
            const quantity = sanitizedData.quantities?.[item.id] || 0;
            const perItemStatus = sanitizedData.itemStatuses?.[item.id] || sanitizedData.status;
            itemsList.push({
              id: `${doc.id}-${index}`, // 고유 ID 생성 (doc.id-index 형태)
              originalItemId: item.id, // 원본 아이템 ID 보존
              ...item,
              quantity: quantity,
              orderSessionId: sanitizedData.orderSessionId,
              office: sanitizedData.office,
              orderDate: sanitizedData.orderDate,
              status: perItemStatus,
              parentDocId: doc.id // 부모 문서 ID 보존
            });
          });
        } else {
          // 기존 구조 (개별 문서) 처리
          itemsList.push({
            id: doc.id,
            ...sanitizedData
          });
        }
      });

      // order 필드로 정렬 (order가 없으면 createdAt으로 정렬)
      itemsList.sort((a, b) => {
        if (a.order && b.order) {
          return a.order - b.order;
        } else if (a.order && !b.order) {
          return -1;
        } else if (!a.order && b.order) {
          return 1;
        } else {
          return new Date(a.createdAt || 0).getTime() - new Date(b.createdAt || 0).getTime();
        }
      });

      // displayId 추가
      const itemsWithDisplayId = itemsList.map((item, index) => ({
        ...item,
        displayId: index + 1
      }));

      setItems(itemsWithDisplayId);
      setFilteredItems(itemsWithDisplayId);
    } catch (error) {
      console.error('Error loading items:', error);
      setError('Failed to load items');
    } finally {
      setLoading(false);
    }
  }, [collectionName]);

  // Firebase 업데이트를 위한 debounce 타이머 저장
  const updateTimersRef = useRef<{ [key: string]: NodeJS.Timeout }>({});

  // 개별 아이템 업데이트 (편집 중인 값을 별도 관리)
  const updateItem = useCallback((itemId: string, field: string, value: string) => {
    // 편집 중인 값을 별도 상태에 저장 (입력이 부드럽게 됨)
    setEditingValues(prev => ({
      ...prev,
      [itemId]: {
        ...(prev[itemId] || {}),
        [field]: value
      }
    }));

    // 이전 타이머 취소
    const timerKey = `${itemId}-${field}`;
    if (updateTimersRef.current[timerKey]) {
      clearTimeout(updateTimersRef.current[timerKey]);
    }

    // 새 타이머 설정 (1초 후 Firebase 및 items 업데이트)
    updateTimersRef.current[timerKey] = setTimeout(async () => {
      try {
        const itemRef = doc(db, collectionName, itemId);
        const updateData = {
          [field]: value,
          lastUpdated: new Date().toISOString()
        };
        const sanitizedUpdateData = sanitizeFirebaseDataClient(updateData);
        await updateDoc(itemRef, sanitizedUpdateData);
        
        // items 상태 업데이트
        setItems(prevItems => 
          prevItems.map(item => 
            item.id === itemId 
              ? { ...item, [field]: value, lastUpdated: new Date().toISOString() }
              : item
          )
        );
        
        // 편집 완료된 값 제거
        setEditingValues(prev => {
          const newValues = { ...prev };
          if (newValues[itemId]) {
            delete newValues[itemId][field];
            if (Object.keys(newValues[itemId]).length === 0) {
              delete newValues[itemId];
            }
          }
          return newValues;
        });
        
        // 타이머 정리
        delete updateTimersRef.current[timerKey];
      } catch (error) {
        console.error('Error updating item:', error);
        setError('Failed to update item');
      }
    }, 1000);
  }, [collectionName]);

  // 새 아이템 추가
  const addItem = useCallback(async () => {
    if (!newItem.item.trim()) {
      setError('Item name is required');
      return;
    }

    try {
      setLoading(true);
      
      let newOrder: number;
      
      // 특정 행 뒤에 삽입하는 경우
      if (insertAfterRow && insertAfterRow.trim() !== '') {
        const insertAfterNumber = parseInt(insertAfterRow);
        
        if (isNaN(insertAfterNumber) || insertAfterNumber < 1 || insertAfterNumber > items.length) {
          setError(`Invalid row number. Please enter a number between 1 and ${items.length}`);
          setLoading(false);
          return;
        }
        
        // 1이면 맨 앞에 추가
        if (insertAfterNumber === 1) {
          newOrder = 0.5;
        } else {
          // 해당 행 찾기 (insertAfterNumber - 1이 실제 행 번호)
          const targetItem = items.find(item => item.displayId === insertAfterNumber - 1);
          if (!targetItem) {
            setError('Target row not found');
            setLoading(false);
            return;
          }
          
          // 해당 행의 order 값 가져오기
          const currentOrder = targetItem.order || (insertAfterNumber - 1);
          
          // 다음 아이템의 order 값 찾기
          const nextItem = items.find(item => (item.order || item.displayId) > currentOrder);
          
          if (nextItem) {
            // 사이에 삽입
            const nextOrder_value = nextItem.order || nextItem.displayId;
            newOrder = (currentOrder + nextOrder_value) / 2;
          } else {
            // 맨 끝에 추가
            newOrder = currentOrder + 1;
          }
        }
        
        // order 값이 너무 가까워지면 전체 재정렬
        const needsReordering = items.some(item => {
          const itemOrder = item.order || 0;
          return Math.abs(itemOrder - newOrder) < 0.001 && itemOrder !== newOrder;
        });
        
        if (needsReordering) {
          // 배치 작업으로 모든 아이템 재정렬
          const batch = writeBatch(db);
          items.forEach((item, index) => {
            const itemRef = doc(db, collectionName, item.id);
            batch.update(itemRef, { order: index + 1 });
          });
          await batch.commit();
          await loadItems();
          
          // 재로드 후 다시 계산
          const reloadedItems = items.map((item, index) => ({ ...item, order: index + 1 }));
          const targetItem = reloadedItems[insertAfterNumber - 1];
          if (targetItem) {
            newOrder = targetItem.order + 0.5;
          } else {
            newOrder = (insertAfterNumber - 1) + 0.5;
          }
        }
      } else {
        // 맨 끝에 추가
        const maxOrder = items.length > 0 ? Math.max(...items.map(item => item.order || 0)) : 0;
        newOrder = maxOrder + 1;
      }

      const itemData = {
        ...newItem,
        order: newOrder,
        createdAt: new Date().toISOString(),
        lastUpdated: new Date().toISOString()
      };

      // 데이터 검증 후 저장
      const sanitizedItemData = sanitizeFirebaseDataClient(itemData);
      const newItemRef = doc(collection(db, collectionName));
      await setDoc(newItemRef, sanitizedItemData);

      setNewItem({
        category: '',
        item: '',
        extraInfo: '',
        price: '',
        seller: '',
        code: '',
        url: ''
      });
      setInsertAfterRow('');

      await loadItems();
    } catch (error) {
      console.error('Error adding item:', error);
      setError('Failed to add item');
    } finally {
      setLoading(false);
    }
  }, [newItem, items, collectionName, loadItems, insertAfterRow]);

  


  // 선택된 아이템들 삭제
  const deleteSelectedItems = useCallback(async () => {
    if (selectedItems.size === 0) {
      setError('Please select items to delete');
      return;
    }

    if (!confirm(`Are you sure you want to delete ${selectedItems.size} items?`)) {
      return;
    }

    try {
      setLoading(true);
      const batch = writeBatch(db);

      selectedItems.forEach((itemId: string) => {
        const itemRef = doc(db, collectionName, itemId);
        batch.delete(itemRef);
      });

      await batch.commit();
      setSelectedItems(new Set());
      await loadItems();
    } catch (error) {
      console.error('Error deleting items:', error);
      setError('Failed to delete items');
    } finally {
      setLoading(false);
    }
  }, [selectedItems, collectionName, loadItems]);

  // 개별 아이템 삭제
  const deleteItem = useCallback(async (itemId: string) => {
    if (typeof itemId !== 'string') {
      console.error('Invalid itemId type:', typeof itemId);
      return;
    }

    if (!confirm('Are you sure you want to delete this item?')) {
      return;
    }

    try {
      setLoading(true);
      await deleteDoc(doc(db, collectionName, itemId));
      await loadItems();
    } catch (error) {
      console.error('Delete item error:', error);
      setError('Failed to delete item');
    } finally {
      setLoading(false);
    }
  }, [collectionName, loadItems]);

  // 아이템 선택 토글
  const toggleItemSelection = useCallback((itemId: string) => {
    setSelectedItems(prev => {
      const newSet = new Set(prev);
      if (newSet.has(itemId)) {
        newSet.delete(itemId);
      } else {
        newSet.add(itemId);
      }
      return newSet;
    });
  }, []);

  // 특정 행 뒤에 삽입하기 위한 함수
  const handleInsertBelow = useCallback((displayId: number) => {
    setInsertAfterRow((displayId + 1).toString());
    // 스크롤을 추가 폼으로 이동
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }, []);

  // 전체 선택/해제
  const toggleSelectAll = useCallback(() => {
    if (selectedItems.size === filteredItems.length) {
      setSelectedItems(new Set());
    } else {
      setSelectedItems(new Set(filteredItems.map(item => item.id)));
    }
  }, [selectedItems.size, filteredItems]);

  // Order Request 그룹화 로직 (orderSessionId 기준)
  const groupedOrders = useCallback(() => {
    if (supplyType !== 'order-request') return {};
    
    const groups: any = {};
    filteredItems.forEach(item => {
      const key = item.orderSessionId || 'unknown';
      if (!groups[key]) {
        groups[key] = {
          orderSessionId: key,
          office: item.office,
          orderDate: item.orderDate,
          items: [],
          totalQuantity: 0,
          itemCount: 0,
          supplyTypes: new Set()  // 여러 supply type을 추적
        };
      }
      groups[key].items.push(item);
      groups[key].totalQuantity += parseInt(item.quantity || 0);
      groups[key].itemCount += 1;
      groups[key].supplyTypes.add(item.supplyType);  // supply type 추가
    });
    
    // supplyTypes Set을 배열로 변환
    Object.values(groups).forEach((group: any) => {
      group.supplyTypesArray = Array.from(group.supplyTypes);
    });
    
    // 주문 날짜 기준으로 정렬 (최신순)
    const sortedGroups = Object.entries(groups).sort(([, a]: [string, any], [, b]: [string, any]) => {
      return new Date(b.orderDate).getTime() - new Date(a.orderDate).getTime();
    });
    
    return Object.fromEntries(sortedGroups);
  }, [filteredItems, supplyType]);

  // 그룹 토글
  const toggleGroup = useCallback((key: string) => {
    setExpandedGroups(prev => {
      const newSet = new Set(prev);
      if (newSet.has(key)) {
        newSet.delete(key);
      } else {
        newSet.add(key);
      }
      return newSet;
    });
  }, []);

  // 주문 상태 변경
  const updateOrderStatus = useCallback(async (itemId: string, newStatus: string) => {
    try {
      // 새로운 구조에서는 개별 item.id가 `${docId}-${index}` 또는 parentDocId를 가짐
      const targetItem = items.find((it) => it.id === itemId);
      let docId = itemId;
      if (targetItem?.parentDocId) {
        docId = targetItem.parentDocId;
      } else if (itemId.includes('-')) {
        docId = itemId.split('-')[0];
      }

      const itemRef = doc(db, 'order-requests', docId);
      
      if (targetItem?.originalItemId) {
        // 새로운 구조: itemStatuses 객체 업데이트
        const docSnap = await getDoc(itemRef);
        if (docSnap.exists()) {
          const currentData = docSnap.data();
          const currentItemStatuses = currentData.itemStatuses || {};
          currentItemStatuses[targetItem.originalItemId] = newStatus;
          
          const statusData = {
            itemStatuses: currentItemStatuses,
            lastUpdated: new Date().toISOString()
          };
          const sanitizedStatusData = sanitizeFirebaseDataClient(statusData);
          await updateDoc(itemRef, sanitizedStatusData);
        }
      } else {
        // 기존 구조 호환
        const statusData = {
          status: newStatus,
          lastUpdated: new Date().toISOString()
        };
        const sanitizedStatusData = sanitizeFirebaseDataClient(statusData);
        await updateDoc(itemRef, sanitizedStatusData);
      }

      // 로컬 상태 업데이트: 해당 아이템만 변경
      setItems(prevItems => 
        prevItems.map(item => 
          item.id === itemId 
            ? { ...item, status: newStatus, lastUpdated: new Date().toISOString() }
            : item
        )
      );
      
    } catch (error) {
      console.error('Error updating order status:', error);
      alert('❌ Failed to update status. Please try again.');
    }
  }, [items]);

  // 주문 그룹 전체 삭제
  const deleteOrderGroup = useCallback(async (orderSessionId: string, itemCount: number) => {
    if (!confirm(`Are you sure you want to delete this entire order (${itemCount} items)?\n\nThis action cannot be undone.`)) {
      return;
    }

    try {
      setLoading(true);
      
      // 해당 orderSessionId의 모든 아이템 찾기
      const itemsToDelete = items.filter(item => item.orderSessionId === orderSessionId);
      
      // 새로운 구조와 기존 구조 모두 처리
      const documentsToDelete = new Set<string>();
      
      itemsToDelete.forEach(item => {
        // 새로운 구조: parentDocId가 있으면 그것을 사용, 없으면 docId-itemId에서 추출
        if (item.parentDocId) {
          documentsToDelete.add(item.parentDocId);
        } else if (item.id.includes('-')) {
          const parts = item.id.split('-');
          const docId = parts[0];
          documentsToDelete.add(docId);
        } else {
          // 기존 구조: 직접 문서 ID 사용
          documentsToDelete.add(item.id);
        }
      });
      
      
      // 배치 삭제
      const batch = writeBatch(db);
      documentsToDelete.forEach(docId => {
        const itemRef = doc(db, 'order-requests', docId);
        batch.delete(itemRef);
      });
      
      await batch.commit();
      
      // 로컬 상태 업데이트
      setItems(prevItems => prevItems.filter(item => item.orderSessionId !== orderSessionId));
      
      alert('✅ Order deleted successfully!');
    } catch (error) {
      console.error('Error deleting order group:', error);
      alert('❌ Failed to delete order. Please try again.');
    } finally {
      setLoading(false);
    }
  }, [items]);

  // 필터링 로직
  useEffect(() => {
    let filtered = [...items];

    if (categoryFilter) {
      filtered = filtered.filter(item => item.category === categoryFilter);
    }

    if (sellerFilter) {
      filtered = filtered.filter(item => item.seller === sellerFilter);
    }

    if (searchInput.trim()) {
      const searchTerm = searchInput.toLowerCase();
      filtered = filtered.filter(item => 
        item.item?.toLowerCase().includes(searchTerm) ||
        item.extraInfo?.toLowerCase().includes(searchTerm) ||
        item.seller?.toLowerCase().includes(searchTerm) ||
        item.code?.toLowerCase().includes(searchTerm)
      );
    }

    // displayId 재계산
    const filteredWithDisplayId = filtered.map((item, index) => ({
      ...item,
      displayId: index + 1
    }));

    setFilteredItems(filteredWithDisplayId);
  }, [items, categoryFilter, sellerFilter, searchInput]);

  // supply type 변경 시 데이터 리로드
  useEffect(() => {
    loadItems();
    setSelectedItems(new Set());
    setCategoryFilter('');
    setSellerFilter('');
    setSearchInput('');
    setEditingValues({});
  }, [supplyType, loadItems]);

  // URL 파라미터에서 supply type 설정
  useEffect(() => {
    const type = searchParams.get('type');
    if (type === 'dental' || type === 'office' || type === 'order-request') {
      setSupplyType(type);
    }
  }, [searchParams]);

  // 초기 로드
  useEffect(() => {
    loadItems();
  }, [loadItems]);

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

  // 컴포넌트 언마운트 시 타이머 정리
  useEffect(() => {
    return () => {
      Object.keys(updateTimersRef.current).forEach((key) => {
        clearTimeout(updateTimersRef.current[key]);
      });
    };
  }, []);

  // 스타일 정의
  const bodyStyle = {
    fontFamily: 'Segoe UI, Tahoma, Geneva, Verdana, sans-serif',
    backgroundColor: '#f8f9fa',
    margin: 0,
    padding: '20px',
    minHeight: '100vh'
  };

  const containerStyle = {
    maxWidth: '95%',
    margin: '0 auto',
    backgroundColor: 'white',
    borderRadius: '8px',
    boxShadow: '0 2px 10px rgba(0,0,0,0.1)',
    padding: '30px'
  };

  const headerStyle: React.CSSProperties = {
    color: '#2c3e50',
    textAlign: 'center',
    marginBottom: '30px',
    fontSize: '2.5em',
    fontWeight: 'bold',
    borderBottom: '3px solid #3498db',
    paddingBottom: '15px'
  };

  const sectionStyle = {
    backgroundColor: 'white',
    padding: '25px',
    borderRadius: '8px',
    boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
    marginBottom: '20px',
    borderWidth: '1px',
    borderStyle: 'solid',
    borderColor: '#e9ecef'
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

  const excelInputStyle: React.CSSProperties = {
    width: '100%',
    padding: '6px 8px',
    borderWidth: '0',
    borderStyle: 'none',
    borderColor: 'transparent',
    backgroundColor: 'transparent',
    fontSize: '15px',
    fontFamily: 'inherit',
    outline: 'none',
    boxSizing: 'border-box'
  };

  const tableStyle: React.CSSProperties = {
    width: '100%',
    borderCollapse: 'collapse',
    backgroundColor: 'white',
    boxShadow: '0 1px 3px rgba(0,0,0,0.1)',
    fontSize: '15px'
  };

  return (
    <div style={bodyStyle}>
      <div style={containerStyle}>
        {/* 헤더 */}
        <h1 style={headerStyle}>📋 Supply Manager</h1>

        {/* Supply Type 선택 */}
        <div style={sectionStyle}>
          <h2 style={{ color: '#0077B6', marginBottom: '15px' }}>📋 Supply Type Selection</h2>
          <div style={{ display: 'flex', gap: '20px', alignItems: 'center' }}>
            <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer' }}>
              <input
                type="radio"
                name="supplyType"
                value="order-request"
                checked={supplyType === 'order-request'}
                onChange={(e) => setSupplyType(e.target.value)}
                style={{ margin: 0 }}
              />
              <span style={{ fontSize: '16px', fontWeight: 'bold' }}>📦 Order Request</span>
            </label>
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
          </div>
        </div>


        {/* 새 아이템 추가 - Order Request가 아닐 때만 표시 */}
        {supplyType !== 'order-request' && (
        <div style={sectionStyle}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px' }}>
            <h2 style={{ color: '#0077B6', margin: 0 }}>➕ Add New Item</h2>
          </div>
          
          <div style={{ 
            display: 'grid', 
            gridTemplateColumns: supplyType === 'dental' 
              ? 'repeat(8, 1fr)' 
              : 'repeat(7, 1fr)', 
            gap: '15px', 
            marginBottom: '15px' 
          }}>
            <div>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>Insert Row Number:</label>
              <input
                type="number"
                min="1"
                max={items.length}
                value={insertAfterRow}
                onChange={(e) => setInsertAfterRow(e.target.value)}
                style={inputStyle}
              />
            </div>
            {supplyType === 'dental' && (
              <div>
                <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>Category:</label>
                <input
                  type="text"
                  value={newItem.category}
                  onChange={(e) => setNewItem({...newItem, category: e.target.value})}
                  placeholder="Category"
                  style={inputStyle}
                />
              </div>
            )}
            <div>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>Item:</label>
              <input
                type="text"
                value={newItem.item}
                onChange={(e) => setNewItem({...newItem, item: e.target.value})}
                placeholder="Item name"
                style={inputStyle}
              />
            </div>
            <div>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>Extra Info:</label>
              <input
                type="text"
                value={newItem.extraInfo}
                onChange={(e) => setNewItem({...newItem, extraInfo: e.target.value})}
                placeholder="Extra information"
                style={inputStyle}
              />
            </div>
            <div>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>Price:</label>
              <input
                type="text"
                value={newItem.price}
                onChange={(e) => setNewItem({...newItem, price: e.target.value})}
                placeholder="Price"
                style={inputStyle}
              />
            </div>
            <div>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>Seller:</label>
              <input
                type="text"
                value={newItem.seller}
                onChange={(e) => setNewItem({...newItem, seller: e.target.value})}
                placeholder="Seller"
                style={inputStyle}
              />
            </div>
            <div>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>Code:</label>
              <input
                type="text"
                value={newItem.code}
                onChange={(e) => setNewItem({...newItem, code: e.target.value})}
                placeholder="Code"
                style={inputStyle}
              />
            </div>
            <div>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>URL:</label>
              <input
                type="text"
                value={newItem.url}
                onChange={(e) => setNewItem({...newItem, url: e.target.value})}
                placeholder="URL"
                style={inputStyle}
              />
            </div>
          </div>
          
          <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
            <button 
              onClick={addItem}
              disabled={loading}
              style={{
                ...buttonStyle,
                backgroundColor: '#28a745'
              }}
            >
              {loading ? 'Adding...' : (insertAfterRow ? `➕ Insert at Position ${insertAfterRow}` : '➕ Add Item')}
            </button>
          </div>
        </div>
        )}


        {/* 에러 메시지 */}
        {error && (
          <div style={{ 
            backgroundColor: '#f8d7da', 
            color: '#721c24', 
            padding: '15px', 
            borderRadius: '4px', 
            marginBottom: '20px',
            borderWidth: '1px',
            borderStyle: 'solid',
            borderColor: '#f5c6cb'
          }}>
            {error}
            <button 
              onClick={() => setError(null)}
              style={{
                float: 'right',
                background: 'none',
                borderWidth: '0',
                borderStyle: 'none',
                borderColor: 'transparent',
                fontSize: '18px',
                cursor: 'pointer'
              }}
            >
              ×
            </button>
          </div>
        )}

        



        {/* 엑셀 스타일 데이터 테이블 */}
        <div style={sectionStyle}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px' }}>
            <h2 style={{ color: '#0077B6', margin: 0 }}>📊 {supplyTypeLabel} {supplyType === 'order-request' ? 'List' : 'Supply Items'}</h2>
            {supplyType !== 'order-request' && (
            <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
              <button
                onClick={toggleSelectAll}
                style={{
                  ...buttonStyle,
                  backgroundColor: selectedItems.size === filteredItems.length ? '#dc3545' : '#6c757d',
                  padding: '8px 16px',
                  fontSize: '14px'
                }}
              >
                {selectedItems.size === filteredItems.length ? 'Deselect All' : 'Select All'}
              </button>
              {selectedItems.size > 0 && (
                <button
                  onClick={deleteSelectedItems}
                  disabled={loading}
                  style={{
                    ...buttonStyle,
                    backgroundColor: '#dc3545',
                    padding: '8px 16px',
                    fontSize: '14px'
                  }}
                >
                  🗑️ Delete Selected ({selectedItems.size})
                </button>
              )}
            </div>
            )}
          </div>

          {/* 필터 및 검색 - Order Request가 아닐 때만 표시 */}
          {supplyType !== 'order-request' && (
          <>
            <div style={{ marginBottom: '20px' }}>
              <h3 style={{ color: '#0077B6', marginBottom: '15px', fontSize: '18px' }}>🔍 Filter & Search</h3>
              
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
                      {categoryOptions.map((category, index) => (
                        <option key={`category-${index}-${category}`} value={category}>{category}</option>
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
                    {[...new Set(items.map(item => item.seller).filter(Boolean))].sort().map((seller, index) => (
                      <option key={`seller-${index}-${seller}`} value={seller}>{seller}</option>
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
                    onClick={loadItems}
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

              <div style={{ 
                fontSize: '14px',
                color: '#666',
                marginBottom: '15px',
                textAlign: 'right'
              }}>
                📊 {filteredItems.length} of {items.length} items
                {categoryFilter && ` • ${categoryFilter}`}
                {sellerFilter && ` • ${sellerFilter}`}
                {searchInput && ` • "${searchInput}"`}
              </div>
            </div>
          </>
          )}
          
          {loading ? (
            <div style={{ textAlign: 'center', padding: '40px', color: '#666' }}>
              Loading items...
            </div>
          ) : filteredItems.length === 0 ? (
            <div style={{ textAlign: 'center', padding: '40px', color: '#666' }}>
              No items found.
            </div>
          ) : supplyType === 'order-request' ? (
            <div style={{ overflowX: 'auto' }}>
              {/* 그룹화된 주문 요약 테이블 */}
              <table style={tableStyle}>
                <thead>
                  <tr style={{ backgroundColor: '#2c3e50', color: 'white' }}>
                    <th style={{ padding: '12px', textAlign: 'center', minWidth: '60px', border: '1px solid #d0d0d0', fontWeight: 'bold' }}>Expand</th>
                    <th style={{ padding: '12px', textAlign: 'center', minWidth: '180px', border: '1px solid #d0d0d0', fontWeight: 'bold' }}>Order Date</th>
                    <th style={{ padding: '12px', textAlign: 'center', minWidth: '150px', border: '1px solid #d0d0d0', fontWeight: 'bold' }}>Office</th>
                    <th style={{ padding: '12px', textAlign: 'center', minWidth: '150px', border: '1px solid #d0d0d0', fontWeight: 'bold' }}>Supply Type</th>
                    <th style={{ padding: '12px', textAlign: 'center', minWidth: '120px', border: '1px solid #d0d0d0', fontWeight: 'bold' }}>Total Items</th>
                    <th style={{ padding: '12px', textAlign: 'center', minWidth: '120px', border: '1px solid #d0d0d0', fontWeight: 'bold' }}>Total Quantity</th>
                    <th style={{ padding: '12px', textAlign: 'center', minWidth: '100px', border: '1px solid #d0d0d0', fontWeight: 'bold' }}>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {Object.entries(groupedOrders()).map(([key, group]: [string, any]) => {
                    const isExpanded = expandedGroups.has(key);
                    return (
                      <React.Fragment key={key}>
                        {/* 그룹 요약 행 */}
                        <tr 
                          onClick={() => toggleGroup(key)}
                          style={{ 
                            backgroundColor: '#f0f8ff',
                            borderBottom: '2px solid #0077B6',
                            cursor: 'pointer',
                            transition: 'background-color 0.2s'
                          }}
                          onMouseEnter={(e) => e.currentTarget.style.backgroundColor = '#e3f2fd'}
                          onMouseLeave={(e) => e.currentTarget.style.backgroundColor = '#f0f8ff'}
                        >
                          <td style={{ padding: '12px', textAlign: 'center', fontSize: '20px' }}>
                            {isExpanded ? '▼' : '▶'}
                          </td>
                          <td style={{ padding: '12px', textAlign: 'center', fontWeight: 'bold', fontSize: '14px', color: '#495057' }}>
                            {group.orderDate ? new Date(group.orderDate).toLocaleString('en-US', {
                              year: 'numeric',
                              month: '2-digit',
                              day: '2-digit',
                              hour: '2-digit',
                              minute: '2-digit'
                            }) : '-'}
                          </td>
                          <td style={{ padding: '12px', textAlign: 'center', fontWeight: 'bold', fontSize: '16px', color: '#0077B6' }}>
                            {group.office}
                          </td>
                          <td style={{ padding: '12px', textAlign: 'center' }}>
                            <div style={{ display: 'flex', gap: '8px', justifyContent: 'center', flexWrap: 'wrap' }}>
                              {group.supplyTypesArray && group.supplyTypesArray.map((type: string, typeIndex: number) => (
                                <span 
                                  key={`${key}-supply-type-${typeIndex}-${type}`}
                                  style={{
                                    padding: '6px 12px',
                                    borderRadius: '4px',
                                    fontSize: '14px',
                                    fontWeight: 'bold',
                                    backgroundColor: type === 'dental' ? '#e7f3ff' : '#f0f8ff',
                                    color: type === 'dental' ? '#0077B6' : '#495057'
                                  }}
                                >
                                  {type === 'dental' ? '🦷 Dental' : '📋 Office'}
                                </span>
                              ))}
                            </div>
                          </td>
                          <td style={{ padding: '12px', textAlign: 'center', fontWeight: 'bold', fontSize: '16px' }}>
                            {group.itemCount} items
                          </td>
                          <td style={{ padding: '12px', textAlign: 'center', fontWeight: 'bold', fontSize: '16px', color: '#28a745' }}>
                            {group.totalQuantity}
                          </td>
                          <td style={{ padding: '12px', textAlign: 'center' }}>
                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                deleteOrderGroup(group.orderSessionId, group.itemCount);
                              }}
                              style={{
                                backgroundColor: '#dc3545',
                                color: 'white',
                                border: 'none',
                                padding: '7px 14px',
                                borderRadius: '4px',
                                fontSize: '14px',
                                fontWeight: 'bold',
                                cursor: 'pointer',
                                transition: 'background-color 0.2s'
                              }}
                              onMouseEnter={(e) => e.currentTarget.style.backgroundColor = '#c82333'}
                              onMouseLeave={(e) => e.currentTarget.style.backgroundColor = '#dc3545'}
                            >
                              🗑️ Delete
                            </button>
                          </td>
                        </tr>
                        
                        {/* 상세 주문 항목들 (확장 시) */}
                        {isExpanded && group.items.map((item: any, itemIndex: number) => (
                          <tr 
                            key={item.id}
                            style={{ 
                              backgroundColor: itemIndex % 2 === 0 ? 'white' : '#f9f9f9',
                              borderBottom: '1px solid #e0e0e0'
                            }}
                          >
                            <td style={{ padding: '8px', textAlign: 'center', color: '#999' }}>•</td>
                            <td colSpan={6} style={{ padding: '0' }}>
                              <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                                <tbody>
                                  <tr>
                                    <td style={{ padding: '8px', width: '5%', textAlign: 'center', color: '#999', fontSize: '14px' }}>
                                      #{itemIndex + 1}
                                    </td>
                                    <td style={{ padding: '8px', width: '12%', fontSize: '15px' }}>
                                      <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                                        <div><strong>Category:</strong> {item.category || '-'}</div>
                                        <span style={{
                                          padding: '3px 8px',
                                          borderRadius: '3px',
                                          fontSize: '13px',
                                          fontWeight: 'bold',
                                          backgroundColor: item.supplyType === 'dental' ? '#e7f3ff' : '#f0f8ff',
                                          color: item.supplyType === 'dental' ? '#0077B6' : '#495057',
                                          alignSelf: 'flex-start'
                                        }}>
                                          {item.supplyType === 'dental' ? '🦷 Dental' : '📋 Office'}
                                        </span>
                                      </div>
                                    </td>
                                    <td style={{ padding: '8px', width: '30%', fontSize: '15px' }}>
                                      <strong>Item:</strong> {item.item}
                                    </td>
                                    <td style={{ padding: '8px', width: '18%', fontSize: '15px' }}>
                                      <strong>Extra Info:</strong> {item.extraInfo || '-'}
                                    </td>
                                    <td style={{ padding: '8px', width: '12%', fontSize: '15px' }}>
                                      <strong>Code:</strong> {item.code || '-'}
                                    </td>
                                    <td style={{ padding: '8px', width: '12%', textAlign: 'center' }}>
                                      <span style={{ 
                                        fontWeight: 'bold', 
                                        fontSize: '15px', 
                                        color: '#28a745',
                                        backgroundColor: '#d4edda',
                                        padding: '4px 10px',
                                        borderRadius: '4px'
                                      }}>
                                        Qty: {item.quantity}
                                      </span>
                                    </td>
                                    <td style={{ padding: '8px', width: '11%', textAlign: 'center' }}>
                                      <select
                                        value={item.status || 'pending'}
                                        onChange={(e) => updateOrderStatus(item.id, e.target.value)}
                                        onClick={(e) => e.stopPropagation()}
                                        style={{
                                          padding: '7px 12px',
                                          borderRadius: '4px',
                                          fontSize: '14px',
                                          fontWeight: 'bold',
                                          cursor: 'pointer',
                                          border: '2px solid',
                                          backgroundColor: item.status === 'pending' ? '#fff3cd' : 
                                                         item.status === 'processing' ? '#cfe2ff' :
                                                         item.status === 'completed' ? '#d4edda' : 
                                                         item.status === 'cancelled' ? '#f8d7da' : '#e9ecef',
                                          color: item.status === 'pending' ? '#856404' : 
                                                item.status === 'processing' ? '#084298' :
                                                item.status === 'completed' ? '#155724' : 
                                                item.status === 'cancelled' ? '#721c24' : '#495057',
                                          borderColor: item.status === 'pending' ? '#ffc107' : 
                                                      item.status === 'processing' ? '#0d6efd' :
                                                      item.status === 'completed' ? '#28a745' : 
                                                      item.status === 'cancelled' ? '#dc3545' : '#6c757d'
                                        }}
                                      >
                                        <option value="pending">⏳ Requested</option>
                                        <option value="processing">🔄 Processing</option>
                                        <option value="completed">✅ Completed</option>
                                        <option value="cancelled">❌ Cancelled</option>
                                      </select>
                                    </td>
                                  </tr>
                                </tbody>
                              </table>
                            </td>
                          </tr>
                        ))}
                      </React.Fragment>
                    );
                  })}
                </tbody>
              </table>
            </div>
          ) : (
            <div style={{ overflowX: 'auto' }}>
              <table style={tableStyle}>
                <thead>
                  <tr style={{ backgroundColor: '#2c3e50', color: 'white' }}>
                    <th style={{ padding: '8px', textAlign: 'center', minWidth: '50px', border: '1px solid #d0d0d0', fontWeight: 'bold' }}>Select</th>
                    <th style={{ padding: '8px', textAlign: 'center', minWidth: '50px', border: '1px solid #d0d0d0', fontWeight: 'bold' }}>No.</th>
                    {supplyType === 'dental' && (
                      <th style={{ padding: '8px', textAlign: 'center', minWidth: '250px', border: '1px solid #d0d0d0', fontWeight: 'bold' }}>Category</th>
                    )}
                    <th style={{ padding: '8px', textAlign: 'center', minWidth: '400px', border: '1px solid #d0d0d0', fontWeight: 'bold' }}>Item</th>
                    <th style={{ padding: '8px', textAlign: 'center', minWidth: '250px', border: '1px solid #d0d0d0', fontWeight: 'bold' }}>Extra Info</th>
                    <th style={{ padding: '8px', textAlign: 'center', minWidth: '100px', border: '1px solid #d0d0d0', fontWeight: 'bold' }}>Price ($)</th>
                    <th style={{ padding: '8px', textAlign: 'center', minWidth: '200px', border: '1px solid #d0d0d0', fontWeight: 'bold' }}>Seller</th>
                    <th style={{ padding: '8px', textAlign: 'center', minWidth: '120px', border: '1px solid #d0d0d0', fontWeight: 'bold' }}>Code</th>
                    <th style={{ padding: '8px', textAlign: 'center', minWidth: '300px', border: '1px solid #d0d0d0', fontWeight: 'bold' }}>URL</th>
                    <th style={{ padding: '8px', textAlign: 'center', minWidth: '100px', border: '1px solid #d0d0d0', fontWeight: 'bold' }}>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredItems.map((item) => (
                    <ItemRow 
                      key={`${item.id}-${item.displayId}`}
                      item={item} 
                      updateItem={updateItem}
                      deleteItem={deleteItem}
                      isSelected={selectedItems.has(item.id)}
                      onSelect={toggleItemSelection}
                      onInsertBelow={handleInsertBelow}
                      excelInputStyle={excelInputStyle}
                      supplyType={supplyType}
                      editingValues={editingValues}
                    />
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

export default function SupplyManagerSystem() {
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
      <SupplyManagerSystemContent />
    </Suspense>
    </>
  );
}
