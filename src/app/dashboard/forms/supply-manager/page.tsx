'use client'

import React, { useState, useEffect, useCallback, useRef, Suspense } from "react";
import { useSearchParams, useRouter } from "next/navigation";
import { doc,setDoc, collection, getDocs, getDoc, updateDoc, deleteDoc, writeBatch, serverTimestamp, } from "firebase/firestore";
import { getStorage, ref, listAll, getMetadata, getDownloadURL, deleteObject, type StorageReference, } from "firebase/storage";
import { db, auth } from "@/lib/firebase.config";
import { onAuthStateChanged } from 'firebase/auth';
const DANGEROUS_KEYS = new Set(['__proto__', 'constructor', 'prototype']);
const MAX_STRING_LENGTH = 10000;

function sanitizeValue(value: unknown): unknown {
  if (value === null || value === undefined) return null;
  if (typeof value === 'string') {
    return value.trim().slice(0, MAX_STRING_LENGTH);
  }
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) return 0;
    return value;
  }
  if (typeof value === 'boolean') return value;
  if (Array.isArray(value)) {
    return value.map(sanitizeValue);
  }
  if (typeof value === 'object') {
    return sanitizeData(value as Record<string, unknown>);
  }
  return null;
}

function sanitizeData<T extends Record<string, unknown>>(data: T): T {
  const result: Record<string, unknown> = {};
  for (const key of Object.keys(data)) {
    if (DANGEROUS_KEYS.has(key)) continue;
    const sanitizedKey = String(key).trim().slice(0, 500);
    if (!sanitizedKey) continue;
    result[sanitizedKey] = sanitizeValue(data[key]);
  }
  return result as T;
}

function normalizeSupplyStoragePath(path: string): string | null {
  if (typeof path !== "string" || !path.trim()) return null;
  const p = path.trim().replace(/^\/*/, "");
  if (!p || p.includes("..") || p.includes("//") || p.length > 1024) return null;
  return p;
}

const SUPPLY_ORDER_PDF_STATUS_COLLECTION = "supply-order-pdf-status";

type ManagerSupplyPdfStatus =
  | "requested"
  | "processing"
  | "completed"
  | "received"
  | "delete";

function buildSupplyOrderPdfStatusDocId(
  office: string,
  dateFolder: string,
  timeSuffix: string
): string | null {
  const o = String(office).trim();
  const d = String(dateFolder).trim();
  const t = String(timeSuffix).trim();
  if (!o || !/^\d{4}-\d{2}-\d{2}$/.test(d) || !t) return null;
  if (!/^[a-zA-Z0-9_-]+$/.test(o)) return null;
  if (!/^[0-9A-Za-z-]+$/.test(t)) return null;
  const id = `${o}_${d}_${t}`;
  return id.length > 800 ? id.slice(0, 800) : id;
}

function supplyOrderPdfStatusDocIdFromStoragePath(storagePath: string): string | null {
  const n = normalizeSupplyStoragePath(storagePath);
  if (!n) return null;
  const segments = n.split("/").filter(Boolean);
  if (segments.length < 4 || segments[0] !== "orders") return null;
  const office = segments[1];
  const dateFolder = segments[2];
  const filename = segments[segments.length - 1];
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateFolder)) return null;
  const base = filename.replace(/\.pdf$/i, "");
  const expectedPrefix = `Supply_Order_${office}_${dateFolder}_`;
  if (!base.startsWith(expectedPrefix)) return null;
  const timeSuffix = base.slice(expectedPrefix.length);
  if (!timeSuffix) return null;
  return buildSupplyOrderPdfStatusDocId(office, dateFolder, timeSuffix);
}

function getDisplayStatusForPdf(pdf: {
  officeStatusFirestore: "requested" | "received";
  managerStatusFirestore: string | null;
}): ManagerSupplyPdfStatus {
  if (pdf.officeStatusFirestore === "received") return "received";
  const m = String(pdf.managerStatusFirestore || "")
    .trim()
    .toLowerCase();
  if (m === "processing" || m === "completed" || m === "delete") return m;
  if (m === "received") return "received";
  return "requested";
}

type StorageOrderPdf = {
  id: string;
  path: string;
  filename: string;
  office: string;
  dateFolder: string;
  createdAt: Date;
  canSyncStatus: boolean;
  officeStatusFirestore: "requested" | "received";
  managerStatusFirestore: string | null;
};

async function collectSupplyOrderPdfRefs(
  currentRef: StorageReference
): Promise<StorageReference[]> {
  const current = await listAll(currentRef);
  const direct = current.items.filter(
    (it) =>
      it.name.toLowerCase().endsWith(".pdf") &&
      it.name.startsWith("Supply_Order_")
  );
  const nested = await Promise.all(
    current.prefixes.map((prefixRef) => collectSupplyOrderPdfRefs(prefixRef))
  );
  return direct.concat(...nested);
}

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
  const [editingValues, setEditingValues] = useState<{ [key: string]: any }>({});
  const [isAuthorized, setIsAuthorized] = useState<boolean | null>(null);
  const [orderRequestPdfs, setOrderRequestPdfs] = useState<StorageOrderPdf[]>([]);
  const [filteredOrderPdfs, setFilteredOrderPdfs] = useState<StorageOrderPdf[]>([]);
  const [viewingStoragePdf, setViewingStoragePdf] = useState<StorageOrderPdf | null>(null);
  const [storagePdfViewerUrl, setStoragePdfViewerUrl] = useState<string | null>(null);
  const [storagePdfViewerLoading, setStoragePdfViewerLoading] = useState(false);
  const [storagePdfViewerError, setStoragePdfViewerError] = useState(false);
  const storagePdfBlobUrlRef = useRef<string | null>(null);
  const [managerStatusUpdatingPath, setManagerStatusUpdatingPath] = useState<string | null>(null);

  const collectionName = supplyType === 'dental' ? 'dental-supplies' : 'office-supplies';
  const supplyTypeLabel = supplyType === 'order-request' ? 'Order Request' : (supplyType === 'dental' ? 'Dental' : 'Office');
  const supplyTypeEmoji = supplyType === 'order-request' ? '📦' : (supplyType === 'dental' ? '🦷' : '📋');

  // 카테고리 옵션 동적 생성
  const categoryOptions = [...new Set(items.map(item => item.category).filter(Boolean))].sort();

  const loadOrderRequestPdfsFromStorage = useCallback(async () => {
    try {
      setLoading(true);
      setError(null);
      const storage = getStorage();
      const rootRef = ref(storage, "orders/");
      const top = await listAll(rootRef);
      const allRefs: StorageReference[] = [];
      for (const prefixRef of top.prefixes) {
        allRefs.push(...(await collectSupplyOrderPdfRefs(prefixRef)));
      }
      const baseList = await Promise.all(
        allRefs.map(async (item) => {
          const meta = await getMetadata(item);
          const createdAt = meta.timeCreated ? new Date(meta.timeCreated) : new Date();
          const parts = item.fullPath.split("/").filter(Boolean);
          const office = parts.length >= 2 ? parts[1] : "";
          const dateFolder =
            parts.length >= 3 && /^\d{4}-\d{2}-\d{2}$/.test(parts[2]) ? parts[2] : "";
          return {
            id: item.fullPath,
            path: item.fullPath,
            filename: item.name,
            office,
            dateFolder,
            createdAt,
          };
        })
      );
      baseList.sort((a, b) => b.createdAt.getTime() - a.createdAt.getTime());

      const list: StorageOrderPdf[] = await Promise.all(
        baseList.map(async (entry) => {
          const docId = supplyOrderPdfStatusDocIdFromStoragePath(entry.path);
          const canSyncStatus = Boolean(docId);
          let officeStatusFirestore: "requested" | "received" = "requested";
          let managerStatusFirestore: string | null = null;
          if (docId) {
            const snap = await getDoc(
              doc(db, SUPPLY_ORDER_PDF_STATUS_COLLECTION, docId)
            );
            if (snap.exists()) {
              const data = sanitizeData(snap.data() || {}) as Record<string, unknown>;
              officeStatusFirestore =
                data.office_status === "received" ? "received" : "requested";
              const ms = data.manager_status;
              managerStatusFirestore =
                typeof ms === "string" && ms.trim() ? ms.trim() : null;
            }
          }
          return {
            ...entry,
            canSyncStatus,
            officeStatusFirestore,
            managerStatusFirestore,
          };
        })
      );

      setOrderRequestPdfs(list);
      setItems([]);
      setFilteredItems([]);
    } catch (e) {
      console.error(e);
      setError("Failed to load supply order PDFs from storage.");
      setOrderRequestPdfs([]);
    } finally {
      setLoading(false);
    }
  }, []);

  const handleManagerSupplyPdfStatusChange = useCallback(
    async (pdf: StorageOrderPdf, value: ManagerSupplyPdfStatus) => {
      const docId = supplyOrderPdfStatusDocIdFromStoragePath(pdf.path);
      if (!docId || !pdf.canSyncStatus) {
        alert("Cannot update status for this file.");
        return;
      }

      if (value === "delete") {
        if (!confirm("Are you sure removing this order?")) return;
        setManagerStatusUpdatingPath(pdf.path);
        try {
          const p = normalizeSupplyStoragePath(pdf.path);
          if (!p) throw new Error("invalid path");
          await deleteObject(ref(getStorage(), p));
          try {
            await deleteDoc(doc(db, SUPPLY_ORDER_PDF_STATUS_COLLECTION, docId));
          } catch {
            /* 문서 없음 */
          }
          setViewingStoragePdf((v) => (v?.path === pdf.path ? null : v));
          await loadOrderRequestPdfsFromStorage();
        } catch (err) {
          console.error(err);
          alert("Failed to delete the PDF.");
        } finally {
          setManagerStatusUpdatingPath(null);
        }
        return;
      }

      setManagerStatusUpdatingPath(pdf.path);
      try {
        const payload: Record<string, unknown> = {
          updateAt: serverTimestamp(),
        };
        if (value === "received" || value === "requested") {
          payload.office_status = value;
          payload.manager_status = value;
        } else {
          payload.manager_status = value;
        }

        await setDoc(
          doc(db, SUPPLY_ORDER_PDF_STATUS_COLLECTION, docId),
          payload as { [key: string]: unknown },
          { merge: true }
        );

        setOrderRequestPdfs((prev) =>
          prev.map((o) => {
            if (o.path !== pdf.path) return o;
            if (value === "received" || value === "requested") {
              return {
                ...o,
                officeStatusFirestore: value,
                managerStatusFirestore: value,
              };
            }
            return {
              ...o,
              managerStatusFirestore: value,
            };
          })
        );
      } catch (err) {
        console.error(err);
        alert("Failed to update status.");
      } finally {
        setManagerStatusUpdatingPath(null);
      }
    },
    [loadOrderRequestPdfsFromStorage]
  );

  // 데이터 로드 (dental / office: Firestore, order-request: Storage PDF만)
  const loadItems = useCallback(async () => {
    if (supplyType === "order-request") {
      await loadOrderRequestPdfsFromStorage();
      return;
    }
    try {
      setLoading(true);
      setError(null);

      const querySnapshot = await getDocs(collection(db, collectionName));
      const itemsList: any[] = [];

      querySnapshot.forEach((docSnap: any) => {
        const rawData = docSnap.data();
        const sanitizedData = sanitizeData(rawData);

        if (sanitizedData.items && sanitizedData.items.length > 0) {
          sanitizedData.items.forEach((item: any, index: number) => {
            const quantity = sanitizedData.quantities?.[item.id] || 0;
            const perItemStatus = sanitizedData.itemStatuses?.[item.id] || sanitizedData.status;
            itemsList.push({
              id: `${docSnap.id}-${index}`,
              originalItemId: item.id,
              ...item,
              quantity: quantity,
              orderSessionId: sanitizedData.orderSessionId,
              office: sanitizedData.office,
              orderDate: sanitizedData.orderDate,
              status: perItemStatus,
              parentDocId: docSnap.id,
            });
          });
        } else {
          itemsList.push({
            id: docSnap.id,
            ...sanitizedData,
          });
        }
      });

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

      const itemsWithDisplayId = itemsList.map((item, index) => ({
        ...item,
        displayId: index + 1,
      }));

      setItems(itemsWithDisplayId);
      setFilteredItems(itemsWithDisplayId);
    } catch (error) {
      setError("Failed to load items");
    } finally {
      setLoading(false);
    }
  }, [collectionName, supplyType, loadOrderRequestPdfsFromStorage]);

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
        const sanitizedUpdateData = sanitizeData(updateData);
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
      const sanitizedItemData = sanitizeData(itemData);
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
      setError('Failed to delete items');
    } finally {
      setLoading(false);
    }
  }, [selectedItems, collectionName, loadItems]);

  // 개별 아이템 삭제
  const deleteItem = useCallback(async (itemId: string) => {
    if (typeof itemId !== 'string') {
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

  useEffect(() => {
    if (supplyType !== "order-request") return;
    let f = [...orderRequestPdfs];
    if (searchInput.trim()) {
      const q = searchInput.trim().toLowerCase();
      f = f.filter(
        (p) =>
          p.filename.toLowerCase().includes(q) ||
          p.office.toLowerCase().includes(q) ||
          (p.dateFolder && p.dateFolder.toLowerCase().includes(q))
      );
    }
    setFilteredOrderPdfs(f);
  }, [supplyType, orderRequestPdfs, searchInput]);

  useEffect(() => {
    const clearViewerUrlRef = () => {
      if (storagePdfBlobUrlRef.current?.startsWith("blob:")) {
        URL.revokeObjectURL(storagePdfBlobUrlRef.current);
      }
      storagePdfBlobUrlRef.current = null;
    };

    if (!viewingStoragePdf?.path) {
      clearViewerUrlRef();
      setStoragePdfViewerUrl(null);
      setStoragePdfViewerLoading(false);
      setStoragePdfViewerError(false);
      return;
    }
    const p = normalizeSupplyStoragePath(viewingStoragePdf.path);
    if (!p) {
      clearViewerUrlRef();
      setStoragePdfViewerError(true);
      setStoragePdfViewerLoading(false);
      return;
    }
    let cancelled = false;
    setStoragePdfViewerLoading(true);
    setStoragePdfViewerError(false);
    setStoragePdfViewerUrl(null);
    clearViewerUrlRef();

    const setViewerSource = (sourceUrl: string, revokePreviousBlob = true) => {
      if (cancelled || !sourceUrl) return;
      if (storagePdfBlobUrlRef.current) {
        if (
          revokePreviousBlob &&
          storagePdfBlobUrlRef.current.startsWith("blob:")
        ) {
          URL.revokeObjectURL(storagePdfBlobUrlRef.current);
        }
        storagePdfBlobUrlRef.current = null;
      }
      storagePdfBlobUrlRef.current = sourceUrl;
      setStoragePdfViewerUrl(sourceUrl);
      setStoragePdfViewerLoading(false);
    };

    const onError = (error?: unknown) => {
      if (!cancelled) {
        if (error) {
          console.error("Failed to open PDF:", error);
        }
        setStoragePdfViewerError(true);
        setStoragePdfViewerLoading(false);
      }
    };

    const storageRef = ref(getStorage(), p);
    getDownloadURL(storageRef)
      .then((downloadUrl) => setViewerSource(downloadUrl, false))
      .catch((urlError) => {
        console.error("Failed to get download URL:", urlError);
        onError(urlError);
      });

    return () => {
      cancelled = true;
      clearViewerUrlRef();
      setStoragePdfViewerUrl(null);
    };
  }, [viewingStoragePdf]);

  useEffect(() => {
    if (!viewingStoragePdf) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        e.preventDefault();
        setViewingStoragePdf(null);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [viewingStoragePdf]);

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

  // 사용자 인증 및 role 확인
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      try {
        if (!currentUser) {
          alert('Please log in.');
          setIsAuthorized(false);
          return;
        }

        const userDoc = await getDoc(doc(db, 'users', currentUser.uid));
        if (!userDoc.exists()) {
          alert('User information could not be found.');
          setIsAuthorized(false);
          return;
        }

        if (userDoc.id !== 'eV4AQK6V6yZQgFPNOfVeRrtz5Eu1' && userDoc.id !== 'rVMFu186CNf6ebSdSsnOg7EUrh63') {
          setIsAuthorized(false);
          if (typeof window !== 'undefined') {
            window.location.href = '/';
          }
          return;
        }

        setIsAuthorized(true);
      } catch (error: any) {
        alert('An error occurred while verifying authentication.');
        setIsAuthorized(false);
      }
    });

    return () => {
      unsubscribe();
    };
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

  // 인증 확인 중
  if (isAuthorized === null) {
    return (
      <div style={{
        display: 'flex',
        justifyContent: 'center',
        alignItems: 'center',
        height: '100vh',
        background: '#f8f9fa',
        fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
      }}>
        <div style={{ textAlign: 'center' }}>
          <div style={{ fontSize: '18px', color: '#2c3e50' }}>Verifying authentication...</div>
        </div>
      </div>
    );
  }

  // 인증 실패
  if (isAuthorized === false) {
    return (
      <div style={{
        display: 'flex',
        justifyContent: 'center',
        alignItems: 'center',
        height: '100vh',
        background: '#f8f9fa',
        fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
      }}>
        <div style={{ textAlign: 'center' }}>
          <div style={{ fontSize: '18px', color: '#d32f2f', marginBottom: '10px' }}>You do not have access to this page.</div>
        </div>
      </div>
    );
  }

  return (
    <div style={bodyStyle}>
      <div style={containerStyle}>
        {/* 헤더 */}
        <h1 style={headerStyle}>Supply Manager</h1>

        {/* Supply Type 선택 */}
        <div style={sectionStyle}>
          <h2 style={{ color: '#0077B6', marginBottom: '15px' }}>Supply Type Selection</h2>
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
              <span style={{ fontSize: '16px', fontWeight: 'bold' }}>Dental Supply</span>
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
              <span style={{ fontSize: '16px', fontWeight: 'bold' }}>Office Supply</span>
            </label>
          </div>
        </div>


        {/* 새 아이템 추가 - Order Request가 아닐 때만 표시 */}
        {supplyType !== 'order-request' && (
        <div style={sectionStyle}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px' }}>
            <h2 style={{ color: '#0077B6', margin: 0 }}>Add New Item</h2>
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
            <h2 style={{ color: '#0077B6', margin: 0 }}>{supplyTypeLabel} {supplyType === 'order-request' ? 'List' : 'Supply Items'}</h2>
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
                {filteredItems.length} of {items.length} items
                {categoryFilter && ` • ${categoryFilter}`}
                {sellerFilter && ` • ${sellerFilter}`}
                {searchInput && ` • "${searchInput}"`}
              </div>
            </div>
          </>
          )}

          {supplyType === "order-request" && (
            <div style={{ marginBottom: "20px" }}>
              <div
                style={{
                  display: "flex",
                  gap: "20px",
                  flexWrap: "wrap",
                  marginBottom: "12px",
                  alignItems: "end",
                }}
              >
                <div style={{ flex: "1", minWidth: "220px" }}>
                  <label
                    style={{ display: "block", marginBottom: "5px", fontWeight: "bold" }}
                  >
                    Search:
                  </label>
                  <input
                    type="text"
                    value={searchInput}
                    onChange={(e) => setSearchInput(e.target.value)}
                    placeholder="Office, file name, or date folder..."
                    style={inputStyle}
                  />
                </div>
                <div style={{ minWidth: "160px" }}>
                  <button
                    type="button"
                    onClick={() => loadOrderRequestPdfsFromStorage()}
                    disabled={loading}
                    style={{
                      ...buttonStyle,
                      backgroundColor: "#28a745",
                      width: "100%",
                    }}
                  >
                    {loading ? "Loading..." : "🔄 Refresh PDFs"}
                  </button>
                </div>
              </div>
              <div
                style={{
                  fontSize: "14px",
                  color: "#666",
                  textAlign: "right",
                }}
              >
                {filteredOrderPdfs.length} PDF
                {filteredOrderPdfs.length !== 1 ? "s" : ""} (all offices)
              </div>
            </div>
          )}
          
          {loading ? (
            <div style={{ textAlign: 'center', padding: '40px', color: '#666' }}>
              {supplyType === "order-request"
                ? "Loading PDFs from storage..."
                : "Loading items..."}
            </div>
          ) : supplyType === "order-request" && filteredOrderPdfs.length === 0 ? (
            <div style={{ textAlign: 'center', padding: '40px', color: '#666' }}>
              No supply orders found.
            </div>
          ) : supplyType !== "order-request" && filteredItems.length === 0 ? (
            <div style={{ textAlign: 'center', padding: '40px', color: '#666' }}>
              No items found.
            </div>
          ) : supplyType === 'order-request' ? (
            <div style={{ overflowX: 'auto' }}>
              <table style={tableStyle}>
                <thead>
                  <tr style={{ backgroundColor: '#2c3e50', color: 'white' }}>
                    <th style={{ padding: '12px', textAlign: 'center', minWidth: '200px', border: '1px solid #d0d0d0', fontWeight: 'bold' }}>Submitted</th>
                    <th style={{ padding: '12px', textAlign: 'center', minWidth: '120px', border: '1px solid #d0d0d0', fontWeight: 'bold' }}>Office</th>
                    <th style={{ padding: '12px', textAlign: 'center', minWidth: '140px', border: '1px solid #d0d0d0', fontWeight: 'bold' }}>View PDF</th>
                    <th style={{ padding: '12px', textAlign: 'center', minWidth: '160px', border: '1px solid #d0d0d0', fontWeight: 'bold' }}>Status</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredOrderPdfs.map((pdf, rowIndex) => {
                    const rowStatus = getDisplayStatusForPdf(pdf);
                    return (
                    <tr
                      key={pdf.id}
                      style={{
                        backgroundColor: rowIndex % 2 === 0 ? '#ffffff' : '#f9f9f9',
                        borderBottom: '1px solid #e0e0e0',
                      }}
                    >
                      <td style={{ padding: '12px', textAlign: 'center', fontSize: '14px', color: '#495057' }}>
                        {pdf.createdAt
                          ? pdf.createdAt.toLocaleString('en-US', {
                              timeZone: 'America/Los_Angeles',
                              dateStyle: 'medium',
                              timeStyle: 'short',
                            })
                          : '—'}
                      </td>
                      <td style={{ padding: '12px', textAlign: 'center', fontWeight: 'bold', fontSize: '16px', color: '#0077B6' }}>
                        {pdf.office || '—'}
                      </td>
                      <td style={{ padding: '12px', textAlign: 'center' }}>
                        <button
                          type="button"
                          onClick={() => {
                            if (!normalizeSupplyStoragePath(pdf.path)) {
                              alert('Cannot open the file.');
                              return;
                            }
                            setViewingStoragePdf(pdf);
                          }}
                          style={{
                            backgroundColor: '#0077B6',
                            color: 'white',
                            border: 'none',
                            padding: '8px 18px',
                            borderRadius: '6px',
                            fontSize: '14px',
                            fontWeight: 'bold',
                            cursor: 'pointer',
                            whiteSpace: 'nowrap',
                          }}
                        >
                          📄 View PDF
                        </button>
                      </td>
                      <td style={{ padding: '12px', textAlign: 'center' }}>
                        {!pdf.canSyncStatus ? (
                          <span style={{ color: '#94a3b8', fontSize: '13px' }}>—</span>
                        ) : (
                          <select
                            value={rowStatus}
                            disabled={managerStatusUpdatingPath === pdf.path}
                            onChange={(e) => {
                              void handleManagerSupplyPdfStatusChange(
                                pdf,
                                e.target.value as ManagerSupplyPdfStatus
                              );
                            }}
                            style={{
                              padding: '7px 12px',
                              borderRadius: '4px',
                              fontSize: '14px',
                              fontWeight: 'bold',
                              cursor:
                                managerStatusUpdatingPath === pdf.path
                                  ? 'wait'
                                  : 'pointer',
                              border: '2px solid',
                              maxWidth: '100%',
                              backgroundColor:
                                rowStatus === 'requested'
                                  ? '#fff3cd'
                                  : rowStatus === 'processing'
                                    ? '#cfe2ff'
                                    : rowStatus === 'completed'
                                      ? '#d4edda'
                                      : rowStatus === 'received'
                                        ? '#e7f3ff'
                                        : rowStatus === 'delete'
                                          ? '#f8d7da'
                                          : '#e9ecef',
                              color:
                                rowStatus === 'requested'
                                  ? '#856404'
                                  : rowStatus === 'processing'
                                    ? '#084298'
                                    : rowStatus === 'completed'
                                      ? '#155724'
                                      : rowStatus === 'received'
                                        ? '#0d6efd'
                                        : rowStatus === 'delete'
                                          ? '#721c24'
                                          : '#495057',
                              borderColor:
                                rowStatus === 'requested'
                                  ? '#ffc107'
                                  : rowStatus === 'processing'
                                    ? '#0d6efd'
                                    : rowStatus === 'completed'
                                      ? '#28a745'
                                      : rowStatus === 'received'
                                        ? '#0d6efd'
                                        : rowStatus === 'delete'
                                          ? '#dc3545'
                                          : '#6c757d',
                            }}
                          >
                            <option value="requested">Requested</option>
                            <option value="processing">Processing</option>
                            <option value="completed">Completed</option>
                            <option value="received">Received</option>
                            <option value="delete">Delete</option>
                          </select>
                        )}
                      </td>
                    </tr>
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

      {viewingStoragePdf && (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: "rgba(0, 0, 0, 0.9)",
            zIndex: 1000,
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            justifyContent: "center",
          }}
          onClick={(e) => {
            if (e.target === e.currentTarget) setViewingStoragePdf(null);
          }}
        >
          <button
            type="button"
            aria-label="Close"
            onClick={(e) => {
              e.stopPropagation();
              setViewingStoragePdf(null);
            }}
            style={{
              position: "absolute",
              top: "20px",
              right: "20px",
              background: "#ff6b6b",
              color: "white",
              border: "none",
              borderRadius: "50%",
              width: "40px",
              height: "40px",
              fontSize: "20px",
              cursor: "pointer",
              zIndex: 1001,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontWeight: "bold",
            }}
          >
            ×
          </button>
          <button
            type="button"
            onClick={(e) => {
              e.stopPropagation();
              setViewingStoragePdf(null);
            }}
            style={{
              position: "absolute",
              top: "20px",
              right: "72px",
              background: "rgba(255,255,255,0.95)",
              color: "#023047",
              border: "1px solid #ccc",
              borderRadius: "8px",
              padding: "8px 16px",
              fontSize: "14px",
              fontWeight: "600",
              cursor: "pointer",
              zIndex: 1001,
            }}
          >
            Close
          </button>
          {storagePdfViewerLoading ? (
            <div
              style={{
                width: "90%",
                height: "90%",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                color: "white",
                fontSize: "18px",
              }}
            >
              Loading PDF…
            </div>
          ) : storagePdfViewerError ? (
            <div
              style={{
                width: "90%",
                height: "90%",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                color: "#ffcccc",
                fontSize: "16px",
                padding: "24px",
                textAlign: "center",
              }}
            >
              Could not load PDF.
            </div>
          ) : storagePdfViewerUrl ? (
            <object
              data={storagePdfViewerUrl}
              type="application/pdf"
              style={{
                width: "90%",
                height: "90%",
                border: "none",
                borderRadius: "8px",
                background: "white",
              }}
              title={viewingStoragePdf.filename}
            >
              <p style={{ padding: "20px", color: "#333" }}>
                Your browser does not support viewing PDFs.
              </p>
            </object>
          ) : null}
        </div>
      )}
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

