'use client'

import React, { useState, useEffect, useCallback, useRef, Suspense } from "react";
import { useSearchParams } from "next/navigation";
import {
  collection,
  getDocs,
  doc,
  getDoc,
  setDoc,
  serverTimestamp,
} from "firebase/firestore";
import {
  getStorage,
  ref,
  uploadBytes,
  listAll,
  getMetadata,
  getDownloadURL,
  type StorageReference,
} from "firebase/storage";
import { db, auth } from "@/lib/firebase.config";
import { onAuthStateChanged } from 'firebase/auth';
import { pdf, Document, Page, View, Text, StyleSheet } from '@react-pdf/renderer';


// Firebase 데이터 sanitization (prototype pollution 방지, 값 검증)
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

// URL 안전성 검증 (XSS 방지)
function getSafeUrl(url: string | undefined | null): string | null {
  if (!url || typeof url !== 'string') return null;
  const trimmed = url.trim();
  if (!trimmed) return null;
  try {
    const parsed = new URL(trimmed);
    if (parsed.protocol === 'https:') {
      return trimmed;
    }
    return null;
  } catch {
    return null;
  }
}

function pdfSafeStr(v: unknown, max: number): string {
  if (v == null) return '';
  return String(v).trim().slice(0, max).replace(/[<>]/g, '');
}

/** Storage 폴더(날짜)와 PDF 파일명 접미사용 — LA 기준, 시간은 HH-mm-ss */
function getCaliforniaDateAndTimeForSupplyPdf(): { dateFolder: string; timeSuffix: string } {
  const now = new Date();
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'America/Los_Angeles',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false,
  }).formatToParts(now);
  const y = parts.find((p) => p.type === 'year')?.value || '';
  const m = parts.find((p) => p.type === 'month')?.value || '';
  const d = parts.find((p) => p.type === 'day')?.value || '';
  const h = parts.find((p) => p.type === 'hour')?.value || '';
  const min = parts.find((p) => p.type === 'minute')?.value || '';
  const s = parts.find((p) => p.type === 'second')?.value || '';
  return {
    dateFolder: `${y}-${m}-${d}`,
    timeSuffix: `${h}${min}${s}`,
  };
}

function sanitizeSupplyPdfFilename(name: string): string {
  return name.replace(/[^a-zA-Z0-9._-]/g, '_').replace(/\.\./g, '_').slice(0, 200);
}

/** 제출 시각(orderDate ISO) → LA 기준 YYYY-MM-DD (PDF 업로드 폴더와 동일) */
type SupplyPdfRow = {
  idx: number;
  supplyType: string;
  category: string;
  item: string;
  extraInfo: string;
  seller: string;
  code: string;
  quantity: number;
};

function normalizeSupplyStoragePath(path: string): string | null {
  if (typeof path !== 'string' || !path.trim()) return null;
  const p = path.trim().replace(/^\/*/, '');
  if (!p || p.includes('..') || p.includes('//') || p.length > 1024) return null;
  return p;
}

/** 문서 ID `{office}_{YYYY-MM-DD}_{시간}`. 필드: `office_status`, `manager_status`(Supply Manager), `updateAt` */
const SUPPLY_ORDER_PDF_STATUS_COLLECTION = "supply-order-pdf-status";

type SupplyOrderPdfStatus = "requested" | "received";

/** 제출 시점의 오피스·날짜(LA)·시간 접미사로 상태 문서 ID 생성 */
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

/** Storage 경로·파일명에서 오피스·날짜·시간을 읽어 동일한 문서 ID 계산 (목록·상태 변경 시) */
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

function parseSupplyOrderPdfStatusFromFirestore(
  data: Record<string, unknown>
): SupplyOrderPdfStatus {
  const raw =
    data.office_status !== undefined ? data.office_status : data.status;
  return raw === "received" ? "received" : "requested";
}

type SupplyOrderPdfEntry = {
  id: string;
  path: string;
  filename: string;
  createdAt: Date;
  dateFolder: string;
  status: SupplyOrderPdfStatus;
};

const supplyPdfStyles = StyleSheet.create({
  page: { padding: 24, fontFamily: 'Helvetica', fontSize: 8 },
  headerBox: { marginBottom: 10 },
  title: { fontSize: 14, fontWeight: 'bold', marginBottom: 6, color: '#000000' },
  meta: { fontSize: 9, marginBottom: 2, color: '#444444' },
  table: { marginTop: 6 },
  rowHeader: {
    flexDirection: 'row',
    backgroundColor: '#e8e8e8',
    borderBottomWidth: 1,
    borderColor: '#000000',
  },
  row: { flexDirection: 'row', borderBottomWidth: 0.5, borderColor: '#cccccc' },
  th: { padding: 5, fontSize: 7, fontWeight: 'bold', color: '#000000' },
  td: { padding: 4, fontSize: 7, color: '#222222' },
  colIdx: { width: '6%' },
  colType: { width: '10%' },
  colCat: { width: '12%' },
  colItem: { width: '24%' },
  colExtra: { width: '18%' },
  colSeller: { width: '10%' },
  colCode: { width: '10%' },
  colQty: { width: '10%' },
  footer: { marginTop: 14, fontSize: 8, color: '#666666' },
});

function SupplyOrderPdfDocument({
  office,
  orderDateFormatted,
  generatedAt,
  rows,
}: {
  office: string;
  orderDateFormatted: string;
  generatedAt: string;
  rows: SupplyPdfRow[];
}) {
  const ROWS_PER_PAGE = 22;
  const pages: SupplyPdfRow[][] = [];
  for (let i = 0; i < rows.length; i += ROWS_PER_PAGE) {
    pages.push(rows.slice(i, i + ROWS_PER_PAGE));
  }
  const s = supplyPdfStyles;

  return (
    <Document>
      {pages.map((chunk, pageIndex) => (
        <Page key={pageIndex} size="A4" style={s.page}>
          <View style={s.headerBox}>
            {pageIndex === 0 ? (
              <>
                <Text style={s.title}>Supply Order</Text>
                <Text style={s.meta}>Office: {pdfSafeStr(office, 80)}</Text>
                <Text style={s.meta}>Order date: {pdfSafeStr(orderDateFormatted, 120)}</Text>
              </>
            ) : (
              <Text style={s.title}>Supply Order (continued)</Text>
            )}
            {pages.length > 1 ? (
              <Text style={s.meta}>
                Page {pageIndex + 1} of {pages.length}
              </Text>
            ) : null}
          </View>
          <View style={s.table}>
            <View style={s.rowHeader}>
              <Text style={[s.th, s.colIdx]}>#</Text>
              <Text style={[s.th, s.colType]}>Type</Text>
              <Text style={[s.th, s.colCat]}>Category</Text>
              <Text style={[s.th, s.colItem]}>Item</Text>
              <Text style={[s.th, s.colExtra]}>Extra</Text>
              <Text style={[s.th, s.colSeller]}>Seller</Text>
              <Text style={[s.th, s.colCode]}>Code</Text>
              <Text style={[s.th, s.colQty]}>Qty</Text>
            </View>
            {chunk.map((r) => (
              <View key={r.idx} style={s.row}>
                <Text style={[s.td, s.colIdx]}>{String(r.idx)}</Text>
                <Text style={[s.td, s.colType]}>{pdfSafeStr(r.supplyType, 20)}</Text>
                <Text style={[s.td, s.colCat]}>{pdfSafeStr(r.category, 40)}</Text>
                <Text style={[s.td, s.colItem]}>{pdfSafeStr(r.item, 200)}</Text>
                <Text style={[s.td, s.colExtra]}>{pdfSafeStr(r.extraInfo, 120)}</Text>
                <Text style={[s.td, s.colSeller]}>{pdfSafeStr(r.seller, 30)}</Text>
                <Text style={[s.td, s.colCode]}>{pdfSafeStr(r.code, 40)}</Text>
                <Text style={[s.td, s.colQty]}>{String(r.quantity)}</Text>
              </View>
            ))}
          </View>
          {pageIndex === pages.length - 1 ? (
            <Text style={s.footer}>Generated: {pdfSafeStr(generatedAt, 120)}</Text>
          ) : null}
        </Page>
      ))}
    </Document>
  );
}

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
      {supplyType === 'dental' && (
        <td style={{ padding: '8px' }}>
          {item.category}
        </td>
      )}
      <td style={{ padding: '8px' }}>
        {item.item}
      </td>
      <td style={{ padding: '8px' }}>
        {item.extraInfo}
      </td>
      <td style={{ padding: '8px' }}>
        {getSafeUrl(item.url) ? (
          <a 
            href={getSafeUrl(item.url)!} 
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
  const [supplyType, setSupplyType] = useState('dental'); // 'dental' | 'office' | 'processing-order'
  
  // 기본 상태
  const [loading, setLoading] = useState(false);
  const [items, setItems] = useState<any[]>([]);
  const [filteredItems, setFilteredItems] = useState<any[]>([]);
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 45;
  
  // Dental과 Office 아이템을 각각 저장
  const [dentalItems, setDentalItems] = useState<any[]>([]);
  const [officeItems, setOfficeItems] = useState<any[]>([]);

  // 필터 상태
  const [categoryFilter, setCategoryFilter] = useState('');
  const [sellerFilter, setSellerFilter] = useState('');
  const [searchInput, setSearchInput] = useState('');
  const [processingDateFilter, setProcessingDateFilter] = useState('');

  const [processingPdfs, setProcessingPdfs] = useState<SupplyOrderPdfEntry[]>([]);
  const [filteredProcessingPdfs, setFilteredProcessingPdfs] = useState<SupplyOrderPdfEntry[]>([]);
  const [ordersLoading, setOrdersLoading] = useState(false);

  const [viewingOrderPdf, setViewingOrderPdf] = useState<SupplyOrderPdfEntry | null>(null);
  const [pdfStatusUpdatingPath, setPdfStatusUpdatingPath] = useState<string | null>(null);
  const [processingPdfViewerUrl, setProcessingPdfViewerUrl] = useState<string | null>(null);
  const [processingPdfViewerLoading, setProcessingPdfViewerLoading] = useState(false);
  const [processingPdfViewerError, setProcessingPdfViewerError] = useState(false);

  // Order Quantity 상태 (오피스별로 분리 저장)
  const [orderQuantitiesByOffice, setOrderQuantitiesByOffice] = useState<{ [office: string]: { [itemId: string]: string | number } }>({});
  
  // 편집 중인 Quantity 값 (오피스별로 분리 저장)
  const [editingQuantitiesByOffice, setEditingQuantitiesByOffice] = useState<{ [office: string]: { [itemId: string]: string | number } }>({});

  // Office 선택 상태
  const [selectedOffice, setSelectedOffice] = useState('');
  const [userOfficeBasedOptions, setuserOfficeBasedOptions] = useState<string[]>([]); // 사용자의 office_based 옵션들
  const [isAuthorized, setIsAuthorized] = useState<boolean | null>(null); // null: 확인 중, true: 인증됨, false: 인증 실패

  // 현재 선택된 오피스의 orderQuantities
  const orderQuantities = orderQuantitiesByOffice[selectedOffice] || {};
  
  // 현재 선택된 오피스의 editingQuantities
  const editingQuantities = editingQuantitiesByOffice[selectedOffice] || {};
  
  // Office 선택 (example 컬렉션 미사용 — 수량은 이 세션 메모리에만 유지)
  const handleOfficeSelect = useCallback((office: string) => {
    if (!office) return;
    setSelectedOffice(office);
  }, []);
  
  // debounce 타이머 저장
  const quantityTimersRef = useRef<{ [key: string]: NodeJS.Timeout }>({});
  const prevSelectedOfficeForTempRef = useRef<string>('');
  const pdfViewerBlobUrlRef = useRef<string | null>(null);

  // 이전 supplyType 추적 (useEffect에서 실제 변경 감지용)
  const prevSupplyTypeRef = useRef(supplyType);
  
  // Office 옵션 목록
  const officeOptions = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];

  const supplyTypeLabel =
    supplyType === 'dental'
      ? 'Dental'
      : supplyType === 'office'
        ? 'Office'
        : 'Processing Order';

  const isProcessingOrders = supplyType === 'processing-order';

  // 카테고리 옵션 (실제 데이터에서 동적으로 생성)
  const categoryOptions = [...new Set(items.map(item => item.category).filter(Boolean))].sort();

  // URL 파라미터에서 supply type 설정
  useEffect(() => {
    const type = searchParams.get('type');
    if (type === 'dental' || type === 'office' || type === 'processing-order') {
      setSupplyType(type);
    }
  }, [searchParams]);

  // supply type 변경 시 items 업데이트
  useEffect(() => {
    const supplyTypeChanged = prevSupplyTypeRef.current !== supplyType;
    prevSupplyTypeRef.current = supplyType;

    // supply type 변경 전에 편집 중인 값들을 로컬 상태에만 반영
    if (supplyTypeChanged && selectedOffice && Object.keys(editingQuantities).length > 0) {
      const updatedQuantities = {
        ...(orderQuantitiesByOffice[selectedOffice] || {}),
        ...editingQuantities
      };
      setOrderQuantitiesByOffice(prev => ({
        ...prev,
        [selectedOffice]: updatedQuantities
      }));
    }

    if (supplyType === 'dental') {
      setItems(dentalItems);
    } else if (supplyType === 'office') {
      setItems(officeItems);
    } else if (supplyType === 'processing-order') {
      setItems([]);
    }

    // supplyType이 실제로 변경되었을 때만 필터/편집 상태 초기화
    if (supplyTypeChanged) {
      setCategoryFilter('');
      setSellerFilter('');
      setSearchInput('');
      setProcessingDateFilter('');
      setEditingQuantitiesByOffice({});
    }
  }, [supplyType, dentalItems, officeItems]);

  // Dental과 Office 데이터 모두 로드하는 함수
  const loadAllItems = useCallback(async () => {
    try {
      setLoading(true);
      
      // Dental items 로드 (sourceType 추가 + 보안 검증)
      const dentalSnapshot = await getDocs(collection(db, 'dental-supplies'));
      const dentalList: any[] = [];
      dentalSnapshot.forEach((doc) => {
        dentalList.push({ id: doc.id, ...sanitizeData(doc.data()), sourceType: 'dental' });
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
        officeList.push({ id: doc.id, ...sanitizeData(doc.data()), sourceType: 'office' });
      });
      officeList.sort((a, b) => {
        if (a.order && b.order) return a.order - b.order;
        if (a.order && !b.order) return -1;
        if (!a.order && b.order) return 1;
        return new Date(a.createdAt || 0).getTime() - new Date(b.createdAt || 0).getTime();
      });
      setOfficeItems(officeList);
      
      if (supplyType === 'dental') {
        setItems(dentalList);
      } else if (supplyType === 'office') {
        setItems(officeList);
      }
    } catch (error) {
      alert('❌ 데이터 로드 중 오류가 발생했습니다.');
    } finally {
      setLoading(false);
    }
  }, [supplyType]);

  const loadSupplyOrderPdfs = useCallback(async () => {
    if (!selectedOffice) {
      setProcessingPdfs([]);
      setFilteredProcessingPdfs([]);
      return;
    }
    setOrdersLoading(true);
    try {
      const storage = getStorage();
      const listRef = ref(storage, `orders/${selectedOffice}/`);
      const collectSupplyOrderPdfs = async (
        currentRef: StorageReference
      ): Promise<StorageReference[]> => {
        const current = await listAll(currentRef);
        const direct = current.items.filter(
          (it) =>
            it.name.toLowerCase().endsWith('.pdf') &&
            it.name.startsWith('Supply_Order_')
        );
        const nestedArrays = await Promise.all(
          current.prefixes.map((prefixRef) => collectSupplyOrderPdfs(prefixRef))
        );
        return direct.concat(...nestedArrays);
      };
      const pdfItems = await collectSupplyOrderPdfs(listRef);
      const baseList: SupplyOrderPdfEntry[] = await Promise.all(
        pdfItems.map(async (item) => {
          const meta = await getMetadata(item);
          const createdAt = meta.timeCreated ? new Date(meta.timeCreated) : new Date();
          const parts = item.fullPath.split('/').filter(Boolean);
          const dateFolder =
            parts.length >= 3 && /^\d{4}-\d{2}-\d{2}$/.test(parts[2]) ? parts[2] : '';
          return {
            id: item.fullPath,
            path: item.fullPath,
            filename: item.name,
            createdAt,
            dateFolder,
            status: "requested" as SupplyOrderPdfStatus,
          };
        })
      );
      const list = await Promise.all(
        baseList.map(async (entry) => {
          const docId = supplyOrderPdfStatusDocIdFromStoragePath(entry.path);
          if (!docId) return entry;
          const snap = await getDoc(doc(db, SUPPLY_ORDER_PDF_STATUS_COLLECTION, docId));
          if (!snap.exists()) return entry;
          const data = sanitizeData(snap.data() || {});
          return {
            ...entry,
            status: parseSupplyOrderPdfStatusFromFirestore(data),
          };
        })
      );
      list.sort((a, b) => b.createdAt.getTime() - a.createdAt.getTime());
      setProcessingPdfs(list);
    } catch (e) {
      console.error('loadSupplyOrderPdfs:', e);
      alert('Failed to load supply order PDFs from storage.');
      setProcessingPdfs([]);
    } finally {
      setOrdersLoading(false);
    }
  }, [selectedOffice]);

  const handleSupplyOrderPdfStatusChange = useCallback(
    async (order: SupplyOrderPdfEntry, value: SupplyOrderPdfStatus) => {
      if (value !== "requested" && value !== "received") return;

      setPdfStatusUpdatingPath(order.path);
      try {
        const docId = supplyOrderPdfStatusDocIdFromStoragePath(order.path);
        if (!docId) throw new Error("invalid path");
        await setDoc(
          doc(db, SUPPLY_ORDER_PDF_STATUS_COLLECTION, docId),
          {
            office_status: value,
            updateAt: serverTimestamp(),
          },
          { merge: true }
        );
        setProcessingPdfs((prev) =>
          prev.map((o) =>
            o.path === order.path ? { ...o, status: value } : o
          )
        );
      } catch (err) {
        console.error(err);
        alert("Failed to update status.");
      } finally {
        setPdfStatusUpdatingPath(null);
      }
    },
    []
  );

  // 컴포넌트 마운트 시 데이터 로드
  useEffect(() => {
    loadAllItems();
  }, [loadAllItems]);

  useEffect(() => {
    if (supplyType === 'processing-order' && selectedOffice) {
      loadSupplyOrderPdfs();
    } else if (supplyType === 'processing-order' && !selectedOffice) {
      setProcessingPdfs([]);
      setFilteredProcessingPdfs([]);
    }
  }, [supplyType, selectedOffice, loadSupplyOrderPdfs]);

  // 컴포넌트 마운트 시 사용자 인증, role 확인 및 office_based 기반 오피스 자동 선택
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      try {
        if (!currentUser) {
          alert('Please log in.');
          setIsAuthorized(false);
          return;
        }

        // Firestore에서 사용자 role 확인
        const userDoc = await getDoc(doc(db, 'users', currentUser.uid));
        if (!userDoc.exists()) {
          alert('User information could not be found.');
          setIsAuthorized(false);
          return;
        }

        const userData = sanitizeData(userDoc.data() || {});

        if (userData?.role !== 'Manager') {
          alert('You do not have access to this page.');
          setIsAuthorized(false);
          if (typeof window !== 'undefined') {
            window.location.href = '/';
          }
          return;
        }

        setIsAuthorized(true);

        // office_based 처리: 배열이거나 단일 값일 수 있음
        if (userData?.office_based) {
          const OfficeBasedArray = Array.isArray(userData.office_based) 
            ? userData.office_based 
            : [userData.office_based];
          
          // officeOptions에 포함된 값들만 필터링
          const validOptions = OfficeBasedArray.filter((g: string) => officeOptions.includes(g));
          
          if (validOptions.length > 0) {
            setuserOfficeBasedOptions(validOptions);
            // 단일 값이면 자동 선택 (비밀번호 없이)
            if (validOptions.length === 1) {
              setSelectedOffice(validOptions[0]);
            }
          }
        }
      } catch (error: any) {
        alert('An error occurred while verifying authentication.');
        setIsAuthorized(false);
      }
    });

    // 프로덕션 환경에서 HTTPS 강제
    if (process.env.NODE_ENV === 'production' && 
        typeof window !== 'undefined' && 
        window.location.protocol !== 'https:') {
      window.location.href = window.location.href.replace('http:', 'https:');
    }

    return () => {
      unsubscribe();
    };
  }, []);

  // 컴포넌트 언마운트 시 debounce 타이머만 정리 (example 저장 없음)
  useEffect(() => {
    return () => {
      Object.keys(quantityTimersRef.current).forEach((key) => {
        clearTimeout(quantityTimersRef.current[key]);
      });
    };
  }, []);

  // Office가 바뀔 때만 temp 버킷 제거 (입력마다 effect 실행되지 않도록 deps에 editing 상태 미포함)
  useEffect(() => {
    const prev = prevSelectedOfficeForTempRef.current;
    prevSelectedOfficeForTempRef.current = selectedOffice;
    if (!selectedOffice || prev === selectedOffice) return;
    setEditingQuantitiesByOffice((p) => {
      if (!p.temp || Object.keys(p.temp).length === 0) return p;
      return { ...p, temp: {} };
    });
  }, [selectedOffice]);

  // 필터 변경 시 데이터 필터링 (Dental / Office)
  useEffect(() => {
    if (supplyType === 'processing-order') return;
    filterItems();
  }, [items, categoryFilter, sellerFilter, searchInput, supplyType]);

  useEffect(() => {
    if (supplyType !== 'processing-order') return;
    let f = [...processingPdfs];
    if (processingDateFilter) {
      f = f.filter((p) => p.dateFolder === processingDateFilter);
    }
    if (searchInput.trim()) {
      const q = searchInput.trim().toLowerCase();
      f = f.filter((p) => p.filename.toLowerCase().includes(q));
    }
    setFilteredProcessingPdfs(f);
  }, [supplyType, processingPdfs, processingDateFilter, searchInput]);

  useEffect(() => {
    setCurrentPage(1);
  }, [categoryFilter, sellerFilter, searchInput, supplyType, processingDateFilter]);

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

  const listForPage = isProcessingOrders ? filteredProcessingPdfs : filteredItems;
  const mainLoading = isProcessingOrders ? ordersLoading : loading;

  const orderDateFolders: string[] = processingPdfs
    .map((p) => p.dateFolder)
    .filter((d): d is string => typeof d === 'string' && d.length > 0);
  const processingDateOptions = [...new Set(orderDateFolders)].sort((a, b) => b.localeCompare(a));

  // 페이지네이션 계산
  const totalPages = Math.ceil(listForPage.length / itemsPerPage);
  const startIndex = (currentPage - 1) * itemsPerPage;
  const endIndex = startIndex + itemsPerPage;
  const currentPageItems = listForPage.slice(startIndex, endIndex);

  // 페이지 변경 함수
  const goToPage = (page: number) => {
    setCurrentPage(page);
  };

  // Order Quantity 변경 핸들러 (편집 중인 값을 별도 관리, DB example 없음)
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
    
    // 새 타이머 설정 (300ms 후 로컬 orderQuantities만 갱신)
    // orderQuantitiesByOffice는 클로저에 두지 않고 함수형 업데이트로 병합 → 다른 행 입력이 서로 덮어쓰지 않음
    quantityTimersRef.current[timerKey] = setTimeout(() => {
      setOrderQuantitiesByOffice((prev) => {
        const officeKey = selectedOffice;
        const prevOffice = prev[officeKey] || {};
        return {
          ...prev,
          [officeKey]: {
            ...prevOffice,
            [itemId]: value,
          },
        };
      });

      setEditingQuantitiesByOffice((prev) => {
        const officeQuantities = prev[selectedOffice] || {};
        const { [itemId]: _, ...restQuantities } = officeQuantities;
        return {
          ...prev,
          [selectedOffice]: restQuantities,
        };
      });

      delete quantityTimersRef.current[timerKey];
    }, 300);
  }, [selectedOffice]);

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

      const orderDate = new Date().toISOString();

      const rows: SupplyPdfRow[] = orderedItems.map((item, index) => ({
        idx: index + 1,
        supplyType: item.sourceType === 'dental' ? 'Dental' : 'Office',
        category: pdfSafeStr(item.category, 80),
        item: pdfSafeStr(item.item, 500),
        extraInfo: pdfSafeStr(item.extraInfo, 300),
        seller: pdfSafeStr(item.seller, 40),
        code: pdfSafeStr(item.code, 80),
        quantity: parseInt(String(currentOrderQuantities[item.id]), 10) || 0,
      }));

      const orderDateFormatted = new Date(orderDate).toLocaleString('en-US', {
        timeZone: 'America/Los_Angeles',
        dateStyle: 'medium',
        timeStyle: 'short',
      });
      const generatedAt = new Date().toLocaleString('en-US', {
        timeZone: 'America/Los_Angeles',
        dateStyle: 'full',
        timeStyle: 'short',
      });

      let pdfUploadOk = true;
      try {
        const storage = getStorage();
        const { dateFolder, timeSuffix } = getCaliforniaDateAndTimeForSupplyPdf();
        const pdfBlob = await pdf(
          <SupplyOrderPdfDocument
            office={selectedOffice}
            orderDateFormatted={orderDateFormatted}
            generatedAt={generatedAt}
            rows={rows}
          />
        ).toBlob();
        const fname = sanitizeSupplyPdfFilename(
          `Supply_Order_${selectedOffice}_${dateFolder}_${timeSuffix}.pdf`
        );
        const storageRef = ref(storage, `orders/${selectedOffice}/${dateFolder}/${fname}`);
        await uploadBytes(storageRef, pdfBlob);
        const fullPath = `orders/${selectedOffice}/${dateFolder}/${fname}`;
        const statusDocId = buildSupplyOrderPdfStatusDocId(
          selectedOffice,
          dateFolder,
          timeSuffix
        );
        if (statusDocId) {
          try {
            await setDoc(
              doc(db, SUPPLY_ORDER_PDF_STATUS_COLLECTION, statusDocId),
              {
                office_status: "requested",
                updateAt: serverTimestamp(),
              },
              { merge: true }
            );
          } catch (statusErr) {
            console.error("Supply order PDF status write failed:", statusErr);
          }
        }
      } catch (pdfErr) {
        pdfUploadOk = false;
        console.error('Supply order PDF upload failed:', pdfErr);
      }

      const dentalCount = orderedItems.filter(item => item.sourceType === 'dental').length;
      const officeCount = orderedItems.filter(item => item.sourceType === 'office').length;

      if (!pdfUploadOk) {
        alert(
          `❌ Could not save the order PDF to storage.\n\nThe order was not recorded. Please try again or contact staff.\n\nPrepared: ${orderedItems.length} item(s) from ${selectedOffice}.`
        );
        return;
      }

      let summary = `✅ Order PDF saved successfully.\n\n`;
      summary += `Total: ${orderedItems.length} item(s) from ${selectedOffice}\n`;
      if (dentalCount > 0) summary += `Dental: ${dentalCount} item(s)\n`;
      if (officeCount > 0) summary += `Office: ${officeCount} item(s)`;

      alert(summary);

      setOrderQuantitiesByOffice(prev => ({
        ...prev,
        [selectedOffice]: {}
      }));

      setEditingQuantitiesByOffice(prev => ({
        ...prev,
        [selectedOffice]: {}
      }));

      setCategoryFilter('');
      setSellerFilter('');
      setSearchInput('');

      setCurrentPage(1);
      
    } catch (error) {
      alert('❌ Failed to submit order. Please try again.');
    } finally {
      setLoading(false);
    }
  }, [selectedOffice, dentalItems, officeItems, orderQuantitiesByOffice, editingQuantitiesByOffice]);

  useEffect(() => {
    const revokeBlob = () => {
      if (pdfViewerBlobUrlRef.current) {
        URL.revokeObjectURL(pdfViewerBlobUrlRef.current);
        pdfViewerBlobUrlRef.current = null;
      }
    };

    if (!viewingOrderPdf?.path) {
      revokeBlob();
      setProcessingPdfViewerUrl(null);
      setProcessingPdfViewerLoading(false);
      setProcessingPdfViewerError(false);
      return;
    }
    const p = normalizeSupplyStoragePath(viewingOrderPdf.path);
    if (!p) {
      revokeBlob();
      setProcessingPdfViewerError(true);
      setProcessingPdfViewerLoading(false);
      return;
    }
    let cancelled = false;
    setProcessingPdfViewerLoading(true);
    setProcessingPdfViewerError(false);
    setProcessingPdfViewerUrl(null);
    revokeBlob();

    const storageRef = ref(getStorage(), p);
    getDownloadURL(storageRef)
      .then(async (url) => {
        const res = await fetch(url);
        if (!res.ok) throw new Error(`PDF fetch failed: ${res.status}`);
        const blob = await res.blob();
        if (cancelled) return;
        revokeBlob();
        const blobUrl = URL.createObjectURL(blob);
        pdfViewerBlobUrlRef.current = blobUrl;
        setProcessingPdfViewerUrl(blobUrl);
        setProcessingPdfViewerLoading(false);
      })
      .catch((err) => {
        console.error(err);
        if (!cancelled) {
          setProcessingPdfViewerError(true);
          setProcessingPdfViewerLoading(false);
        }
      });
    return () => {
      cancelled = true;
      revokeBlob();
    };
  }, [viewingOrderPdf]);

  useEffect(() => {
    if (!viewingOrderPdf) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        setViewingOrderPdf(null);
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [viewingOrderPdf]);

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

  // 인증 확인 중
  if (isAuthorized === null) {
    return (
      <div style={{
        display: 'flex',
        justifyContent: 'center',
        alignItems: 'center',
        height: '100vh',
        background: 'linear-gradient(to bottom, #a2d2ff, #f0f8ff)',
        fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
      }}>
        <div style={{ textAlign: 'center' }}>
          <div style={{ fontSize: '24px', marginBottom: '20px' }}>🔐</div>
          <div style={{ fontSize: '18px', color: '#023047' }}>Verifying authentication...</div>
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
        background: 'linear-gradient(to bottom, #a2d2ff, #f0f8ff)',
        fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
      }}>
        <div style={{ textAlign: 'center' }}>
          <div style={{ fontSize: '24px', marginBottom: '20px' }}>🚫</div>
          <div style={{ fontSize: '18px', color: '#d32f2f', marginBottom: '10px' }}>You do not have access to this page.</div>
          <div style={{ fontSize: '14px', color: '#666' }}>You do not have access to this page.</div>
        </div>
      </div>
    );
  }

  return (
    <>
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
          }}>Supply View</h1>
          
          {/* Office 선택 - 헤더 옆에 배치 */}
          <div style={{ flex: '0 0 auto', minWidth: '300px' }}>
            <label style={{ display: 'block', marginBottom: '8px', fontWeight: 'bold', color: '#495057', fontSize: '14px' }}>
              Select Office:
            </label>
            {userOfficeBasedOptions.length === 1 ? (
              <span style={{
                display: 'inline-flex',
                alignItems: 'center',
                padding: '10px 12px',
                backgroundColor: '#e9ecef',
                borderRadius: '4px',
                fontWeight: '600',
                color: '#0077B6',
                fontSize: '16px',
                width: '100%',
                boxSizing: 'border-box' as const
              }}>
                {selectedOffice}
              </span>
            ) : (
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
                {(userOfficeBasedOptions.length > 0 ? userOfficeBasedOptions : officeOptions).map(office => (
                  <option key={office} value={office}>{office}</option>
                ))}
              </select>
            )}
          </div>
        </div>

        {/* Supply Type 선택 */}
        <div style={sectionStyle}>
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
            <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer' }}>
              <input
                type="radio"
                name="supplyType"
                value="processing-order"
                checked={supplyType === 'processing-order'}
                onChange={(e) => setSupplyType(e.target.value)}
                disabled={!selectedOffice}
                style={{ margin: 0 }}
              />
              <span
                style={{
                  fontSize: '16px',
                  fontWeight: 'bold',
                  color: !selectedOffice ? '#ccc' : 'inherit',
                }}
              >
                Processing Order
                {!selectedOffice && ' (Select Office First)'}
              </span>
            </label>
          </div>
        </div>

        {/* 통합 아이템 관리 섹션 */}
        <div style={sectionStyle}>
          <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap', marginBottom: '20px' }}>
            {isProcessingOrders ? (
              <>
                <div style={{ flex: '1', minWidth: '200px' }}>
                  <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                    Order date (folder):
                  </label>
                  <select
                    value={processingDateFilter}
                    onChange={(e) => setProcessingDateFilter(e.target.value)}
                    style={inputStyle}
                  >
                    <option value="">All dates</option>
                    {processingDateOptions.map((d) => (
                      <option key={d} value={d}>
                        {d}
                      </option>
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
                    placeholder="Search by file name..."
                    style={inputStyle}
                  />
                </div>
              </>
            ) : (
              <>
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
              </>
            )}

            <div style={{ flex: '1', minWidth: '200px', display: 'flex', alignItems: 'end' }}>
              <button 
                type="button"
                onClick={() => (isProcessingOrders ? loadSupplyOrderPdfs() : loadAllItems())}
                disabled={mainLoading}
                style={{
                  ...buttonStyle,
                  backgroundColor: '#28a745',
                  width: '100%'
                }}
              >
                {mainLoading ? 'Loading...' : '🔄 Refresh Data'}
              </button>
            </div>
          </div>
          
          {mainLoading ? (
            <div style={{ textAlign: 'center', padding: '40px', color: '#666' }}>
              {isProcessingOrders ? 'Loading order PDFs...' : 'Loading items...'}
            </div>
          ) : listForPage.length === 0 ? (
            <div style={{ textAlign: 'center', padding: '40px', color: '#666' }}>
              {isProcessingOrders
                ? 'No supply order PDFs found for this office. Submit an order from Dental or Office Supply first.'
                : 'No items found for the selected criteria.'}
            </div>
          ) : isProcessingOrders ? (
            <div style={{ overflowX: 'auto' }}>
              {currentPageItems.map((order, orderIndex) => (
                <div
                  key={order.id}
                  style={{
                    marginBottom: '20px',
                    border: '1px solid #e0e0e0',
                    borderRadius: '8px',
                    overflow: 'hidden',
                  }}
                >
                  <div
                    style={{
                      backgroundColor: '#eef2f6',
                      color: '#334155',
                      padding: '15px 20px',
                      display: 'flex',
                      justifyContent: 'space-between',
                      alignItems: 'center',
                      flexWrap: 'wrap',
                      gap: '12px',
                      borderBottom: '1px solid #dde5ee',
                    }}
                  >
                    <div style={{ flex: '1', minWidth: '200px' }}>
                      <div
                        style={{
                          fontSize: '12px',
                          color: '#94a3b8',
                          marginBottom: '4px',
                          wordBreak: 'break-all',
                        }}
                      >
                        {pdfSafeStr(order.filename, 180)}
                      </div>
                      <div
                        style={{
                          fontSize: '14px',
                          color: '#64748b',
                        }}
                      >
                        {order.createdAt.toLocaleString('en-US', {
                          timeZone: 'America/Los_Angeles',
                          dateStyle: 'medium',
                          timeStyle: 'short',
                        })}
                      </div>
                    </div>
                    <div
                      style={{
                        display: 'flex',
                        alignItems: 'center',
                        flexWrap: 'wrap',
                        gap: '10px',
                      }}
                    >
                      <button
                        type="button"
                        onClick={() => {
                          if (!normalizeSupplyStoragePath(order.path)) {
                            alert('Cannot open the file.');
                            return;
                          }
                          setViewingOrderPdf(order);
                        }}
                        style={{
                          backgroundColor: '#ffffff',
                          color: '#475569',
                          border: '1px solid #cbd5e1',
                          padding: '8px 16px',
                          borderRadius: '6px',
                          fontSize: '14px',
                          fontWeight: 'bold',
                          cursor: 'pointer',
                          whiteSpace: 'nowrap',
                          boxShadow: '0 1px 2px rgba(15, 23, 42, 0.06)',
                        }}
                      >
                        📄 View PDF
                      </button>
                      <label
                        style={{
                          fontSize: '13px',
                          fontWeight: 600,
                          color: '#475569',
                          whiteSpace: 'nowrap',
                        }}
                      >
                        Status
                      </label>
                      <select
                        value={order.status}
                        disabled={pdfStatusUpdatingPath === order.path}
                        onChange={(e) => {
                          void handleSupplyOrderPdfStatusChange(
                            order,
                            e.target.value as SupplyOrderPdfStatus
                          );
                        }}
                        style={{
                          padding: '8px 12px',
                          borderRadius: '6px',
                          border: '1px solid #cbd5e1',
                          fontSize: '14px',
                          color: '#334155',
                          backgroundColor: '#ffffff',
                          minWidth: '130px',
                          cursor:
                            pdfStatusUpdatingPath === order.path ? 'wait' : 'pointer',
                        }}
                      >
                        <option value="requested">Requested</option>
                        <option value="received">Received</option>
                      </select>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          ) : (
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
                  {(currentPageItems as any[]).map((item, index) => (
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
            {isProcessingOrders
              ? `Showing ${startIndex + 1} to ${Math.min(endIndex, listForPage.length)} of ${listForPage.length} orders`
              : `Showing ${startIndex + 1} to ${Math.min(endIndex, listForPage.length)} of ${listForPage.length} items`}
          </div>
        </div>

        {/* Submit Order Button - 맨 아래 중앙에 위치 */}
        {selectedOffice && !isProcessingOrders && (
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
              {loading ? 'Submitting...' : '📦 Submit Order'}
            </button>
          </div>
        )}

      </div>
    </div>

    {viewingOrderPdf && (
      <div
        style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          backgroundColor: 'rgba(0, 0, 0, 0.9)',
          zIndex: 1000,
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'center',
          justifyContent: 'center',
        }}
        onClick={(e) => {
          if (e.target === e.currentTarget) setViewingOrderPdf(null);
        }}
      >
        <button
          type="button"
          aria-label="Close"
          onClick={(e) => {
            e.stopPropagation();
            setViewingOrderPdf(null);
          }}
          style={{
            position: 'absolute',
            top: '20px',
            right: '20px',
            background: '#ff6b6b',
            color: 'white',
            border: 'none',
            borderRadius: '50%',
            width: '40px',
            height: '40px',
            fontSize: '20px',
            cursor: 'pointer',
            zIndex: 1001,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            fontWeight: 'bold',
          }}
        >
          ×
        </button>
        <button
          type="button"
          onClick={(e) => {
            e.stopPropagation();
            setViewingOrderPdf(null);
          }}
          style={{
            position: 'absolute',
            top: '20px',
            right: '72px',
            background: 'rgba(255,255,255,0.95)',
            color: '#023047',
            border: '1px solid #ccc',
            borderRadius: '8px',
            padding: '8px 16px',
            fontSize: '14px',
            fontWeight: '600',
            cursor: 'pointer',
            zIndex: 1001,
          }}
        >
          Close
        </button>
        {processingPdfViewerLoading ? (
          <div
            style={{
              width: '90%',
              height: '90%',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              color: 'white',
              fontSize: '18px',
            }}
          >
            Loading PDF…
          </div>
        ) : processingPdfViewerError ? (
          <div
            style={{
              width: '90%',
              height: '90%',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              color: '#ffcccc',
              fontSize: '16px',
              padding: '24px',
              textAlign: 'center',
            }}
          >
            Could not load PDF.
          </div>
        ) : processingPdfViewerUrl ? (
          <object
            data={processingPdfViewerUrl}
            type="application/pdf"
            style={{
              width: '90%',
              height: '90%',
              border: 'none',
              borderRadius: '8px',
              background: 'white',
            }}
            title={pdfSafeStr(viewingOrderPdf.filename, 200)}
          >
            <p style={{ padding: '20px', color: '#333' }}>
              Your browser does not support viewing PDFs.
            </p>
          </object>
        ) : null}
      </div>
    )}
    </>
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