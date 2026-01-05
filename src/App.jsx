import React, { useMemo, useRef, useState } from 'react';
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip as RechartsTooltip,
  Legend,
  ResponsiveContainer,
  LabelList,
} from 'recharts';
import { Upload, Activity, Users, Clock, Award, CheckCircle, AlertCircle, X, ListChecks, PhoneCall } from 'lucide-react';
import * as XLSX from 'xlsx';

// ================= 1) CONFIG =====================

// وزن كل نوع مهمة (fallback فقط)
const TASK_WEIGHTS = {
  voice_whatsapp: 15,
  adb_retention_omu: 10,
  vip: 30,
  complaint: 30,
  application: 15,
  mystery_calls: 12,
  adb_transaction: 45,
  coaching: 10,
};

const DEFAULT_PAID_DAY_MINUTES = 480;
const DEFAULT_BUSINESS_DAY_MINUTES = 480;

const SHORT_CALL_THRESHOLD_MIN = 2; // <= 2 min يعتبر Short

// خرائط أسماء الأعمدة (Utilization)
const COLUMN_ALIASES = {
  date: ['تاريخ باليوم', 'Date', 'التاريخ'],
  shift: ['الشيفت', 'Shift'],
  employee: ['اسم الموظف', 'Name'],

  fullDayTraining: ['Full Day Training', 'Training', 'Full training'],

  voice_whatsapp: ["Voice - What's App", 'Voice - WhatsApp', 'Voice & WhatsApp'],
  adb_retention_omu: ['ADB , Retention , OMU and Social Media', 'ADB & Social'],
  vip: ['VIP'],
  complaint: ['Complaint', 'Complaints'],
  application: ['Application'],
  mystery_calls: ['Mysetry Calls', 'Mystery Calls', 'Mysetry Chats', 'Mystery Chats'],
  adb_transaction: ['ADB Transcation', 'ADB Transaction', 'Transactions'],
  coaching: ['Coaching'],

  otherTasks: ['Other Tasks (Min)', 'Other Tasks Minutes'],
  otherTasksComment: [
    'Other Tasks Clarificatoin or other comments',
    'Other Tasks Clarification or other comments',
    'Other Tasks Clarificatoin',
  ],

  paidDay: ['paid day', 'Paid Day', 'Paid day'],
  businessDay: ['Business day', 'Business Day'],
  adhoc: ['Ad-Hoc', 'Adhoc', 'ADHOC'],

  actualMonitoringTime: [
    'Actual Monitoring time',
    'Actual Monitoring Time',
    'Actual Monitoring time (Min)',
    'Actual Monitoring Time (Min)',
    'Actual Monitoring time (min)',
    'Actual Monitoring Time (min)',
    'Actual Monitoring time (Minutes)',
    'Actual Monitoring Time (Minutes)',
    'Actual Monitoring time (Min.)',
    'Actual Monitoring Time (Min.)',
  ],
  actualTasksTime: [
    'Actual Tasks time',
    'Actual Tasks Time',
    'Actual Tasks time (Min)',
    'Actual Tasks Time (Min)',
    'Actual Tasks time (min)',
    'Actual Tasks Time (min)',
    'Actual Tasks time (Minutes)',
    'Actual Tasks Time (Minutes)',
    'Actual Tasks time (Min.)',
    'Actual Tasks Time (Min.)',
  ],
  overallActualTime: [
    'Overall Actual time',
    'Overall Actual Time',
    'Overall Actual time (Min)',
    'Overall Actual Time (Min)',
    'Overall Actual time (min)',
    'Overall Actual Time (min)',
    'Overall Actual time (Minutes)',
    'Overall Actual Time (Minutes)',
    'Overall Actual time (Min.)',
    'Overall Actual Time (Min.)',
  ],

  accuracy: ['Acurracy', 'Accuracy'],
  internalUtilization: ['intrenal utilization', 'internal utilization', 'Internal utilization'],
  publicUtilization: ['Public Utilization', 'public utilization'],
};

// Aliases - Calls file
const CALL_COLUMN_ALIASES = {
  monitoredBy: ['Monitored By', 'monitored by', 'Monitor By', 'MonitoredBy'],
  duration: ['Duration', 'duration', 'Call Duration', 'Talk Time', 'CallDuration'],
  callDate: ['Call Date', 'call date', 'Date', 'date', 'CallDate'],
};

// ================= HELPERS =====================

function isDailySheetHeader(headers) {
  const line = headers.join(' ').toLowerCase();
  const hasDate = line.includes('date') || line.includes('تاريخ');
  const hasName = line.includes('name') || line.includes('اسم الموظف');
  return hasDate && hasName;
}

function pickByAliases(row, aliases) {
  for (const key of aliases) {
    if (row[key] !== undefined && row[key] !== null && row[key] !== '') return row[key];
    const trimmed = typeof key === 'string' ? key.trim() : key;
    if (row[trimmed] !== undefined && row[trimmed] !== null && row[trimmed] !== '') return row[trimmed];
  }
  return '';
}

function toNumber(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const s = String(v).trim().replace(/,/g, '');
  const p = s.endsWith('%') ? s.slice(0, -1) : s;
  const n = Number(p);
  return isFinite(n) ? n : 0;
}

function normalizeShift(shift) {
  return String(shift || '').trim().toLowerCase();
}

function isLeaveShift(shift) {
  const s = normalizeShift(shift);
  return (
    s === 'sick' ||
    s === 'annual' ||
    s === 'casual' ||
    s === 'off' ||
    s === 'vacation' ||
    s === 'permission' ||
    s === 'instead of' ||
    s === 'mission'
  );
}

function normalizeDate(v) {
  if (!v) return '';
  if (v instanceof Date && !isNaN(v.getTime())) return v.toISOString().split('T')[0];
  if (typeof v === 'number') {
    const dateObj = new Date(Math.round((v - 25569) * 86400 * 1000));
    return isNaN(dateObj.getTime()) ? '' : dateObj.toISOString().split('T')[0];
  }
  const s = String(v).trim();
  const m = s.match(/^(\d{4}-\d{2}-\d{2})/);
  if (m) return m[1];
  return s;
}

function ratioToPercent(r) {
  const n = toNumber(r);
  if (n === 0) return 0;
  if (n > 0 && n <= 3) return n * 100;
  return n;
}

function statusFromPublicUtil(publicPct) {
  if (publicPct < 85) return { key: 'low', label: 'Low', badge: 'bg-red-100 text-red-700' };
  if (publicPct >= 85 && publicPct <= 89) return { key: 'normal', label: 'Normal', badge: 'bg-amber-100 text-amber-700' };
  if (publicPct >= 90 && publicPct <= 95) return { key: 'good', label: 'Good', badge: 'bg-blue-100 text-blue-700' };
  return { key: 'excellent', label: 'Excellent', badge: 'bg-green-100 text-green-700' };
}

// Duration parser => minutes
// يدعم:
// - رقم مباشر (minutes)
// - "mm:ss"
// - "hh:mm:ss"
// - "hh:mm"
function parseDurationToMinutes(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;

  const s = String(v).trim();
  if (!s) return 0;

  // لو رقم في نص
  const asNum = Number(s);
  if (isFinite(asNum)) return asNum;

  const parts = s.split(':').map((x) => x.trim());
  if (parts.length < 2 || parts.length > 3) return 0;

  const nums = parts.map((p) => Number(p));
  if (nums.some((n) => !isFinite(n))) return 0;

  if (parts.length === 2) {
    const [a, b] = nums; // mm:ss OR hh:mm
    // نفترض mm:ss لو b <= 59 (ده الأغلب)
    if (b <= 59) return a + b / 60;
    // غير كده نخليه hh:mm (نادر)
    return a * 60 + b;
  }

  const [h, m, sec] = nums;
  return h * 60 + m + sec / 60;
}

function fmtHours(min) {
  const h = min / 60;
  return Number(h.toFixed(2));
}

function round2(n) {
  return Number((n || 0).toFixed(2));
}

// ✅ duration → minute bucket (1..10) + (>10)
// - 0:01 .. 0:59 => 1m
// - 1:00 .. 1:59 => 1m
// - 2:00 .. 2:59 => 2m ... إلخ
function durationToMinuteBucket(durMin) {
  if (!(durMin > 0)) return null;
  const b = Math.max(1, Math.floor(durMin));
  return b; // could be > 10
}

// ================= UI COMPONENTS =====================

const Card = ({ title, value, subtext, icon: Icon, colorClass }) => (
  <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-6 flex items-start justify-between transition-all hover:shadow-md">
    <div>
      <p className="text-slate-500 text-sm font-medium mb-1">{title}</p>
      <h3 className="text-2xl font-bold text-slate-800">{value}</h3>
      {subtext && <p className={`text-xs mt-2 ${colorClass || 'text-slate-400'}`}>{subtext}</p>}
    </div>
    <div
      className={`p-3 rounded-lg ${
        colorClass ? colorClass.replace('text-', 'bg-').replace('700', '100').replace('600', '100') : 'bg-slate-100'
      }`}
    >
      <Icon className={`w-6 h-6 ${colorClass || 'text-slate-600'}`} />
    </div>
  </div>
);

function Modal({ open, title, onClose, children }) {
  if (!open) return null;
  return (
    <div className="fixed inset-0 z-50">
      <div className="absolute inset-0 bg-black/40" onClick={onClose} />
      <div className="absolute inset-0 flex items-center justify-center p-4">
        <div className="w-full max-w-3xl bg-white rounded-2xl shadow-xl border border-slate-200 overflow-hidden">
          <div className="flex items-center justify-between px-5 py-4 border-b border-slate-200">
            <h3 className="font-bold text-slate-800">{title}</h3>
            <button onClick={onClose} className="p-2 rounded-lg hover:bg-slate-100 transition-colors text-slate-600" aria-label="close">
              <X size={18} />
            </button>
          </div>
          <div className="p-5 max-h-[70vh] overflow-auto">{children}</div>
        </div>
      </div>
    </div>
  );
}

function SortTh({ label, colKey, sortKey, sortDir, onSort, className }) {
  const active = sortKey === colKey;
  const arrow = !active ? '↕' : sortDir === 'asc' ? '↑' : '↓';
  return (
    <th
      className={`px-4 py-3 border-b whitespace-nowrap cursor-pointer select-none hover:bg-slate-100 ${className || ''}`}
      onClick={() => onSort(colKey)}
      title="اضغط للترتيب"
    >
      <div className="flex items-center justify-center gap-2">
        <span>{label}</span>
        <span className={`text-xs ${active ? 'text-blue-700 font-bold' : 'text-slate-400'}`}>{arrow}</span>
      </div>
    </th>
  );
}

// ================= MAIN APP =====================

export default function App() {
  // Utilization data
  const [data, setData] = useState([]);
  const [uploadError, setUploadError] = useState(null);
  const [selectedRow, setSelectedRow] = useState(null);

  // Calls raw rows
  const [callsRows, setCallsRows] = useState([]);
  const [callsError, setCallsError] = useState(null);

  // Tabs
  const [activeTab, setActiveTab] = useState('dashboard'); // dashboard | data | calls

  // Filters (shared)
  const [datePreset, setDatePreset] = useState('all');
  const [customFrom, setCustomFrom] = useState('');
  const [customTo, setCustomTo] = useState('');
  const [employeeFilter, setEmployeeFilter] = useState('all');

  // Sort states (existing)
  const [utilSortDir, setUtilSortDir] = useState('desc');
  const [lobSortDir, setLobSortDir] = useState('desc');
  const [leaveSortMode, setLeaveSortMode] = useState('totalDesc');
  const [dataSortKey, setDataSortKey] = useState('publicPct');
  const [dataSortDir, setDataSortDir] = useState('desc');

  // Calls table sorting (click header)
  const [callsSortKey, setCallsSortKey] = useState('totalCalls');
  const [callsSortDir, setCallsSortDir] = useState('desc');

  const fileInputRef = useRef(null);
  const callsFileInputRef = useRef(null);

  // ============ Upload Utilization ============
  const handleFileUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target.result;
        const workbook = XLSX.read(bstr, { type: 'binary' });

        const allRows = [];
        const maxRowsPerSheet = 50000;

        workbook.SheetNames.forEach((sheetName) => {
          const ws = workbook.Sheets[sheetName];
          if (!ws || !ws['!ref']) return;

          const fullRange = XLSX.utils.decode_range(ws['!ref']);
          const headerRowIndex = fullRange.s.r;

          const headers = [];
          for (let c = fullRange.s.c; c <= Math.min(fullRange.e.c, fullRange.s.c + 60); c++) {
            const cellAddress = XLSX.utils.encode_cell({ r: headerRowIndex, c });
            const cell = ws[cellAddress];
            headers.push(cell ? String(cell.v).trim() : '');
          }

          if (!isDailySheetHeader(headers)) return;

          const limitedRange = {
            s: { r: fullRange.s.r, c: fullRange.s.c },
            e: { r: Math.min(fullRange.s.r + maxRowsPerSheet, fullRange.e.r), c: fullRange.e.c },
          };

          const originalRef = ws['!ref'];
          ws['!ref'] = XLSX.utils.encode_range(limitedRange);
          const jsonData = XLSX.utils.sheet_to_json(ws, { defval: '' });
          ws['!ref'] = originalRef;

          jsonData.forEach((row) => {
            const employee = pickByAliases(row, COLUMN_ALIASES.employee || []);
            const dateVal = pickByAliases(row, COLUMN_ALIASES.date || []);
            if (!employee && !dateVal) return;

            allRows.push({
              sheetName,
              date: normalizeDate(dateVal),
              shift: pickByAliases(row, COLUMN_ALIASES.shift || []),
              employee,
              fullDayTraining: pickByAliases(row, COLUMN_ALIASES.fullDayTraining || []),

              voice_whatsapp: pickByAliases(row, COLUMN_ALIASES.voice_whatsapp || []),
              adb_retention_omu: pickByAliases(row, COLUMN_ALIASES.adb_retention_omu || []),
              vip: pickByAliases(row, COLUMN_ALIASES.vip || []),
              complaint: pickByAliases(row, COLUMN_ALIASES.complaint || []),
              application: pickByAliases(row, COLUMN_ALIASES.application || []),
              mystery_calls: pickByAliases(row, COLUMN_ALIASES.mystery_calls || []),
              adb_transaction: pickByAliases(row, COLUMN_ALIASES.adb_transaction || []),
              coaching: pickByAliases(row, COLUMN_ALIASES.coaching || []),

              otherTasks: pickByAliases(row, COLUMN_ALIASES.otherTasks || []),
              otherTasksComment: pickByAliases(row, COLUMN_ALIASES.otherTasksComment || []),

              paidDay: pickByAliases(row, COLUMN_ALIASES.paidDay || []),
              businessDay: pickByAliases(row, COLUMN_ALIASES.businessDay || []),
              adhoc: pickByAliases(row, COLUMN_ALIASES.adhoc || []),

              actualMonitoringTime: pickByAliases(row, COLUMN_ALIASES.actualMonitoringTime || []),
              actualTasksTime: pickByAliases(row, COLUMN_ALIASES.actualTasksTime || []),
              overallActualTime: pickByAliases(row, COLUMN_ALIASES.overallActualTime || []),

              accuracy: pickByAliases(row, COLUMN_ALIASES.accuracy || []),
              internalUtilization: pickByAliases(row, COLUMN_ALIASES.internalUtilization || []),
              publicUtilization: pickByAliases(row, COLUMN_ALIASES.publicUtilization || []),
            });
          });
        });

        if (allRows.length === 0) {
          setUploadError('لم يتم العثور على بيانات صالحة في أي شيت يومي.');
          return;
        }

        setData(allRows);
        setUploadError(null);
        alert(`تم رفع ملف الـ Utilization ✅\nتم تحميل ${allRows.length} صف.`);
      } catch (error) {
        console.error(error);
        setUploadError('حدث خطأ أثناء قراءة ملف الـ Utilization. تأكد من تطابق أسماء الأعمدة.');
      }
    };
    reader.readAsBinaryString(file);
  };

  const triggerUpload = () => {
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
      fileInputRef.current.click();
    }
  };

  // ============ Upload Calls ============
  const handleCallsUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target.result;
        const workbook = XLSX.read(bstr, { type: 'binary' });

        const all = [];
        workbook.SheetNames.forEach((sheetName) => {
          const ws = workbook.Sheets[sheetName];
          if (!ws || !ws['!ref']) return;

          const jsonData = XLSX.utils.sheet_to_json(ws, { defval: '' });
          jsonData.forEach((row) => {
            const monitoredBy = pickByAliases(row, CALL_COLUMN_ALIASES.monitoredBy);
            const durationRaw = pickByAliases(row, CALL_COLUMN_ALIASES.duration);
            const callDateRaw = pickByAliases(row, CALL_COLUMN_ALIASES.callDate);

            if (!monitoredBy && !durationRaw && !callDateRaw) return;

            const durationMin = parseDurationToMinutes(durationRaw);
            const date = normalizeDate(callDateRaw);

            all.push({
              monitoredBy: String(monitoredBy || '').trim(),
              durationMin: durationMin,
              callDate: date,
              rawDuration: durationRaw,
              sheetName,
            });
          });
        });

        const hasAny = all.some((r) => r.monitoredBy || r.rawDuration !== '');
        if (!hasAny) {
          setCallsError('لم يتم العثور على بيانات صالحة في ملف المكالمات. تأكد من وجود (Monitored By) و (Duration).');
          return;
        }

        setCallsRows(all);
        setCallsError(null);
        alert(`تم رفع ملف المكالمات ✅\nتم تحميل ${all.length} صف.`);
      } catch (err) {
        console.error(err);
        setCallsError('حدث خطأ أثناء قراءة ملف المكالمات. تأكد من وجود الأعمدة (Monitored By) و (Duration) ويفضل (Call Date).');
      }
    };

    reader.readAsBinaryString(file);
  };

  const triggerCallsUpload = () => {
    if (callsFileInputRef.current) {
      callsFileInputRef.current.value = '';
      callsFileInputRef.current.click();
    }
  };

  // ============ Date range helper ============
  function getPresetRange() {
    if (datePreset === 'all') return null;

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let from, to;

    switch (datePreset) {
      case 'today':
        from = new Date(today);
        to = new Date(today);
        break;
      case '7days':
        to = new Date(today);
        from = new Date(today);
        from.setDate(from.getDate() - 6);
        break;
      case 'thisMonth': {
        const year = today.getFullYear();
        const month = today.getMonth();
        from = new Date(year, month, 1);
        to = new Date(year, month + 1, 0);
        break;
      }
      case 'lastMonth': {
        const year = today.getFullYear();
        const month = today.getMonth() - 1;
        from = new Date(year, month, 1);
        to = new Date(year, month + 1, 0);
        break;
      }
      case 'custom':
        if (!customFrom || !customTo) return null;
        from = new Date(customFrom);
        to = new Date(customTo);
        break;
      default:
        return null;
    }
    from.setHours(0, 0, 0, 0);
    to.setHours(23, 59, 59, 999);
    return { from, to };
  }

  // ============ Process Utilization ============
  const processedData = useMemo(() => {
    return data
      .map((row) => {
        const shift = row.shift;
        const leave = isLeaveShift(shift);
        const fullDayTraining = toNumber(row.fullDayTraining) === 1;

        const paidDay = toNumber(row.paidDay) || (leave ? 0 : DEFAULT_PAID_DAY_MINUTES);
        const businessDay = toNumber(row.businessDay) || (leave ? 0 : DEFAULT_BUSINESS_DAY_MINUTES);
        const adhoc = toNumber(row.adhoc);

        let actualMonitoring = toNumber(row.actualMonitoringTime);
        if (!actualMonitoring) {
          actualMonitoring =
            toNumber(row.voice_whatsapp) * TASK_WEIGHTS.voice_whatsapp +
            toNumber(row.adb_retention_omu) * TASK_WEIGHTS.adb_retention_omu +
            toNumber(row.vip) * TASK_WEIGHTS.vip +
            toNumber(row.complaint) * TASK_WEIGHTS.complaint +
            toNumber(row.application) * TASK_WEIGHTS.application +
            toNumber(row.mystery_calls) * TASK_WEIGHTS.mystery_calls +
            toNumber(row.adb_transaction) * TASK_WEIGHTS.adb_transaction +
            toNumber(row.coaching) * TASK_WEIGHTS.coaching;
        }

        let actualTasks = toNumber(row.actualTasksTime);
        if (!actualTasks) actualTasks = toNumber(row.otherTasks);

        let overallActual = toNumber(row.overallActualTime);
        if (!overallActual) {
          overallActual = actualMonitoring + actualTasks + (fullDayTraining ? DEFAULT_PAID_DAY_MINUTES : 0);
        }

        if (leave) {
          actualMonitoring = 0;
          actualTasks = 0;
          overallActual = 0;
        }

        const accuracyRatio = toNumber(row.accuracy) || (paidDay ? overallActual / paidDay : 0);
        const internalRatio = toNumber(row.internalUtilization) || (businessDay ? (overallActual + adhoc) / businessDay : 0);
        const publicRatio = toNumber(row.publicUtilization) || (paidDay ? (overallActual + adhoc) / paidDay : 0);

        const accuracyPct = ratioToPercent(accuracyRatio);
        const internalPct = ratioToPercent(internalRatio);
        const publicPct = ratioToPercent(publicRatio);

        const status = statusFromPublicUtil(publicPct);

        const monitoringPct = overallActual ? (actualMonitoring / overallActual) * 100 : 0;
        const tasksPct = overallActual ? (actualTasks / overallActual) * 100 : 0;

        const taskDetails = {
          voice_whatsapp: { count: toNumber(row.voice_whatsapp), minutes: toNumber(row.voice_whatsapp) * TASK_WEIGHTS.voice_whatsapp },
          adb_retention_omu: { count: toNumber(row.adb_retention_omu), minutes: toNumber(row.adb_retention_omu) * TASK_WEIGHTS.adb_retention_omu },
          vip: { count: toNumber(row.vip), minutes: toNumber(row.vip) * TASK_WEIGHTS.vip },
          complaint: { count: toNumber(row.complaint), minutes: toNumber(row.complaint) * TASK_WEIGHTS.complaint },
          application: { count: toNumber(row.application), minutes: toNumber(row.application) * TASK_WEIGHTS.application },
          mystery_calls: { count: toNumber(row.mystery_calls), minutes: toNumber(row.mystery_calls) * TASK_WEIGHTS.mystery_calls },
          adb_transaction: { count: toNumber(row.adb_transaction), minutes: toNumber(row.adb_transaction) * TASK_WEIGHTS.adb_transaction },
          coaching: { count: toNumber(row.coaching), minutes: toNumber(row.coaching) * TASK_WEIGHTS.coaching },
        };

        return {
          ...row,
          date: row.date,
          shift,
          leave,
          paidDay,
          businessDay,
          adhoc,

          actualMonitoring,
          actualTasks,
          overallActual,

          accuracyPct: Number(accuracyPct.toFixed(1)),
          internalPct: Number(internalPct.toFixed(1)),
          publicPct: Number(publicPct.toFixed(1)),

          monitoringPct: Number(monitoringPct.toFixed(1)),
          tasksPct: Number(tasksPct.toFixed(1)),

          status,
          taskDetails,
        };
      })
      .filter((r) => r.employee || r.date);
  }, [data]);

  // Filters (Utilization)
  const allEmployees = useMemo(() => Array.from(new Set(processedData.map((d) => d.employee).filter(Boolean))).sort(), [processedData]);

  const filteredData = useMemo(() => {
    const range = getPresetRange();
    return processedData.filter((row) => {
      if (employeeFilter !== 'all' && row.employee !== employeeFilter) return false;
      if (range && row.date) {
        const d = new Date(row.date);
        if (isNaN(d.getTime())) return false;
        if (d < range.from || d > range.to) return false;
      }
      return true;
    });
  }, [processedData, employeeFilter, datePreset, customFrom, customTo]);

  // ============ KPIs (Utilization) ============
  const stats = useMemo(() => {
    if (filteredData.length === 0) {
      return { totalEmployees: 0, avgPublicUtil: 0, totalHours: 0, topPerformer: { name: '-', score: 0 }, monitoringShare: 0, tasksShare: 0 };
    }

    const totalEmployees = new Set(filteredData.map((d) => d.employee).filter(Boolean)).size;
    const rowsWithPaid = filteredData.filter((d) => d.paidDay > 0);
    const avgPublicUtil = (rowsWithPaid.reduce((acc, curr) => acc + (curr.publicPct || 0), 0) / Math.max(rowsWithPaid.length, 1)) || 0;

    const overallMinutes = filteredData.reduce((acc, curr) => acc + (curr.overallActual || 0), 0);
    const totalHours = overallMinutes / 60;

    const empPerformance = {};
    rowsWithPaid.forEach((d) => {
      if (!d.employee) return;
      if (!empPerformance[d.employee]) empPerformance[d.employee] = { total: 0, count: 0 };
      empPerformance[d.employee].total += d.publicPct || 0;
      empPerformance[d.employee].count += 1;
    });

    let topPerformer = { name: '-', score: 0 };
    Object.keys(empPerformance).forEach((emp) => {
      const avg = empPerformance[emp].total / empPerformance[emp].count;
      if (avg > topPerformer.score) topPerformer = { name: emp, score: avg };
    });

    const monitoringMinutes = filteredData.reduce((acc, curr) => acc + (curr.actualMonitoring || 0), 0);
    const tasksMinutes = filteredData.reduce((acc, curr) => acc + (curr.actualTasks || 0), 0);
    const denom = monitoringMinutes + tasksMinutes;
    const monitoringShare = denom ? (monitoringMinutes / denom) * 100 : 0;
    const tasksShare = denom ? (tasksMinutes / denom) * 100 : 0;

    return { totalEmployees, avgPublicUtil, totalHours, topPerformer, monitoringShare, tasksShare };
  }, [filteredData]);

  // ============ Charts (Utilization) ============
  const chartData = useMemo(() => {
    const empData = {};
    const rowsWithPaid = filteredData.filter((d) => d.paidDay > 0);
    rowsWithPaid.forEach((d) => {
      if (!d.employee) return;
      if (!empData[d.employee]) empData[d.employee] = { name: d.employee, sum: 0, count: 0 };
      empData[d.employee].sum += d.publicPct || 0;
      empData[d.employee].count += 1;
    });

    return Object.values(empData)
      .map((d) => ({ name: d.name, avgPublic: Number((d.sum / d.count).toFixed(1)) }))
      .sort((a, b) => (utilSortDir === 'desc' ? b.avgPublic - a.avgPublic : a.avgPublic - b.avgPublic));
  }, [filteredData, utilSortDir]);

  const taskDistributionLOBData = useMemo(() => {
    const totals = {
      "Voice - What's App": 0,
      'ADB , Retention , OMU and Social Media': 0,
      VIP: 0,
      Complaint: 0,
      Application: 0,
      'Mysetry Calls': 0,
      'ADB Transcation': 0,
      Coaching: 0,
      'Other Tasks (Min)': 0,
    };

    filteredData.forEach((row) => {
      if (row.leave) return;
      totals["Voice - What's App"] += toNumber(row.voice_whatsapp) * TASK_WEIGHTS.voice_whatsapp;
      totals['ADB , Retention , OMU and Social Media'] += toNumber(row.adb_retention_omu) * TASK_WEIGHTS.adb_retention_omu;
      totals.VIP += toNumber(row.vip) * TASK_WEIGHTS.vip;
      totals.Complaint += toNumber(row.complaint) * TASK_WEIGHTS.complaint;
      totals.Application += toNumber(row.application) * TASK_WEIGHTS.application;
      totals['Mysetry Calls'] += toNumber(row.mystery_calls) * TASK_WEIGHTS.mystery_calls;
      totals['ADB Transcation'] += toNumber(row.adb_transaction) * TASK_WEIGHTS.adb_transaction;
      totals.Coaching += toNumber(row.coaching) * TASK_WEIGHTS.coaching;
      totals['Other Tasks (Min)'] += toNumber(row.otherTasks);
    });

    const arr = Object.entries(totals).map(([name, minutes]) => ({ name, hours: Number((minutes / 60).toFixed(1)) }));
    arr.sort((a, b) => (lobSortDir === 'desc' ? b.hours - a.hours : a.hours - b.hours));
    return arr;
  }, [filteredData, lobSortDir]);

  const leaveData = useMemo(() => {
    const map = {};
    filteredData.forEach((row) => {
      const emp = row.employee || '-';
      if (!map[emp]) map[emp] = { name: emp, Casual: 0, Annual: 0, Sick: 0 };
      const s = normalizeShift(row.shift);
      if (s === 'casual') map[emp].Casual += 1;
      if (s === 'annual') map[emp].Annual += 1;
      if (s === 'sick') map[emp].Sick += 1;
    });

    const arr = Object.values(map).filter((r) => r.name !== '-' && (r.Casual + r.Annual + r.Sick) > 0);
    const total = (r) => r.Casual + r.Annual + r.Sick;

    switch (leaveSortMode) {
      case 'totalAsc':
        arr.sort((a, b) => total(a) - total(b));
        break;
      case 'nameAsc':
        arr.sort((a, b) => a.name.localeCompare(b.name));
        break;
      case 'nameDesc':
        arr.sort((a, b) => b.name.localeCompare(a.name));
        break;
      case 'totalDesc':
      default:
        arr.sort((a, b) => total(b) - total(a));
        break;
    }
    return arr;
  }, [filteredData, leaveSortMode]);

  // ============ Calls: filter by date & employee ============
  const filteredCalls = useMemo(() => {
    const range = getPresetRange();
    return callsRows.filter((r) => {
      if (employeeFilter !== 'all' && (r.monitoredBy || '') !== employeeFilter) return false;

      // لو مفيش callDate في الملف، هنعدّي الفلترة الزمنية (مش هنمنع)
      if (range && r.callDate) {
        const d = new Date(r.callDate);
        if (isNaN(d.getTime())) return false;
        if (d < range.from || d > range.to) return false;
      }
      return true;
    });
  }, [callsRows, employeeFilter, datePreset, customFrom, customTo]);

  // ============ Calls: aggregation per employee ============
  const callsReport = useMemo(() => {
    const map = {};

    filteredCalls.forEach((r) => {
      const emp = (r.monitoredBy || '').trim();
      if (!emp) return;

      if (!map[emp]) {
        map[emp] = {
          name: emp,
          totalCalls: 0, // كل الصفوف
          validCalls: 0, // Duration > 0 (دا داخلي فقط)
          totalDurationMin: 0, // على valid فقط
          shortCalls: 0, // على valid فقط

          // ✅ minute buckets (valid only)
          m1: 0,
          m2: 0,
          m3: 0,
          m4: 0,
          m5: 0,
          m6: 0,
          m7: 0,
          m8: 0,
          m9: 0,
          m10: 0,
          gt10: 0,
        };
      }

      map[emp].totalCalls += 1;

      const dur = r.durationMin;
      if (dur > 0) {
        map[emp].validCalls += 1;
        map[emp].totalDurationMin += dur;

        if (dur <= SHORT_CALL_THRESHOLD_MIN) map[emp].shortCalls += 1;

        const b = durationToMinuteBucket(dur);
        if (b !== null) {
          if (b > 10) map[emp].gt10 += 1;
          else map[emp][`m${b}`] += 1;
        }
      }
    });

    const arr = Object.values(map).map((x) => {
      const avg = x.validCalls ? x.totalDurationMin / x.validCalls : 0;
      const shortPct = x.validCalls ? (x.shortCalls / x.validCalls) * 100 : 0;

      // ✅ Groups (مش تراكمي)
      const g13 = x.m1 + x.m2 + x.m3; // 1–3 only
      const g46 = x.m4 + x.m5 + x.m6; // 4–6 only
      const g710 = x.m7 + x.m8 + x.m9 + x.m10; // 7–10 only

      return {
        name: x.name,
        totalCalls: x.totalCalls,
        avgDurationMin: round2(avg),
        totalDurationMin: round2(x.totalDurationMin),
        totalDurationHours: round2(fmtHours(x.totalDurationMin)),
        shortCalls: x.shortCalls,
        shortPct: round2(shortPct),

        m1: x.m1,
        m2: x.m2,
        m3: x.m3,
        g13,
        m4: x.m4,
        m5: x.m5,
        m6: x.m6,
        g46,
        m7: x.m7,
        m8: x.m8,
        m9: x.m9,
        m10: x.m10,
        g710,
        gt10: x.gt10,
      };
    });

    // sorting by selected column
    const getVal = (row, k) => {
      if (k === 'name') return String(row.name || '');
      return Number(row[k] || 0);
    };

    arr.sort((a, b) => {
      const va = getVal(a, callsSortKey);
      const vb = getVal(b, callsSortKey);

      if (callsSortKey === 'name') {
        return callsSortDir === 'asc' ? va.localeCompare(vb) : vb.localeCompare(va);
      }
      return callsSortDir === 'asc' ? va - vb : vb - va;
    });

    return arr;
  }, [filteredCalls, callsSortKey, callsSortDir]);

  // Calls charts
  const callsAvgChart = useMemo(() => {
    return [...callsReport]
      .sort((a, b) => b.avgDurationMin - a.avgDurationMin)
      .map((r) => ({ name: r.name, avg: r.avgDurationMin }));
  }, [callsReport]);

  const callsShortPctChart = useMemo(() => {
    return [...callsReport]
      .sort((a, b) => b.shortPct - a.shortPct)
      .map((r) => ({ name: r.name, shortPct: r.shortPct }));
  }, [callsReport]);

  function onCallsSort(colKey) {
    if (callsSortKey === colKey) {
      setCallsSortDir((d) => (d === 'asc' ? 'desc' : 'asc'));
    } else {
      setCallsSortKey(colKey);
      setCallsSortDir(colKey === 'name' ? 'asc' : 'desc');
    }
  }

  // ============ Utilization detailed table ============
  const tableColumns = [
    { key: 'date', label: 'Date' },
    { key: 'shift', label: 'Shift' },
    { key: 'employee', label: 'Name' },

    { key: 'fullDayTraining', label: 'Full Day Training' },
    { key: 'voice_whatsapp', label: "Voice - What's App" },
    { key: 'adb_retention_omu', label: 'ADB , Retention , OMU and Social Media' },
    { key: 'vip', label: 'VIP' },
    { key: 'complaint', label: 'Complaint' },
    { key: 'application', label: 'Application' },
    { key: 'mystery_calls', label: 'Mysetry Calls' },
    { key: 'adb_transaction', label: 'ADB Transcation' },
    { key: 'coaching', label: 'Coaching' },
    { key: 'otherTasks', label: 'Other Tasks (Min)' },
    { key: 'otherTasksComment', label: 'Other Tasks Clarificatoin or other comments' },

    { key: 'actualMonitoring', label: 'Actual Monitoring time (Min)', isComputed: true },
    { key: 'accuracyPct', label: 'Acurracy %', isComputed: true },
    { key: 'internalPct', label: 'intrenal utilization %', isComputed: true },
    { key: 'publicPct', label: 'Public Utilization %', isComputed: true },

    { key: 'status', label: 'Status', isComputed: true },
    { key: 'taskDetails', label: 'Task details', isComputed: true },
  ];

  const sortedFilteredData = useMemo(() => {
    const rows = [...filteredData];

    const getVal = (row, key) => {
      if (key === 'actualMonitoring') return toNumber(row.actualMonitoring);
      if (key === 'accuracyPct') return toNumber(row.accuracyPct);
      if (key === 'internalPct') return toNumber(row.internalPct);
      if (key === 'publicPct') return toNumber(row.publicPct);
      if (key === 'date') return new Date(row.date || '1970-01-01').getTime();
      if (key === 'employee') return String(row.employee || '');
      return row[key];
    };

    rows.sort((a, b) => {
      const va = getVal(a, dataSortKey);
      const vb = getVal(b, dataSortKey);

      if (typeof va === 'number' && typeof vb === 'number') {
        return dataSortDir === 'desc' ? vb - va : va - vb;
      }
      return dataSortDir === 'desc' ? String(vb).localeCompare(String(va)) : String(va).localeCompare(String(vb));
    });

    return rows;
  }, [filteredData, dataSortKey, dataSortDir]);

  // highlights
  const thHi13 = 'bg-amber-50 text-amber-800 font-extrabold';
  const thHi46 = 'bg-blue-50 text-blue-800 font-extrabold';
  const thHi710 = 'bg-green-50 text-green-800 font-extrabold';

  const tdHi13 = 'bg-amber-50 text-amber-900 font-extrabold text-base';
  const tdHi46 = 'bg-blue-50 text-blue-900 font-extrabold text-base';
  const tdHi710 = 'bg-green-50 text-green-900 font-extrabold text-base';

  // ================= RENDER =====================
  return (
    <div dir="rtl" className="min-h-screen bg-slate-50 font-sans text-slate-800">
      {/* Top Nav */}
      <nav className="bg-white border-b border-slate-200 px-6 py-4 flex items-center justify-between sticky top-0 z-10 shadow-sm">
        <div className="flex items-center gap-3">
          <div className="bg-blue-600 p-2 rounded-lg text-white">
            <Activity size={24} />
          </div>
          <div>
            <h1 className="text-xl font-bold text-slate-900 tracking-tight">لوحة تحكم الأداء</h1>
            <p className="text-xs text-slate-500">Utilization Dashboard</p>
          </div>
        </div>

        <div className="flex items-center gap-3">
          {/* Utilization Upload */}
          <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} ref={fileInputRef} className="hidden" />
          <button
            onClick={triggerUpload}
            className="flex items-center gap-2 px-4 py-2 bg-slate-900 text-white rounded-lg hover:bg-slate-800 transition-colors shadow-lg shadow-slate-900/20 text-sm font-medium"
          >
            <Upload size={16} />
            <span>رفع ملف Utilization</span>
          </button>

          {/* Calls Upload */}
          <input type="file" accept=".xlsx, .xls" onChange={handleCallsUpload} ref={callsFileInputRef} className="hidden" />
          <button
            onClick={triggerCallsUpload}
            className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors shadow-lg shadow-blue-600/20 text-sm font-medium"
          >
            <PhoneCall size={16} />
            <span>رفع ملف المكالمات</span>
          </button>
        </div>
      </nav>

      <div className="p-6 max-w-screen-2xl mx-auto space-y-6">
        {/* Errors */}
        {uploadError && (
          <div className="bg-red-50 text-red-700 p-4 rounded-lg flex items-center gap-2 border border-red-200">
            <AlertCircle size={20} />
            <span>{uploadError}</span>
          </div>
        )}
        {callsError && (
          <div className="bg-red-50 text-red-700 p-4 rounded-lg flex items-center gap-2 border border-red-200">
            <AlertCircle size={20} />
            <span>{callsError}</span>
          </div>
        )}

        {/* Tabs + Filters */}
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          {/* Tabs */}
          <div className="flex items-center bg-white rounded-lg p-1 border border-slate-200 w-fit shadow-sm">
            <button
              onClick={() => setActiveTab('dashboard')}
              className={`px-4 py-2 rounded-md text-sm font-medium transition-all ${
                activeTab === 'dashboard' ? 'bg-blue-50 text-blue-700' : 'text-slate-600 hover:bg-slate-50'
              }`}
            >
              نظرة عامة
            </button>
            <button
              onClick={() => setActiveTab('data')}
              className={`px-4 py-2 rounded-md text-sm font-medium transition-all ${
                activeTab === 'data' ? 'bg-blue-50 text-blue-700' : 'text-slate-600 hover:bg-slate-50'
              }`}
            >
              البيانات التفصيلية
            </button>
            <button
              onClick={() => setActiveTab('calls')}
              className={`px-4 py-2 rounded-md text-sm font-medium transition-all ${
                activeTab === 'calls' ? 'bg-blue-50 text-blue-700' : 'text-slate-600 hover:bg-slate-50'
              }`}
            >
              تقرير المكالمات
            </button>
          </div>

          {/* Filters */}
          <div className="flex flex-wrap items-center gap-3">
            <div className="flex flex-col">
              <span className="text-xs text-slate-500 mb-1">الفترة الزمنية</span>
              <select
                className="bg-white border border-slate-300 text-slate-700 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 px-2 py-1"
                value={datePreset}
                onChange={(e) => setDatePreset(e.target.value)}
              >
                <option value="all">كل الفترة</option>
                <option value="today">اليوم</option>
                <option value="7days">آخر 7 أيام</option>
                <option value="thisMonth">الشهر الحالي</option>
                <option value="lastMonth">الشهر السابق</option>
                <option value="custom">مخصص</option>
              </select>
            </div>

            {datePreset === 'custom' && (
              <div className="flex items-center gap-2">
                <div className="flex flex-col">
                  <span className="text-xs text-slate-500 mb-1">من</span>
                  <input
                    type="date"
                    className="bg-white border border-slate-300 text-slate-700 text-sm rounded-lg px-2 py-1"
                    value={customFrom}
                    onChange={(e) => setCustomFrom(e.target.value)}
                  />
                </div>
                <div className="flex flex-col">
                  <span className="text-xs text-slate-500 mb-1">إلى</span>
                  <input
                    type="date"
                    className="bg-white border border-slate-300 text-slate-700 text-sm rounded-lg px-2 py-1"
                    value={customTo}
                    onChange={(e) => setCustomTo(e.target.value)}
                  />
                </div>
              </div>
            )}

            <div className="flex flex-col">
              <span className="text-xs text-slate-500 mb-1">الموظف</span>
              <select
                className="bg-white border border-slate-300 text-slate-700 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 px-2 py-1 min-w-[170px]"
                value={employeeFilter}
                onChange={(e) => setEmployeeFilter(e.target.value)}
              >
                <option value="all">جميع الموظفين</option>
                {allEmployees.map((emp) => (
                  <option key={emp} value={emp}>
                    {emp}
                  </option>
                ))}
              </select>
            </div>
          </div>
        </div>

        {/* DASHBOARD */}
        {activeTab === 'dashboard' && (
          <>
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-4">
              <Card
                title="متوسط Public Utilization"
                value={`${stats.avgPublicUtil.toFixed(1)}%`}
                subtext="حسب الفلترة الحالية"
                colorClass={stats.avgPublicUtil >= 90 ? 'text-green-700' : stats.avgPublicUtil >= 85 ? 'text-amber-700' : 'text-red-700'}
                icon={Activity}
              />
              <Card title="إجمالي ساعات العمل" value={Math.round(stats.totalHours)} subtext="ساعة" colorClass="text-blue-700" icon={Clock} />
              <Card title="الموظف الأفضل" value={stats.topPerformer.name} subtext={`Avg ${stats.topPerformer.score.toFixed(1)}%`} colorClass="text-purple-700" icon={Award} />
              <Card title="عدد الموظفين" value={stats.totalEmployees} subtext="نشطين في الداتا" colorClass="text-slate-700" icon={Users} />
              <Card title="% Monitoring / % Tasks" value={`${stats.monitoringShare.toFixed(0)}% / ${stats.tasksShare.toFixed(0)}%`} subtext="من إجمالي الدقائق" colorClass="text-slate-700" icon={ListChecks} />
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              <div className="lg:col-span-2 bg-white p-6 rounded-xl border border-slate-100 shadow-sm">
                <div className="flex items-center justify-between mb-6">
                  <h3 className="text-lg font-bold text-slate-800">Public Utilization (Average %) لكل موظف</h3>
                  <div className="flex items-center gap-2">
                    <span className="text-xs text-slate-500">Sort</span>
                    <select
                      className="bg-white border border-slate-300 text-slate-700 text-xs rounded-lg px-2 py-1"
                      value={utilSortDir}
                      onChange={(e) => setUtilSortDir(e.target.value)}
                    >
                      <option value="desc">الأكبر → الأصغر</option>
                      <option value="asc">الأصغر → الأكبر</option>
                    </select>
                  </div>
                </div>

                <div className="h-80 w-full" style={{ direction: 'ltr' }}>
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={chartData} margin={{ top: 20, right: 30, left: 20, bottom: 20 }}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                      <XAxis dataKey="name" interval={0} angle={-10} textAnchor="end" height={60} />
                      <YAxis domain={[0, 140]} />
                      <RechartsTooltip formatter={(v) => [`${v}%`, 'Public Utilization']} />
                      <Bar dataKey="avgPublic" fill="#3b82f6" radius={[4, 4, 0, 0]} barSize={36}>
                        <LabelList dataKey="avgPublic" position="top" formatter={(v) => `${v}%`} />
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>

              <div className="bg-white p-6 rounded-xl border border-slate-100 shadow-sm">
                <div className="flex items-center justify-between mb-6">
                  <h3 className="text-lg font-bold text-slate-800">Task Distribution (Hours)</h3>
                  <div className="flex items-center gap-2">
                    <span className="text-xs text-slate-500">Sort</span>
                    <select
                      className="bg-white border border-slate-300 text-slate-700 text-xs rounded-lg px-2 py-1"
                      value={lobSortDir}
                      onChange={(e) => setLobSortDir(e.target.value)}
                    >
                      <option value="desc">الأكبر → الأصغر</option>
                      <option value="asc">الأصغر → الأكبر</option>
                    </select>
                  </div>
                </div>

                <div className="h-[420px] w-full" style={{ direction: 'ltr' }}>
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={taskDistributionLOBData} margin={{ top: 20, right: 20, left: 10, bottom: 110 }}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                      <XAxis dataKey="name" interval={0} angle={-90} textAnchor="end" height={120} tickMargin={10} />
                      <YAxis />
                      <RechartsTooltip formatter={(v) => [`${v} h`, 'Hours']} />
                      <Bar dataKey="hours" fill="#10b981" radius={[4, 4, 0, 0]} name="Hours">
                        <LabelList dataKey="hours" position="top" />
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>

            <div className="bg-white p-6 rounded-xl border border-slate-100 shadow-sm">
              <div className="flex items-center justify-between mb-4">
                <h3 className="text-lg font-bold text-slate-800">Casual / Annual / Sick Days per Specialist</h3>
                <div className="flex items-center gap-2">
                  <span className="text-xs text-slate-500">Sort</span>
                  <select
                    className="bg-white border border-slate-300 text-slate-700 text-xs rounded-lg px-2 py-1"
                    value={leaveSortMode}
                    onChange={(e) => setLeaveSortMode(e.target.value)}
                  >
                    <option value="totalDesc">Most Leave</option>
                    <option value="totalAsc">Least Leave</option>
                    <option value="nameAsc">Name A → Z</option>
                    <option value="nameDesc">Name Z → A</option>
                  </select>
                </div>
              </div>

              {leaveData.length === 0 ? (
                <p className="text-sm text-slate-500">لا توجد أيام (Casual / Annual / Sick) ضمن الفلترة الحالية.</p>
              ) : (
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                  <div className="overflow-x-auto">
                    <table className="w-full text-sm text-right">
                      <thead className="bg-slate-50 text-slate-600 font-medium">
                        <tr>
                          <th className="px-4 py-3 border-b">Name</th>
                          <th className="px-4 py-3 border-b">Casual</th>
                          <th className="px-4 py-3 border-b">Annual</th>
                          <th className="px-4 py-3 border-b">Sick</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {leaveData.map((r) => (
                          <tr key={r.name} className="hover:bg-slate-50">
                            <td className="px-4 py-3 font-medium text-slate-800">{r.name}</td>
                            <td className="px-4 py-3 text-slate-700">{r.Casual}</td>
                            <td className="px-4 py-3 text-slate-700">{r.Annual}</td>
                            <td className="px-4 py-3 text-slate-700">{r.Sick}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  <div className="h-72 w-full" style={{ direction: 'ltr' }}>
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={leaveData} margin={{ top: 20, right: 20, left: 10, bottom: 30 }}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                        <XAxis dataKey="name" interval={0} angle={-15} textAnchor="end" height={70} />
                        <YAxis allowDecimals={false} />
                        <RechartsTooltip />
                        <Legend />
                        <Bar dataKey="Casual" fill="#f59e0b" radius={[4, 4, 0, 0]} />
                        <Bar dataKey="Annual" fill="#3b82f6" radius={[4, 4, 0, 0]} />
                        <Bar dataKey="Sick" fill="#ef4444" radius={[4, 4, 0, 0]} />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              )}
            </div>
          </>
        )}

        {/* DATA */}
        {activeTab === 'data' && (
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
            <div className="p-4 border-b border-slate-200 bg-slate-50 flex flex-col md:flex-row md:items-center md:justify-between gap-3">
              <div>
                <h3 className="font-bold text-slate-800">البيانات (بنفس ترتيب الشيت)</h3>
                <p className="text-xs text-slate-500 mt-1">
                  ✅ Date / Shift / Name … ثم في الآخر: Actual Monitoring time + Acurracy + intrenal utilization + Public Utilization
                </p>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                <span className="text-xs text-slate-500">Sort by</span>

                <select
                  className="bg-white border border-slate-300 text-slate-700 text-xs rounded-lg px-2 py-1"
                  value={dataSortKey}
                  onChange={(e) => setDataSortKey(e.target.value)}
                >
                  <option value="date">Date</option>
                  <option value="employee">Name</option>
                  <option value="actualMonitoring">Actual Monitoring time</option>
                  <option value="accuracyPct">Acurracy %</option>
                  <option value="internalPct">intrenal utilization %</option>
                  <option value="publicPct">Public Utilization %</option>
                </select>

                <select
                  className="bg-white border border-slate-300 text-slate-700 text-xs rounded-lg px-2 py-1"
                  value={dataSortDir}
                  onChange={(e) => setDataSortDir(e.target.value)}
                >
                  <option value="desc">Desc (High→Low)</option>
                  <option value="asc">Asc (Low→High)</option>
                </select>

                <span className="text-xs text-slate-500">Rows: {sortedFilteredData.length}</span>
              </div>
            </div>

            <div className="overflow-x-auto max-h-[560px]">
              <table className="w-full text-sm text-right">
                <thead className="bg-slate-50 text-slate-600 font-medium sticky top-0">
                  <tr>
                    {tableColumns.map((c) => (
                      <th key={c.key} className="px-4 py-3 border-b whitespace-nowrap">
                        {c.label}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {sortedFilteredData.map((row, idx) => (
                    <tr key={idx} className="hover:bg-blue-50/50 transition-colors">
                      {tableColumns.map((col) => {
                        if (col.key === 'actualMonitoring') {
                          return (
                            <td key={col.key} className="px-4 py-3 text-slate-700 font-semibold whitespace-nowrap">
                              {Math.round(row.actualMonitoring || 0)}
                            </td>
                          );
                        }
                        if (col.key === 'accuracyPct')
                          return (
                            <td key={col.key} className="px-4 py-3 text-slate-700 whitespace-nowrap">
                              {row.accuracyPct}%
                            </td>
                          );
                        if (col.key === 'internalPct')
                          return (
                            <td key={col.key} className="px-4 py-3 text-slate-700 whitespace-nowrap">
                              {row.internalPct}%
                            </td>
                          );
                        if (col.key === 'publicPct')
                          return (
                            <td key={col.key} className="px-4 py-3 text-slate-700 whitespace-nowrap font-semibold">
                              {row.publicPct}%
                            </td>
                          );
                        if (col.key === 'status') {
                          return (
                            <td key={col.key} className="px-4 py-3 whitespace-nowrap">
                              <span className={`px-2 py-1 rounded-full text-xs font-bold ${row.status.badge}`}>{row.status.label}</span>
                            </td>
                          );
                        }
                        if (col.key === 'taskDetails') {
                          return (
                            <td key={col.key} className="px-4 py-3 whitespace-nowrap">
                              <button
                                onClick={() => setSelectedRow(row)}
                                className="inline-flex items-center gap-2 px-3 py-1.5 rounded-lg border border-slate-200 hover:bg-slate-50 text-slate-700 text-xs font-semibold"
                              >
                                <ListChecks size={14} />
                                Details
                              </button>
                            </td>
                          );
                        }

                        const v = row[col.key];
                        return (
                          <td key={col.key} className="px-4 py-3 text-slate-700 whitespace-nowrap">
                            {v === undefined || v === null || v === '' ? '-' : String(v)}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* CALLS REPORT */}
        {activeTab === 'calls' && (
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
            <div className="p-4 border-b border-slate-200 bg-slate-50">
              <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-3">
                <div>
                  <h3 className="font-bold text-slate-800">تقرير المكالمات (Monitored By + Duration)</h3>
                  <p className="text-xs text-slate-500 mt-1">
                    ✅ Avg / Total Duration محسوبين على المكالمات الصالحة فقط (Duration &gt; 0).<br />
                    ✅ Buckets محسوبة على المكالمات الصالحة فقط، و Group Totals منفصلة: (1–3) و (4–6) و (7–10) و (&gt;10).
                  </p>
                </div>
                <div className="text-xs text-slate-500">Rows: {callsReport.length}</div>
              </div>
            </div>

            {/* Table */}
            <div className="overflow-x-auto max-h-[560px]">
              <table className="w-full text-sm text-right">
                <thead className="bg-slate-50 text-slate-600 font-medium sticky top-0">
                  <tr>
                    {/* ✅ Sticky Name header */}
                    <SortTh
                      label="Name"
                      colKey="name"
                      sortKey={callsSortKey}
                      sortDir={callsSortDir}
                      onSort={onCallsSort}
                      className="sticky right-0 z-30 bg-slate-50 border-l border-slate-200"
                    />

                    <SortTh label="Total Calls" colKey="totalCalls" sortKey={callsSortKey} sortDir={callsSortDir} onSort={onCallsSort} />

                    <SortTh label="Avg Duration (min)" colKey="avgDurationMin" sortKey={callsSortKey} sortDir={callsSortDir} onSort={onCallsSort} />
                    <SortTh label="Total Duration (min)" colKey="totalDurationMin" sortKey={callsSortKey} sortDir={callsSortDir} onSort={onCallsSort} />
                    <SortTh label="Total Duration (hours)" colKey="totalDurationHours" sortKey={callsSortKey} sortDir={callsSortDir} onSort={onCallsSort} />

                    <SortTh label={`Short Calls (≤ ${SHORT_CALL_THRESHOLD_MIN}m)`} colKey="shortCalls" sortKey={callsSortKey} sortDir={callsSortDir} onSort={onCallsSort} />
                    <SortTh label="Short %" colKey="shortPct" sortKey={callsSortKey} sortDir={callsSortDir} onSort={onCallsSort} />

                    {/* ✅ Group 1 */}
                    <SortTh label="=1m" colKey="m1" sortKey={callsSortKey} sortDir={callsSortDir} onSort={onCallsSort} />
                    <SortTh label="=2m" colKey="m2" sortKey={callsSortKey} sortDir={callsSortDir} onSort={onCallsSort} />
                    <SortTh label="=3m" colKey="m3" sortKey={callsSortKey} sortDir={callsSortDir} onSort={onCallsSort} />
                    <SortTh label="1–3 Total" colKey="g13" sortKey={callsSortKey} sortDir={callsSortDir} onSort={onCallsSort} className={thHi13} />

                    {/* ✅ Group 2 */}
                    <SortTh label="=4m" colKey="m4" sortKey={callsSortKey} sortDir={callsSortDir} onSort={onCallsSort} />
                    <SortTh label="=5m" colKey="m5" sortKey={callsSortKey} sortDir={callsSortDir} onSort={onCallsSort} />
                    <SortTh label="=6m" colKey="m6" sortKey={callsSortKey} sortDir={callsSortDir} onSort={onCallsSort} />
                    <SortTh label="4–6 Total" colKey="g46" sortKey={callsSortKey} sortDir={callsSortDir} onSort={onCallsSort} className={thHi46} />

                    {/* ✅ Group 3 */}
                    <SortTh label="=7m" colKey="m7" sortKey={callsSortKey} sortDir={callsSortDir} onSort={onCallsSort} />
                    <SortTh label="=8m" colKey="m8" sortKey={callsSortKey} sortDir={callsSortDir} onSort={onCallsSort} />
                    <SortTh label="=9m" colKey="m9" sortKey={callsSortKey} sortDir={callsSortDir} onSort={onCallsSort} />
                    <SortTh label="=10m" colKey="m10" sortKey={callsSortKey} sortDir={callsSortDir} onSort={onCallsSort} />
                    <SortTh label="7–10 Total" colKey="g710" sortKey={callsSortKey} sortDir={callsSortDir} onSort={onCallsSort} className={thHi710} />

                    {/* ✅ Last */}
                    <SortTh label=">10m" colKey="gt10" sortKey={callsSortKey} sortDir={callsSortDir} onSort={onCallsSort} />
                  </tr>
                </thead>

                <tbody className="divide-y divide-slate-100">
                  {callsReport.length === 0 ? (
                    <tr>
                      <td colSpan={21} className="px-4 py-6 text-center text-slate-500">
                        ارفع ملف المكالمات أولاً، أو غيّر الفلترة.
                      </td>
                    </tr>
                  ) : (
                    callsReport.map((r) => (
                      <tr key={r.name} className="group hover:bg-blue-50/50 transition-colors">
                        {/* ✅ Sticky Name cell */}
                        <td className="sticky right-0 z-20 bg-white border-l border-slate-200 px-4 py-3 font-semibold text-slate-800 whitespace-nowrap group-hover:bg-blue-50/50">
                          {r.name}
                        </td>

                        <td className="px-4 py-3 text-slate-700 whitespace-nowrap text-center">{r.totalCalls}</td>

                        <td className="px-4 py-3 text-slate-700 whitespace-nowrap text-center">{r.avgDurationMin}</td>
                        <td className="px-4 py-3 text-slate-700 whitespace-nowrap text-center">{r.totalDurationMin}</td>
                        <td className="px-4 py-3 text-slate-700 whitespace-nowrap text-center">{r.totalDurationHours}</td>

                        <td className="px-4 py-3 text-slate-700 whitespace-nowrap text-center">{r.shortCalls}</td>
                        <td className="px-4 py-3 text-slate-700 whitespace-nowrap text-center font-semibold">{r.shortPct}%</td>

                        {/* Group 1 */}
                        <td className="px-4 py-3 text-slate-700 whitespace-nowrap text-center">{r.m1}</td>
                        <td className="px-4 py-3 text-slate-700 whitespace-nowrap text-center">{r.m2}</td>
                        <td className="px-4 py-3 text-slate-700 whitespace-nowrap text-center">{r.m3}</td>
                        <td className={`px-4 py-3 whitespace-nowrap text-center ${tdHi13}`}>{r.g13}</td>

                        {/* Group 2 */}
                        <td className="px-4 py-3 text-slate-700 whitespace-nowrap text-center">{r.m4}</td>
                        <td className="px-4 py-3 text-slate-700 whitespace-nowrap text-center">{r.m5}</td>
                        <td className="px-4 py-3 text-slate-700 whitespace-nowrap text-center">{r.m6}</td>
                        <td className={`px-4 py-3 whitespace-nowrap text-center ${tdHi46}`}>{r.g46}</td>

                        {/* Group 3 */}
                        <td className="px-4 py-3 text-slate-700 whitespace-nowrap text-center">{r.m7}</td>
                        <td className="px-4 py-3 text-slate-700 whitespace-nowrap text-center">{r.m8}</td>
                        <td className="px-4 py-3 text-slate-700 whitespace-nowrap text-center">{r.m9}</td>
                        <td className="px-4 py-3 text-slate-700 whitespace-nowrap text-center">{r.m10}</td>
                        <td className={`px-4 py-3 whitespace-nowrap text-center ${tdHi710}`}>{r.g710}</td>

                        {/* >10 */}
                        <td className="px-4 py-3 text-slate-700 whitespace-nowrap text-center font-semibold">{r.gt10}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>

            {/* Analysis Charts */}
            {callsReport.length > 0 && (
              <div className="p-6 grid grid-cols-1 lg:grid-cols-2 gap-6 border-t border-slate-200">
                <div className="bg-white p-4 rounded-xl border border-slate-100">
                  <h4 className="font-bold text-slate-800 mb-3">Avg Duration (Min) لكل موظف</h4>
                  <div className="h-80 w-full" style={{ direction: 'ltr' }}>
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={callsAvgChart} margin={{ top: 20, right: 20, left: 10, bottom: 60 }}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                        <XAxis dataKey="name" interval={0} angle={-20} textAnchor="end" height={70} />
                        <YAxis />
                        <RechartsTooltip formatter={(v) => [`${v} min`, 'Avg']} />
                        <Bar dataKey="avg" fill="#3b82f6" radius={[4, 4, 0, 0]}>
                          <LabelList dataKey="avg" position="top" />
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                <div className="bg-white p-4 rounded-xl border border-slate-100">
                  <h4 className="font-bold text-slate-800 mb-3">% Short Calls (≤ 2m) لكل موظف</h4>
                  <div className="h-80 w-full" style={{ direction: 'ltr' }}>
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={callsShortPctChart} margin={{ top: 20, right: 20, left: 10, bottom: 60 }}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                        <XAxis dataKey="name" interval={0} angle={-20} textAnchor="end" height={70} />
                        <YAxis />
                        <RechartsTooltip formatter={(v) => [`${v}%`, 'Short %']} />
                        <Bar dataKey="shortPct" fill="#ef4444" radius={[4, 4, 0, 0]}>
                          <LabelList dataKey="shortPct" position="top" formatter={(v) => `${v}%`} />
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                <div className="lg:col-span-2 bg-slate-50 border border-slate-200 rounded-xl p-4 text-sm text-slate-700">
                  <ul className="list-disc pr-5 space-y-1">
                  </ul>
                </div>
              </div>
            )}
          </div>
        )}
      </div>

      {/* Task Details Modal (Utilization) */}
      <Modal
        open={!!selectedRow}
        title={selectedRow ? `Task Details - ${selectedRow.employee || ''} (${selectedRow.date || ''})` : ''}
        onClose={() => setSelectedRow(null)}
      >
        {selectedRow &&
          (() => {
            const td = selectedRow.taskDetails || {};
            const safe = (k) => td[k] || { count: 0, minutes: 0 };

            return (
              <div className="space-y-4">
                <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                  <div className="bg-slate-50 border border-slate-200 rounded-xl p-3">
                    <p className="text-xs text-slate-500">Shift</p>
                    <p className="font-semibold text-slate-800">{selectedRow.shift || '-'}</p>
                  </div>
                  <div className="bg-slate-50 border border-slate-200 rounded-xl p-3">
                    <p className="text-xs text-slate-500">Monitoring % / Tasks %</p>
                    <p className="font-semibold text-slate-800">
                      {selectedRow.monitoringPct}% / {selectedRow.tasksPct}%
                    </p>
                  </div>
                  <div className="bg-slate-50 border border-slate-200 rounded-xl p-3">
                    <p className="text-xs text-slate-500">Public Utilization</p>
                    <p className="font-semibold text-slate-800">
                      {selectedRow.publicPct}% <span className="text-xs text-slate-500">({selectedRow.status.label})</span>
                    </p>
                  </div>
                </div>

                <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
                  <div className="px-4 py-3 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
                    <h4 className="font-bold text-slate-800">Monitoring Tasks (Counts + Minutes)</h4>
                    <span className="text-xs text-slate-500">Actual Monitoring time: {Math.round(selectedRow.actualMonitoring || 0)} min</span>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-sm text-right">
                      <thead className="bg-white text-slate-600 font-medium">
                        <tr>
                          <th className="px-4 py-3 border-b">Task</th>
                          <th className="px-4 py-3 border-b">Count</th>
                          <th className="px-4 py-3 border-b">Minutes</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {[
                          ['Voice - WhatsApp', safe('voice_whatsapp')],
                          ['ADB / Retention / OMU / Social', safe('adb_retention_omu')],
                          ['VIP', safe('vip')],
                          ['Complaint', safe('complaint')],
                          ['Application', safe('application')],
                          ['Mysetry Calls', safe('mystery_calls')],
                          ['ADB Transaction', safe('adb_transaction')],
                          ['Coaching', safe('coaching')],
                        ].map(([label, obj]) => (
                          <tr key={label}>
                            <td className="px-4 py-3 text-slate-700">{label}</td>
                            <td className="px-4 py-3 text-slate-700">{obj.count || 0}</td>
                            <td className="px-4 py-3 font-semibold text-slate-800">{Math.round(obj.minutes || 0)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                  <div className="bg-white border border-slate-200 rounded-xl p-4">
                    <h4 className="font-bold text-slate-800 mb-2">Other Tasks</h4>
                    <p className="text-sm text-slate-700">
                      Minutes: <span className="font-semibold">{Math.round(selectedRow.actualTasks || 0)}</span>
                    </p>
                    <p className="text-sm text-slate-700 mt-2">
                      Comment: <span className="font-semibold">{selectedRow.otherTasksComment || '-'}</span>
                    </p>
                  </div>
                  <div className="bg-white border border-slate-200 rounded-xl p-4">
                    <h4 className="font-bold text-slate-800 mb-2">Calculated Summary</h4>
                    <div className="text-sm text-slate-700 space-y-1">
                      <div>
                        Overall Actual time: <span className="font-semibold">{Math.round(selectedRow.overallActual || 0)} min</span>
                      </div>
                      <div>
                        Acurracy: <span className="font-semibold">{selectedRow.accuracyPct}%</span>
                      </div>
                      <div>
                        intrenal utilization: <span className="font-semibold">{selectedRow.internalPct}%</span>
                      </div>
                      <div>
                        Public Utilization: <span className="font-semibold">{selectedRow.publicPct}%</span>
                      </div>
                      <div className="text-xs text-slate-500 pt-2">
                        Paid day: {Math.round(selectedRow.paidDay || 0)} • Business day: {Math.round(selectedRow.businessDay || 0)} • Ad-Hoc: {Math.round(selectedRow.adhoc || 0)}
                      </div>
                    </div>
                  </div>
                </div>

                {selectedRow.leave && (
                  <div className="bg-amber-50 border border-amber-200 text-amber-800 rounded-xl p-4 flex items-center gap-2">
                    <CheckCircle size={18} />
                    <span className="text-sm">هذا الصف مُصنّف كـ Leave/Off حسب قيمة Shift، لذلك تم اعتبار الأوقات = 0.</span>
                  </div>
                )}
              </div>
            );
          })()}
      </Modal>
    </div>
  );
}