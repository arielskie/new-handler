import React, { useState, useEffect, useRef } from 'react';

// Power Automate Endpoints
const READ_URL = "https://defaulteaa689b48f8740e09c6f7228de4d75.4a.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/33cda42ea60b487ea12bc6d6b2fb094d/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=-W97iRQOOSQ0lRc9j2WefBnqQegAXIPRvQHX1I2bNOU";
const WRITE_URL = "https://defaulteaa689b48f8740e09c6f7228de4d75.4a.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/60eeb9c1879249f79dea375e12bfb527/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=aFIZCX9EK8H4QTmltyJtonb_mLDhEi27kQHK_EoglAw";
const DELETE_URL = "https://defaulteaa689b48f8740e09c6f7228de4d75.4a.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/b4e3ca791183425e99995266ab810738/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Vy2Jtiuem6SIKGtMW2bSujFxI-9HbF18hq2rQ8v8KoU";

const EXCEL_SHAREPOINT_LINK = "https://analog-my.sharepoint.com/:x:/r/personal/ariel_mentawan_analog_com/_layouts/15/Doc.aspx?sourcedoc=%7B3febd084-a91a-4d93-814e-e86736350784%7D&action=edit";

const ADMIN_USER = "admin";
const ADMIN_PASS = "admin123";

interface RepairRecord {
  ID?: string;
  WW?: string;
  DATE?: string;
  TESTER_ID?: string;
  HANDLER_ID?: string;
  DETAILS?: string;
  DEVICE?: string;
  PACKAGE?: string;
  MajorPart?: string;
  PROBLEM?: string;
  ROOT_CAUSE?: string;
  MATERIAL_CODE?: string;
  StockItems?: string;
  PENDING_ACTION?: string;
  STATUS?: string;
  [key: string]: any;
}

interface SparePart {
  id: string;
  name: string;
  code: string;
  handler: string;
  quantity: number;
  image?: string;
}

const DEFAULT_SPARES: SparePart[] = [
  {
    id: 'sp-1',
    name: 'Test Socket Pin Block Assembly',
    code: 'SP-MT99-042',
    handler: 'MT99',
    quantity: 18,
    image: 'https://images.unsplash.com/photo-1581092160607-ee22621dd758?w=300&auto=format&fit=crop&q=80'
  },
  {
    id: 'sp-2',
    name: 'High-Precision Pick & Place Rubber Nozzle',
    code: 'SP-MT93-108',
    handler: 'MT93',
    quantity: 5,
    image: 'https://images.unsplash.com/photo-1581092335397-9583fe92d232?w=300&auto=format&fit=crop&q=80'
  },
  {
    id: 'sp-3',
    name: 'High Vacuum Suction Cup Set (Pack of 10)',
    code: 'SP-MT-COM-011',
    handler: 'MT99, MT93',
    quantity: 42,
    image: 'https://images.unsplash.com/photo-1581092580497-e0d23cbdf1dc?w=300&auto=format&fit=crop&q=80'
  },
  {
    id: 'sp-4',
    name: 'Optical Position Sensor Module',
    code: 'SP-MT99-SEN-08',
    handler: 'MT99',
    quantity: 2,
    image: 'https://images.unsplash.com/photo-1581091226825-a6a2a5aee158?w=300&auto=format&fit=crop&q=80'
  },
  {
    id: 'sp-5',
    name: 'Thermal Chuck Heater Cartridge 24V',
    code: 'SP-MT93-TH-99',
    handler: 'MT93',
    quantity: 0,
    image: 'https://images.unsplash.com/photo-1581092162384-8987c1d64718?w=300&auto=format&fit=crop&q=80'
  }
];

export default function App() {
  const [theme, setTheme] = useState<'dark' | 'light'>(() => {
    return (localStorage.getItem('user_theme_pref') as 'dark' | 'light') || 'dark';
  });
  const [density, setDensity] = useState<'compact' | 'normal'>(() => {
    return (localStorage.getItem('user_density_pref') as 'compact' | 'normal') || 'normal';
  });
  const [isCollapsed, setIsCollapsed] = useState<boolean>(() => {
    return localStorage.getItem('sidebar_collapsed_pref') === 'true';
  });

  const [isLoggedIn, setIsLoggedIn] = useState<boolean>(() => {
    return sessionStorage.getItem('app_user_logged_in') === 'true';
  });
  const [showLoginModal, setShowLoginModal] = useState(false);
  const [loginUsername, setLoginUsername] = useState('');
  const [loginPassword, setLoginPassword] = useState('');

  const [activeTab, setActiveTab] = useState<'repair' | 'spareparts'>('repair');
  const [statusFilter, setStatusFilter] = useState<string>('');
  const [spareFilter, setSpareFilter] = useState<string>('');
  const [searchQuery, setSearchQuery] = useState<string>('');

  const [rawData, setRawData] = useState<RepairRecord[]>([]);
  const [spareData, setSpareData] = useState<SparePart[]>(() => {
    try {
      const saved = localStorage.getItem('spare_parts_inventory_data');
      return saved ? JSON.parse(saved) : DEFAULT_SPARES;
    } catch {
      return DEFAULT_SPARES;
    }
  });

  const [selectedRows, setSelectedRows] = useState<Set<number>>(new Set());
  const [toastMessage, setToastMessage] = useState<string | null>(null);
  const [lightboxImg, setLightboxImg] = useState<string | null>(null);
  const [showGithubModal, setShowGithubModal] = useState<boolean>(false);
  const [copiedCmd, setCopiedCmd] = useState<boolean>(false);

  // Drawers & Modals
  const [showRepairDrawer, setShowRepairDrawer] = useState(false);
  const [editingRepairIndex, setEditingRepairIndex] = useState<number | null>(null);
  const [repairForm, setRepairForm] = useState({
    id: '',
    ww: '',
    date: '',
    testerId: '',
    handlerId: '',
    details: '',
    device: '',
    package: '',
    majorPart: '',
    problem: '',
    rootCause: '',
    materialCode: '',
    stockItems: 'Stock Items',
    pendingAction: '',
    status: 'Active'
  });

  const [showSpareDrawer, setShowSpareDrawer] = useState(false);
  const [editingSpareId, setEditingSpareId] = useState<string | null>(null);
  const [spareForm, setSpareForm] = useState({
    name: '',
    code: '',
    handler: 'MT99, MT93',
    quantity: 10,
    image: ''
  });

  const [showStockModal, setShowStockModal] = useState(false);
  const [stockAdjustId, setStockAdjustId] = useState<string | null>(null);
  const [stockAdjustType, setStockAdjustType] = useState<'add' | 'withdraw'>('add');
  const [stockAdjustQty, setStockAdjustQty] = useState<number>(1);
  const [stockAdjustReason, setStockAdjustReason] = useState<string>('');

  const tableWrapRef = useRef<HTMLDivElement>(null);

  // Sync theme, density, and auth classes on body
  useEffect(() => {
    document.body.classList.remove('theme-dark', 'theme-light');
    document.body.classList.add(`theme-${theme}`);
    localStorage.setItem('user_theme_pref', theme);
  }, [theme]);

  useEffect(() => {
    document.body.classList.remove('density-compact', 'density-normal');
    document.body.classList.add(`density-${density}`);
    localStorage.setItem('user_density_pref', density);
  }, [density]);

  useEffect(() => {
    if (isLoggedIn) {
      document.body.classList.remove('is-readonly');
      document.body.classList.add('is-admin');
    } else {
      document.body.classList.remove('is-admin');
      document.body.classList.add('is-readonly');
    }
  }, [isLoggedIn]);

  useEffect(() => {
    localStorage.setItem('sidebar_collapsed_pref', isCollapsed ? 'true' : 'false');
  }, [isCollapsed]);

  useEffect(() => {
    try {
      localStorage.setItem('spare_parts_inventory_data', JSON.stringify(spareData));
    } catch (e) {
      console.error(e);
    }
  }, [spareData]);

  const showToast = (msg: string) => {
    setToastMessage(msg);
    setTimeout(() => {
      setToastMessage((c) => (c === msg ? null : c));
    }, 3200);
  };

  const sanitizeValue = (v: any): string => {
    if (v === undefined || v === null) return '';
    const str = String(v).trim();
    if (
      str.toUpperCase().includes('#BLOCKED') ||
      str.startsWith('#REF!') ||
      str.startsWith('#VALUE!') ||
      str.startsWith('#N/A') ||
      str.startsWith('#NAME?')
    ) {
      return '';
    }
    return str;
  };

  const getRowVal = (row: any, key: string): string => {
    if (!row) return '';
    if (row[key] !== undefined && row[key] !== null) {
      return sanitizeValue(row[key]);
    }
    const aliases: Record<string, string[]> = {
      WW: ['WW', 'ww'],
      DATE: ['DATE', 'Date'],
      'TESTER ID': ['TESTER ID', 'TESTER_ID', 'TESTER_x0020_ID', 'Tester ID'],
      'HANDLER ID': ['HANDLER ID', 'HANDLER_ID', 'HANDLER_x0020_ID', 'Handler ID', 'EventName'],
      DETAILS: [
        'DETAILS (Image as Evidence)',
        'DETAILS\n(Image as Evidence)',
        'Details (Image as Evidence)',
        'DETAILS',
        'Details'
      ],
      DEVICE: ['DEVICE', 'Device', 'device'],
      PACKAGE: ['PACKAGE', 'Package', 'package'],
      'Handler Major Part': ['Handler Major Part', 'HANDLER MAJOR PART', 'Handler_x0020_Major_x0020_Part', 'MajorPart'],
      PROBLEM: ['PROBLEM', 'Problem', 'problem'],
      'ROOT CAUSE': ['ROOT CAUSE', 'ROOT CASUE', 'ROOT_x0020_CAUSE', 'Root Cause'],
      'MATERIAL CODE': ['MATERIAL CODE', 'Material Code', 'MATERIAL_CODE'],
      'STOCK ITEMS': ['Stock Items\nNon Stock Items', 'Stock Items / Non Stock Items', 'Stock Items', 'StockItems'],
      'PENDING ACTION': ['PENDING ACTION', 'Pending Action', 'PENDING_ACTION'],
      STATUS: ['STATUS', 'Status', 'status']
    };

    const checks = aliases[key] || [key];
    for (const k of checks) {
      if (row[k] !== undefined && row[k] !== null) {
        const v = sanitizeValue(row[k]);
        if (v !== '') return v;
      }
    }

    const cleanTarget = key.toLowerCase().replace(/[^a-z0-9]/g, '');
    for (const actualKey of Object.keys(row)) {
      const cleanActual = actualKey
        .toLowerCase()
        .replace(/_x[0-9a-f]{4}_/gi, '')
        .replace(/[^a-z0-9]/g, '');
      if (cleanActual === cleanTarget && row[actualKey] !== undefined && row[actualKey] !== null) {
        return sanitizeValue(row[actualKey]);
      }
    }
    return '';
  };

  const getRowId = (row: any): string => {
    if (!row) return '';
    if (row['ID']) return String(row['ID']).trim();
    if (row['id']) return String(row['id']).trim();
    if (row['__PowerAppsId__']) return String(row['__PowerAppsId__']);
    if (row['itemInternalId']) return String(row['itemInternalId']);
    return '';
  };

  const loadTableData = async () => {
    showToast("Fetching records from database...");
    try {
      const response = await fetch(READ_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({})
      });
      const textData = await response.text();
      let data = JSON.parse(textData);
      let list = Array.isArray(data) ? data : data.value || (data.body && data.body.value) || [];

      const filtered = list.filter((r: any) => {
        const str = Object.values(r).join('').trim();
        return str.length > 0 && str !== 'null';
      });

      setRawData(filtered);
      showToast(`Loaded ${filtered.length} active rows`);
    } catch (err) {
      console.error(err);
      showToast("Error loading Excel records");
    }
  };

  useEffect(() => {
    loadTableData();
  }, []);

  const compressImage = (file: File, maxWidth = 600, quality = 0.6): Promise<string> => {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = (event) => {
        const img = new Image();
        img.src = event.target?.result as string;
        img.onload = () => {
          let width = img.width;
          let height = img.height;
          if (width > maxWidth) {
            height = Math.round((height * maxWidth) / width);
            width = maxWidth;
          }
          const canvas = document.createElement('canvas');
          canvas.width = width;
          canvas.height = height;
          const ctx = canvas.getContext('2d');
          ctx?.drawImage(img, 0, 0, width, height);
          resolve(canvas.toDataURL('image/jpeg', quality));
        };
      };
    });
  };

  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    if (loginUsername === ADMIN_USER && loginPassword === ADMIN_PASS) {
      setIsLoggedIn(true);
      sessionStorage.setItem('app_user_logged_in', 'true');
      setShowLoginModal(false);
      setLoginUsername('');
      setLoginPassword('');
      showToast("Logged in successfully!");
    } else {
      alert("Invalid username or password");
    }
  };

  const handleLogout = () => {
    setIsLoggedIn(false);
    sessionStorage.removeItem('app_user_logged_in');
    setSelectedRows(new Set());
    showToast("Logged out");
  };

  // Filtered repair records
  const filteredRepairs = rawData.filter((row) => {
    const status = getRowVal(row, 'STATUS');
    const matchStatus = !statusFilter || status === statusFilter;
    if (!matchStatus) return false;

    if (searchQuery) {
      const q = searchQuery.toLowerCase();
      const match = Object.values(row).some((val) =>
        String(val).toLowerCase().includes(q)
      );
      if (!match) return false;
    }
    return true;
  });

  // Filtered spare parts
  const filteredSpares = spareData.filter((part) => {
    if (spareFilter === 'MT99' && !(part.handler || '').toUpperCase().includes('MT99')) return false;
    if (spareFilter === 'MT93' && !(part.handler || '').toUpperCase().includes('MT93')) return false;
    if (spareFilter === 'low' && Number(part.quantity || 0) > 5) return false;

    if (searchQuery) {
      const q = searchQuery.toLowerCase();
      const combined = `${part.name} ${part.code} ${part.handler} ${part.quantity}`.toLowerCase();
      if (!combined.includes(q)) return false;
    }
    return true;
  });

  // Counters
  const countTotal = rawData.length;
  const countActive = rawData.filter((r) => getRowVal(r, 'STATUS') === 'Active').length;
  const countPending = rawData.filter((r) => getRowVal(r, 'STATUS') === 'Pending').length;
  const countResolved = rawData.filter((r) => getRowVal(r, 'STATUS') === 'Resolved').length;

  const countSpareTotal = spareData.length;
  const countSpareMT99 = spareData.filter((s) => (s.handler || '').toUpperCase().includes('MT99')).length;
  const countSpareMT93 = spareData.filter((s) => (s.handler || '').toUpperCase().includes('MT93')).length;
  const countSpareLow = spareData.filter((s) => Number(s.quantity || 0) <= 5).length;

  // Single / Bulk Delete
  const deleteSingleRow = async (idx: number) => {
    if (!isLoggedIn) {
      setShowLoginModal(true);
      return;
    }
    const row = rawData[idx];
    const rId = getRowId(row);
    const hndId = getRowVal(row, 'HANDLER ID');

    if (!confirm(`Permanently delete row for Handler "${hndId || rId}" from Excel?`)) return;
    if (!rId) {
      alert("Could not find a valid ID for this row.");
      return;
    }

    showToast(`Deleting row from Excel...`);
    try {
      const response = await fetch(DELETE_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ROW_ID: String(rId) })
      });

      if (!response.ok) {
        const err = await response.text();
        alert(`Power Automate returned an error deleting ID ${rId}:\n${err}`);
        return;
      }

      showToast("Row deleted from Excel successfully!");
      setRawData((prev) => prev.filter((_, i) => i !== idx));
      setTimeout(loadTableData, 2000);
    } catch (err) {
      console.error(err);
      alert("Network error: Could not reach Delete Flow.");
    }
  };

  const deleteSelectedRows = async () => {
    if (!isLoggedIn || selectedRows.size === 0) return;
    const count = selectedRows.size;
    if (!confirm(`Permanently delete ${count} row(s) directly from Excel?`)) return;

    const idsToDelete: string[] = [];
    selectedRows.forEach((idx) => {
      const rId = getRowId(rawData[idx]);
      if (rId) idsToDelete.push(rId);
    });

    if (idsToDelete.length === 0) {
      alert("Could not find a valid ID for the selected row(s).");
      return;
    }

    showToast(`Deleting ${idsToDelete.length} row(s)...`);
    try {
      for (let rId of idsToDelete) {
        await fetch(DELETE_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ ROW_ID: String(rId) })
        });
      }
      showToast("Deleted successfully! Updating...");
      setSelectedRows(new Set());
      setTimeout(loadTableData, 2000);
    } catch (err) {
      console.error(err);
      alert("Network error: Could not reach Delete Flow.");
    }
  };

  const openAddRowDrawer = () => {
    if (!isLoggedIn) {
      setShowLoginModal(true);
      return;
    }
    setEditingRepairIndex(null);
    setRepairForm({
      id: '',
      ww: 'WW' + new Date().getFullYear().toString().slice(-2) + '01',
      date: new Date().toISOString().split('T')[0],
      testerId: '',
      handlerId: '',
      details: '',
      device: '',
      package: '',
      majorPart: '',
      problem: '',
      rootCause: '',
      materialCode: '',
      stockItems: 'Stock Items',
      pendingAction: '',
      status: 'Active'
    });
    setShowRepairDrawer(true);
  };

  const openEditDrawer = (idx: number) => {
    if (!isLoggedIn) {
      setShowLoginModal(true);
      return;
    }
    const row = rawData[idx];
    setEditingRepairIndex(idx);
    setRepairForm({
      id: getRowId(row),
      ww: getRowVal(row, 'WW'),
      date: getRowVal(row, 'DATE') || new Date().toISOString().split('T')[0],
      testerId: getRowVal(row, 'TESTER ID'),
      handlerId: getRowVal(row, 'HANDLER ID'),
      details: getRowVal(row, 'DETAILS'),
      device: getRowVal(row, 'DEVICE'),
      package: getRowVal(row, 'PACKAGE'),
      majorPart: getRowVal(row, 'Handler Major Part'),
      problem: getRowVal(row, 'PROBLEM'),
      rootCause: getRowVal(row, 'ROOT CAUSE'),
      materialCode: getRowVal(row, 'MATERIAL CODE'),
      stockItems: getRowVal(row, 'STOCK ITEMS') || 'Stock Items',
      pendingAction: getRowVal(row, 'PENDING ACTION'),
      status: getRowVal(row, 'STATUS') || 'Active'
    });
    setShowRepairDrawer(true);
  };

  const submitDrawerRecord = async (e: React.FormEvent) => {
    e.preventDefault();
    const editId = repairForm.id;
    const uniqueId = editId ? editId : 'REC-' + Date.now();

    const recordPayload = {
      ID: uniqueId,
      WW: repairForm.ww.trim(),
      DATE: repairForm.date,
      TESTER_ID: repairForm.testerId.trim(),
      HANDLER_ID: repairForm.handlerId.trim(),
      DETAILS: repairForm.details,
      DEVICE: repairForm.device.trim(),
      PACKAGE: repairForm.package.trim(),
      MajorPart: repairForm.majorPart.trim(),
      PROBLEM: repairForm.problem.trim(),
      ROOT_CAUSE: repairForm.rootCause.trim(),
      MATERIAL_CODE: repairForm.materialCode.trim(),
      StockItems: repairForm.stockItems,
      PENDING_ACTION: repairForm.pendingAction.trim(),
      STATUS: repairForm.status
    };

    showToast("Syncing with Excel Online...");
    try {
      const response = await fetch(WRITE_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(recordPayload)
      });

      if (response.ok) {
        showToast(editId ? "Record updated! Reloading..." : "Record added! Reloading...");
        setShowRepairDrawer(false);
        setTimeout(loadTableData, 2000);
      } else {
        const err = await response.text();
        alert("Power Automate Error on Save:\n" + err);
      }
    } catch (err) {
      console.error(err);
      alert("Network error: Could not reach Write Flow.");
    }
  };

  const openAddSpareDrawer = () => {
    if (!isLoggedIn) {
      setShowLoginModal(true);
      return;
    }
    setEditingSpareId(null);
    setSpareForm({
      name: '',
      code: '',
      handler: 'MT99, MT93',
      quantity: 10,
      image: ''
    });
    setShowSpareDrawer(true);
  };

  const openEditSpareDrawer = (id: string) => {
    if (!isLoggedIn) {
      setShowLoginModal(true);
      return;
    }
    const part = spareData.find((s) => s.id === id);
    if (!part) return;
    setEditingSpareId(id);
    setSpareForm({
      name: part.name || '',
      code: part.code || '',
      handler: part.handler || '',
      quantity: part.quantity || 0,
      image: part.image || ''
    });
    setShowSpareDrawer(true);
  };

  const handleSpareFormSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (editingSpareId) {
      setSpareData((prev) =>
        prev.map((s) => (s.id === editingSpareId ? { ...s, ...spareForm } : s))
      );
      showToast(`Updated "${spareForm.name}"`);
    } else {
      const newPart = {
        id: 'sp-' + Date.now(),
        ...spareForm
      };
      setSpareData((prev) => [newPart, ...prev]);
      showToast(`Added "${spareForm.name}" to inventory`);
    }
    setShowSpareDrawer(false);
  };

  const deleteSparePart = (id: string) => {
    if (!isLoggedIn) return;
    const part = spareData.find((s) => s.id === id);
    if (!confirm(`Are you sure you want to delete "${part?.name || 'this part'}" from inventory?`)) return;
    setSpareData((prev) => prev.filter((s) => s.id !== id));
    showToast("Spare part deleted");
  };

  const openStockAdjustModal = (id: string, actionType: 'add' | 'withdraw') => {
    setStockAdjustId(id);
    setStockAdjustType(actionType);
    setStockAdjustQty(1);
    setStockAdjustReason('');
    setShowStockModal(true);
  };

  const handleStockAdjustSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!stockAdjustId) return;
    const part = spareData.find((s) => s.id === stockAdjustId);
    if (!part) return;

    let currentQty = Number(part.quantity || 0);
    if (stockAdjustType === 'withdraw') {
      if (stockAdjustQty > currentQty) {
        alert(`Cannot withdraw ${stockAdjustQty} units. Only ${currentQty} units currently in stock.`);
        return;
      }
      setSpareData((prev) =>
        prev.map((s) =>
          s.id === stockAdjustId ? { ...s, quantity: currentQty - stockAdjustQty } : s
        )
      );
      showToast(`Withdrew ${stockAdjustQty} units of ${part.code}`);
    } else {
      setSpareData((prev) =>
        prev.map((s) =>
          s.id === stockAdjustId ? { ...s, quantity: currentQty + stockAdjustQty } : s
        )
      );
      showToast(`Added ${stockAdjustQty} units of ${part.code}`);
    }
    setShowStockModal(false);
  };

  const exportCSV = () => {
    const cols = [
      'WW', 'DATE', 'TESTER ID', 'HANDLER ID', 'DETAILS', 'DEVICE', 'PACKAGE',
      'Handler Major Part', 'PROBLEM', 'ROOT CAUSE', 'MATERIAL CODE', 'Stock Items', 'PENDING ACTION', 'STATUS'
    ];
    let csv = cols.map((c) => `"${c}"`).join(',') + '\n';
    rawData.forEach((row) => {
      csv += cols.map((c) => `"${(getRowVal(row, c) || '').toString().replace(/"/g, '""')}"`).join(',') + '\n';
    });
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'handler_monitoring_export.csv';
    a.click();
  };

  const exportSpareCSV = () => {
    let csv = '"Name","Product Code","Handler Type","Stock Quantity"\n';
    spareData.forEach((p) => {
      csv += `"${(p.name || '').replace(/"/g, '""')}","${(p.code || '').replace(/"/g, '""')}","${(p.handler || '').replace(/"/g, '""')}","${p.quantity || 0}"\n`;
    });
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'spare_parts_inventory_export.csv';
    a.click();
  };

  const toggleSelectAll = (checked: boolean) => {
    if (!isLoggedIn) return;
    if (checked) {
      setSelectedRows(new Set(filteredRepairs.map((_, i) => i)));
    } else {
      setSelectedRows(new Set());
    }
  };

  const toggleRowSelect = (idx: number) => {
    if (!isLoggedIn) return;
    setSelectedRows((prev) => {
      const next = new Set(prev);
      if (next.has(idx)) next.delete(idx);
      else next.add(idx);
      return next;
    });
  };

  return (
    <div className="app-layout">
      {/* Ambient background shapes */}
      <div className="ambient-shape shape-1"></div>
      <div className="ambient-shape shape-2"></div>
      <div className="ambient-shape shape-3"></div>
      <div className="ambient-shape shape-4"></div>

      {/* ==================================================== */}
      {/* LEFT SIDEBAR                                         */}
      {/* ==================================================== */}
      <aside className={`app-sidebar ${isCollapsed ? 'collapsed' : ''}`}>
        <div className="sidebar-top">
          {/* Brand */}
          <div className="sidebar-brand-row">
            <div className="brand-group" onClick={() => setIsCollapsed(!isCollapsed)} title="Toggle Sidebar">
              <div className="brand-icon">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                  <rect x="3" y="3" width="18" height="18" rx="2" />
                  <path d="M3 9h18" />
                  <path d="M3 15h18" />
                  <path d="M9 3v18" />
                  <path d="M15 3v18" />
                </svg>
              </div>
              <div className="brand-text">
                <h2>Handler Monitoring</h2>
                <span>Intest 2</span>
              </div>
            </div>
          </div>

          {/* Quick Sidebar Search */}
          <div className="sidebar-search">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <circle cx="11" cy="11" r="8"></circle>
              <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
            </svg>
            <input
              type="text"
              placeholder={activeTab === 'repair' ? "Quick search repair..." : "Quick search spares..."}
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
            />
          </div>

          {/* Category Switcher Tabs */}
          <div className="sidebar-tabs-row">
            <button
              type="button"
              className={`sidebar-tab-btn ${activeTab === 'repair' ? 'active' : ''}`}
              onClick={() => { setActiveTab('repair'); setSelectedRows(new Set()); }}
            >
              Handler Repair
            </button>
            <button
              type="button"
              className={`sidebar-tab-btn ${activeTab === 'spareparts' ? 'active' : ''}`}
              onClick={() => { setActiveTab('spareparts'); setSelectedRows(new Set()); }}
            >
              Spare Parts
            </button>
          </div>

          {/* Handler Repair Navigation */}
          {activeTab === 'repair' ? (
            <nav className="sidebar-nav">
              <div className="nav-section-label">Main Menu</div>

              <button
                type="button"
                className={`nav-item ${statusFilter === '' ? 'active' : ''}`}
                onClick={() => setStatusFilter('')}
              >
                <div className="nav-item-left">
                  <div className="nav-icon">
                    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                      <ellipse cx="12" cy="5" rx="9" ry="3"></ellipse>
                      <path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"></path>
                      <path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"></path>
                    </svg>
                  </div>
                  <span>All Records</span>
                </div>
                <span className="nav-badge nav-badge-pill">{countTotal}</span>
              </button>

              <button
                type="button"
                className={`nav-item ${statusFilter === 'Active' ? 'active' : ''}`}
                onClick={() => setStatusFilter('Active')}
              >
                <div className="nav-item-left">
                  <div className="nav-icon">
                    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#ef4444" strokeWidth="2"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>
                  </div>
                  <span>Active Issues</span>
                </div>
                <span className="nav-badge nav-badge-active">{countActive}</span>
              </button>

              <button
                type="button"
                className={`nav-item ${statusFilter === 'Pending' ? 'active' : ''}`}
                onClick={() => setStatusFilter('Pending')}
              >
                <div className="nav-item-left">
                  <div className="nav-icon">
                    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#f59e0b" strokeWidth="2"><circle cx="12" cy="12" r="10"></circle><polyline points="12 6 12 12 16 14"></polyline></svg>
                  </div>
                  <span>Pending Review</span>
                </div>
                <span className="nav-badge nav-badge-pending">{countPending}</span>
              </button>

              <button
                type="button"
                className={`nav-item ${statusFilter === 'Resolved' ? 'active' : ''}`}
                onClick={() => setStatusFilter('Resolved')}
              >
                <div className="nav-item-left">
                  <div className="nav-icon">
                    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#10b981" strokeWidth="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>
                  </div>
                  <span>Resolved</span>
                </div>
                <span className="nav-badge nav-badge-resolved">{countResolved}</span>
              </button>

              {isLoggedIn && (
                <>
                  <div className="nav-section-label" style={{ marginTop: '10px' }}>Quick Actions</div>
                  <button type="button" className="nav-item" onClick={exportCSV}>
                    <div className="nav-item-left">
                      <div className="nav-icon">
                        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                      </div>
                      <span>Export CSV</span>
                    </div>
                  </button>
                </>
              )}
            </nav>
          ) : (
            /* Spare Parts Navigation */
            <nav className="sidebar-nav">
              <div className="nav-section-label">Inventory Filter</div>

              <button
                type="button"
                className={`nav-item ${spareFilter === '' ? 'active' : ''}`}
                onClick={() => setSpareFilter('')}
              >
                <div className="nav-item-left">
                  <div className="nav-icon">
                    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                      <rect x="2" y="7" width="20" height="14" rx="2" ry="2"></rect>
                      <path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"></path>
                    </svg>
                  </div>
                  <span>All Inventory</span>
                </div>
                <span className="nav-badge nav-badge-pill">{countSpareTotal}</span>
              </button>

              <button
                type="button"
                className={`nav-item ${spareFilter === 'MT99' ? 'active' : ''}`}
                onClick={() => setSpareFilter('MT99')}
              >
                <div className="nav-item-left">
                  <div className="nav-icon">
                    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="11" width="18" height="10" rx="2" /><circle cx="12" cy="5" r="2" /><path d="M12 7v4" /></svg>
                  </div>
                  <span>MT99 Spares</span>
                </div>
                <span className="nav-badge nav-badge-num">{countSpareMT99}</span>
              </button>

              <button
                type="button"
                className={`nav-item ${spareFilter === 'MT93' ? 'active' : ''}`}
                onClick={() => setSpareFilter('MT93')}
              >
                <div className="nav-item-left">
                  <div className="nav-icon">
                    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="11" width="18" height="10" rx="2" /><circle cx="12" cy="5" r="2" /><path d="M12 7v4" /></svg>
                  </div>
                  <span>MT93 Spares</span>
                </div>
                <span className="nav-badge nav-badge-num">{countSpareMT93}</span>
              </button>

              <button
                type="button"
                className={`nav-item ${spareFilter === 'low' ? 'active' : ''}`}
                onClick={() => setSpareFilter('low')}
              >
                <div className="nav-item-left">
                  <div className="nav-icon">
                    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#f59e0b" strokeWidth="2"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>
                  </div>
                  <span>Low / Out of Stock</span>
                </div>
                <span className="nav-badge nav-badge-active">{countSpareLow}</span>
              </button>

              {isLoggedIn && (
                <>
                  <div className="nav-section-label" style={{ marginTop: '10px' }}>Quick Actions</div>
                  <button type="button" className="nav-item" onClick={exportSpareCSV}>
                    <div className="nav-item-left">
                      <div className="nav-icon">
                        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                      </div>
                      <span>Export Inventory</span>
                    </div>
                  </button>
                </>
              )}
            </nav>
          )}
        </div>

        {/* Sidebar Bottom: Theme Toggle & Login/Logout Action */}
        <div className="sidebar-bottom">
          <div className="user-profile-card">
            <div
              className="sidebar-theme-toggle"
              onClick={() => setTheme(theme === 'dark' ? 'light' : 'dark')}
              title="Switch Light / Dark theme"
            >
              <div className={`theme-toggle-dot ${theme === 'light' ? 'active' : ''}`} id="dotLight" title="Light Mode">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="4" /><path d="M12 2v2" /><path d="M12 20v2" /><path d="m4.93 4.93 1.41 1.41" /><path d="m17.66 17.66 1.41 1.41" /><path d="M2 12h2" /><path d="M20 12h2" /><path d="m6.34 17.66-1.41 1.41" /><path d="m19.07 4.93-1.41 1.41" /></svg>
              </div>
              <div className={`theme-toggle-dot ${theme === 'dark' ? 'active' : ''}`} id="dotDark" title="Dark Mode">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z" /></svg>
              </div>
            </div>

            <span className="sidebar-theme-text" id="themeModeText">
              {theme === 'dark' ? 'Dark Mode' : 'Light Mode'}
            </span>

            <button
              type="button"
              id="authBtn"
              className={`sidebar-auth-action ${isLoggedIn ? 'btn-logout' : 'btn-login'}`}
              onClick={() => {
                if (isLoggedIn) handleLogout();
                else setShowLoginModal(true);
              }}
              title={isLoggedIn ? "Logout" : "Login to Edit"}
            >
              {isLoggedIn ? (
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"></path>
                  <polyline points="16 17 21 12 16 7"></polyline>
                  <line x1="21" y1="12" x2="9" y2="12"></line>
                </svg>
              ) : (
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"></path>
                  <polyline points="10 17 15 12 10 7"></polyline>
                  <line x1="15" y1="12" x2="3" y2="12"></line>
                </svg>
              )}
            </button>
          </div>
        </div>
      </aside>

      {/* ==================================================== */}
      {/* MAIN CONTENT AREA                                    */}
      {/* ==================================================== */}
      <main className="app-main-content">
        {/* Header */}
        <header className="app-header">
          <div className="header-left">
            <button
              type="button"
              className="sidebar-toggle-btn"
              onClick={() => setIsCollapsed(!isCollapsed)}
              title="Toggle Sidebar (Collapse / Expand)"
            >
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect>
                <line x1="9" y1="3" x2="9" y2="21"></line>
              </svg>
            </button>
            <div className="header-search">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <circle cx="11" cy="11" r="8"></circle>
                <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
              </svg>
              <input
                type="text"
                id="searchInput"
                placeholder="Search across all fields..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
              />
            </div>
          </div>

          <div className="header-actions">
            <div className="dock-pill-group">
              <button
                className={`dock-btn ${density === 'compact' ? 'active' : ''}`}
                id="btnDensityCompact"
                onClick={() => setDensity('compact')}
              >
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M4 14h6v6"></path>
                  <path d="M20 10h-6V4"></path>
                  <path d="M14 10l7-7"></path>
                  <path d="M3 21l7-7"></path>
                </svg>
                <span>Compact</span>
              </button>
              <button
                className={`dock-btn ${density === 'normal' ? 'active' : ''}`}
                id="btnDensityNormal"
                onClick={() => setDensity('normal')}
              >
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <rect x="3" y="3" width="18" height="18" rx="2"></rect>
                  <path d="M3 9h18"></path>
                  <path d="M3 15h18"></path>
                </svg>
                <span>Normal</span>
              </button>
            </div>

            {isLoggedIn && activeTab === 'repair' && selectedRows.size > 0 && (
              <button
                className="btn-ui btn-danger"
                id="bulkDeleteBtn"
                onClick={deleteSelectedRows}
              >
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <polyline points="3 6 5 6 21 6"></polyline>
                  <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
                </svg>
                <span>Delete Selected ({selectedRows.size})</span>
              </button>
            )}

            {/* Open Excel File Button */}
            <a
              href={EXCEL_SHAREPOINT_LINK}
              target="_blank"
              rel="noopener noreferrer"
              className="btn-ui"
              title="Open Original Excel Sheet in SharePoint"
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#10b981" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z" />
                <polyline points="14 2 14 8 20 8" />
                <line x1="8" y1="13" x2="16" y2="13" />
                <line x1="8" y1="17" x2="16" y2="17" />
                <line x1="10" y1="9" x2="8" y2="9" />
              </svg>
              <span>Open Excel</span>
            </a>

            {/* Publish to GitHub Button */}
            <button
              type="button"
              className="btn-ui"
              onClick={() => setShowGithubModal(true)}
              title="Publish on your GitHub"
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M9 19c-5 1.5-5-2.5-7-3m14 6v-3.87a3.37 3.37 0 0 0-.94-2.61c3.14-.35 6.44-1.54 6.44-7A5.44 5.44 0 0 0 20 4.77 5.07 5.07 0 0 0 19.91 1S18.73.65 16 2.48a13.38 13.38 0 0 0-7 0C6.27.65 5.09 1 5.09 1A5.07 5.07 0 0 0 5 4.77a5.44 5.44 0 0 0-1.5 3.78c0 5.42 3.3 6.61 6.44 7A3.37 3.37 0 0 0 9 18.13V22"></path>
              </svg>
              <span>Publish to GitHub</span>
            </button>

            {isLoggedIn && activeTab === 'repair' && (
              <button
                className="btn-ui btn-primary"
                id="addRowBtn"
                onClick={openAddRowDrawer}
              >
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                  <line x1="12" y1="5" x2="12" y2="19"></line>
                  <line x1="5" y1="12" x2="19" y2="12"></line>
                </svg>
                <span>Add Row</span>
              </button>
            )}

            {isLoggedIn && activeTab === 'spareparts' && (
              <button
                className="btn-ui btn-primary"
                id="addSpareBtn"
                onClick={openAddSpareDrawer}
              >
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                  <line x1="12" y1="5" x2="12" y2="19"></line>
                  <line x1="5" y1="12" x2="19" y2="12"></line>
                </svg>
                <span>Add Spare Part</span>
              </button>
            )}
          </div>
        </header>

        {/* ==================================================== */}
        {/* TABLE CANVAS - HANDLER REPAIR                        */}
        {/* ==================================================== */}
        {activeTab === 'repair' && (
          <div className="table-wrap" id="tableWrap" ref={tableWrapRef}>
            <table className="sheet-table">
              <thead>
                <tr>
                  {isLoggedIn && (
                    <th className="col-check">
                      <input
                        type="checkbox"
                        id="selectAll"
                        checked={filteredRepairs.length > 0 && selectedRows.size === filteredRepairs.length}
                        onChange={(e) => toggleSelectAll(e.target.checked)}
                      />
                    </th>
                  )}
                  <th className="col-num">#</th>
                  <th><div className="th-content"><span className="type-icon-svg"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z" /><line x1="16" y1="2" x2="16" y2="6" /><line x1="8" y1="2" x2="8" y2="6" /><line x1="2" y1="10" x2="22" y2="10" /></svg></span> WW</div></th>
                  <th><div className="th-content"><span className="type-icon-svg"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="4" width="18" height="18" rx="2" ry="2" /><line x1="16" y1="2" x2="16" y2="6" /><line x1="8" y1="2" x2="8" y2="6" /><line x1="3" y1="10" x2="21" y2="10" /></svg></span> DATE</div></th>
                  <th><div className="th-content"><span className="type-icon-svg"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="4" y="4" width="16" height="16" rx="2" /><rect x="9" y="9" width="6" height="6" /><line x1="9" y1="1" x2="9" y2="4" /><line x1="15" y1="1" x2="15" y2="4" /></svg></span> TESTER ID</div></th>
                  <th><div className="th-content"><span className="type-icon-svg"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="11" width="18" height="10" rx="2" /><circle cx="12" cy="5" r="2" /><path d="M12 7v4" /></svg></span> HANDLER ID</div></th>
                  <th><div className="th-content"><span className="type-icon-svg"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="3" width="18" height="18" rx="2" ry="2" /><circle cx="8.5" cy="8.5" r="1.5" /><polyline points="21 15 16 10 5 21" /></svg></span> DETAILS (Image as Evidence)</div></th>
                  <th><div className="th-content"><span className="type-icon-svg"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="5" y="2" width="14" height="20" rx="2" ry="2" /><line x1="12" y1="18" x2="12.01" y2="18" /></svg></span> DEVICE</div></th>
                  <th><div className="th-content"><span className="type-icon-svg"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z" /><polyline points="3.27 6.96 12 12.01 20.73 6.96" /></svg></span> PACKAGE</div></th>
                  <th><div className="th-content"><span className="type-icon-svg"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polygon points="12 2 2 7 12 12 22 7 12 2" /><polyline points="2 17 12 22 22 17" /></svg></span> Handler Major Part</div></th>
                  <th><div className="th-content"><span className="type-icon-svg"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z" /><line x1="12" y1="9" x2="12" y2="13" /></svg></span> PROBLEM</div></th>
                  <th><div className="th-content"><span className="type-icon-svg"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="11" cy="11" r="8" /><line x1="21" y1="21" x2="16.65" y2="16.65" /></svg></span> ROOT CAUSE</div></th>
                  <th><div className="th-content"><span className="type-icon-svg"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M20.59 13.41l-7.17 7.17a2 2 0 0 1-2.83 0L2 12V2h10l8.59 8.59a2 2 0 0 1 0 2.82z" /></svg></span> MATERIAL CODE</div></th>
                  <th><div className="th-content"><span className="type-icon-svg"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="2" y="7" width="20" height="14" rx="2" ry="2" /><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16" /></svg></span> Stock Items</div></th>
                  <th><div className="th-content"><span className="type-icon-svg"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="10" /><polyline points="12 6 12 12 14 14" /></svg></span> PENDING ACTION</div></th>
                  <th><div className="th-content"><span className="type-icon-svg"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="10" /><circle cx="12" cy="12" r="3" fill="currentColor" /></svg></span> STATUS</div></th>
                  {isLoggedIn && (
                    <th className="col-actions"><div className="th-content"><span className="type-icon-svg"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="3" /><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z" /></svg></span> ACTIONS</div></th>
                  )}
                </tr>
              </thead>
              <tbody id="dataTableBody">
                {filteredRepairs.length === 0 ? (
                  <tr>
                    <td colSpan={17} style={{ textAlign: 'center', padding: '36px', color: 'var(--text-muted)' }}>
                      No matching handler records found
                    </td>
                  </tr>
                ) : (
                  filteredRepairs.map((row, idx) => {
                    const st = getRowVal(row, 'STATUS') || 'Active';
                    const detailsImg = getRowVal(row, 'DETAILS');
                    const isImg = detailsImg && (detailsImg.startsWith('data:image') || detailsImg.startsWith('http') || detailsImg.startsWith('blob:'));

                    return (
                      <tr key={idx} className={selectedRows.has(idx) ? 'selected' : ''}>
                        {isLoggedIn && (
                          <td className="col-check">
                            <input
                              type="checkbox"
                              checked={selectedRows.has(idx)}
                              onChange={() => toggleRowSelect(idx)}
                            />
                          </td>
                        )}
                        <td className="col-num">{idx + 1}</td>
                        <td>{getRowVal(row, 'WW') || '-'}</td>
                        <td>{getRowVal(row, 'DATE') || '-'}</td>
                        <td>{getRowVal(row, 'TESTER ID') || '-'}</td>
                        <td style={{ fontWeight: 700 }}>{getRowVal(row, 'HANDLER ID') || '-'}</td>
                        <td>
                          {isImg ? (
                            <img
                              src={detailsImg}
                              className="cell-img-thumb"
                              onClick={() => setLightboxImg(detailsImg)}
                              alt="Evidence"
                            />
                          ) : (
                            <span style={{ color: 'var(--text-muted)' }}>-</span>
                          )}
                        </td>
                        <td>{getRowVal(row, 'DEVICE') || '-'}</td>
                        <td>{getRowVal(row, 'PACKAGE') || '-'}</td>
                        <td>{getRowVal(row, 'Handler Major Part') || '-'}</td>
                        <td>{getRowVal(row, 'PROBLEM') || '-'}</td>
                        <td>{getRowVal(row, 'ROOT CAUSE') || '-'}</td>
                        <td>{getRowVal(row, 'MATERIAL CODE') || '-'}</td>
                        <td>{getRowVal(row, 'STOCK ITEMS') || '-'}</td>
                        <td>{getRowVal(row, 'PENDING ACTION') || '-'}</td>
                        <td><span className={`badge status-${st}`}>{st}</span></td>
                        {isLoggedIn && (
                          <td className="col-actions">
                            <button className="row-btn row-btn-edit" onClick={() => openEditDrawer(idx)} title="Edit Record">
                              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path></svg>
                              <span>Edit</span>
                            </button>
                            <button className="row-btn row-btn-delete" onClick={() => deleteSingleRow(idx)} title="Delete Record">
                              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>
                              <span>Delete</span>
                            </button>
                          </td>
                        )}
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        )}

        {/* ==================================================== */}
        {/* TABLE CANVAS - SPARE PARTS INVENTORY                 */}
        {/* ==================================================== */}
        {activeTab === 'spareparts' && (
          <div className="table-wrap" id="sparePartsView">
            <table className="sheet-table">
              <thead>
                <tr>
                  <th className="col-num">#</th>
                  <th><div className="th-content"><span className="type-icon-svg"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M20.59 13.41l-7.17 7.17a2 2 0 0 1-2.83 0L2 12V2h10l8.59 8.59a2 2 0 0 1 0 2.82z" /><line x1="7" y1="7" x2="7.01" y2="7" /></svg></span> NAME</div></th>
                  <th><div className="th-content"><span className="type-icon-svg"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="3" width="18" height="18" rx="2" ry="2" /><line x1="9" y1="3" x2="9" y2="21" /><line x1="15" y1="3" x2="15" y2="21" /></svg></span> PRODUCT CODE</div></th>
                  <th><div className="th-content"><span className="type-icon-svg"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="11" width="18" height="10" rx="2" /><circle cx="12" cy="5" r="2" /><path d="M12 7v4" /></svg></span> HANDLER TYPE</div></th>
                  <th><div className="th-content"><span className="type-icon-svg"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="2" y="7" width="20" height="14" rx="2" ry="2" /><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16" /></svg></span> STOCKS QUANTITY</div></th>
                  <th><div className="th-content"><span className="type-icon-svg"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="3" width="18" height="18" rx="2" ry="2" /><circle cx="8.5" cy="8.5" r="1.5" /><polyline points="21 15 16 10 5 21" /></svg></span> IMAGE</div></th>
                  <th className="col-actions"><div className="th-content"><span className="type-icon-svg"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M12 20v-6M6 20V10M18 20V4" /></svg></span> STOCK ADJUST</div></th>
                  {isLoggedIn && (
                    <th className="col-actions"><div className="th-content"><span className="type-icon-svg"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="3" /><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z" /></svg></span> ACTIONS</div></th>
                  )}
                </tr>
              </thead>
              <tbody>
                {filteredSpares.length === 0 ? (
                  <tr>
                    <td colSpan={8} style={{ textAlign: 'center', padding: '36px', color: 'var(--text-muted)' }}>
                      No spare parts match the filter.
                    </td>
                  </tr>
                ) : (
                  filteredSpares.map((part, idx) => {
                    const qty = Number(part.quantity || 0);
                    let badgeClass = 'in-stock';
                    let badgeLabel = `${qty} in stock`;
                    if (qty === 0) {
                      badgeClass = 'out-of-stock';
                      badgeLabel = 'Out of Stock';
                    } else if (qty <= 5) {
                      badgeClass = 'low-stock';
                      badgeLabel = `${qty} (Low Stock)`;
                    }

                    const handlers = (part.handler || 'MT99').split(',').map((h, i) => (
                      <span key={i} className="stock-handler-tag">{h.trim()}</span>
                    ));

                    return (
                      <tr key={part.id}>
                        <td className="col-num">{idx + 1}</td>
                        <td style={{ fontWeight: 700, color: 'var(--text-main)' }}>{part.name}</td>
                        <td>
                          <code style={{ fontSize: '12px', background: 'var(--control-bg)', padding: '3px 6px', borderRadius: '6px' }}>
                            {part.code}
                          </code>
                        </td>
                        <td>{handlers}</td>
                        <td><span className={`stock-qty-badge ${badgeClass}`}>{badgeLabel}</span></td>
                        <td>
                          {part.image ? (
                            <img
                              src={part.image}
                              className="cell-img-thumb"
                              onClick={() => setLightboxImg(part.image || null)}
                              alt={part.name}
                            />
                          ) : (
                            <span style={{ color: 'var(--text-muted)' }}>-</span>
                          )}
                        </td>
                        <td className="col-actions">
                          <div className="stock-ctrl-group">
                            <button
                              type="button"
                              className="btn-stock-adj btn-minus"
                              title="Withdraw Stocks (-)"
                              onClick={() => openStockAdjustModal(part.id, 'withdraw')}
                            >
                              &minus;
                            </button>
                            <span style={{ fontWeight: 800, fontSize: '13px', minWidth: '28px', textAlign: 'center' }}>
                              {qty}
                            </span>
                            <button
                              type="button"
                              className="btn-stock-adj btn-plus"
                              title="Add / Replenish Stocks (+)"
                              onClick={() => openStockAdjustModal(part.id, 'add')}
                            >
                              &plus;
                            </button>
                          </div>
                        </td>
                        {isLoggedIn && (
                          <td className="col-actions">
                            <button
                              type="button"
                              className="row-btn row-btn-edit"
                              onClick={() => openEditSpareDrawer(part.id)}
                              title="Edit Part"
                            >
                              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path></svg>
                              <span>Edit</span>
                            </button>
                            <button
                              type="button"
                              className="row-btn row-btn-delete"
                              onClick={() => deleteSparePart(part.id)}
                              title="Delete Part"
                            >
                              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>
                              <span>Delete</span>
                            </button>
                          </td>
                        )}
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        )}

        <footer className="app-footer">
          <span>
            {activeTab === 'repair'
              ? `Showing ${filteredRepairs.length} of ${rawData.length} items`
              : `Showing ${filteredSpares.length} of ${spareData.length} spare parts`}
          </span>
          <span style={{ opacity: 0.75, fontSize: '11px' }}>
            Scroll horizontally to view all columns
          </span>
        </footer>
      </main>

      {/* ==================================================== */}
      {/* MODALS & DRAWERS                                     */}
      {/* ==================================================== */}

      {/* Login Modal */}
      {showLoginModal && (
        <div className="modal-overlay" onClick={() => setShowLoginModal(false)}>
          <div className="modal-card" onClick={(e) => e.stopPropagation()}>
            <h3 style={{ marginBottom: '6px', fontSize: '16px', fontWeight: 800, color: 'var(--text-main)' }}>
              Login to Edit
            </h3>
            <p style={{ fontSize: '12px', color: 'var(--text-muted)', marginBottom: '18px' }}>
              Sign in with credentials to modify records.
            </p>
            <form onSubmit={handleLogin}>
              <input
                type="text"
                placeholder="Username (Default: admin)"
                value={loginUsername}
                onChange={(e) => setLoginUsername(e.target.value)}
                required
              />
              <input
                type="password"
                placeholder="Password (Default: admin123)"
                value={loginPassword}
                onChange={(e) => setLoginPassword(e.target.value)}
                required
              />
              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px' }}>
                <button type="button" className="btn-ui" onClick={() => setShowLoginModal(false)}>
                  Cancel
                </button>
                <button type="submit" className="btn-ui btn-primary">
                  Login
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Handler Record Add / Edit Drawer */}
      {showRepairDrawer && (
        <>
          <div className="drawer-overlay" onClick={() => setShowRepairDrawer(false)}></div>
          <div className="drawer-panel">
            <div className="drawer-header">
              <h2>
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#f97316" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                  <line x1="12" y1="5" x2="12" y2="19"></line>
                  <line x1="5" y1="12" x2="19" y2="12"></line>
                </svg>
                <span>{editingRepairIndex !== null ? 'Edit Record' : 'Add New Row / Record'}</span>
              </h2>
              <button className="drawer-close" onClick={() => setShowRepairDrawer(false)}>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <line x1="18" y1="6" x2="6" y2="18"></line>
                  <line x1="6" y1="6" x2="18" y2="18"></line>
                </svg>
              </button>
            </div>
            <div className="drawer-body">
              <form onSubmit={submitDrawerRecord}>
                <div className="form-grid">
                  <div className="form-group">
                    <label>WW</label>
                    <input
                      type="text"
                      className="form-control"
                      value={repairForm.ww}
                      onChange={(e) => setRepairForm({ ...repairForm, ww: e.target.value })}
                      required
                    />
                  </div>
                  <div className="form-group">
                    <label>Date</label>
                    <input
                      type="date"
                      className="form-control"
                      value={repairForm.date}
                      onChange={(e) => setRepairForm({ ...repairForm, date: e.target.value })}
                      required
                    />
                  </div>

                  <div className="form-group">
                    <label>Tester ID</label>
                    <input
                      type="text"
                      className="form-control"
                      placeholder="Enter tester id"
                      value={repairForm.testerId}
                      onChange={(e) => setRepairForm({ ...repairForm, testerId: e.target.value })}
                    />
                  </div>
                  <div className="form-group">
                    <label>Handler ID</label>
                    <input
                      type="text"
                      className="form-control"
                      placeholder="Enter handler id"
                      value={repairForm.handlerId}
                      onChange={(e) => setRepairForm({ ...repairForm, handlerId: e.target.value })}
                      required
                    />
                  </div>

                  {/* DETAILS (IMAGE AS EVIDENCE) */}
                  <div className="form-group col-span-2">
                    <label>DETAILS (Image as Evidence)</label>
                    <div className="photo-picker-wrap">
                      <div className="photo-thumb-box">
                        {repairForm.details ? (
                          <img src={repairForm.details} alt="Evidence" />
                        ) : (
                          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8">
                            <rect x="3" y="3" width="18" height="18" rx="2" ry="2" />
                            <circle cx="8.5" cy="8.5" r="1.5" />
                            <polyline points="21 15 16 10 5 21" />
                          </svg>
                        )}
                      </div>
                      <input
                        type="file"
                        id="form-details-file"
                        accept="image/*"
                        style={{ display: 'none' }}
                        onChange={async (e) => {
                          if (e.target.files && e.target.files[0]) {
                            const b64 = await compressImage(e.target.files[0]);
                            setRepairForm({ ...repairForm, details: b64 });
                          }
                        }}
                      />
                      <button
                        type="button"
                        className="btn-ui"
                        onClick={() => document.getElementById('form-details-file')?.click()}
                      >
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                          <polyline points="17 8 12 3 7 8"></polyline>
                          <line x1="12" y1="3" x2="12" y2="15"></line>
                        </svg>
                        <span>Upload / Select Photo</span>
                      </button>
                      <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
                        {repairForm.details ? 'Photo Attached' : 'No photo chosen'}
                      </span>
                    </div>
                  </div>

                  <div className="form-group">
                    <label>Device</label>
                    <input
                      type="text"
                      className="form-control"
                      value={repairForm.device}
                      onChange={(e) => setRepairForm({ ...repairForm, device: e.target.value })}
                    />
                  </div>
                  <div className="form-group">
                    <label>Package</label>
                    <input
                      type="text"
                      className="form-control"
                      value={repairForm.package}
                      onChange={(e) => setRepairForm({ ...repairForm, package: e.target.value })}
                    />
                  </div>

                  <div className="form-group col-span-2">
                    <label>Handler Major Part</label>
                    <input
                      type="text"
                      className="form-control"
                      value={repairForm.majorPart}
                      onChange={(e) => setRepairForm({ ...repairForm, majorPart: e.target.value })}
                    />
                  </div>
                  <div className="form-group col-span-2">
                    <label>Problem</label>
                    <input
                      type="text"
                      className="form-control"
                      value={repairForm.problem}
                      onChange={(e) => setRepairForm({ ...repairForm, problem: e.target.value })}
                      required
                    />
                  </div>

                  <div className="form-group col-span-2">
                    <label>Root Cause</label>
                    <input
                      type="text"
                      className="form-control"
                      value={repairForm.rootCause}
                      onChange={(e) => setRepairForm({ ...repairForm, rootCause: e.target.value })}
                    />
                  </div>

                  <div className="form-group">
                    <label>Material Code</label>
                    <input
                      type="text"
                      className="form-control"
                      value={repairForm.materialCode}
                      onChange={(e) => setRepairForm({ ...repairForm, materialCode: e.target.value })}
                    />
                  </div>
                  <div className="form-group">
                    <label>Stock Items</label>
                    <select
                      className="form-control"
                      value={repairForm.stockItems}
                      onChange={(e) => setRepairForm({ ...repairForm, stockItems: e.target.value })}
                    >
                      <option value="Stock Items">Stock Items</option>
                      <option value="Non Stock Items">Non Stock Items</option>
                    </select>
                  </div>

                  <div className="form-group">
                    <label>Pending Action</label>
                    <input
                      type="text"
                      className="form-control"
                      value={repairForm.pendingAction}
                      onChange={(e) => setRepairForm({ ...repairForm, pendingAction: e.target.value })}
                    />
                  </div>
                  <div className="form-group">
                    <label>Status</label>
                    <select
                      className="form-control"
                      value={repairForm.status}
                      onChange={(e) => setRepairForm({ ...repairForm, status: e.target.value })}
                    >
                      <option value="Active">Active</option>
                      <option value="Pending">Pending</option>
                      <option value="Resolved">Resolved</option>
                    </select>
                  </div>
                </div>

                <div className="drawer-footer" style={{ marginTop: '26px', marginLeft: '-26px', marginRight: '-26px', marginBottom: '-26px' }}>
                  <button type="button" className="btn-ui" onClick={() => setShowRepairDrawer(false)}>
                    Cancel
                  </button>
                  <button type="submit" className="btn-ui btn-primary">
                    {editingRepairIndex !== null ? 'Update Record' : 'Save Record'}
                  </button>
                </div>
              </form>
            </div>
          </div>
        </>
      )}

      {/* Spare Part Add / Edit Drawer */}
      {showSpareDrawer && (
        <>
          <div className="drawer-overlay" onClick={() => setShowSpareDrawer(false)}></div>
          <div className="drawer-panel">
            <div className="drawer-header">
              <div>
                <h3 style={{ fontSize: '16.5px', fontWeight: 800, color: 'var(--text-main)' }}>
                  {editingSpareId ? 'Edit Spare Part' : 'Add Spare Part'}
                </h3>
                <p style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
                  Register stock equipment or spare part into inventory.
                </p>
              </div>
              <button type="button" className="drawer-close" onClick={() => setShowSpareDrawer(false)}>
                &times;
              </button>
            </div>
            <div className="drawer-body">
              <form onSubmit={handleSpareFormSubmit}>
                <div className="form-grid">
                  <div className="form-group col-span-2">
                    <label>Part Name *</label>
                    <input
                      type="text"
                      className="form-control"
                      placeholder="e.g. Test Socket Pin Block"
                      value={spareForm.name}
                      onChange={(e) => setSpareForm({ ...spareForm, name: e.target.value })}
                      required
                    />
                  </div>

                  <div className="form-group">
                    <label>Product Code *</label>
                    <input
                      type="text"
                      className="form-control"
                      placeholder="e.g. SP-MT99-042"
                      value={spareForm.code}
                      onChange={(e) => setSpareForm({ ...spareForm, code: e.target.value })}
                      required
                    />
                  </div>

                  <div className="form-group">
                    <label>Handler Compatibility *</label>
                    <input
                      type="text"
                      className="form-control"
                      placeholder="e.g. MT99, MT93"
                      value={spareForm.handler}
                      onChange={(e) => setSpareForm({ ...spareForm, handler: e.target.value })}
                      required
                    />
                  </div>

                  <div className="form-group col-span-2">
                    <label>Initial Stock Quantity *</label>
                    <input
                      type="number"
                      className="form-control"
                      min="0"
                      value={spareForm.quantity}
                      onChange={(e) => setSpareForm({ ...spareForm, quantity: parseInt(e.target.value, 10) || 0 })}
                      required
                    />
                  </div>

                  <div className="form-group col-span-2">
                    <label>Product Image</label>
                    <div className="photo-picker-wrap">
                      <div className="photo-thumb-box">
                        {spareForm.image ? (
                          <img src={spareForm.image} alt="Part" />
                        ) : (
                          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                            <rect x="3" y="3" width="18" height="18" rx="2" ry="2" />
                            <circle cx="8.5" cy="8.5" r="1.5" />
                            <polyline points="21 15 16 10 5 21" />
                          </svg>
                        )}
                      </div>
                      <input
                        type="file"
                        id="spare-form-file"
                        accept="image/*"
                        style={{ display: 'none' }}
                        onChange={async (e) => {
                          if (e.target.files && e.target.files[0]) {
                            const b64 = await compressImage(e.target.files[0]);
                            setSpareForm({ ...spareForm, image: b64 });
                          }
                        }}
                      />
                      <button
                        type="button"
                        className="btn-ui"
                        onClick={() => document.getElementById('spare-form-file')?.click()}
                      >
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                          <polyline points="17 8 12 3 7 8"></polyline>
                          <line x1="12" y1="3" x2="12" y2="15"></line>
                        </svg>
                        <span>Upload Image</span>
                      </button>
                      <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
                        {spareForm.image ? 'Image Attached' : 'No image chosen'}
                      </span>
                    </div>
                  </div>
                </div>

                <div className="drawer-footer" style={{ marginTop: '26px', marginLeft: '-26px', marginRight: '-26px', marginBottom: '-26px' }}>
                  <button type="button" className="btn-ui" onClick={() => setShowSpareDrawer(false)}>
                    Cancel
                  </button>
                  <button type="submit" className="btn-ui btn-primary">
                    {editingSpareId ? 'Update Spare Part' : 'Save Spare Part'}
                  </button>
                </div>
              </form>
            </div>
          </div>
        </>
      )}

      {/* Stock Adjust Modal */}
      {showStockModal && (
        <div className="modal-overlay" onClick={() => setShowStockModal(false)}>
          <div className="modal-card" onClick={(e) => e.stopPropagation()}>
            <h3 style={{ marginBottom: '6px', fontSize: '16px', fontWeight: 800, color: 'var(--text-main)' }}>
              {stockAdjustType === 'add' ? 'Add / Replenish Stocks' : 'Withdraw Stocks'}
            </h3>
            <p style={{ fontSize: '12px', color: 'var(--text-muted)', marginBottom: '18px' }}>
              Adjust inventory units for selected spare equipment.
            </p>
            <form onSubmit={handleStockAdjustSubmit}>
              <div className="form-group" style={{ marginBottom: '14px' }}>
                <label>Quantity to {stockAdjustType === 'add' ? 'Add' : 'Withdraw'}</label>
                <input
                  type="number"
                  className="form-control"
                  min="1"
                  value={stockAdjustQty}
                  onChange={(e) => setStockAdjustQty(parseInt(e.target.value, 10) || 1)}
                  required
                  style={{ fontSize: '16px', fontWeight: 700, textAlign: 'center' }}
                />
              </div>

              <div className="form-group" style={{ marginBottom: '18px' }}>
                <label>Reason / Note (Optional)</label>
                <input
                  type="text"
                  className="form-control"
                  placeholder="e.g. Preventive Maintenance MT99"
                  value={stockAdjustReason}
                  onChange={(e) => setStockAdjustReason(e.target.value)}
                />
              </div>

              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px' }}>
                <button type="button" className="btn-ui" onClick={() => setShowStockModal(false)}>
                  Cancel
                </button>
                <button type="submit" className={`btn-ui ${stockAdjustType === 'add' ? 'btn-primary' : 'btn-danger'}`}>
                  Confirm
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Lightbox Modal */}
      {lightboxImg && (
        <div className="lightbox-modal" onClick={() => setLightboxImg(null)}>
          <img src={lightboxImg} alt="Preview" />
        </div>
      )}

      {/* GitHub Instructions Modal */}
      {showGithubModal && (
        <div className="modal-overlay" onClick={() => setShowGithubModal(false)}>
          <div className="modal-card" style={{ width: '480px' }} onClick={(e) => e.stopPropagation()}>
            <h3 style={{ marginBottom: '6px', fontSize: '17px', fontWeight: 800, color: 'var(--text-main)', display: 'flex', alignItems: 'center', gap: '8px' }}>
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M9 19c-5 1.5-5-2.5-7-3m14 6v-3.87a3.37 3.37 0 0 0-.94-2.61c3.14-.35 6.44-1.54 6.44-7A5.44 5.44 0 0 0 20 4.77 5.07 5.07 0 0 0 19.91 1S18.73.65 16 2.48a13.38 13.38 0 0 0-7 0C6.27.65 5.09 1 5.09 1A5.07 5.07 0 0 0 5 4.77a5.44 5.44 0 0 0-1.5 3.78c0 5.42 3.3 6.61 6.44 7A3.37 3.37 0 0 0 9 18.13V22"></path>
              </svg>
              <span>Publish to your GitHub</span>
            </h3>
            <p style={{ fontSize: '12.5px', color: 'var(--text-muted)', marginBottom: '14px', lineHeight: 1.5 }}>
              Your Git repository is initialized with GitHub Actions workflow and compiled <code>docs/</code> bundle.
            </p>

            <div style={{ background: 'rgba(0,0,0,0.35)', padding: '12px 14px', borderRadius: '10px', fontSize: '12px', fontFamily: 'monospace', marginBottom: '14px', overflowX: 'auto', border: '1px solid rgba(255,255,255,0.06)' }}>
              <div>git remote add origin https://github.com/arielskie12345/new-handler.git</div>
              <div>git branch -M main</div>
              <div>git push -u origin main</div>
            </div>

            <div style={{ fontSize: '12px', color: 'var(--text-secondary)', background: 'var(--control-bg)', padding: '10px 12px', borderRadius: '8px', marginBottom: '16px', lineHeight: 1.5 }}>
              <strong>To fix blank page on GitHub Pages:</strong>
              <div style={{ marginTop: '4px' }}>
                1. Go to your repo <strong>Settings &rarr; Pages</strong>.
              </div>
              <div>
                2. Under <em>Build and deployment &gt; Source</em>, select <strong>GitHub Actions</strong> (automated) OR choose branch <strong>main</strong> with folder <strong>/docs</strong>.
              </div>
            </div>

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px' }}>
              <button
                type="button"
                className="btn-ui"
                onClick={() => {
                  navigator.clipboard.writeText("git remote add origin https://github.com/arielskie12345/new-handler.git\ngit branch -M main\ngit push -u origin main");
                  setCopiedCmd(true);
                  setTimeout(() => setCopiedCmd(false), 2000);
                }}
              >
                <span>{copiedCmd ? "Copied Commands!" : "Copy Commands"}</span>
              </button>
              <button type="button" className="btn-ui btn-primary" onClick={() => setShowGithubModal(false)}>
                Got it
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Toast Notification */}
      {toastMessage && <div id="toast">{toastMessage}</div>}
    </div>
  );
}
