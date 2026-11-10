import { useEffect, useRef, useState } from 'react';
import './App.css';
import {
  ArrowUpRight,
  Download,
  FileSpreadsheet,
  FileText,
  LayoutDashboard,
  Pencil,
  Plus,
  Printer,
  Search,
  Settings,
  Trash2,
} from 'lucide-react';
import type { WorkbookPass, WorkbookSettings } from './workbookStore';

type MaterialRow = {
  id: number;
  description: string;
  pieces: string | number;
  feet: string | number;
  rate: string | number;
};

type GatePass = Omit<WorkbookPass, 'items'> & { items: MaterialRow[] };

type SettingsState = WorkbookSettings;

const defaultMaterialRow = (id = 1): MaterialRow => ({
  id,
  description: '',
  pieces: '',
  feet: '',
  rate: '',
});

const createEmptyForm = () => ({
  gate_pass_number: '',
  date: new Date().toISOString().slice(0, 10),
  person_name: '',
  mobile_number: '',
  company_name: '',
  taken_by: '',
  authorized_by: '',
  status: 'Returnable' as 'Returnable' | 'Non-Returnable',
  items: [defaultMaterialRow()],
});

function localDateString(date = new Date()) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function App() {
  const [activePage, setActivePage] = useState('dashboard');
  const [gates, setGates] = useState<GatePass[]>([]);
  const [workbookName, setWorkbookName] = useState('');
  const [workbookBusy, setWorkbookBusy] = useState(false);
  const [settings, setSettings] = useState<SettingsState>({
    company_name: 'ISLAM MARBLE',
    address: '',
    contact_number: '',
    logo_path: '',
    gate_pass_prefix: 'GP-',
    default_status: 'Returnable',
  });
  const [form, setForm] = useState(createEmptyForm());
  const [editingId, setEditingId] = useState<number | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [recordFilter, setRecordFilter] = useState('all');
  const [reportFilter, setReportFilter] = useState('all');
  const [reportRange, setReportRange] = useState({ start: '', end: '' });
  const [previewPass, setPreviewPass] = useState<any>(null);
  const [toast, setToast] = useState('');
  const [exportingPass, setExportingPass] = useState(false);
  const printDocumentRef = useRef<HTMLDivElement>(null);

  const today = localDateString();
  const dashboard = {
    totalGatePasses: gates.length,
    todaysPasses: gates.filter((pass) => pass.date === today).length,
    recent: [...gates].sort((a, b) => b.date.localeCompare(a.date) || b.id - a.id).slice(0, 5),
  };
  const filteredGates = gates.filter((pass) => {
    const search = searchTerm.trim().toLowerCase();
    const searchMatches = !search || [
      pass.person_name, pass.mobile_number, pass.company_name, pass.gate_pass_number,
      ...(pass.items || []).map((item) => item.description),
    ].join(' ').toLowerCase().includes(search);
    const passDate = pass.date;
    const todayDate = new Date(`${today}T00:00:00`);
    const filterMatches = recordFilter === 'all'
      || (recordFilter === 'today' && passDate === today)
      || (recordFilter === 'yesterday' && passDate === localDateString(new Date(todayDate.getTime() - 86400000)))
      || (recordFilter === 'this-week' && passDate >= localDateString(new Date(todayDate.getFullYear(), todayDate.getMonth(), todayDate.getDate() - todayDate.getDay()))
        && passDate <= localDateString(new Date(todayDate.getFullYear(), todayDate.getMonth(), todayDate.getDate() + 6 - todayDate.getDay())))
      || (recordFilter === 'this-month' && passDate.startsWith(today.slice(0, 7)));
    return searchMatches && filterMatches;
  }).sort((a, b) => b.date.localeCompare(a.date) || b.id - a.id);

  useEffect(() => {
    if (!toast) return;
    const timer = window.setTimeout(() => setToast(''), 3000);
    return () => window.clearTimeout(timer);
  }, [toast]);

  const statusClass = (status: string) =>
    status === 'Returnable' ? 'status-pill returnable' : 'status-pill non-returnable';

  const selectWorkbook = async (action: 'open' | 'create') => {
    setWorkbookBusy(true);
    try {
      const store = await import('./workbookStore');
      const result = action === 'open'
        ? await store.openWorkbookFile()
        : await store.createWorkbookFile();
      setWorkbookName(result.fileName);
      setSettings(result.settings);
      setGates(result.passes);
      setForm(createEmptyForm());
      setEditingId(null);
      setToast(`Excel file ready: ${result.fileName}`);
    } catch (error: any) {
      if (error?.name !== 'AbortError') {
        setToast(error.message || 'Could not open the Excel file.');
      }
    } finally {
      setWorkbookBusy(false);
    }
  };

  const requireWorkbook = () => {
    if (!workbookName) {
      setToast('Open or create an Excel file first. Gate passes will be saved only in that file.');
      return false;
    }
    return true;
  };

  const handleFormUpdate = (field: string, value: string) => {
    setForm((current) => ({ ...current, [field]: value }));
  };

  const handleMaterialUpdate = (index: number, field: keyof MaterialRow, value: string) => {
    setForm((current) => ({
      ...current,
      items: current.items.map((item, itemIndex) =>
        itemIndex === index ? { ...item, [field]: value } : item,
      ),
    }));
  };

  const addMaterialRow = () => {
    setForm((current) => ({
      ...current,
      items: [...current.items, defaultMaterialRow(current.items.length + 1)],
    }));
  };

  const removeMaterialRow = (index: number) => {
    setForm((current) => ({
      ...current,
      items: current.items.filter((_, idx) => idx !== index).map((item, updateIndex) => ({ ...item, id: updateIndex + 1 })),
    }));
  };

  const resetForm = () => {
    setForm(createEmptyForm());
    setEditingId(null);
  };

  const validateForm = () => {
    if (!form.person_name || !form.mobile_number) {
      throw new Error('Person name and mobile number are required.');
    }

    const cleaned = form.items.filter((item) => item.description.trim() || item.pieces || item.feet || item.rate);
    if (cleaned.length === 0) {
      throw new Error('At least one material item is required.');
    }

    for (const item of cleaned) {
      if (!item.description.trim()) {
        throw new Error('Enter a description for each material row.');
      }
      for (const [label, value] of [
        ['Pieces', item.pieces],
        ['Feet', item.feet],
        ['Rate', item.rate],
      ]) {
        const quantity = Number(value || 0);
        if (!Number.isFinite(quantity) || quantity < 0) {
          throw new Error(`${label} must be a valid non-negative number.`);
        }
      }
      if (Number(item.feet || 0) > 0 && !item.rate.trim()) {
        throw new Error('Enter a rate for each row with feet.');
      }
    }
  };

  const saveGatePass = async (printAfter = false) => {
    try {
      if (!requireWorkbook()) return;
      validateForm();
      const payload = {
        ...form,
        company_name: form.company_name || settings.company_name,
        status: form.status || settings.default_status,
        items: form.items
          .filter((item) => item.description.trim() || item.pieces || item.feet || item.rate)
          .map((item) => ({
            id: item.id,
            description: item.description,
            pieces: Number(item.pieces || 0),
            feet: Number(item.feet || 0),
            rate: Number(item.rate || 0),
            amount: Math.round((Number(item.feet || 0) * Number(item.rate || 0) + Number.EPSILON) * 100) / 100,
          })),
      };

      const savedPass = await (await import('./workbookStore')).saveGatePass({
        ...payload,
        id: editingId || 0,
        items: payload.items,
      }, editingId);
      setGates((current) => {
        const next = editingId
          ? current.map((pass) => pass.id === savedPass.id ? savedPass : pass)
          : [...current, savedPass];
        return next as GatePass[];
      });
      setToast(`Saved gate pass ${savedPass.gate_pass_number} in ${workbookName}`);
      setEditingId(null);
      setForm(createEmptyForm());

      if (printAfter) {
        setPreviewPass(savedPass);
        setTimeout(() => window.print(), 400);
      }
    } catch (error: any) {
      setToast(error.message || 'Could not save gate pass');
    }
  };

  const previewForm = () => {
    try {
      validateForm();
      setPreviewPass(buildPassDocument(form));
    } catch (error: any) {
      setToast(error.message || 'Please complete the required gate pass details.');
    }
  };

  const downloadPass = async (format: 'pdf' | 'png') => {
    const element = printDocumentRef.current;
    if (!element || exportingPass) return;

    setExportingPass(true);
    try {
      const [{ default: html2canvas }, { jsPDF }] = await Promise.all([
        import('html2canvas'),
        import('jspdf'),
      ]);
      await document.fonts.ready;
      const canvas = await html2canvas(element, {
        backgroundColor: '#ffffff',
        scale: 2,
        useCORS: true,
      });
      const passNumber = buildPassDocument(previewPass).gatePassNumber;
      const filename = `gate-pass-${passNumber.replace(/[^a-z0-9-]/gi, '-')}`;

      if (format === 'png') {
        const link = document.createElement('a');
        link.download = `${filename}.png`;
        link.href = canvas.toDataURL('image/png');
        link.click();
        return;
      }

      const pdf = new jsPDF('p', 'mm', 'a4');
      const margin = 10;
      const pageWidth = pdf.internal.pageSize.getWidth() - margin * 2;
      const pageHeight = pdf.internal.pageSize.getHeight() - margin * 2;
      const sourcePageHeight = Math.floor((canvas.width * pageHeight) / pageWidth);
      const pageCount = Math.ceil(canvas.height / sourcePageHeight);

      for (let pageIndex = 0; pageIndex < pageCount; pageIndex += 1) {
        const sourceY = pageIndex * sourcePageHeight;
        const sliceHeight = Math.min(sourcePageHeight, canvas.height - sourceY);
        const pageCanvas = document.createElement('canvas');
        pageCanvas.width = canvas.width;
        pageCanvas.height = sliceHeight;
        const context = pageCanvas.getContext('2d');
        if (!context) throw new Error('Could not prepare the gate pass PDF image.');
        context.drawImage(canvas, 0, sourceY, canvas.width, sliceHeight, 0, 0, canvas.width, sliceHeight);

        if (pageIndex > 0) pdf.addPage();
        pdf.addImage(
          pageCanvas.toDataURL('image/jpeg', 0.94),
          'JPEG',
          margin,
          margin,
          pageWidth,
          (sliceHeight * pageWidth) / canvas.width,
          undefined,
          'FAST',
        );
      }

      pdf.save(`${filename}.pdf`);
    } catch (error: any) {
      setToast(error.message || `Could not download gate pass as ${format.toUpperCase()}.`);
    } finally {
      setExportingPass(false);
    }
  };

  const fillFormForEdit = (id: number) => {
    const pass = gates.find((gate) => gate.id === id);
    if (!pass) {
      setToast('Gate pass could not be found in the selected Excel file.');
      return;
    }
    setForm({
      gate_pass_number: pass.gate_pass_number,
      date: pass.date,
      person_name: pass.person_name,
      mobile_number: pass.mobile_number,
      company_name: pass.company_name || '',
      taken_by: pass.taken_by || '',
      authorized_by: pass.authorized_by || '',
      status: pass.status || 'Returnable',
      items: (pass.items || []).map((item: any, index: number) => ({
        id: index + 1,
        description: item.description || item.item_serial_number || item.purpose || '',
        pieces: String(item.pieces ?? (item.unit === 'Pcs' ? item.quantity ?? '' : '')),
        feet: String(item.feet ?? (item.unit === 'Feet' ? item.quantity ?? '' : '')),
        rate: String(item.rate ?? ''),
      })),
    });
    setEditingId(id);
    setActivePage('create');
  };

  const deleteGatePass = async (id: number) => {
    const confirmed = window.confirm('Delete this gate pass?');
    if (!confirmed) return;
    if (!requireWorkbook()) return;
    try {
      await (await import('./workbookStore')).deleteGatePass(id);
      setGates((current) => current.filter((pass) => pass.id !== id));
      setToast('Gate pass deleted from the Excel file.');
    } catch (error: any) {
      setToast(error.message || 'Delete failed.');
    }
  };

  const duplicateGatePass = async (id: number) => {
    if (!requireWorkbook()) return;
    try {
      const pass = await (await import('./workbookStore')).duplicateGatePass(id);
      setGates((current) => [...current, pass as GatePass]);
      setToast(`Created ${pass.gate_pass_number} in the Excel file.`);
    } catch (error: any) {
      setToast(error.message || 'Duplicate failed.');
    }
  };

  const exportExcel = async () => {
    if (!requireWorkbook()) return;
    try {
      const reportPasses = gates.filter((pass) => {
        const statusMatches = reportFilter === 'returnable'
          ? pass.status === 'Returnable'
          : reportFilter === 'nonreturnable' ? pass.status === 'Non-Returnable' : true;
        const rangeMatches = reportFilter === 'today'
          ? pass.date === today
          : (!reportRange.start || pass.date >= reportRange.start)
            && (!reportRange.end || pass.date <= reportRange.end);
        return statusMatches && rangeMatches;
      });
      await (await import('./workbookStore')).downloadReportWorkbook(reportPasses);
      setToast('Excel report downloaded.');
    } catch (error: any) {
      setToast(error.message || 'Excel export failed.');
    }
  };

  const saveSettings = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (!requireWorkbook()) return;
    try {
      let logoPath = settings.logo_path;
      const logoInput = document.getElementById('logo-upload') as HTMLInputElement | null;
      const logoFile = logoInput?.files?.[0];
      if (logoFile) {
        if (logoFile.size > 30 * 1024) throw new Error('Logo must be 30 KB or smaller.');
        logoPath = await new Promise<string>((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => typeof reader.result === 'string'
            ? resolve(reader.result)
            : reject(new Error('Could not read the logo file.'));
          reader.onerror = () => reject(new Error('Could not read the logo file.'));
          reader.readAsDataURL(logoFile);
        });
      }
      const nextSettings = { ...settings, logo_path: logoPath };
      await (await import('./workbookStore')).saveWorkbookSettings(nextSettings);
      setSettings(nextSettings);
      if (logoInput) logoInput.value = '';
      setToast('Settings saved in the selected Excel file.');
    } catch (error: any) {
      setToast(error.message || 'Settings update failed.');
    }
  };

  const downloadWorkbook = async () => {
    if (!requireWorkbook()) return;
    try {
      await (await import('./workbookStore')).downloadWorkbookCopy();
      setToast('Excel copy downloaded.');
    } catch (error: any) {
      setToast(error.message || 'Could not download Excel file.');
    }
  };

  const buildPassDocument = (pass: any) => ({
    gatePassNumber: pass.gate_pass_number || 'DRAFT',
    personName: pass.person_name || '—',
    mobileNumber: pass.mobile_number || '—',
    companyName: pass.company_name || settings.company_name || 'ISLAM MARBLE',
    date: pass.date || new Date().toISOString().slice(0, 10),
    status: pass.status || 'Returnable',
    takenBy: pass.taken_by || '—',
    authorizedBy: pass.authorized_by || '—',
    items: pass.items || [],
  });

  return (
    <>
      <div className="app-shell">
        <aside className="sidebar">
          <div className="brand">
            <div className="brand-badge">IM</div>
            <div>
              <p>ISLAM MARBLE</p>
              <small>Gate Pass System</small>
            </div>
          </div>

          <nav className="nav-menu">
            <button className={activePage === 'dashboard' ? 'nav-item active' : 'nav-item'} onClick={() => setActivePage('dashboard')}>
              <LayoutDashboard size={18} /> Dashboard
            </button>
            <button className={activePage === 'create' ? 'nav-item active' : 'nav-item'} onClick={() => setActivePage('create')}>
              <Plus size={18} /> Create Gate Pass
            </button>
            <button className={activePage === 'records' ? 'nav-item active' : 'nav-item'} onClick={() => setActivePage('records')}>
              <FileText size={18} /> All Gate Passes
            </button>
            <button className={activePage === 'reports' ? 'nav-item active' : 'nav-item'} onClick={() => setActivePage('reports')}>
              <FileSpreadsheet size={18} /> Excel Reports
            </button>
            <button className={activePage === 'settings' ? 'nav-item active' : 'nav-item'} onClick={() => setActivePage('settings')}>
              <Settings size={18} /> Settings
            </button>
          </nav>

        </aside>

        <main className="main-panel">
          <header className="topbar">
            <div>
              <p className="eyebrow">Islam Marble</p>
              <h1>{activePage === 'dashboard' ? 'Dashboard' : activePage === 'create' ? 'Create Gate Pass' : activePage === 'records' ? 'All Gate Passes' : activePage === 'reports' ? 'Excel Reports' : 'Settings'}</h1>
            </div>
            <div className="top-actions">
              <span className={workbookName ? 'workbook-label active' : 'workbook-label'}>
                {workbookName || 'No Excel file selected'}
              </span>
              <button className="secondary-btn" type="button" onClick={() => selectWorkbook('open')} disabled={workbookBusy}>
                <FileSpreadsheet size={16} /> Open Excel
              </button>
              <button className="secondary-btn" type="button" onClick={() => selectWorkbook('create')} disabled={workbookBusy}>
                <Plus size={16} /> New Excel
              </button>
              <button className="primary-btn" type="button" onClick={() => setActivePage('create')}>
                <Plus size={16} /> New Gate Pass
              </button>
            </div>
          </header>

          {toast && <div className="toast">{toast}</div>}
          {!workbookName && (
            <div className="workbook-notice">
              Open or create an Excel file to load and save gate passes. Data stays in that file on your computer.
            </div>
          )}

          {activePage === 'dashboard' && (
            <section className="page-section">
              <div className="dashboard-summary">
                <div><span>Total gate passes</span><strong>{dashboard?.totalGatePasses ?? 0}</strong></div>
                <div><span>Issued today</span><strong>{dashboard?.todaysPasses ?? 0}</strong></div>
              </div>

              <div className="panel">
                <div className="section-header">
                  <h2>Recent Gate Passes</h2>
                  <button className="text-btn" onClick={() => setActivePage('records')}>View all</button>
                </div>
                <div className="table-wrap">
                  <table>
                    <thead>
                      <tr>
                        <th>Gate Pass Number</th>
                        <th>Date</th>
                        <th>Person Name</th>
                        <th>Total Items</th>
                        <th>Status</th>
                        <th>Actions</th>
                      </tr>
                    </thead>
                    <tbody>
                      {(dashboard?.recent || []).map((pass: any) => (
                        <tr key={pass.id}>
                          <td>{pass.gate_pass_number}</td>
                          <td>{pass.date}</td>
                          <td>{pass.person_name}</td>
                          <td>{pass.total_items}</td>
                          <td><span className={statusClass(pass.status)}>{pass.status}</span></td>
                          <td className="table-actions">
                            <button className="icon-btn" onClick={() => setPreviewPass(pass)}><Printer size={15} /></button>
                            <button className="icon-btn" onClick={() => fillFormForEdit(pass.id)}><Pencil size={15} /></button>
                          </td>
                        </tr>
                      ))}
                      {!dashboard?.recent?.length && (
                        <tr><td className="empty-state" colSpan={6}>No gate passes yet.</td></tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>
          )}

          {activePage === 'create' && (
            <section className="page-section form-section">
              <div className="panel">
                <div className="section-header">
                  <h2>{editingId ? 'Edit Gate Pass' : 'Create New Gate Pass'}</h2>
                  <button className="text-btn" onClick={resetForm}>Reset form</button>
                </div>

                <div className="form-grid two-col">
                  <label>
                    Person Name
                    <input value={form.person_name} onChange={(e) => handleFormUpdate('person_name', e.target.value)} />
                  </label>
                  <label>
                    Mobile Number
                    <input value={form.mobile_number} onChange={(e) => handleFormUpdate('mobile_number', e.target.value)} />
                  </label>
                  <label>
                    Company Name
                    <input value={form.company_name} onChange={(e) => handleFormUpdate('company_name', e.target.value)} placeholder={settings.company_name} />
                  </label>
                  <label>
                    Taken By
                    <input value={form.taken_by} onChange={(e) => handleFormUpdate('taken_by', e.target.value)} />
                  </label>
                  <label>
                    Authorized By
                    <input value={form.authorized_by} onChange={(e) => handleFormUpdate('authorized_by', e.target.value)} />
                  </label>
                  <label>
                    Status
                    <select value={form.status} onChange={(e) => handleFormUpdate('status', e.target.value)}>
                      <option value="Returnable">Returnable</option>
                      <option value="Non-Returnable">Non-Returnable</option>
                    </select>
                  </label>
                  <label>
                    Gate Pass Number
                    <input value={form.gate_pass_number || 'Auto generated'} readOnly />
                  </label>
                  <label>
                    Date
                    <input type="date" value={form.date} onChange={(e) => handleFormUpdate('date', e.target.value)} />
                  </label>
                </div>

                <div className="table-wrap material-table">
                  <table>
                    <thead>
                      <tr>
                        <th>S. No.</th>
                        <th>Description</th>
                        <th>Pieces</th>
                        <th>Feet</th>
                        <th>Rate / Ft</th>
                        <th>Amount</th>
                        <th>Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {form.items.map((item, index) => (
                        <tr key={item.id}>
                          <td>{index + 1}</td>
                          <td><input value={item.description} onChange={(e) => handleMaterialUpdate(index, 'description', e.target.value)} /></td>
                          <td><input type="number" min="0" step="1" value={item.pieces} onChange={(e) => handleMaterialUpdate(index, 'pieces', e.target.value)} /></td>
                          <td><input type="number" min="0" step="0.01" value={item.feet} onChange={(e) => handleMaterialUpdate(index, 'feet', e.target.value)} /></td>
                          <td><input type="number" min="0" step="0.01" value={item.rate} onChange={(e) => handleMaterialUpdate(index, 'rate', e.target.value)} /></td>
                          <td>{(Number(item.feet || 0) * Number(item.rate || 0)).toFixed(2)}</td>
                          <td><button className="danger-btn" type="button" onClick={() => removeMaterialRow(index)}>Remove</button></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                <div className="form-actions">
                  <button className="primary-btn" type="button" onClick={() => saveGatePass(false)}>Save Gate Pass</button>
                  <button className="secondary-btn" type="button" onClick={() => saveGatePass(true)}>Save & Print</button>
                  <button className="secondary-btn" type="button" onClick={previewForm}>Preview</button>
                  <button className="secondary-btn" type="button" onClick={resetForm}>Reset</button>
                  <button className="ghost-btn" type="button" onClick={() => setActivePage('dashboard')}>Cancel</button>
                  <button className="secondary-btn" type="button" onClick={addMaterialRow}>Add Row</button>
                </div>
              </div>
            </section>
          )}

          {activePage === 'records' && (
            <section className="page-section">
              <div className="panel">
                <div className="section-header">
                  <h2>All Gate Passes</h2>
                  <div className="record-toolbar">
                    <div className="search-box">
                      <Search size={15} />
                      <input value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} placeholder="Search by name, number, mobile, item" />
                    </div>
                    <select value={recordFilter} onChange={(e) => setRecordFilter(e.target.value)}>
                      <option value="all">All records</option>
                      <option value="today">Today</option>
                      <option value="yesterday">Yesterday</option>
                      <option value="this-week">This Week</option>
                      <option value="this-month">This Month</option>
                    </select>
                  </div>
                </div>

                <div className="table-wrap">
                  <table>
                    <thead>
                      <tr>
                        <th>Gate Pass Number</th>
                        <th>Date</th>
                        <th>Person Name</th>
                        <th>Mobile Number</th>
                        <th>Company Name</th>
                        <th>Total Items</th>
                        <th>Total Pcs</th>
                        <th>Total Feet</th>
                        <th>Total Amount</th>
                        <th>Status</th>
                        <th>Actions</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredGates.map((pass) => (
                        <tr key={pass.id}>
                          <td>{pass.gate_pass_number}</td>
                          <td>{pass.date}</td>
                          <td>{pass.person_name}</td>
                          <td>{pass.mobile_number}</td>
                          <td>{pass.company_name}</td>
                          <td>{pass.total_items || 0}</td>
                          <td>{pass.total_pcs || 0}</td>
                          <td>{pass.total_feet || 0}</td>
                          <td>{Number(pass.total_amount || 0).toFixed(2)}</td>
                          <td><span className={statusClass(pass.status)}>{pass.status}</span></td>
                          <td className="row-actions">
                            <button className="icon-btn" onClick={() => setPreviewPass(pass)}><Printer size={15} /></button>
                            <button className="icon-btn" onClick={() => fillFormForEdit(pass.id)}><Pencil size={15} /></button>
                            <button className="icon-btn" onClick={() => duplicateGatePass(pass.id)}><ArrowUpRight size={15} /></button>
                            <button className="icon-btn" onClick={() => deleteGatePass(pass.id)}><Trash2 size={15} /></button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>
          )}

          {activePage === 'reports' && (
            <section className="page-section">
              <div className="panel export-panel">
                <h2>Excel Reports</h2>
                <div className="report-filters">
                  <select value={reportFilter} onChange={(e) => setReportFilter(e.target.value)}>
                    <option value="all">All Records</option>
                    <option value="today">Today</option>
                    <option value="this-month">This Month</option>
                    <option value="returnable">Returnable Only</option>
                    <option value="nonreturnable">Non-Returnable Only</option>
                  </select>
                  <input type="date" value={reportRange.start} onChange={(e) => setReportRange((prev) => ({ ...prev, start: e.target.value }))} />
                  <input type="date" value={reportRange.end} onChange={(e) => setReportRange((prev) => ({ ...prev, end: e.target.value }))} />
                </div>
                <button className="primary-btn" type="button" onClick={exportExcel}>
                  <Download size={16} /> Export Workbook
                </button>
              </div>
            </section>
          )}

          {activePage === 'settings' && (
            <section className="page-section">
              <div className="panel settings-panel">
                <h2>Company Settings</h2>
                <form onSubmit={saveSettings} className="settings-form">
                  <label>Company Name<input value={settings.company_name} onChange={(e) => setSettings((prev) => ({ ...prev, company_name: e.target.value }))} /></label>
                  <label>Company Address<textarea value={settings.address} onChange={(e) => setSettings((prev) => ({ ...prev, address: e.target.value }))} /></label>
                  <label>Contact Number<input value={settings.contact_number} onChange={(e) => setSettings((prev) => ({ ...prev, contact_number: e.target.value }))} /></label>
                  <label>Gate Pass Prefix<input value={settings.gate_pass_prefix} onChange={(e) => setSettings((prev) => ({ ...prev, gate_pass_prefix: e.target.value }))} /></label>
                  <label>
                    Default Status
                    <select value={settings.default_status} onChange={(e) => setSettings((prev) => ({ ...prev, default_status: e.target.value as 'Returnable' | 'Non-Returnable' }))}>
                      <option value="Returnable">Returnable</option>
                      <option value="Non-Returnable">Non-Returnable</option>
                    </select>
                  </label>
                  <label className="file-upload">
                    Company Logo (PNG/JPEG/WebP, max 30 KB)
                    <input id="logo-upload" type="file" accept="image/png,image/jpeg,image/webp" />
                  </label>
                  {settings.logo_path && <img className="company-logo" src={settings.logo_path} alt="company logo" />}
                  <div className="settings-actions">
                    <button className="primary-btn" type="submit">Save Settings</button>
                    <button className="secondary-btn" type="button" onClick={downloadWorkbook}>Download Excel Copy</button>
                  </div>
                </form>
              </div>
            </section>
          )}
        </main>
      </div>

      {previewPass && (
        <div className="print-overlay" onClick={() => setPreviewPass(null)}>
          <div className="print-modal" onClick={(e) => e.stopPropagation()}>
            <div className="print-toolbar">
              <button className="secondary-btn" type="button" onClick={() => window.print()}>Print</button>
              <button className="secondary-btn" type="button" onClick={() => downloadPass('pdf')} disabled={exportingPass}>
                <Download size={16} /> {exportingPass ? 'Preparing...' : 'Download PDF'}
              </button>
              <button className="secondary-btn" type="button" onClick={() => downloadPass('png')} disabled={exportingPass}>
                <Download size={16} /> {exportingPass ? 'Preparing...' : 'Download Image'}
              </button>
              <button className="ghost-btn" type="button" onClick={() => setPreviewPass(null)}>Close</button>
            </div>
            <div className="print-document print-through" ref={printDocumentRef}>
              <div className="doc-header">
                <div className="doc-brand">
                  {settings.logo_path ? <img src={settings.logo_path} alt="Islam Marble logo" className="doc-logo" /> : <div className="doc-logo text">IM</div>}
                  <div className="doc-company">
                    <h2>{settings.company_name || 'ISLAM MARBLE'}</h2>
                  </div>
                </div>
                <div className="doc-contact">
                  <p><span>Address</span><strong>{settings.address || '—'}</strong></p>
                  <p><span>Number</span><strong>{settings.contact_number || '—'}</strong></p>
                </div>
              </div>

              <div className="doc-title-row">
                <h3 className="doc-title">MATERIAL OUTWARD GATE PASS</h3>
                <span className={`doc-status ${buildPassDocument(previewPass).status === 'Returnable' ? 'returnable' : 'non-returnable'}`}>
                  {buildPassDocument(previewPass).status}
                </span>
              </div>

              <div className="doc-meta">
                <div><span>Gate Pass No.</span><strong>{buildPassDocument(previewPass).gatePassNumber}</strong></div>
                <div><span>Date</span><strong>{buildPassDocument(previewPass).date}</strong></div>
                <div><span>Person Name</span><strong>{buildPassDocument(previewPass).personName}</strong></div>
                <div><span>Mobile Number</span><strong>{buildPassDocument(previewPass).mobileNumber}</strong></div>
                <div><span>Company</span><strong>{buildPassDocument(previewPass).companyName}</strong></div>
                <div><span>Taken By</span><strong>{buildPassDocument(previewPass).takenBy}</strong></div>
                <div><span>Authorized By</span><strong>{buildPassDocument(previewPass).authorizedBy}</strong></div>
              </div>

              <table className="gate-table">
                <colgroup>
                  <col className="gate-col-serial" />
                  <col className="gate-col-description" />
                  <col className="gate-col-pieces" />
                  <col className="gate-col-feet" />
                  <col className="gate-col-rate" />
                  <col className="gate-col-amount" />
                </colgroup>
                <thead>
                  <tr>
                    <th>S. No.</th>
                    <th>Description</th>
                    <th>Pieces</th>
                    <th>Feet</th>
                    <th>Rate / Ft</th>
                    <th>Amount</th>
                  </tr>
                </thead>
                <tbody>
                  {(buildPassDocument(previewPass).items || []).map((row: any, index: number) => (
                    <tr key={`${row.id ?? index}`}>
                      <td>{index + 1}</td>
                      <td>{row.description || row.item_serial_number || row.purpose || ''}</td>
                      <td>{row.pieces ?? (row.unit === 'Pcs' ? row.quantity : '')}</td>
                      <td>{row.feet ?? (row.unit === 'Feet' ? row.quantity : '')}</td>
                      <td>{row.rate ?? ''}</td>
                      <td>{(Number(row.feet ?? (row.unit === 'Feet' ? row.quantity : 0) ?? 0) * Number(row.rate || 0)).toFixed(2)}</td>
                    </tr>
                  ))}
                  {Array.from({ length: Math.max(0, 15 - (buildPassDocument(previewPass).items || []).length) }).map((_, idx) => (
                    <tr key={`empty-${idx}`}>
                      <td>&nbsp;</td><td></td><td></td><td></td><td></td><td></td>
                    </tr>
                  ))}
                  <tr>
                    <td colSpan={2}><strong>Totals</strong></td>
                    <td><strong>{(buildPassDocument(previewPass).items || []).reduce((total: number, row: any) => (
                      total + Number(row.pieces ?? (row.unit === 'Pcs' ? row.quantity : 0) ?? 0)
                    ), 0)}</strong></td>
                    <td><strong>{(buildPassDocument(previewPass).items || []).reduce((total: number, row: any) => (
                      total + Number(row.feet ?? (row.unit === 'Feet' ? row.quantity : 0) ?? 0)
                    ), 0)}</strong></td>
                    <td></td>
                    <td><strong>{(buildPassDocument(previewPass).items || []).reduce((total: number, row: any) => (
                      total + Number(row.amount ?? Number(row.feet ?? (row.unit === 'Feet' ? row.quantity : 0) ?? 0) * Number(row.rate || 0))
                    ), 0).toFixed(2)}</strong></td>
                  </tr>
                </tbody>
              </table>

              <div className="doc-signatures">
                <div><span>Prepared By</span></div>
                <div><span>Received By</span></div>
                <div><span>Authorized Signature</span></div>
              </div>
            </div>
          </div>
        </div>
      )}
    </>
  );
}

export default App;
