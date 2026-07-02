import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  Activity,
  FileText,
  Calendar,
  ClipboardList,
  Pill,
  Stethoscope,
  Search,
  FileSpreadsheet,
  ChevronLeft,
  ChevronRight,
  Printer,
  Settings,
  User,
  Sun,
  Moon,
  Upload,
  AlertCircle,
  ExternalLink,
  HeartPulse,
  Clock,
  ShieldAlert,
  Download
} from 'lucide-react';
import type { Patient, DatosActa } from './types';
import demoPatients from './data/demo_patients.json';
import './App.css';

// Checkbox items for visual filtering
interface FieldConfig {
  key: keyof typeof DEFAULT_VISIBLE_FIELDS;
  label: string;
}

const DEFAULT_VISIBLE_FIELDS = {
  edad: true,
  genero: true,
  nivel_riesgo: true,
  analisis: true,
  antecedentes: true,
  metricas_relevantes: true,
  historias_anteriores: true,
  atenciones_programadas: true,
  prog_especialidad_acta: true,
  diagnosticos: true,
  medicamentos_prescritos: true
};

const FIELDS_CONFIG: FieldConfig[] = [
  { key: 'edad', label: 'Edad' },
  { key: 'genero', label: 'Género' },
  { key: 'nivel_riesgo', label: 'Nivel de Riesgo' },
  { key: 'analisis', label: 'Análisis y Notas' },
  { key: 'antecedentes', label: 'Antecedentes' },
  { key: 'metricas_relevantes', label: 'Métricas Relevantes' },
  { key: 'historias_anteriores', label: 'Historias Anteriores' },
  { key: 'prog_especialidad_acta', label: 'Programación de Especialidades' },
  { key: 'diagnosticos', label: 'Diagnósticos' },
  { key: 'medicamentos_prescritos', label: 'Medicamentos Prescritos' }
];

function App() {
  // Theme state
  const [theme, setTheme] = useState<'light' | 'dark'>(() => {
    const savedTheme = localStorage.getItem('theme');
    if (savedTheme === 'light' || savedTheme === 'dark') return savedTheme;
    return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  });

  // Patient data state
  const [patients, setPatients] = useState<Patient[]>([]);
  const [selectedPatientIndex, setSelectedPatientIndex] = useState<number>(0);
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [activeTab, setActiveTab] = useState<string>('generales');

  // Checklist of visible fields state
  const [visibleFields, setVisibleFields] = useState<Record<string, boolean>>(() => {
    const saved = localStorage.getItem('visibleFields');
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        // Fallback
      }
    }
    return DEFAULT_VISIBLE_FIELDS;
  });

  // Sync theme to body element
  useEffect(() => {
    const root = document.documentElement;
    if (theme === 'dark') {
      root.classList.add('dark-theme');
    } else {
      root.classList.remove('dark-theme');
    }
    localStorage.setItem('theme', theme);
  }, [theme]);

  // Sync visibleFields to localStorage
  useEffect(() => {
    localStorage.setItem('visibleFields', JSON.stringify(visibleFields));
  }, [visibleFields]);

  const toggleTheme = () => {
    setTheme(prev => prev === 'light' ? 'dark' : 'light');
  };

  // Helper to load mock demo data
  const loadDemoData = () => {
    const typedDemo = (demoPatients as any[]).map(p => ({
      identificacion: p.identificacion,
      nombre_completo: p.nombre_completo,
      nombre_convenio: p.nombre_convenio,
      datos_acta: p.datos_acta as DatosActa | null
    }));
    setPatients(typedDemo);
    setSelectedPatientIndex(0);
  };

  // Excel parser
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const binaryString = event.target?.result;
        const workbook = XLSX.read(binaryString, { type: 'binary' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json<any>(worksheet);

        console.log('Parsed rows:', jsonData.length);

        const parsedPatients: Patient[] = jsonData.map((row) => {
          let parsedActa: DatosActa | null = null;
          if (row.datos_acta) {
            try {
              // Handle string JSON vs object
              parsedActa = typeof row.datos_acta === 'string' ? JSON.parse(row.datos_acta) : row.datos_acta;
            } catch (err) {
              console.error('Error parsing datos_acta cell JSON:', err);
            }
          }
          return {
            identificacion: row.identificacion ? String(row.identificacion).trim() : '',
            nombre_completo: row.nombre_completo ? String(row.nombre_completo).trim() : '',
            nombre_convenio: row.nombre_convenio ? String(row.nombre_convenio).trim() : '',
            datos_acta: parsedActa
          };
        });

        if (parsedPatients.length > 0) {
          setPatients(parsedPatients);
          setSelectedPatientIndex(0);
        } else {
          alert('No se encontraron filas de datos válidas en el Excel.');
        }
      } catch (err) {
        console.error(err);
        alert('Error al leer el archivo Excel. Asegúrese de que el formato sea correcto.');
      }
    };
    reader.readAsBinaryString(file);
  };

  // Filter patients list by search query
  const filteredPatients = useMemo(() => {
    if (!searchQuery.trim()) return patients;
    const query = searchQuery.toLowerCase().trim();
    return patients.filter(
      p =>
        p.nombre_completo.toLowerCase().includes(query) ||
        p.identificacion.toLowerCase().includes(query)
    );
  }, [patients, searchQuery]);

  // Active patient object
  const activePatient = useMemo(() => {
    if (filteredPatients.length === 0) return null;
    // Bound check
    const index = Math.min(selectedPatientIndex, filteredPatients.length - 1);
    return filteredPatients[index >= 0 ? index : 0];
  }, [filteredPatients, selectedPatientIndex]);

  const handleCheckboxChange = (key: string) => {
    setVisibleFields(prev => ({
      ...prev,
      [key]: !prev[key]
    }));
  };

  // Navigation handlers
  const handlePrevPatient = () => {
    setSelectedPatientIndex(prev => Math.max(0, prev - 1));
  };

  const handleNextPatient = () => {
    setSelectedPatientIndex(prev => Math.min(filteredPatients.length - 1, prev + 1));
  };

  // Helper to render clinical values cleanly
  const renderVal = (val: any) => {
    if (val === null || val === undefined || String(val).trim() === '') {
      return <span className="kv-value null-value">Sin información</span>;
    }

    const strVal = String(val).trim();
    // Check if it's a PDF link
    if (strVal.startsWith('http://') || strVal.startsWith('https://')) {
      if (strVal.toLowerCase().endsWith('.pdf') || strVal.includes('pdf')) {
        return (
          <a href={strVal} target="_blank" rel="noopener noreferrer" className="btn-pdf">
            <Download size={14} /> Abrir PDF
          </a>
        );
      }
      return (
        <a href={strVal} target="_blank" rel="noopener noreferrer" style={{ display: 'inline-flex', alignItems: 'center', gap: '0.25rem' }}>
          Enlace <ExternalLink size={12} />
        </a>
      );
    }

    return <span className="kv-value">{strVal}</span>;
  };

  // Risk display helper
  const getRiskBadge = (level: number | null | undefined) => {
    if (level === null || level === undefined) {
      return (
        <span className="risk-badge risk-level-null">
          <span className="risk-dot"></span> Riesgo: Sin información
        </span>
      );
    }

    switch (level) {
      case 1:
        return (
          <span className="risk-badge risk-level-1">
            <span className="risk-dot" style={{ backgroundColor: 'var(--risk-low)' }}></span> Riesgo: Bajo (1)
          </span>
        );
      case 2:
        return (
          <span className="risk-badge risk-level-2">
            <span className="risk-dot" style={{ backgroundColor: 'var(--risk-medium)' }}></span> Riesgo: Medio (2)
          </span>
        );
      case 3:
        return (
          <span className="risk-badge risk-level-3">
            <span className="risk-dot" style={{ backgroundColor: 'var(--risk-high)' }}></span> Riesgo: Alto (3)
          </span>
        );
      case 4:
        return (
          <span className="risk-badge risk-level-4">
            <span className="risk-dot" style={{ backgroundColor: 'var(--risk-very-high)' }}></span> Riesgo: Muy Alto (4)
          </span>
        );
      default:
        return (
          <span className="risk-badge risk-level-null">
            <span className="risk-dot"></span> Riesgo: {level}
          </span>
        );
    }
  };

  // Format Antecedentes key titles
  const formatAntecedenteTitle = (key: string) => {
    const titles: Record<string, string> = {
      antecedentes_familiares: 'Familiares',
      antecedentes_patologicos: 'Patológicos',
      antecedentes_personales: 'Personales',
      antecedentes_quirurgicos: 'Quirúrgicos',
      antecedentes_traumaticos: 'Traumáticos',
      antecedentes_farmacologicos: 'Farmacológicos'
    };
    return titles[key] || key.replace('antecedentes_', '').toUpperCase();
  };

  // Trigger browser print modal
  const handlePrint = () => {
    window.print();
  };

  // Render core sections based on checkbox settings
  const renderSection = (key: string, data: DatosActa) => {
    if (!visibleFields[key]) return null;

    switch (key) {
      case 'edad':
        return (
          <div key="edad" className="kv-item">
            <span className="kv-label">Edad</span>
            {renderVal(data.edad)}
          </div>
        );
      case 'genero':
        return (
          <div key="genero" className="kv-item">
            <span className="kv-label">Género</span>
            {renderVal(data.genero)}
          </div>
        );
      case 'nivel_riesgo':
        return (
          <div key="nivel_riesgo" className="kv-item">
            <span className="kv-label">Riesgo Clínico</span>
            <div>{getRiskBadge(data.nivel_riesgo)}</div>
          </div>
        );
      case 'metricas_relevantes':
        return (
          <div key="metricas_relevantes" className="info-card">
            <div className="info-card-header">
              <Activity className="info-card-icon" size={18} />
              <span className="info-card-title">Métricas Relevantes</span>
            </div>
            <div className="kv-grid">
              <div className="kv-item">
                <span className="kv-label">Presión Arterial</span>
                {renderVal(data.metricas_relevantes?.presion_arterial)}
              </div>
              <div className="kv-item">
                <span className="kv-label">Nivel de Azúcar</span>
                {renderVal(data.metricas_relevantes?.nivel_azucar)}
              </div>
              <div className="kv-item">
                <span className="kv-label">HbA1c</span>
                {renderVal(data.metricas_relevantes?.hba1c)}
              </div>
            </div>
          </div>
        );
      case 'analisis':
        return (
          <div key="analisis" className="info-card">
            <div className="info-card-header">
              <ClipboardList className="info-card-icon" size={18} />
              <span className="info-card-title">Análisis Médico y Notas</span>
            </div>
            <div className="kv-grid" style={{ gridTemplateColumns: '1fr' }}>
              {data.analisis && (
                <div className="kv-item">
                  <span className="kv-label">Análisis Clínico</span>
                  <div className="text-block">{data.analisis}</div>
                </div>
              )}
              {data.observaciones_clinicas && (
                <div className="kv-item">
                  <span className="kv-label">Observaciones Clínicas</span>
                  <div className="text-block">{data.observaciones_clinicas}</div>
                </div>
              )}
              {data.observaciones_operativas && (
                <div className="kv-item">
                  <span className="kv-label">Observaciones Operativas</span>
                  <div className="text-block">{data.observaciones_operativas}</div>
                </div>
              )}
              {!data.analisis && !data.observaciones_clinicas && !data.observaciones_operativas && (
                <span className="kv-value null-value">Sin registros de notas o análisis.</span>
              )}
            </div>
          </div>
        );
      case 'antecedentes':
        return (
          <div key="antecedentes" className="info-card">
            <div className="info-card-header">
              <ShieldAlert className="info-card-icon" size={18} />
              <span className="info-card-title">Antecedentes Clínicos</span>
            </div>
            <div className="kv-grid">
              {data.antecedentes && Object.entries(data.antecedentes).map(([antKey, value]) => (
                <div key={antKey} className="kv-item">
                  <span className="kv-label">Antecedentes {formatAntecedenteTitle(antKey)}</span>
                  {renderVal(value)}
                </div>
              ))}
              {!data.antecedentes && (
                <span className="kv-value null-value">Sin antecedentes registrados.</span>
              )}
            </div>
          </div>
        );
      case 'atenciones_programadas':
        return (
          <div key="atenciones_programadas" className="info-card">
            <div className="info-card-header">
              <Calendar className="info-card-icon" size={18} />
              <span className="info-card-title">Atenciones Programadas</span>
            </div>
            {data.atenciones_programadas && data.atenciones_programadas.length > 0 ? (
              <div className="table-wrapper">
                <table className="clinical-table">
                  <thead>
                    <tr>
                      <th>Código</th>
                      <th>Fecha de Cita</th>
                      <th>Especialidad</th>
                      <th>Profesional</th>
                      <th>Estado</th>
                    </tr>
                  </thead>
                  <tbody>
                    {data.atenciones_programadas.map((atencion, idx) => (
                      <tr key={idx}>
                        <td><code>{atencion.codigo_cita}</code></td>
                        <td>{atencion.fecha_cita}</td>
                        <td><strong>{atencion.nombre_especialidad}</strong></td>
                        <td>{atencion.nombre_profesional}</td>
                        <td>
                          <span className={`risk-badge`} style={{
                            padding: '0.15rem 0.5rem', fontSize: '0.75rem', display: 'inline-flex',
                            backgroundColor: atencion.estado_cita.includes('Pendiente') ? 'rgba(245, 158, 11, 0.1)' :
                              atencion.estado_cita.includes('Confirm') || atencion.estado_cita.includes('Pagada') ? 'rgba(16, 185, 129, 0.1)' : 'rgba(148, 163, 184, 0.1)',
                            color: atencion.estado_cita.includes('Pendiente') ? 'var(--risk-medium)' :
                              atencion.estado_cita.includes('Confirm') || atencion.estado_cita.includes('Pagada') ? 'var(--risk-low)' : 'var(--text-secondary)',
                            border: 'none', boxShadow: 'none'
                          }}>
                            {atencion.estado_cita}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <span className="kv-value null-value">No hay atenciones programadas.</span>
            )}
          </div>
        );
      case 'prog_especialidad_acta':
        return (
          <div key="prog_especialidad_acta" className="info-card">
            <div className="info-card-header">
              <Clock className="info-card-icon" size={18} />
              <span className="info-card-title">Programación de Especialidades</span>
            </div>
            {data.prog_especialidad_acta && data.prog_especialidad_acta.length > 0 ? (
              <div className="table-wrapper">
                <table className="clinical-table">
                  <thead>
                    <tr>
                      <th>ID</th>
                      <th>Especialidad</th>
                      <th>Fecha Inicio</th>
                      <th>Días Previstos</th>
                    </tr>
                  </thead>
                  <tbody>
                    {data.prog_especialidad_acta.map((prog, idx) => (
                      <tr key={idx}>
                        <td><code>{prog.id}</code></td>
                        <td><strong>{prog.nombre_especialidad}</strong></td>
                        <td>{prog.fecha_inicio}</td>
                        <td>{prog.dias || 'Sin días definidos'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <span className="kv-value null-value">Sin programación de especialidad registrada.</span>
            )}
          </div>
        );
      case 'historias_anteriores':
        return (
          <div key="historias_anteriores" className="info-card">
            <div className="info-card-header">
              <FileText className="info-card-icon" size={18} />
              <span className="info-card-title">Historias Anteriores y Archivos</span>
            </div>
            {data.historias_anteriores && data.historias_anteriores.length > 0 ? (
              <div className="table-wrapper">
                <table className="clinical-table">
                  <thead>
                    <tr>
                      <th>Fecha</th>
                      <th>Especialidad</th>
                      <th>Aliado</th>
                      <th>Acción</th>
                    </tr>
                  </thead>
                  <tbody>
                    {data.historias_anteriores.map((hist, idx) => (
                      <tr key={idx}>
                        <td>{hist.fecha_atencion}</td>
                        <td><strong>{hist.nombre_especialidad}</strong></td>
                        <td>{hist.nombre_aliado}</td>
                        <td>{renderVal(hist.url_archivo)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <span className="kv-value null-value">No se encontraron archivos de historias anteriores.</span>
            )}
          </div>
        );
      case 'diagnosticos':
        return (
          <div key="diagnosticos" className="info-card">
            <div className="info-card-header">
              <Stethoscope className="info-card-icon" size={18} />
              <span className="info-card-title">Diagnósticos Activos</span>
            </div>
            <div className="text-block">
              {data.diagnosticos ? renderVal(data.diagnosticos) : <span className="null-value">Sin diagnósticos activos en esta acta.</span>}
            </div>
          </div>
        );
      case 'medicamentos_prescritos':
        return (
          <div key="medicamentos_prescritos" className="info-card">
            <div className="info-card-header">
              <Pill className="info-card-icon" size={18} />
              <span className="info-card-title">Medicamentos Prescritos</span>
            </div>
            <div className="text-block">
              {data.medicamentos_prescritos ? renderVal(data.medicamentos_prescritos) : <span className="null-value">Sin medicamentos recetados vigentes.</span>}
            </div>
          </div>
        );
      default:
        return null;
    }
  };

  // Filter tabs content visibility based on visibleFields
  const isTabSectionVisible = (sections: string[]) => {
    return sections.some(sec => visibleFields[sec]);
  };

  return (
    <div className="app-container">
      {/* Header bar */}
      <header className="app-header">
        <div className="header-title-container">
          <div className="header-logo">
            <HeartPulse size={28} strokeWidth={2.5} />
          </div>
          <div>
            <h1 className="header-title">Actas Médicas</h1>
            <p style={{ fontSize: '0.75rem', color: 'var(--text-tertiary)' }}>Visualizador de Historias Clínicas</p>
          </div>
        </div>

        <div className="header-actions">
          {/* File Upload Wrapper */}
          <div className="upload-btn-wrapper">
            <button className="btn-upload">
              <Upload size={16} /> Cargar Excel
            </button>
            <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
          </div>

          {/* Theme switcher */}
          <button className="icon-btn" onClick={toggleTheme} aria-label="Toggle Theme">
            {theme === 'light' ? <Moon size={18} /> : <Sun size={18} />}
          </button>
        </div>
      </header>

      {/* Main dashboard content */}
      <div className="dashboard-body">
        {/* Sidebar panels */}
        <aside className="sidebar">
          {/* Search bar */}
          <div className="search-box">
            <Search className="search-icon" size={16} />
            <input
              type="text"
              className="search-input"
              placeholder="Buscar paciente (ID o nombre)..."
              value={searchQuery}
              onChange={(e) => {
                setSearchQuery(e.target.value);
                setSelectedPatientIndex(0); // Reset index on search
              }}
            />
          </div>

          {/* Patient navigation list */}
          <div>
            <h3 className="section-title">Pacientes ({filteredPatients.length})</h3>
            <div className="patient-list">
              {filteredPatients.length > 0 ? (
                filteredPatients.map((patient, idx) => (
                  <div
                    key={idx}
                    className={`patient-item ${activePatient && activePatient.identificacion === patient.identificacion ? 'active' : ''}`}
                    onClick={() => setSelectedPatientIndex(idx)}
                  >
                    <span className="patient-item-name">{patient.nombre_completo}</span>
                    <span className="patient-item-id">CC {patient.identificacion}</span>
                    <span className="patient-item-convenio">{patient.nombre_convenio}</span>
                  </div>
                ))
              ) : (
                <div className="no-results">
                  {patients.length === 0 ? 'Por favor cargue un archivo o los datos de prueba.' : 'No se encontraron resultados.'}
                </div>
              )}
            </div>
          </div>

          {/* Fields Selection Checkboxes */}
          <div className="fields-filter-card">
            <h3 className="section-title" style={{ display: 'flex', alignItems: 'center', gap: '0.35rem', marginBottom: '0.75rem' }}>
              <Settings size={14} /> Campos a Visualizar
            </h3>
            <div className="fields-checkbox-list">
              {FIELDS_CONFIG.map(field => (
                <label key={field.key} className="field-checkbox-label">
                  <input
                    type="checkbox"
                    checked={!!visibleFields[field.key]}
                    onChange={() => handleCheckboxChange(field.key)}
                  />
                  {field.label}
                </label>
              ))}
            </div>
          </div>
        </aside>

        {/* Clinical History Details view */}
        <main className="detail-panel">
          {activePatient ? (
            <>
              {/* Patient Profile Card Header */}
              <div className="patient-banner">
                <div className="patient-banner-info">
                  <div className="patient-badge-id">
                    <User size={12} /> CC {activePatient.identificacion}
                  </div>
                  <h2 className="patient-banner-name">{activePatient.nombre_completo}</h2>
                  <div className="patient-banner-convenio">
                    <strong>Convenio:</strong> {activePatient.nombre_convenio}
                  </div>
                </div>

                <div className="patient-banner-actions">
                  {activePatient.datos_acta && getRiskBadge(activePatient.datos_acta.nivel_riesgo)}

                  <div className="patient-navigation-export">
                    <button
                      className="btn-nav"
                      onClick={handlePrevPatient}
                      disabled={selectedPatientIndex === 0}
                      title="Paciente Anterior"
                    >
                      <ChevronLeft size={16} />
                    </button>
                    <button
                      className="btn-nav"
                      onClick={handleNextPatient}
                      disabled={selectedPatientIndex === filteredPatients.length - 1}
                      title="Siguiente Paciente"
                    >
                      <ChevronRight size={16} />
                    </button>
                    <button className="btn-action" onClick={handlePrint}>
                      <Printer size={16} /> Exportar a PDF
                    </button>
                  </div>
                </div>
              </div>

              {/* Check if patient has datos_acta */}
              {activePatient.datos_acta ? (
                <>
                  {/* Clinical Tabs Header */}
                  <div className="clinical-tabs no-print">
                    <button
                      className={`clinical-tab ${activeTab === 'generales' ? 'active' : ''}`}
                      onClick={() => setActiveTab('generales')}
                    >
                      <User size={16} /> Datos Generales
                    </button>
                    <button
                      className={`clinical-tab ${activeTab === 'antecedentes' ? 'active' : ''}`}
                      onClick={() => setActiveTab('antecedentes')}
                    >
                      <ShieldAlert size={16} /> Antecedentes
                    </button>
                    <button
                      className={`clinical-tab ${activeTab === 'analisis' ? 'active' : ''}`}
                      onClick={() => setActiveTab('analisis')}
                    >
                      <ClipboardList size={16} /> Análisis y Notas
                    </button>

                    <button
                      className={`clinical-tab ${activeTab === 'historias' ? 'active' : ''}`}
                      onClick={() => setActiveTab('historias')}
                    >
                      <FileText size={16} /> Historias Anteriores
                    </button>
                    <button
                      className={`clinical-tab ${activeTab === 'medicinas' ? 'active' : ''}`}
                      onClick={() => setActiveTab('medicinas')}
                    >
                      <Pill size={16} /> Diagnósticos / Meds
                    </button>
                  </div>

                  {/* Clinical Tab Content (VISIBLE ON SCREEN, HIDDEN ON PRINT) */}
                  <div className="clinical-content no-print">
                    {/* Generales Tab */}
                    {activeTab === 'generales' && (
                      <div className="clinical-tab-section">
                        {isTabSectionVisible(['edad', 'genero', 'nivel_riesgo']) && (
                          <div className="info-card">
                            <div className="info-card-header">
                              <User className="info-card-icon" size={18} />
                              <span className="info-card-title">Filiación General</span>
                            </div>
                            <div className="kv-grid">
                              {renderSection('edad', activePatient.datos_acta)}
                              {renderSection('genero', activePatient.datos_acta)}
                              {renderSection('nivel_riesgo', activePatient.datos_acta)}
                            </div>
                          </div>
                        )}
                        {renderSection('metricas_relevantes', activePatient.datos_acta)}
                        {!isTabSectionVisible(['edad', 'genero', 'nivel_riesgo', 'metricas_relevantes']) && (
                          <div className="info-card">
                            <span className="null-value text-center">Todos los campos de esta sección están ocultos por los filtros de visualización.</span>
                          </div>
                        )}
                      </div>
                    )}

                    {/* Antecedentes Tab */}
                    {activeTab === 'antecedentes' && (
                      <div className="clinical-tab-section">
                        {visibleFields.antecedentes ? (
                          renderSection('antecedentes', activePatient.datos_acta)
                        ) : (
                          <div className="info-card">
                            <span className="null-value text-center">La sección de Antecedentes está desactivada en los filtros de visualización.</span>
                          </div>
                        )}
                      </div>
                    )}

                    {/* Analisis Tab */}
                    {activeTab === 'analisis' && (
                      <div className="clinical-tab-section">
                        {visibleFields.analisis ? (
                          renderSection('analisis', activePatient.datos_acta)
                        ) : (
                          <div className="info-card">
                            <span className="null-value text-center">La sección de Análisis y Notas está desactivada en los filtros de visualización.</span>
                          </div>
                        )}
                      </div>
                    )}

                    {/* Atenciones Tab */}
                    {activeTab === 'atenciones' && (
                      <div className="clinical-tab-section">
                        {renderSection('atenciones_programadas', activePatient.datos_acta)}
                        {renderSection('prog_especialidad_acta', activePatient.datos_acta)}
                        {!isTabSectionVisible(['atenciones_programadas', 'prog_especialidad_acta']) && (
                          <div className="info-card">
                            <span className="null-value text-center">Las secciones de Atenciones y Especialidades están desactivadas en los filtros.</span>
                          </div>
                        )}
                      </div>
                    )}

                    {/* Historias Tab */}
                    {activeTab === 'historias' && (
                      <div className="clinical-tab-section">
                        {visibleFields.historias_anteriores ? (
                          renderSection('historias_anteriores', activePatient.datos_acta)
                        ) : (
                          <div className="info-card">
                            <span className="null-value text-center">La sección de Historias Anteriores está desactivada en los filtros.</span>
                          </div>
                        )}
                      </div>
                    )}

                    {/* Medicinas Tab */}
                    {activeTab === 'medicinas' && (
                      <div className="clinical-tab-section">
                        {renderSection('diagnosticos', activePatient.datos_acta)}
                        {renderSection('medicamentos_prescritos', activePatient.datos_acta)}
                        {!isTabSectionVisible(['diagnosticos', 'medicamentos_prescritos']) && (
                          <div className="info-card">
                            <span className="null-value text-center">Las secciones de Diagnósticos y Medicamentos están desactivadas en los filtros.</span>
                          </div>
                        )}
                      </div>
                    )}
                  </div>

                  {/* Print-only layout (HIDDEN ON SCREEN, VISIBLE ON PRINT) containing ALL checked sections in a single stream */}
                  <div className="print-only-container">
                    <div style={{ borderBottom: '2px solid #000', paddingBottom: '0.5rem', marginBottom: '1.5rem' }}>
                      <h3 style={{ margin: 0, textTransform: 'uppercase' }}>Reporte de Historia Clínica General</h3>
                      <p style={{ fontSize: '0.8rem', margin: 0 }}>Generado el {new Date().toLocaleDateString()}</p>
                    </div>

                    {/* Filiación General Section */}
                    {isTabSectionVisible(['edad', 'genero', 'nivel_riesgo']) && (
                      <div className="info-card" style={{ pageBreakInside: 'avoid' }}>
                        <div className="info-card-header">
                          <span className="info-card-title">Datos Demográficos</span>
                        </div>
                        <div className="kv-grid" style={{ gridTemplateColumns: 'repeat(3, 1fr)' }}>
                          {visibleFields.edad && (
                            <div className="kv-item">
                              <span className="kv-label">Edad</span>
                              <span className="kv-value">{activePatient.datos_acta.edad}</span>
                            </div>
                          )}
                          {visibleFields.genero && (
                            <div className="kv-item">
                              <span className="kv-label">Género</span>
                              <span className="kv-value">{activePatient.datos_acta.genero}</span>
                            </div>
                          )}
                          {visibleFields.nivel_riesgo && (
                            <div className="kv-item">
                              <span className="kv-label">Nivel de Riesgo</span>
                              <span className="kv-value">{activePatient.datos_acta.nivel_riesgo}</span>
                            </div>
                          )}
                        </div>
                      </div>
                    )}

                    {/* Sequential print blocks of all other components */}
                    {renderSection('metricas_relevantes', activePatient.datos_acta)}
                    {renderSection('antecedentes', activePatient.datos_acta)}
                    {renderSection('analisis', activePatient.datos_acta)}
                    {renderSection('atenciones_programadas', activePatient.datos_acta)}
                    {renderSection('prog_especialidad_acta', activePatient.datos_acta)}
                    {renderSection('historias_anteriores', activePatient.datos_acta)}
                    {renderSection('diagnosticos', activePatient.datos_acta)}
                    {renderSection('medicamentos_prescritos', activePatient.datos_acta)}
                  </div>
                </>
              ) : (
                <div className="info-card">
                  <div style={{ textAlign: 'center', padding: '2rem 0', color: 'var(--text-secondary)' }}>
                    <AlertCircle size={40} style={{ color: 'var(--risk-medium)', marginBottom: '1rem' }} />
                    <p style={{ fontWeight: 600 }}>No hay información estructurada disponible en el campo `datos_acta` para este paciente.</p>
                  </div>
                </div>
              )}
            </>
          ) : (
            <div className="welcome-panel">
              <FileSpreadsheet className="welcome-icon" size={64} />
              <h2 className="welcome-title">Bienvenido al Visualizador de Actas</h2>
              <p className="welcome-desc">
                Cargue un archivo Excel que contenga los campos de actas médicas.
              </p>

              <div className="welcome-actions">
                <div className="upload-btn-wrapper">
                  <button className="btn-upload">
                    <Upload size={16} /> Cargar Archivo Excel
                  </button>
                  <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
                </div>

              </div>
            </div>
          )}
        </main>
      </div>
    </div>
  );
}

export default App;
