import React, { useState, useRef, useEffect } from 'react';
import {
  UploadCloud,
  FileSpreadsheet,
  Download,
  Search,
  AlertTriangle,
  RefreshCw,
  Trophy,
  Percent,
  Users,
  BookOpen,
  ArrowUpDown,
  CheckCircle,
  HelpCircle,
  XCircle
} from 'lucide-react';
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip as ChartTooltip,
  ResponsiveContainer,
  Cell
} from 'recharts';
import * as XLSX from 'xlsx';

import { parseExcelFile } from './utils/analyzer';
import type { StudentRecord, AnalysisResult } from './utils/analyzer';
import { generateSampleExcel } from './utils/sampleGenerator';

export default function App() {
  // Parsing states
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const [data, setData] = useState<AnalysisResult | null>(null);

  // Drag & drop state
  const [dragActive, setDragActive] = useState<boolean>(false);

  // Filters & Sorting state
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [statusFilter, setStatusFilter] = useState<string>('All');
  const [highlightRegNo, setHighlightRegNo] = useState<string>('');
  const [sortKey, setSortKey] = useState<keyof StudentRecord>('name');
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc'>('asc');

  // References
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Auto-scroll and highlight selected student
  useEffect(() => {
    if (highlightRegNo && data) {
      // Find if student is in the current search/filtered list
      const timer = setTimeout(() => {
        const rowElement = document.getElementById(`student-row-${highlightRegNo}`);
        if (rowElement) {
          rowElement.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
      }, 100);
      return () => clearTimeout(timer);
    }
  }, [highlightRegNo, searchQuery, statusFilter, sortKey, sortDirection]);

  // Handle Drag Over / Enter
  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  // Handle Drag Drop
  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const droppedFile = e.dataTransfer.files[0];
      if (
        droppedFile.type === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
        droppedFile.name.endsWith('.xlsx')
      ) {
        processFile(droppedFile);
      } else {
        setError("Only Excel (.xlsx) files are supported.");
      }
    }
  };

  // Handle File Input Change
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      processFile(e.target.files[0]);
    }
  };

  // Parse and process uploaded file
  const processFile = async (selectedFile: File) => {
    setFile(selectedFile);
    setLoading(true);
    setError(null);
    setData(null);

    try {
      const result = await parseExcelFile(selectedFile);
      setData(result);
    } catch (err: any) {
      console.error(err);
      setError(err.message || "Failed to parse the Excel file. Please verify its structure.");
      setFile(null);
    } finally {
      setLoading(false);
    }
  };

  // Export results to Excel
  const handleExport = () => {
    if (!data) return;

    // Prepare table format
    const exportRows = data.students.map(s => ({
      'Rank': s.rank,
      'Register Number': s.regNo,
      'Student Name': s.name,
      'SGPA': s.sgpa === null ? 'N/A' : s.sgpa,
      'Attempted Credits': s.attemptedCredits,
      'Earned Credits': s.earnedCredits,
      'Arrears Count': s.arrearsCount,
      'Arrear Subjects': s.arrearSubjects.join(', ') || 'None',
      'Status': s.status,
      'Warnings': s.warnings.join(' | ') || 'None'
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Calculated SGPA Ranks');
    
    // Auto-fit column widths
    const maxLens = Object.keys(exportRows[0] || {}).map(key => {
      let maxLen = key.length;
      exportRows.forEach(row => {
        const val = String((row as any)[key] || '');
        if (val.length > maxLen) maxLen = val.length;
      });
      return { wch: maxLen + 3 };
    });
    worksheet['!cols'] = maxLens;

    XLSX.writeFile(workbook, `Class_Semester_SGPA_Ranks_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  // Reset analysis session
  const handleReset = () => {
    setFile(null);
    setData(null);
    setError(null);
    setSearchQuery('');
    setStatusFilter('All');
    setHighlightRegNo('');
    setSortKey('name');
    setSortDirection('asc');
  };

  // Dropzone click handler
  const onUploadClick = () => {
    fileInputRef.current?.click();
  };

  // Handle student highlighting and clear filters to guarantee visibility
  const handleHighlightSelect = (regNo: string) => {
    setHighlightRegNo(regNo);
    if (regNo !== '') {
      // Clear filters so the highlighted student actually exists in the table DOM
      setSearchQuery('');
      setStatusFilter('All');
    }
  };

  // Sort handler
  const requestSort = (key: keyof StudentRecord) => {
    let direction: 'asc' | 'desc' = 'asc';
    if (sortKey === key && sortDirection === 'asc') {
      direction = 'desc';
    }
    setSortKey(key);
    setSortDirection(direction);
  };

  // Render components helper
  const getSortedStudents = (students: StudentRecord[]) => {
    return [...students].sort((a, b) => {
      let aVal: any = a[sortKey];
      let bVal: any = b[sortKey];

      // Handle null cases for sorting SGPAs/ranks
      if (aVal === null || aVal === undefined || aVal === '—') {
        return sortDirection === 'asc' ? 1 : -1;
      }
      if (bVal === null || bVal === undefined || bVal === '—') {
        return sortDirection === 'asc' ? -1 : 1;
      }

      if (typeof aVal === 'string') {
        return sortDirection === 'asc'
          ? aVal.localeCompare(bVal)
          : bVal.localeCompare(aVal);
      } else {
        return sortDirection === 'asc'
          ? aVal - bVal
          : bVal - aVal;
      }
    });
  };

  // Filtering
  const filteredStudents = data
    ? getSortedStudents(
        data.students.filter(student => {
          const matchesSearch =
            student.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
            student.regNo.includes(searchQuery);
          
          const matchesStatus =
            statusFilter === 'All' ||
            (statusFilter === 'Pass' && student.status === 'Pass') ||
            (statusFilter === 'Arrear' && student.status === 'Arrear') ||
            (statusFilter === 'N/A' && student.status === 'N/A');

          return matchesSearch && matchesStatus;
        })
      )
    : [];

  // Theme colors for charts
  const GRADE_COLORS: Record<string, string> = {
    'O': '#10b981',
    'A+': '#34d399',
    'A': '#60a5fa',
    'B+': '#6366f1',
    'B': '#8b5cf6',
    'C': '#c084fc',
    'U/RA': '#f43f5e',
    'SA/W': '#9ca3af',
    'Invalid': '#fbbf24'
  };

  const ARREAR_COLORS = ['#10b981', '#fbbf24', '#f97316', '#f43f5e'];

  return (
    <div className="app-container">
      {/* HEADER SECTION */}
      <header className="app-header animate-fade-in">
        <div className="header-title-area">
          <h1>Anna University Semester Result Analyzer</h1>
          <p>
            {data && file 
              ? `Currently analyzing: ${file.name}` 
              : "Instantly calculate SGPAs, assign ranks, and extract class-wide statistics (Regulation 2021)"}
          </p>
        </div>
        {data && (
          <div style={{ display: 'flex', gap: '0.75rem' }}>
            <button className="btn btn-outline" onClick={handleReset}>
              <RefreshCw size={16} />
              Reset & Upload New
            </button>
            <button className="btn btn-primary" onClick={handleExport}>
              <Download size={16} />
              Export Results (.xlsx)
            </button>
          </div>
        )}
      </header>

      {/* UPLOAD SCREEN */}
      {!data && !loading && (
        <div className="animate-fade-in">
          <div className="upload-wrapper glass-card">
            {error && (
              <div className="alert-banner alert-danger">
                <AlertTriangle size={20} style={{ flexShrink: 0 }} />
                <div>
                  <strong>Parsing Error:</strong>
                  <div style={{ marginTop: '0.25rem', whiteSpace: 'pre-line' }}>{error}</div>
                </div>
              </div>
            )}

            <div
              className={`upload-dropzone ${dragActive ? 'drag-active' : ''}`}
              onDragEnter={handleDrag}
              onDragOver={handleDrag}
              onDragLeave={handleDrag}
              onDrop={handleDrop}
              onClick={onUploadClick}
            >
              <div className="upload-icon-wrapper">
                <UploadCloud size={40} />
              </div>
              <div>
                <h3>Drag & Drop Result Sheet</h3>
                <p>Upload the classroom results Excel file (.xlsx) directly from your computer</p>
              </div>
              <button className="btn btn-outline" type="button">
                Browse Files
              </button>
              <input
                ref={fileInputRef}
                type="file"
                className="hidden-file-input"
                accept=".xlsx"
                onChange={handleFileChange}
              />
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem', alignItems: 'center' }}>
              <p style={{ fontSize: '0.85rem', color: 'var(--text-muted)' }}>
                Don't have a structured file? Download a sample template with sample students, electives, arrear cases, and formatting warnings.
              </p>
              <button className="btn btn-outline" onClick={generateSampleExcel}>
                <FileSpreadsheet size={16} />
                Download Sample Excel Template
              </button>
            </div>
          </div>
        </div>
      )}

      {/* LOADING SCREEN */}
      {loading && (
        <div className="loading-panel glass-card animate-fade-in">
          <div className="spinner"></div>
          <div>
            <h3 style={{ textAlign: 'center', marginBottom: '0.5rem' }}>Analyzing Result Sheet...</h3>
            <p style={{ color: 'var(--text-secondary)', fontSize: '0.9rem' }}>
              Calculating SGPAs, credits, pass rates, and compiling charts in real time.
            </p>
          </div>
        </div>
      )}

      {/* DASHBOARD SCREEN */}
      {data && (
        <main className="dashboard-grid animate-fade-in">
          
          {/* KPI SUMMARY CARDS */}
          <section className="kpi-row">
            <div className="kpi-card glass-card">
              <div className="kpi-icon-container" style={{ background: 'rgba(99, 102, 241, 0.1)', color: 'var(--primary)' }}>
                <Users size={24} />
              </div>
              <div className="kpi-info">
                <h4>Total Strength</h4>
                <p>{data.analytics.totalStudents}</p>
                <div className="kpi-subtext">Students in classroom</div>
              </div>
            </div>

            <div className="kpi-card glass-card">
              <div className="kpi-icon-container" style={{ background: 'var(--success-glow)', color: 'var(--success-text)' }}>
                <Percent size={24} />
              </div>
              <div className="kpi-info">
                <h4>Pass Percentage</h4>
                <p>{data.analytics.passPercentage.toFixed(2)}%</p>
                <div className="kpi-subtext">
                  {data.students.filter(s => s.status === 'Pass' && s.attemptedCredits > 0).length} of {data.students.filter(s => s.attemptedCredits > 0).length} passed
                </div>
              </div>
            </div>

            <div className="kpi-card glass-card">
              <div className="kpi-icon-container" style={{ background: 'rgba(139, 92, 246, 0.1)', color: 'var(--secondary)' }}>
                <Trophy size={24} />
              </div>
              <div className="kpi-info">
                <h4>Class Average SGPA</h4>
                <p>{data.analytics.classAverageSgpa !== null ? data.analytics.classAverageSgpa.toFixed(2) : 'N/A'}</p>
                <div className="kpi-subtext">Excludes non-registered students</div>
              </div>
            </div>

            <div className="kpi-card glass-card" style={{ gridColumn: 'span 1' }}>
              <div className="kpi-icon-container" style={{ background: 'var(--warning-glow)', color: 'var(--warning-text)' }}>
                <Trophy size={24} />
              </div>
              <div className="kpi-info" style={{ overflow: 'hidden', textOverflow: 'ellipsis' }}>
                <h4>Class Topper(s)</h4>
                <p style={{ fontSize: '1.25rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {data.analytics.toppers.length > 0 
                    ? data.analytics.toppers.map(t => `${t.name} (${t.sgpa?.toFixed(2)})`).join(', ') 
                    : 'N/A'}
                </p>
                <div className="kpi-subtext">Highest SGPA in class</div>
              </div>
            </div>
          </section>

          {/* CHARTS CONTAINER */}
          <section className="charts-grid-row">
            {/* Grade Distribution */}
            <div className="chart-widget glass-card">
              <h3>Overall Grade Distribution</h3>
              <div className="chart-container-inner">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={data.analytics.gradeDistribution} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" />
                    <XAxis dataKey="grade" stroke="var(--text-secondary)" fontSize={11} />
                    <YAxis stroke="var(--text-secondary)" fontSize={11} allowDecimals={false} />
                    <ChartTooltip 
                      contentStyle={{ background: '#0d1426', borderColor: 'var(--border-color)', borderRadius: '8px' }}
                      labelStyle={{ color: '#fff', fontWeight: 'bold' }}
                    />
                    <Bar dataKey="count" radius={[4, 4, 0, 0]}>
                      {data.analytics.gradeDistribution.map((entry, index) => (
                        <Cell key={`cell-${index}`} fill={GRADE_COLORS[entry.grade] || '#6366f1'} />
                      ))}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>

            {/* Arrear Count Distribution */}
            <div className="chart-widget glass-card">
              <h3>Arrears Count Distribution</h3>
              <div className="chart-container-inner">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={data.analytics.arrearDistribution} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" />
                    <XAxis dataKey="arrears" stroke="var(--text-secondary)" fontSize={11} />
                    <YAxis stroke="var(--text-secondary)" fontSize={11} allowDecimals={false} />
                    <ChartTooltip 
                      contentStyle={{ background: '#0d1426', borderColor: 'var(--border-color)', borderRadius: '8px' }}
                      labelStyle={{ color: '#fff', fontWeight: 'bold' }}
                    />
                    <Bar dataKey="count" radius={[4, 4, 0, 0]}>
                      {data.analytics.arrearDistribution.map((_, index) => (
                        <Cell key={`cell-${index}`} fill={ARREAR_COLORS[index % ARREAR_COLORS.length]} />
                      ))}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </section>

          {/* SUBJECT ANALYSIS & ELECTIVES TABLE */}
          <section className="chart-widget glass-card" style={{ minHeight: 'auto' }}>
            <h3 style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
              <BookOpen size={18} />
              Subject-wise Analytics (Registration-adjusted)
            </h3>
            <div className="subject-stats-table-wrapper">
              <table className="subject-stats-table">
                <thead>
                  <tr>
                    <th>Subject Code</th>
                    <th>Credits</th>
                    <th>Students Registered</th>
                    <th>Students Passed</th>
                    <th>Pass Percentage</th>
                    <th>Avg Grade Point</th>
                    <th>Grade Counts (O / A+ / A / B+ / B / C / U-RA / SA-W / Invalid)</th>
                  </tr>
                </thead>
                <tbody>
                  {data.analytics.subjectStats.map(subject => (
                    <tr key={subject.code}>
                      <td style={{ fontWeight: '600', color: '#fff' }}>{subject.code}</td>
                      <td>{subject.credits}</td>
                      <td>{subject.registeredCount}</td>
                      <td>{subject.passedCount}</td>
                      <td style={{ fontWeight: '600', color: subject.passPercentage >= 75 ? 'var(--success-text)' : 'var(--danger-text)' }}>
                        {subject.passPercentage.toFixed(2)}%
                      </td>
                      <td>{subject.averageGradePoint.toFixed(2)}</td>
                      <td>
                        <div style={{ display: 'flex', gap: '0.4rem', flexWrap: 'wrap' }}>
                          <span className="grade-badge-pill" style={{ color: GRADE_COLORS['O'] }}>O: {subject.gradeDistribution['O']}</span>
                          <span className="grade-badge-pill" style={{ color: GRADE_COLORS['A+'] }}>A+: {subject.gradeDistribution['A+']}</span>
                          <span className="grade-badge-pill" style={{ color: GRADE_COLORS['A'] }}>A: {subject.gradeDistribution['A']}</span>
                          <span className="grade-badge-pill" style={{ color: GRADE_COLORS['B+'] }}>B+: {subject.gradeDistribution['B+']}</span>
                          <span className="grade-badge-pill" style={{ color: GRADE_COLORS['B'] }}>B: {subject.gradeDistribution['B']}</span>
                          <span className="grade-badge-pill" style={{ color: GRADE_COLORS['C'] }}>C: {subject.gradeDistribution['C']}</span>
                          <span className="grade-badge-pill fail-grade">U/RA: {subject.gradeDistribution['U/RA']}</span>
                          {subject.gradeDistribution['SA/W'] > 0 && (
                            <span className="grade-badge-pill" style={{ color: GRADE_COLORS['SA/W'] }}>SA/W: {subject.gradeDistribution['SA/W']}</span>
                          )}
                          {subject.gradeDistribution['Invalid'] > 0 && (
                            <span className="grade-badge-pill" style={{ color: GRADE_COLORS['Invalid'], background: 'var(--warning-glow)' }}>
                              Inv: {subject.gradeDistribution['Invalid']}
                            </span>
                          )}
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>

          {/* CONTROLS & FILTERING BAR */}
          <section className="control-bar glass-card">
            <div className="controls-left-group">
              {/* Search */}
              <div className="search-input-wrapper">
                <Search size={16} />
                <input
                  type="text"
                  placeholder="Search student by name or Reg No..."
                  className="input-field"
                  value={searchQuery}
                  onChange={(e) => {
                    setSearchQuery(e.target.value);
                    // Clear highlighted item if searching
                    setHighlightRegNo('');
                  }}
                />
              </div>

              {/* Status Filter */}
              <select
                className="select-field"
                value={statusFilter}
                onChange={(e) => {
                  setStatusFilter(e.target.value);
                  // Clear highlight
                  setHighlightRegNo('');
                }}
              >
                <option value="All">All Statuses</option>
                <option value="Pass">Pass (0 Arrears)</option>
                <option value="Arrear">Arrear (&gt;0 Arrears)</option>
                <option value="N/A">Not Registered (N/A)</option>
              </select>

              {/* Highlight Specific Reg No */}
              <div className="highlight-field-wrapper">
                <span style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', whiteSpace: 'nowrap' }}>Highlight my Reg No:</span>
                <select
                  className="select-field highlight-input"
                  value={highlightRegNo}
                  onChange={(e) => handleHighlightSelect(e.target.value)}
                >
                  <option value="">-- Select Register No --</option>
                  {data.students.map(s => (
                    <option key={s.regNo} value={s.regNo}>
                      {s.regNo} ({s.name})
                    </option>
                  ))}
                </select>
                {highlightRegNo && (
                  <button className="btn btn-outline" style={{ padding: '0.5rem', borderRadius: '10px' }} onClick={() => setHighlightRegNo('')}>
                    Clear
                  </button>
                )}
              </div>
            </div>

            <div style={{ color: 'var(--text-muted)', fontSize: '0.85rem' }}>
              Showing <strong>{filteredStudents.length}</strong> of <strong>{data.students.length}</strong> students
            </div>
          </section>

          {/* MAIN RESULTS TABLE */}
          <section className="table-section glass-card">
            <div className="table-header-row">
              <h2>Rank List & Calculations</h2>
              <div style={{ display: 'flex', gap: '0.5rem', alignItems: 'center' }}>
                <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>* Click table headers to sort</span>
              </div>
            </div>
            
            <div className="table-responsive">
              <table className="students-data-table">
                <thead>
                  <tr>
                    <th onClick={() => requestSort('rank')}>
                      Rank
                      <ArrowUpDown size={12} className="sort-icon-inline" />
                    </th>
                    <th onClick={() => requestSort('regNo')}>
                      Reg. Number
                      <ArrowUpDown size={12} className="sort-icon-inline" />
                    </th>
                    <th onClick={() => requestSort('name')}>
                      Student Name
                      <ArrowUpDown size={12} className="sort-icon-inline" />
                    </th>
                    <th onClick={() => requestSort('sgpa')}>
                      SGPA
                      <ArrowUpDown size={12} className="sort-icon-inline" />
                    </th>
                    <th>Credits Earned / Attempted</th>
                    <th onClick={() => requestSort('arrearsCount')}>
                      Arrears List
                      <ArrowUpDown size={12} className="sort-icon-inline" />
                    </th>
                    <th onClick={() => requestSort('status')}>
                      Status
                      <ArrowUpDown size={12} className="sort-icon-inline" />
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {filteredStudents.length > 0 ? (
                    filteredStudents.map(student => {
                      const isHighlighted = student.regNo === highlightRegNo;
                      return (
                        <tr
                          key={student.regNo}
                          id={`student-row-${student.regNo}`}
                          className={isHighlighted ? 'highlighted-student-row' : ''}
                        >
                          <td style={{ fontWeight: '700', color: student.rank === 1 ? '#fbbf24' : '#fff' }}>
                            {student.rank === 1 ? '🥇 1' : student.rank}
                          </td>
                          <td style={{ fontFamily: 'monospace', fontSize: '0.95rem' }}>{student.regNo}</td>
                          <td>
                            <div style={{ display: 'flex', alignItems: 'center' }}>
                              <span>{student.name}</span>
                              {student.warnings.length > 0 && (
                                <span className="warnings-indicator tooltip-trigger">
                                  <AlertTriangle size={16} />
                                  <div className="tooltip-content">
                                    {student.warnings.map((w, idx) => (
                                      <div key={idx} style={{ marginBottom: '2px' }}>⚠️ {w}</div>
                                    ))}
                                  </div>
                                </span>
                              )}
                            </div>
                          </td>
                          <td style={{ fontWeight: '700', fontSize: '1rem' }}>
                            {student.sgpa !== null ? student.sgpa.toFixed(2) : 'N/A'}
                          </td>
                          <td>
                            <span style={{ color: '#fff', fontWeight: '500' }}>{student.earnedCredits}</span>
                            <span style={{ color: 'var(--text-muted)' }}> / {student.attemptedCredits}</span>
                          </td>
                          <td>
                            {student.arrearsCount > 0 ? (
                              <div>
                                <span style={{ color: 'var(--danger-text)', fontWeight: '600' }}>
                                  {student.arrearsCount} Arrear{student.arrearsCount > 1 ? 's' : ''}
                                </span>
                                <div className="grades-summary-grid" style={{ marginTop: '0.25rem' }}>
                                  {student.arrearSubjects.map(subCode => (
                                    <span key={subCode} className="grade-badge-pill fail-grade">
                                      {subCode}: {student.grades[subCode]}
                                    </span>
                                  ))}
                                </div>
                              </div>
                            ) : student.status === 'N/A' ? (
                              <span style={{ color: 'var(--text-muted)' }}>—</span>
                            ) : (
                              <span style={{ color: 'var(--success-text)' }}>None</span>
                            )}
                          </td>
                          <td>
                            {student.status === 'Pass' && (
                              <span className="badge badge-success">
                                <CheckCircle size={12} style={{ marginRight: '4px' }} />
                                Pass
                              </span>
                            )}
                            {student.status === 'Arrear' && (
                              <span className="badge badge-danger">
                                <XCircle size={12} style={{ marginRight: '4px' }} />
                                Arrear
                              </span>
                            )}
                            {student.status === 'N/A' && (
                              <span className="badge badge-neutral">
                                N/A
                              </span>
                            )}
                          </td>
                        </tr>
                      );
                    })
                  ) : (
                    <tr>
                      <td colSpan={7}>
                        <div className="empty-table-state">
                          <HelpCircle size={32} style={{ color: 'var(--text-muted)' }} />
                          <p>No student results match your search and filter criteria.</p>
                        </div>
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </section>
        </main>
      )}
    </div>
  );
}
