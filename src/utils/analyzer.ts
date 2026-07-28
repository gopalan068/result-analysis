import * as XLSX from 'xlsx';

export interface Subject {
  code: string;         // e.g. "CB3491"
  originalCode: string;   // e.g. "CB3491.1"
  credits: number;      // e.g. 3
  colIdx: number;
}

export interface StudentRecord {
  regNo: string;
  name: string;
  grades: Record<string, string>; // subjectCode -> grade string
  sgpa: number | null; // null if zero credits attempted
  attemptedCredits: number;
  earnedCredits: number;
  arrearsCount: number;
  arrearSubjects: string[];
  status: 'Pass' | 'Arrear' | 'N/A';
  rank: number | string;
  warnings: string[];
  rowIdx: number;
}

export interface SubjectStats {
  code: string;
  credits: number;
  registeredCount: number; // students who had a non-blank grade for this subject
  passedCount: number;     // students who passed (grade !== U/RA/SA/W and valid)
  passPercentage: number;  // (passedCount / registeredCount) * 100
  averageGradePoint: number; // sum of GP for registered students / registeredCount
  gradeDistribution: Record<string, number>; // grade -> count
}

export interface ArrearDistributionItem {
  arrears: string; // "0 Arrears", "1 Arrear", "2 Arrears", "3+ Arrears"
  count: number;
}

export interface GradeDistributionItem {
  grade: string; // "O", "A+", "A", "B+", "B", "C", "U/RA", "SA/W", "Invalid"
  count: number;
}

export interface ClassAnalytics {
  classAverageSgpa: number | null;
  totalStudents: number;
  passPercentage: number;
  toppers: StudentRecord[];
  subjectStats: SubjectStats[];
  arrearDistribution: ArrearDistributionItem[];
  gradeDistribution: GradeDistributionItem[];
}

export interface AnalysisResult {
  subjects: Subject[];
  students: StudentRecord[];
  analytics: ClassAnalytics;
}

export const GRADE_POINTS: Record<string, number> = {
  'O': 10,
  'A+': 9,
  'A': 8,
  'B+': 7,
  'B': 6,
  'C': 5,
  'U': 0,
  'RA': 0,
  'SA': 0,
  'W': 0
};

export const cleanSubjectCode = (code: any): string => {
  if (code === null || code === undefined) return '';
  return String(code).replace(/\.\d+$/, '').trim();
};

export const analyzeData = (rows: any[][]): AnalysisResult => {
  if (rows.length < 3) {
    throw new Error(
      "The Excel sheet has insufficient rows. It must contain at least:\n" +
      "Row 1: Subject Codes (e.g. CB3491, CCS371)\n" +
      "Row 2: Credit Values (e.g. 3, 4, 1.5)\n" +
      "Row 3: Header Row (Reg. Number, Stud. Name, Grade, Grade...)"
    );
  }

  const row1 = rows[0] || [];
  const row2 = rows[1] || [];
  const row3 = rows[2] || [];

  // 1. Identify Register Number and Student Name columns
  let regIdx = -1;
  let nameIdx = -1;

  for (let c = 0; c < Math.max(row3.length, 5); c++) {
    const val = String(row3[c] || '').trim().toLowerCase();
    if (
      val.includes('reg') ||
      val.includes('roll') ||
      val.includes('number') ||
      val.includes('id')
    ) {
      regIdx = c;
    } else if (val.includes('name') || val.includes('student') || val.includes('stud')) {
      nameIdx = c;
    }
  }

  // Fallbacks if not detected
  if (regIdx === -1) regIdx = 0;
  if (nameIdx === -1) nameIdx = 1;

  // 2. Identify Subject Columns
  const subjects: Subject[] = [];
  const maxCols = Math.max(row1.length, row2.length, row3.length);

  for (let c = 0; c < maxCols; c++) {
    if (c === regIdx || c === nameIdx) continue;

    const rawSubCode = row1[c];
    const rawCredits = row2[c];

    if (rawSubCode !== undefined && rawSubCode !== null && String(rawSubCode).trim() !== '') {
      const codeStr = String(rawSubCode).trim();
      const code = cleanSubjectCode(codeStr);
      
      const credits = parseFloat(String(rawCredits).trim());
      if (!isNaN(credits) && credits > 0) {
        subjects.push({
          code,
          originalCode: codeStr,
          credits,
          colIdx: c
        });
      }
    }
  }

  if (subjects.length === 0) {
    throw new Error(
      "Could not find any valid subjects. Make sure Row 1 contains Subject Codes (e.g. CS3501) " +
      "and Row 2 contains numeric credit values (e.g. 3) directly under each subject."
    );
  }

  // 3. Parse Student Records from Row 4 (index 3) onwards
  const students: StudentRecord[] = [];
  const regNoPattern = /^\d+$/;

  for (let r = 3; r < rows.length; r++) {
    const row = rows[r] || [];
    const rawReg = row[regIdx];
    const regNo = String(rawReg || '').trim();

    // Skip trailing rows, empty rows, or summary footer rows.
    // A student record must have a valid numeric register number.
    if (regNo === '' || !regNoPattern.test(regNo) || regNo.length < 5) {
      continue;
    }

    const name = String(row[nameIdx] || '').trim();
    const grades: Record<string, string> = {};
    const warnings: string[] = [];

    let totalGpCredits = 0;
    let totalAttemptedCredits = 0;
    let totalEarnedCredits = 0;
    let arrearsCount = 0;
    const arrearSubjects: string[] = [];

    subjects.forEach(sub => {
      const cellVal = row[sub.colIdx];
      const rawGrade = cellVal !== undefined && cellVal !== null ? String(cellVal).trim() : '';

      if (rawGrade === '') {
        // Excluded: student did not register or take this course (blank grade cell)
        return;
      }

      const grade = rawGrade.toUpperCase();
      grades[sub.code] = grade;

      const isFail = (grade === 'U' || grade === 'RA');
      
      const gp = GRADE_POINTS[grade];
      if (gp !== undefined) {
        totalGpCredits += gp * sub.credits;
        totalAttemptedCredits += sub.credits;
        if (gp > 0) {
          totalEarnedCredits += sub.credits;
        }
        if (isFail) {
          arrearsCount++;
          arrearSubjects.push(sub.code);
        }
      } else {
        // Register unrecognized grades as warning
        warnings.push(`Unrecognized grade '${rawGrade}' in subject ${sub.code}`);
      }
    });

    let sgpa: number | null = null;
    let status: 'Pass' | 'Arrear' | 'N/A' = 'N/A';

    if (totalAttemptedCredits > 0) {
      status = arrearsCount === 0 ? 'Pass' : 'Arrear';
      if (status === 'Pass') {
        sgpa = Math.round((totalGpCredits / totalAttemptedCredits) * 100) / 100;
      }
    }

    students.push({
      regNo,
      name,
      grades,
      sgpa,
      attemptedCredits: totalAttemptedCredits,
      earnedCredits: totalEarnedCredits,
      arrearsCount,
      arrearSubjects,
      status,
      rank: '', // computed below
      warnings,
      rowIdx: r + 1
    });
  }

  if (students.length === 0) {
    throw new Error(
      "No student records found. Check if Row 3 contains columns like 'Reg. Number' " +
      "and Row 4 onwards contains numeric student register IDs."
    );
  }

  // 4. Rank Students (Standard Competition Ranking: 1, 2, 2, 4...)
  const validStudents = students.filter(s => s.sgpa !== null);
  const invalidStudents = students.filter(s => s.sgpa === null);

  // Sort valid students by SGPA descending
  validStudents.sort((a, b) => (b.sgpa || 0) - (a.sgpa || 0));

  let currentRank = 1;
  for (let i = 0; i < validStudents.length; i++) {
    if (i > 0 && validStudents[i].sgpa !== validStudents[i - 1].sgpa) {
      currentRank = i + 1;
    }
    validStudents[i].rank = currentRank;
  }

  invalidStudents.forEach(s => {
    s.rank = '—';
  });

  const allStudents = [...students];
  allStudents.sort((a, b) => a.name.localeCompare(b.name));

  // 5. Generate Class-wide Analytics
  // Class Average SGPA (excluding students who didn't take any exams/N/A)
  const validSgpaList = validStudents.map(s => s.sgpa as number);
  const classAverageSgpa = validSgpaList.length > 0
    ? Math.round((validSgpaList.reduce((sum, val) => sum + val, 0) / validSgpaList.length) * 100) / 100
    : null;

  // Class Toppers
  const highestSgpa = validSgpaList.length > 0 ? Math.max(...validSgpaList) : null;
  const toppers = highestSgpa !== null
    ? validStudents.filter(s => s.sgpa === highestSgpa)
    : [];

  // Pass Percentage: Students with 0 arrears / Total students who took exams
  const studentsWithExams = students.filter(s => s.attemptedCredits > 0);
  const passCount = studentsWithExams.filter(s => s.status === 'Pass').length;
  const passPercentage = studentsWithExams.length > 0
    ? Math.round((passCount / studentsWithExams.length) * 10000) / 100
    : 0;

  // Subject-wise stats
  const subjectStats: SubjectStats[] = subjects.map(sub => {
    let registeredCount = 0;
    let passedCount = 0;
    let totalGp = 0;
    const gradeDistribution: Record<string, number> = {
      'O': 0, 'A+': 0, 'A': 0, 'B+': 0, 'B': 0, 'C': 0, 'U/RA': 0, 'SA/W': 0, 'Invalid': 0
    };

    students.forEach(s => {
      const grade = s.grades[sub.code];
      if (grade !== undefined && grade !== '') {
        registeredCount++;
        const gp = GRADE_POINTS[grade];
        if (gp !== undefined) {
          totalGp += gp;
          if (gp > 0) {
            passedCount++;
          }
          if (['O', 'A+', 'A', 'B+', 'B', 'C'].includes(grade)) {
            gradeDistribution[grade]++;
          } else if (['U', 'RA'].includes(grade)) {
            gradeDistribution['U/RA']++;
          } else if (['SA', 'W'].includes(grade)) {
            gradeDistribution['SA/W']++;
          }
        } else {
          gradeDistribution['Invalid']++;
        }
      }
    });

    const passPct = registeredCount > 0 ? Math.round((passedCount / registeredCount) * 10000) / 100 : 0;
    const avgGp = registeredCount > 0 ? Math.round((totalGp / registeredCount) * 100) / 100 : 0;

    return {
      code: sub.code,
      credits: sub.credits,
      registeredCount,
      passedCount,
      passPercentage: passPct,
      averageGradePoint: avgGp,
      gradeDistribution
    };
  });

  // Arrear Distribution (0, 1, 2, 3+ arrears)
  let zeroArrears = 0;
  let oneArrear = 0;
  let twoArrears = 0;
  let threeOrMoreArrears = 0;

  studentsWithExams.forEach(s => {
    if (s.arrearsCount === 0) zeroArrears++;
    else if (s.arrearsCount === 1) oneArrear++;
    else if (s.arrearsCount === 2) twoArrears++;
    else threeOrMoreArrears++;
  });

  const arrearDistribution = [
    { arrears: '0 Arrears', count: zeroArrears },
    { arrears: '1 Arrear', count: oneArrear },
    { arrears: '2 Arrears', count: twoArrears },
    { arrears: '3+ Arrears', count: threeOrMoreArrears }
  ];

  // Combined Grade Distribution across all subjects & students
  const overallGradeDistribution: Record<string, number> = {
    'O': 0, 'A+': 0, 'A': 0, 'B+': 0, 'B': 0, 'C': 0, 'U/RA': 0, 'SA/W': 0, 'Invalid': 0
  };

  subjectStats.forEach(stat => {
    Object.keys(stat.gradeDistribution).forEach(k => {
      overallGradeDistribution[k] += stat.gradeDistribution[k];
    });
  });

  const gradeDistribution = Object.keys(overallGradeDistribution).map(grade => ({
    grade,
    count: overallGradeDistribution[grade]
  }));

  return {
    subjects,
    students: allStudents,
    analytics: {
      classAverageSgpa,
      totalStudents: students.length,
      passPercentage,
      toppers,
      subjectStats,
      arrearDistribution,
      gradeDistribution
    }
  };
};

export const parseExcelFile = (file: File): Promise<AnalysisResult> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        if (workbook.SheetNames.length === 0) {
          throw new Error("The Excel file doesn't contain any sheets.");
        }

        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        const rows = XLSX.utils.sheet_to_json<any[]>(sheet, { header: 1, defval: '' });
        resolve(analyzeData(rows));
      } catch (err: any) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("Failed to read the file. Please make sure the file is not corrupted."));
    reader.readAsArrayBuffer(file);
  });
};
