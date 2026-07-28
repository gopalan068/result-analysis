import * as XLSX from 'xlsx';

export const generateSampleExcel = () => {
  const data = [
    // Row 1: Subject Codes
    ['Subject Code ->', '', 'CB3491', 'CCS371', 'CCS375', 'CS3501', 'CS3551', 'CS3591', 'MX3084', 'NM1122', 'CCS343', 'CCS354'],
    // Row 2: Credits
    ['Credits ->', '', '3', '3', '2', '4', '4', '1.5', '3', '3', '3', '3'],
    // Row 3: Headers
    ['Reg. Number', 'Stud. Name', 'Grade', 'Grade', 'Grade', 'Grade', 'Grade', 'Grade', 'Grade', 'Grade', 'Grade', 'Grade'],
    // Row 4 onwards: Students
    ['312221104001', 'AARAV SHARMA', 'O', 'A+', 'A', 'O', 'B+', 'A+', 'A', 'B', 'B+', ''],
    ['312221104002', 'ADITI SHENOY', 'A+', 'A', 'B+', 'A', 'B', 'C', 'O', 'A+', '', 'B+'],
    ['312221104003', 'ANANYA IYER', 'B', 'C', 'U', 'B+', 'A', 'B', 'C', 'B', 'A', ''],
    ['312221104004', 'ARJUN NAIR', 'O', 'O', 'A+', 'O', 'O', 'O', 'A+', 'O', '', 'O'],
    ['312221104005', 'DEVENDRA SINGH', 'A', 'B+', 'C', 'RA', 'B', 'A', 'B+', 'C', 'C', ''],
    ['312221104006', 'MEERA PILLAI', 'B+', 'B', 'B', 'C', 'U', 'B', 'B', 'C', '', 'B'],
    ['312221104007', 'PRANAV KUMAR', 'O', 'A+', 'A', 'A+', 'A', 'A+', 'O', 'A+', 'A+', ''],
    ['312221104008', 'RIYA PATEL', 'A', 'A', 'A', 'B+', 'B', 'C', 'A', 'A', '', 'A'],
    ['312221104009', 'SANJAY SUBRAMANIAN', 'A+', 'O', 'A', 'A+', 'A+', 'A', 'O', 'A', 'B', ''],
    ['312221104010', 'VIKRAM SENTHIL', 'U', 'RA', 'U', 'U', 'U', 'U', 'U', 'U', '', ''],
    // Student with an unrecognized grade (warning demo)
    ['312221104011', 'YASMIN BEGUM (WARNING DEMO)', 'A+', 'A', 'A+', 'O', 'B+', 'A+', 'A', 'B', 'O', ''],
    // Student with no registered exams (empty demo)
    ['312221104012', 'ZACHARY MILLER (EMPTY DEMO)', '', '', '', '', '', '', '', '', '', ''],
    // Junk trailing rows to verify parsing robustness
    ['', '', '', '', '', '', '', '', '', '', '', ''],
    ['Class Total Students:', '12', '', '', '', '', '', '', '', '', '', ''],
    ['Prepared by: Exam Coordinator', '', '', '', '', '', '', '', '', '', '', ''],
    ['Approved by: Head of Department', '', '', '', '', '', '', '', '', '', '', '']
  ];

  const worksheet = XLSX.utils.aoa_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Semester Results');

  // Write and download
  XLSX.writeFile(workbook, 'Anna_University_Semester_Results_Sample.xlsx');
};
