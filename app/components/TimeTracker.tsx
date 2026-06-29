'use client';

import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { Upload, Download, Edit2, Save, X, AlertCircle, FileText } from 'lucide-react';

interface TimeEntry {
  id: string;
  date: string;
  endDate: string;
  name: string;
  startTime: string;
  endTime: string;
  hoursWorked?: string;
  location?: string;
  latCheckin?: string;
  lonCheckin?: string;
  latCheckout?: string;
  lonCheckout?: string;
  workGroup?: string;
  notes?: string;
  isMissing?: boolean;
  dayNumber?: number;
  missingDateLabel?: string;
}

interface AnalysisResult {
  missingDates: string[];
  shortDays: TimeEntry[];
  totalHours: number;
}

interface ExcelHeader {
  title: string;
  subtitle: string;
}

interface SignatureStaff {
  name: string;
  position: string;
  signature: string;
}

const STAFF_LIST: SignatureStaff[] = [
  {
    name: 'นายวโรดม ศรีใส',
    position: 'พอต.ชง.',
    signature: '/assets/img/1.png',
  },
  {
    name: 'นายชัยวัฒน์ อัศวภูมิ',
    position: 'พอต.ชง.',
    signature: '/assets/img/3.png',
  },
  {
    name: 'นางสาวอทิติยา ไกรสีห์',
    position: 'พอต.ชง.',
    signature: '/assets/img/4.png',
  },
  {
    name: 'นางชญานิศ ภูจอมดาว',
    position: 'พอต.ชง.',
    signature: '/assets/img/2.png',
  },
];

const SUPERVISOR_STAFF: SignatureStaff = {
  name: 'นายสามารถ ปลอดกระโทก',
  position: 'ผส.ฝส.',
  signature: '/assets/img/5.png',
};

const findSignatureStaff = (name?: string) =>
  STAFF_LIST.find((staff) => staff.name === (name || '').trim());

const THAI_DATES = [
  'จันทร์ 1 พฤษภาคม 2569', 'อังคาร 2 พฤษภาคม 2569', 'พุธ 3 พฤษภาคม 2569',
  'พฤหัสบดี 4 พฤษภาคม 2569', 'ศุกร์ 5 พฤษภาคม 2569', 'เสาร์ 6 พฤษภาคม 2569',
  'อาทิตย์ 7 พฤษภาคม 2569', 'จันทร์ 8 พฤษภาคม 2569', 'อังคาร 9 พฤษภาคม 2569',
  'พุธ 10 พฤษภาคม 2569', 'พฤหัสบดี 11 พฤษภาคม 2569', 'ศุกร์ 12 พฤษภาคม 2569',
  'เสาร์ 13 พฤษภาคม 2569', 'อาทิตย์ 14 พฤษภาคม 2569', 'จันทร์ 15 พฤษภาคม 2569',
  'อังคาร 16 พฤษภาคม 2569', 'พุธ 17 พฤษภาคม 2569', 'พฤหัสบดี 18 พฤษภาคม 2569',
  'ศุกร์ 19 พฤษภาคม 2569', 'เสาร์ 20 พฤษภาคม 2569', 'อาทิตย์ 21 พฤษภาคม 2569',
  'จันทร์ 22 พฤษภาคม 2569', 'อังคาร 23 พฤษภาคม 2569', 'พุธ 24 พฤษภาคม 2569',
  'พฤหัสบดี 25 พฤษภาคม 2569', 'ศุกร์ 26 พฤษภาคม 2569', 'เสาร์ 27 พฤษภาคม 2569',
  'อาทิตย์ 28 พฤษภาคม 2569', 'จันทร์ 29 พฤษภาคม 2569', 'อังคาร 30 พฤษภาคม 2569',
  'พุธ 31 พฤษภาคม 2569'
];

export default function TimeTracker() {
  const [data, setData] = useState<TimeEntry[]>([]);
  const [fileName, setFileName] = useState('');
  const [excelHeader, setExcelHeader] = useState<ExcelHeader>({ title: '', subtitle: '' });
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editValues, setEditValues] = useState<Partial<TimeEntry>>({});
  const [analysis, setAnalysis] = useState<AnalysisResult | null>(null);
  const [showMissingDates, setShowMissingDates] = useState(false);
  const [showShortDays, setShowShortDays] = useState(false);
  const [selectedStaffName, setSelectedStaffName] = useState('');
  const [showAddRowDialog, setShowAddRowDialog] = useState(false);
  const [newRowData, setNewRowData] = useState<Partial<TimeEntry>>({
    date: '',
    endDate: '',
    name: '',
    startTime: '',
    endTime: '',
    location: '',
    latCheckin: '',
    lonCheckin: '',
    latCheckout: '',
    lonCheckout: '',
    workGroup: '',
    notes: ''
  });
  const selectedOperatorStaff = findSignatureStaff(selectedStaffName);

  // Parse Excel file
  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const content = e.target?.result;
        const workbook = XLSX.read(content, { type: 'binary' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        
        // Get all data from sheet
        const allData = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[];
        
        // Extract title from A1
        const title = allData[0]?.[0] || '';
        // Extract subtitle from A2  
        const subtitle = allData[1]?.[0] || '';
        
        setExcelHeader({ title, subtitle });
        
        // Extract headers from row 4 (index 3)
        const headerRow = allData[3] || [];
        
        // Map column indices
        const columnMap: { [key: string]: number } = {};
        headerRow.forEach((header: string, index: number) => {
          columnMap[header?.trim() || ''] = index;
        });

        // Extract data starting from row 5 (index 4)
        const transformedData: TimeEntry[] = [];
        for (let i = 4; i < allData.length; i++) {
          const row = allData[i];
          if (!row || row.every((cell: any) => !cell)) continue; // Skip empty rows

          const startDateTimeStr = getCellValue(row, columnMap, ['วันเวลาเข้างาน', 'วันที่เข้างาน']);
          const startTimeStr = getCellValue(row, columnMap, ['เวลาเข้างาน']);
          const endDateTimeStr = getCellValue(row, columnMap, ['วันเวลาออกงาน', 'วันที่ออกงาน']);
          const endTimeStr = getCellValue(row, columnMap, ['เวลาออกงาน']);
          const startParts = splitDateTime(startDateTimeStr, startTimeStr);
          const endParts = splitDateTime(endDateTimeStr, endTimeStr);
          const dayNumber = getDayNumber(startParts.date);

          transformedData.push({
            id: `${Date.now()}-${i}`,
            date: startParts.date,
            endDate: endParts.date,
            name: row[columnMap['ชื่อ-นามสกุล']] || '',
            startTime: startParts.time,
            endTime: endParts.time,
            hoursWorked: calculateHours(startParts.time, endParts.time),
            location: row[columnMap['สถานที่']] || '',
            latCheckin: row[columnMap['ละติจูด(เข้า)']] || '',
            lonCheckin: row[columnMap['ลองจิจูด(เข้า)']] || '',
            latCheckout: row[columnMap['ละติจูด(ออก)']] || '',
            lonCheckout: row[columnMap['ลองจิจูด(ออก)']] || '',
            workGroup: row[columnMap['กลุ่มงาน']] || '',
            notes: row[columnMap['บันทึกงาน/หมายเหตุ']] || '',
            dayNumber,
            isMissing: false
          });
        }

        const firstOperatorStaff = transformedData
          .map((entry) => findSignatureStaff(entry.name))
          .find((staff): staff is SignatureStaff => Boolean(staff));
        if (!selectedStaffName && firstOperatorStaff) {
          setSelectedStaffName(firstOperatorStaff.name);
        }

        setData(transformedData);
        analyzeDataAndAddMissing(transformedData);
      } catch (error) {
        alert('Error reading file: ' + error);
        console.error(error);
      }
    };

    reader.readAsBinaryString(file);
  };

  const getCellValue = (
    row: any[],
    columnMap: { [key: string]: number },
    headers: string[]
  ): any => {
    for (const header of headers) {
      const columnIndex = columnMap[header];
      if (columnIndex !== undefined) {
        return row[columnIndex] ?? '';
      }
    }

    return '';
  };

  const splitDateTime = (dateTimeValue: any, timeValue?: any): { date: string; time: string } => {
    const rawDateTime = String(dateTimeValue || '').trim();
    const dateTimeTime = parseTime(rawDateTime);
    const separateTime = parseTime(timeValue);
    const time = dateTimeTime.split(':').length === 3 ? dateTimeTime : (separateTime || dateTimeTime);
    const date = rawDateTime
      .replace(/\s+\d{1,2}:\d{2}(?::\d{2})?\s*$/, '')
      .trim();

    return { date, time };
  };

  // Extract day number from date string
  const getDayNumber = (dateStr: string): number => {
    const match = String(dateStr).match(/(\d{1,2})\s/);
    return match ? parseInt(match[1]) : 0;
  };

  // Parse time from various formats
  const parseTime = (timeValue: any): string => {
    if (!timeValue) return '';

    // If it's a number (Excel serial time)
    if (typeof timeValue === 'number') {
      const totalSeconds = Math.round(timeValue * 24 * 3600);
      const hours = Math.floor(totalSeconds / 3600);
      const minutes = Math.floor((totalSeconds % 3600) / 60);
      const seconds = totalSeconds % 60;
      return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
    }

    const timeMatch = String(timeValue).match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
    if (timeMatch) {
      const hours = String(timeMatch[1]).padStart(2, '0');
      const minutes = timeMatch[2];
      const seconds = timeMatch[3];
      return seconds ? `${hours}:${minutes}:${seconds}` : `${hours}:${minutes}`;
    }
    
    return '';
  };

  // Convert time to Thai format
  const formatTimeToThai = (timeStr: string): string => {
    if (!timeStr) return '';
    
    const [hours, minutes, seconds] = timeStr.split(':').map(Number);
    const h = hours || 0;
    const m = minutes || 0;
    const s = seconds || 0;
    
    if (h === 0 && m === 0 && s === 0) return '';
    
    let result = '';
    if (h > 0) {
      result += `${h} ชั่วโมง`;
    }
    if (m > 0) {
      if (result) result += ' ';
      result += `${m} นาที`;
    }
    if (s > 0) {
      if (result) result += ' ';
      result += `${s} วินาที`;
    } else if (h > 0 || m > 0) {
      if (result) result += ' ';
      result += '0 วินาที';
    }
    
    return result;
  };

  const getDurationSeconds = (timeStr: string): number => {
    const [hours, minutes, seconds] = (timeStr || '0:00:00').split(':').map(Number);
    return (hours || 0) * 3600 + (minutes || 0) * 60 + (seconds || 0);
  };

  const normalizeTimeInput = (timeStr: string): string => {
    const parts = String(timeStr || '').trim().split(':');
    if (parts.length < 2) return timeStr;

    const [hours, minutes, seconds = '0'] = parts;
    const h = Math.min(23, Math.max(0, Number(hours) || 0));
    const m = Math.min(59, Math.max(0, Number(minutes) || 0));
    const s = Math.min(59, Math.max(0, Number(seconds) || 0));

    return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
  };

  // Calculate hours between two times
  const calculateHours = (startTime: string, endTime: string): string => {
    if (!startTime || !endTime) return '0:00:00';
    
    try {
      const startParts = startTime.split(':').map(Number);
      const endParts = endTime.split(':').map(Number);
      
      if (startParts.length < 2 || endParts.length < 2) return '0:00:00';
      
      const startH = startParts[0] || 0;
      const startM = startParts[1] || 0;
      const startS = startParts[2] || 0;
      const endH = endParts[0] || 0;
      const endM = endParts[1] || 0;
      const endS = endParts[2] || 0;
      
      const start = startH * 3600 + startM * 60 + startS;
      const end = endH * 3600 + endM * 60 + endS;
      let diff = end - start;
      
      // Handle cases where end time is next day
      if (diff < 0) {
        diff += 24 * 3600; // Add 24 hours
      }
      
      const hours = Math.floor(diff / 3600);
      const minutes = Math.floor((diff % 3600) / 60);
      const seconds = diff % 60;
      
      return `${hours}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
    } catch {
      return '0:00:00';
    }
  };

  // Analyze data for missing dates and short work days
  const analyzeData = (entries: TimeEntry[]) => {
    // Extract day numbers from date strings
    const daysWorked = new Set<number>();
    entries.forEach(entry => {
      if (entry.isMissing) return;
      // Extract day number from date string like "เสาร์ 30 เมษายนคม 2569"
      const dayNumber = entry.dayNumber || getDayNumber(entry.date);
      if (dayNumber) {
        daysWorked.add(dayNumber);
      }
    });

    // Find missing dates in May
    const missingDates: string[] = [];
    const dayNames = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์', 'อาทิตย์'];
    
    for (let day = 1; day <= 31; day++) {
      if (!daysWorked.has(day)) {
        missingDates.push(`${day} พฤษภาคม 2569`);
      }
    }

    // Find days with less than 12 hours
    const shortDays = entries.filter(entry => {
      if (entry.isMissing) return false;
      const timeStr = entry.hoursWorked || '0:00';
      const totalSeconds = getDurationSeconds(timeStr);
      return totalSeconds < 12 * 3600 && totalSeconds > 0;
    });

    // Calculate total hours
    const totalHours = entries.reduce((sum, entry) => {
      if (entry.isMissing) return sum;
      const timeStr = entry.hoursWorked || '0:00';
      return sum + getDurationSeconds(timeStr) / 3600;
    }, 0);

    setAnalysis({
      missingDates,
      shortDays,
      totalHours
    });
  };

  // Analyze and add missing date rows
  const analyzeDataAndAddMissing = (entries: TimeEntry[]) => {
    const baseEntries = entries.filter(entry => !entry.isMissing);
    // Extract day numbers from date strings
    const daysWorked = new Set<number>();
    baseEntries.forEach(entry => {
      const dayNumber = entry.dayNumber || getDayNumber(entry.date);
      if (dayNumber) {
        daysWorked.add(dayNumber);
      }
    });

    // Find missing dates in May and create empty rows
    const missingDates: string[] = [];
    const missingEntries: TimeEntry[] = [];
    
    for (let day = 1; day <= 31; day++) {
      if (!daysWorked.has(day)) {
        const dateStr = `${day} พฤษภาคม 2569`;
        missingDates.push(dateStr);
        
        // Create empty row for missing date
        missingEntries.push({
          id: `missing-${day}`,
          date: '',
          endDate: '',
          missingDateLabel: dateStr,
          name: '',
          startTime: '',
          endTime: '',
          hoursWorked: '',
          location: '',
          latCheckin: '',
          lonCheckin: '',
          latCheckout: '',
          lonCheckout: '',
          workGroup: '',
          notes: '',
          isMissing: true,
          dayNumber: day
        });
      }
    }

    // Combine and sort by day number
    const combinedData = [...baseEntries, ...missingEntries].sort((a, b) => {
      const dayA = a.dayNumber || getDayNumber(a.date) || 0;
      const dayB = b.dayNumber || getDayNumber(b.date) || 0;
      return dayA - dayB;
    });

    setData(combinedData);

    // Find days with less than 12 hours
    const shortDays = combinedData.filter(entry => {
      if (entry.isMissing) return false;
      const timeStr = entry.hoursWorked || '0:00';
      const totalSeconds = getDurationSeconds(timeStr);
      return totalSeconds < 12 * 3600 && totalSeconds > 0;
    });

    // Calculate total hours
    const totalHours = combinedData.reduce((sum, entry) => {
      if (entry.isMissing) return sum;
      const timeStr = entry.hoursWorked || '0:00';
      return sum + getDurationSeconds(timeStr) / 3600;
    }, 0);

    setAnalysis({
      missingDates,
      shortDays,
      totalHours
    });
  };

  // Start editing
  const startEdit = (entry: TimeEntry) => {
    setEditingId(entry.id);
    setEditValues({ ...entry });
  };

  // Save changes
  const saveEdit = () => {
    if (!editingId) return;

    const updatedData = data.map(entry => {
      if (entry.id === editingId) {
        const updated = { ...entry, ...editValues };
        updated.dayNumber = getDayNumber(updated.date) || updated.dayNumber;
        updated.hoursWorked = calculateHours(
          updated.startTime,
          updated.endTime
        );
        if (updated.isMissing && (updated.name || updated.date || updated.endDate || updated.startTime || updated.endTime)) {
          updated.isMissing = false;
          updated.missingDateLabel = '';
        }
        return updated;
      }
      return entry;
    });

    setData(updatedData);
    setEditingId(null);
    analyzeDataAndAddMissing(updatedData);
  };

  // Cancel edit
  const cancelEdit = () => {
    setEditingId(null);
    setEditValues({});
  };

  // Download modified data
  const downloadExcel = () => {
    const escapeXml = (value: any) =>
      String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');

    const headers = [
      'ลำดับ',
      'ชื่อ-นามสกุล',
      'วันเวลาเข้างาน',
      'วันเวลาออกงาน',
      'รวมเวลางาน',
      'สถานที่',
      'ละติจูด(เข้า)',
      'ลองจิจูด(เข้า)',
      'ละติจูด(ออก)',
      'ลองจิจูด(ออก)',
      'กลุ่มงาน',
      'บันทึกงาน/หมายเหตุ'
    ];
    const columnWidths = [40, 110, 205, 205, 120, 90, 90, 90, 90, 90, 138, 115];
    const title = excelHeader.title || 'ระบบการเข้าและออกเวลางาน';
    const subtitle = excelHeader.subtitle || '';

    const cell = (value: any, styleId = 'Data', type: 'String' | 'Number' = 'String') =>
      `<Cell ss:StyleID="${styleId}"><Data ss:Type="${type}">${escapeXml(value)}</Data></Cell>`;
    const row = (cells: string[], height?: number) =>
      `<Row${height ? ` ss:Height="${height}"` : ''}>${cells.join('')}</Row>`;

    const headerRow = row(headers.map((header) => cell(header, 'Header')), 32);
    const dataRows = data.map((entry, idx) => {
      const style = entry.isMissing ? 'Missing' : 'Data';
      const hourStyle = entry.isMissing ? 'Missing' : 'Hours';
      const values = [
        idx + 1,
        entry.name,
        entry.isMissing ? '' : `${entry.date} ${entry.startTime}`.trim(),
        entry.isMissing ? '' : `${entry.endDate} ${entry.endTime}`.trim(),
        formatTimeToThai(entry.hoursWorked || ''),
        entry.location,
        entry.latCheckin,
        entry.lonCheckin,
        entry.latCheckout,
        entry.lonCheckout,
        entry.workGroup,
        entry.notes
      ];

      return row(
        values.map((value, valueIdx) =>
          cell(value, valueIdx === 4 ? hourStyle : style, typeof value === 'number' ? 'Number' : 'String')
        ),
        24
      );
    });
    const contentRowsBeforeFooter = 4 + dataRows.length;
    const footerGapRows = 2;
    const footerStartRow = Math.max(22, contentRowsBeforeFooter + footerGapRows + 1);
    const blankRowsBeforeFooter = Array.from(
      { length: Math.max(0, footerStartRow - contentRowsBeforeFooter - 1) },
      () => '<Row ss:Height="18"></Row>'
    );
    const operatorStaff = selectedOperatorStaff || data
      .map((entry) => findSignatureStaff(entry.name))
      .find((staff): staff is SignatureStaff => Boolean(staff));
    const operatorName = operatorStaff ? `(${operatorStaff.name})` : '(.......................................................)';
    const operatorPosition = operatorStaff?.position || '-';
    const supervisorName = `(${SUPERVISOR_STAFF.name})`;
    const footerCell = (columnIndex: number, value: string, styleId = 'FooterText', mergeAcross = 1) =>
      `<Cell ss:Index="${columnIndex}" ss:MergeAcross="${mergeAcross}" ss:StyleID="${styleId}"><Data ss:Type="String">${escapeXml(value)}</Data></Cell>`;
    const footerRows = [
      `<Row ss:Height="22">${footerCell(3, 'ลงชื่อ........................................')}${footerCell(9, 'ลงชื่อ........................................', 'FooterText', 2)}</Row>`,
      `<Row ss:Height="22">${footerCell(3, operatorName)}${footerCell(9, supervisorName, 'FooterText', 2)}</Row>`,
      `<Row ss:Height="22">${footerCell(3, `ตำแหน่ง...${operatorPosition}...`)}${footerCell(9, `ตำแหน่ง...${SUPERVISOR_STAFF.position}...`, 'FooterText', 2)}</Row>`,
      `<Row ss:Height="24">${footerCell(3, '', 'FooterRole')}${footerCell(9, 'ผส./ผู้ควบคุมและรับรองการปฏิบัติงาน', 'FooterRole', 2)}</Row>`
    ];

    const workbookXml = `<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
  <Styles>
    <Style ss:ID="Title">
      <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
      <Font ss:FontName="TH SarabunPSK" ss:Size="20"/>
      <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
    </Style>
    <Style ss:ID="Subtitle">
      <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
      <Font ss:FontName="TH SarabunPSK" ss:Size="16"/>
      <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
    </Style>
    <Style ss:ID="Header">
      <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
      <Font ss:FontName="TH SarabunPSK" ss:Size="12" ss:Bold="1" ss:Color="#000000"/>
      <Interior ss:Color="#92D050" ss:Pattern="Solid"/>
      <Borders>
        <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
        <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
        <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
        <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
      </Borders>
    </Style>
    <Style ss:ID="Data">
      <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
      <Font ss:FontName="TH SarabunPSK" ss:Size="12" ss:Bold="1"/>
      <Interior ss:Color="#FFFFCC" ss:Pattern="Solid"/>
      <Borders>
        <Border ss:Position="Bottom" ss:LineStyle="Dot" ss:Weight="1" ss:Color="#000000"/>
        <Border ss:Position="Left" ss:LineStyle="Dot" ss:Weight="1" ss:Color="#000000"/>
        <Border ss:Position="Right" ss:LineStyle="Dot" ss:Weight="1" ss:Color="#000000"/>
        <Border ss:Position="Top" ss:LineStyle="Dot" ss:Weight="1" ss:Color="#000000"/>
      </Borders>
    </Style>
    <Style ss:ID="Hours">
      <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
      <Font ss:FontName="TH SarabunPSK" ss:Size="12" ss:Bold="1"/>
      <Interior ss:Color="#FFFFCC" ss:Pattern="Solid"/>
      <Borders>
        <Border ss:Position="Bottom" ss:LineStyle="Dot" ss:Weight="1" ss:Color="#000000"/>
        <Border ss:Position="Left" ss:LineStyle="Dot" ss:Weight="1" ss:Color="#000000"/>
        <Border ss:Position="Right" ss:LineStyle="Dot" ss:Weight="1" ss:Color="#000000"/>
        <Border ss:Position="Top" ss:LineStyle="Dot" ss:Weight="1" ss:Color="#000000"/>
      </Borders>
    </Style>
    <Style ss:ID="Missing">
      <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
      <Font ss:FontName="TH SarabunPSK" ss:Size="12" ss:Bold="1"/>
      <Interior ss:Color="#FEE2E2" ss:Pattern="Solid"/>
      <Borders>
        <Border ss:Position="Bottom" ss:LineStyle="Dot" ss:Weight="1" ss:Color="#000000"/>
        <Border ss:Position="Left" ss:LineStyle="Dot" ss:Weight="1" ss:Color="#000000"/>
        <Border ss:Position="Right" ss:LineStyle="Dot" ss:Weight="1" ss:Color="#000000"/>
        <Border ss:Position="Top" ss:LineStyle="Dot" ss:Weight="1" ss:Color="#000000"/>
      </Borders>
    </Style>
    <Style ss:ID="FooterText">
      <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
      <Font ss:FontName="TH SarabunPSK" ss:Size="14" ss:Bold="1" ss:Color="#000000"/>
      <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
    </Style>
    <Style ss:ID="FooterRole">
      <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
      <Font ss:FontName="TH SarabunPSK" ss:Size="14" ss:Bold="1" ss:Color="#000000"/>
      <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
    </Style>
  </Styles>
  <Worksheet ss:Name="เวลางาน">
    <Table>
      ${columnWidths.map((width) => `<Column ss:Width="${width}"/>`).join('')}
      <Row ss:Height="30"><Cell ss:MergeAcross="${headers.length - 1}" ss:StyleID="Title"><Data ss:Type="String">${escapeXml(title)}</Data></Cell></Row>
      <Row ss:Height="24"><Cell ss:MergeAcross="${headers.length - 1}" ss:StyleID="Subtitle"><Data ss:Type="String">${escapeXml(subtitle)}</Data></Cell></Row>
      <Row ss:Height="8"></Row>
      ${headerRow}
      ${dataRows.join('')}
      ${blankRowsBeforeFooter.join('')}
      ${footerRows.join('')}
    </Table>
    <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
      <FreezePanes/>
      <FrozenNoSplit/>
      <SplitHorizontal>4</SplitHorizontal>
      <TopRowBottomPane>4</TopRowBottomPane>
      <ActivePane>2</ActivePane>
    </WorksheetOptions>
  </Worksheet>
</Workbook>`;

    const blob = new Blob([workbookXml], {
      type: 'application/vnd.ms-excel;charset=utf-8;'
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${fileName.replace(/\.(xlsx|xls)$/i, '') || 'timetrack'}_modified.xls`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const downloadExcelWithSignatures = async () => {
    const headers = [
      'ลำดับ',
      'ชื่อ-นามสกุล',
      'วันเวลาเข้างาน',
      'วันเวลาออกงาน',
      'รวมเวลางาน',
      'สถานที่',
      'ละติจูด(เข้า)',
      'ลองจิจูด(เข้า)',
      'ละติจูด(ออก)',
      'ลองจิจูด(ออก)',
      'กลุ่มงาน',
      'บันทึกงาน/หมายเหตุ'
    ];
    const columnWidths = [40, 110, 205, 205, 120, 90, 90, 90, 90, 90, 138, 115];
    const title = excelHeader.title || 'ระบบการเข้าและออกเวลางาน';
    const subtitle = excelHeader.subtitle || '';
    const operatorStaff = selectedOperatorStaff || data
      .map((entry) => findSignatureStaff(entry.name))
      .find((staff): staff is SignatureStaff => Boolean(staff));

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'worktmd';
    workbook.created = new Date();
    const worksheet = workbook.addWorksheet('เวลางาน', {
      pageSetup: { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
      views: [{ state: 'frozen', ySplit: 4 }]
    });

    worksheet.columns = columnWidths.map((width) => ({ width: Math.max(6, Math.round(width / 7)) }));
    worksheet.mergeCells('A1:L1');
    worksheet.mergeCells('A2:L2');
    worksheet.getCell('A1').value = title;
    worksheet.getCell('A2').value = subtitle;
    worksheet.getRow(1).height = 30;
    worksheet.getRow(2).height = 24;
    worksheet.getRow(3).height = 8;

    const thinBlack = { style: 'thin' as const, color: { argb: 'FF000000' } };
    const mediumBlack = { style: 'medium' as const, color: { argb: 'FF000000' } };
    const dashDotBlack = { style: 'dashDot' as const, color: { argb: 'FF000000' } };
    const headerBorder: Partial<ExcelJS.Borders> = { top: mediumBlack, left: mediumBlack, bottom: mediumBlack, right: mediumBlack };
    const dataBorder: Partial<ExcelJS.Borders> = { top: thinBlack, left: mediumBlack, bottom: dashDotBlack, right: mediumBlack };
    const center = { horizontal: 'center' as const, vertical: 'middle' as const, wrapText: true };
    const sarabunBold = { name: 'TH SarabunPSK', size: 12, bold: true };

    const styleCell = (
      cell: ExcelJS.Cell,
      fillColor: string,
      border: Partial<ExcelJS.Borders> | null = dataBorder,
      font: Partial<ExcelJS.Font> = sarabunBold
    ) => {
      cell.alignment = center;
      cell.font = font;
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
      if (border) {
        cell.border = border;
      } else {
        cell.border = {};
      }
    };

    styleCell(worksheet.getCell('A1'), 'FFFFFFFF', null, { name: 'TH SarabunPSK', size: 20, bold: true });
    styleCell(worksheet.getCell('A2'), 'FFFFFFFF', null, { name: 'TH SarabunPSK', size: 16, bold: true });

    const headerRow = worksheet.getRow(4);
    headerRow.height = 32;
    headers.forEach((header, index) => {
      const cell = headerRow.getCell(index + 1);
      cell.value = header;
      styleCell(cell, 'FF92D050', headerBorder, { ...sarabunBold, color: { argb: 'FF000000' } });
    });

    data.forEach((entry, index) => {
      const row = worksheet.getRow(index + 5);
      row.height = 24;
      const fillColor = entry.isMissing ? 'FFFEE2E2' : 'FFFFFFCC';
      const values = [
        index + 1,
        entry.name,
        entry.isMissing ? '' : `${entry.date} ${entry.startTime}`.trim(),
        entry.isMissing ? '' : `${entry.endDate} ${entry.endTime}`.trim(),
        formatTimeToThai(entry.hoursWorked || ''),
        entry.location,
        entry.latCheckin,
        entry.lonCheckin,
        entry.latCheckout,
        entry.lonCheckout,
        entry.workGroup,
        entry.notes
      ];

      values.forEach((value, valueIndex) => {
        const cell = row.getCell(valueIndex + 1);
        cell.value = value ?? '';
        styleCell(cell, fillColor);
      });
    });

    const tableStartRow = 4;
    const tableEndRow = data.length + 4;
    const tableStartCol = 1;
    const tableEndCol = headers.length;

    for (let rowNumber = tableStartRow; rowNumber <= tableEndRow; rowNumber += 1) {
      for (let colNumber = tableStartCol; colNumber <= tableEndCol; colNumber += 1) {
        const cell = worksheet.getCell(rowNumber, colNumber);
        const border = { ...(cell.border || {}) } as Partial<ExcelJS.Borders>;

        border.left = colNumber === tableStartCol ? mediumBlack : mediumBlack;
        border.right = colNumber === tableEndCol ? mediumBlack : mediumBlack;
        border.top = rowNumber === tableStartRow ? mediumBlack : thinBlack;
        border.bottom = rowNumber === tableEndRow ? mediumBlack : dashDotBlack;

        if (rowNumber === tableStartRow) {
          border.bottom = mediumBlack;
        }

        cell.border = border;
      }
    }

    const contentRowsBeforeFooter = 4 + data.length;
    const footerGapRows = 2;
    const footerStartRow = Math.max(22, contentRowsBeforeFooter + footerGapRows + 1);
    const operatorName = operatorStaff ? `(${operatorStaff.name})` : '(.......................................................)';
    const operatorPosition = operatorStaff?.position || '-';
    const supervisorName = `(${SUPERVISOR_STAFF.name})`;
    const signatureLabel = 'ลงชื่อ........................................';

    for (let rowNumber = contentRowsBeforeFooter + 1; rowNumber < footerStartRow; rowNumber += 1) {
      worksheet.getRow(rowNumber).height = 18;
    }

    const setFooter = (rowNumber: number, leftValue: string, rightValue: string) => {
      worksheet.mergeCells(`C${rowNumber}:D${rowNumber}`);
      worksheet.mergeCells(`I${rowNumber}:K${rowNumber}`);
      worksheet.getRow(rowNumber).height = rowNumber === footerStartRow ? 30 : 22;
      const leftCell = worksheet.getCell(`C${rowNumber}`);
      const rightCell = worksheet.getCell(`I${rowNumber}`);
      leftCell.value = leftValue;
      rightCell.value = rightValue;
      styleCell(leftCell, 'FFFFFFFF', null, { name: 'TH SarabunPSK', size: 14, bold: true });
      styleCell(rightCell, 'FFFFFFFF', null, { name: 'TH SarabunPSK', size: 14, bold: true });
    };

    setFooter(footerStartRow, signatureLabel, signatureLabel);
    setFooter(footerStartRow + 1, operatorName, supervisorName);
    setFooter(footerStartRow + 2, `ตำแหน่ง...${operatorPosition}...`, `ตำแหน่ง...${SUPERVISOR_STAFF.position}...`);
    setFooter(footerStartRow + 3, '', 'ผส./ผู้ควบคุมและรับรองการปฏิบัติงาน');

    const addSignatureImage = async (staff: SignatureStaff | undefined, col: number) => {
      if (!staff) return;
      const response = await fetch(staff.signature);
      const imageBuffer = await response.arrayBuffer();
      const imageId = workbook.addImage({ buffer: imageBuffer, extension: 'png' });
      worksheet.addImage(imageId, {
        tl: { col, row: footerStartRow - 0.95 },
        ext: { width: 120, height: 36 },
        editAs: 'oneCell'
      });
    };

    await addSignatureImage(operatorStaff, 2.85);
    await addSignatureImage(SUPERVISOR_STAFF, 8.85);

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer as BlobPart], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${fileName.replace(/\.(xlsx|xls)$/i, '') || 'timetrack'}_modified.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const downloadPDF = () => {
    const printWindow = window.open('', '_blank');
    if (!printWindow) {
      alert('Please allow popups to export PDF');
      return;
    }

    const escapeHtml = (value: any) =>
      String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');

    const headers = [
      'ลำดับ',
      'ชื่อ-นามสกุล',
      'วันเวลาเข้างาน(วันที่)',
      'เวลาเข้างาน(เวลา)',
      'วันเวลาออกงาน(วันที่)',
      'เวลาออกงาน(เวลา)',
      'รวมเวลางาน',
      'สถานที่',
      'ละติจูด(เข้า)',
      'ลองจิจูด(เข้า)',
      'ละติจูด(ออก)',
      'ลองจิจูด(ออก)',
      'กลุ่มงาน',
      'บันทึกงาน/หมายเหตุ'
    ];

    const rows = data.map((entry, idx) => {
      const values = [
        idx + 1,
        entry.name,
        entry.isMissing ? '' : entry.date,
        entry.isMissing ? '' : entry.startTime,
        entry.isMissing ? '' : entry.endDate,
        entry.isMissing ? '' : entry.endTime,
        formatTimeToThai(entry.hoursWorked || ''),
        entry.location,
        entry.latCheckin,
        entry.lonCheckin,
        entry.latCheckout,
        entry.lonCheckout,
        entry.workGroup,
        entry.notes
      ];

      return `<tr class="${entry.isMissing ? 'missing' : ''}">
        ${values.map((value, valueIdx) => `<td class="${valueIdx === 6 ? 'hours' : ''}">${escapeHtml(value)}</td>`).join('')}
      </tr>`;
    });

    const title = excelHeader.title || 'ระบบการเข้าและออกเวลางาน';
    const subtitle = excelHeader.subtitle || '';
    const operatorStaff = selectedOperatorStaff || data
      .map((entry) => findSignatureStaff(entry.name))
      .find((staff): staff is SignatureStaff => Boolean(staff));
    const signatureBlock = (staff: SignatureStaff | undefined, role: string, showRole = false) => {
      const name = staff?.name || '.......................................................';
      const position = staff?.position || '-';
      const signature = staff
        ? `<img class="signature-image" src="${escapeHtml(staff.signature)}" alt="${escapeHtml(staff.name)}" />`
        : '';

      return `<div class="signature-block">
        <div class="signature-image-row">${signature}</div>
        <div>ลงชื่อ..................................................</div>
        <div>(${escapeHtml(name)})</div>
        <div>ตำแหน่ง...${escapeHtml(position)}...</div>
        ${showRole ? `<div class="signature-role">${escapeHtml(role)}</div>` : ''}
      </div>`;
    };

    printWindow.document.write(`<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>${escapeHtml(title)}</title>
  <style>
    @page { size: A4 landscape; margin: 8mm; }
    * { box-sizing: border-box; }
    body {
      font-family: Tahoma, Arial, sans-serif;
      color: #111827;
      margin: 0;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }
    .title {
      background: #ffffff;
      border: 1px solid #94a3b8;
      padding: 10px 12px 6px;
      text-align: center;
      font-size: 20px;
      font-weight: 700;
    }
    .subtitle {
      background: #ffffff;
      border-left: 1px solid #94a3b8;
      border-right: 1px solid #94a3b8;
      border-bottom: 1px solid #94a3b8;
      padding: 4px 12px 10px;
      text-align: center;
      font-size: 16px;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      table-layout: fixed;
      margin-top: 8px;
      font-size: 8px;
    }
    th, td {
      padding: 4px 3px;
      text-align: center;
      vertical-align: middle;
      overflow-wrap: anywhere;
    }
    th {
      border: 1px solid #000000;
      background: #00a63e;
      color: #ffffff;
      font-weight: 700;
    }
    td {
      border: 1px dashed #000000;
      background: #fefce8;
      font-weight: 700;
    }
    td.hours {
      background: #dbeafe;
      font-weight: 700;
    }
    tr.missing td {
      background: #fee2e2;
      border-color: #000000;
    }
    tr.missing td.hours {
      background: #fee2e2;
    }
    .signature-footer {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 48px;
      margin-top: 46px;
      page-break-inside: avoid;
      font-size: 14px;
      font-weight: 700;
    }
    .signature-block {
      text-align: center;
      line-height: 1.45;
      min-width: 260px;
    }
    .signature-image-row {
      height: 38px;
      display: flex;
      align-items: flex-end;
      justify-content: center;
      margin-bottom: -2px;
    }
    .signature-image {
      width: 120px;
      height: 36px;
      object-fit: contain;
    }
    .signature-role {
      font-weight: 700;
    }
  </style>
</head>
<body>
  <div class="title">${escapeHtml(title)}</div>
  <div class="subtitle">${escapeHtml(subtitle)}</div>
  <table>
    <thead>
      <tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join('')}</tr>
    </thead>
    <tbody>${rows.join('')}</tbody>
  </table>
  <div class="signature-footer">
    ${signatureBlock(operatorStaff, 'ผู้ปฏิบัติงาน')}
    ${signatureBlock(SUPERVISOR_STAFF, 'ผส./ผู้ควบคุมและรับรองการปฏิบัติงาน', true)}
  </div>
  <script>
    window.addEventListener('load', () => {
      setTimeout(() => window.print(), 250);
    });
  </script>
</body>
</html>`);
    printWindow.document.close();
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-gray-50 to-gray-100">
      <div className="w-full px-4 py-6">
        <div className="bg-white border-4 border-gray-400 p-0 shadow-2xl rounded-lg overflow-hidden">
          {/* Header - Thai Government Style */}
          <div className="border-b-4 border-gray-400 p-8 bg-gradient-to-r from-blue-600 via-blue-500 to-blue-400 text-center shadow-lg">
            {excelHeader.title ? (
              <>
                <h1 className="text-3xl font-bold text-white mb-2 drop-shadow-md">
                  {excelHeader.title}
                </h1>
                <p className="text-sm text-blue-100 drop-shadow-sm">
                  {excelHeader.subtitle}
                </p>
              </>
            ) : (
              <>
                <h1 className="text-3xl font-bold text-white mb-2 drop-shadow-md">
                  ระบบการเข้าและออกเวลางาน
                </h1>
                <p className="text-sm text-blue-100 drop-shadow-sm">
                  วอร์ดม ศรีใสตั้งแต่ 1 พฤษภาคม 2569 ถึง 31 พฤษภาคม 2569
                </p>
              </>
            )}
          </div>

          {/* File Upload */}
          <div className="p-8 border-b-4 border-gray-300 bg-gradient-to-b from-white to-gray-50">
            <div className="mb-6 flex gap-4">
              <label className="flex-1">
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleFileUpload}
                  className="hidden"
                />
                <div className="flex items-center justify-center w-full px-6 py-12 border-3 border-dashed border-blue-400 rounded-xl hover:border-blue-600 cursor-pointer transition-all duration-300 bg-gradient-to-br from-blue-50 to-blue-100 hover:shadow-lg hover:bg-gradient-to-br hover:from-blue-100 hover:to-blue-200">
                  <div className="text-center">
                    <Upload className="w-16 h-16 text-blue-600 mx-auto mb-4 drop-shadow" />
                    <p className="text-xl font-bold text-gray-800">
                      คลิกเพื่ออัพโหลดไฟล์ Excel
                    </p>
                    <p className="text-sm text-gray-600 mt-3 font-medium">
                      (.xlsx หรือ .xls)
                    </p>
                  </div>
                </div>
              </label>
            </div>

            <div className="text-center">
              <button
                onClick={() => window.location.href = '/api/sample-excel'}
                className="bg-gradient-to-r from-blue-500 to-blue-600 hover:from-blue-600 hover:to-blue-700 text-white px-8 py-3 rounded-lg text-sm transition-all duration-300 shadow-md hover:shadow-lg hover:scale-105 font-semibold flex items-center gap-2 mx-auto"
              >
                <Download className="w-4 h-4" />
                ดาวน์โหลด Sample Excel
              </button>
            </div>
          </div>

          {/* Analysis Summary */}
          {analysis && data.length > 0 && (
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6 p-8 border-b-4 border-gray-300 bg-gradient-to-r from-gray-50 to-white">
              {/* Total Hours */}
              <div className="bg-gradient-to-br from-green-50 to-green-100 border-2 border-green-400 rounded-xl p-6 shadow-md hover:shadow-lg transition-all hover:scale-105">
                <h3 className="text-sm font-bold text-green-800 mb-3 uppercase tracking-wider">
                  รวมชั่วโมงทั้งหมด
                </h3>
                <p className="text-4xl font-bold text-green-600 drop-shadow">
                  {analysis.totalHours.toFixed(1)}
                </p>
                <p className="text-xs text-green-700 mt-1 font-medium">ชั่วโมง</p>
              </div>

              {/* Missing Dates */}
              <div className="bg-gradient-to-br from-yellow-50 to-yellow-100 border-2 border-yellow-400 rounded-xl p-6 shadow-md hover:shadow-lg transition-all hover:scale-105 cursor-pointer" onClick={() => setShowMissingDates(!showMissingDates)}>
                <h3 className="text-sm font-bold text-yellow-800 mb-3 uppercase tracking-wider flex items-center">
                  <AlertCircle className="w-5 h-5 mr-2" />
                  วันที่ขาด
                </h3>
                <p className="text-4xl font-bold text-yellow-600 drop-shadow">
                  {analysis.missingDates.length}
                </p>
                <p className="text-xs text-yellow-700 mt-1 font-medium">วัน</p>
                {showMissingDates && analysis.missingDates.length > 0 && (
                  <div className="mt-4 pt-4 border-t border-yellow-300 max-h-32 overflow-y-auto">
                    <p className="text-xs font-bold text-yellow-800 mb-2">รายละเอียด:</p>
                    {analysis.missingDates.map((date, idx) => (
                      <p key={idx} className="text-xs text-yellow-700 py-1">• {date}</p>
                    ))}
                  </div>
                )}
              </div>

              {/* Short Days */}
              <div className="bg-gradient-to-br from-red-50 to-red-100 border-2 border-red-400 rounded-xl p-6 shadow-md hover:shadow-lg transition-all hover:scale-105 cursor-pointer" onClick={() => setShowShortDays(!showShortDays)}>
                <h3 className="text-sm font-bold text-red-800 mb-3 uppercase tracking-wider flex items-center">
                  <AlertCircle className="w-5 h-5 mr-2" />
                  น้อยกว่า 12 ชั่วโมง
                </h3>
                <p className="text-4xl font-bold text-red-600 drop-shadow">
                  {analysis.shortDays.length}
                </p>
                <p className="text-xs text-red-700 mt-1 font-medium">วัน</p>
                {showShortDays && analysis.shortDays.length > 0 && (
                  <div className="mt-4 pt-4 border-t border-red-300 max-h-32 overflow-y-auto">
                    <p className="text-xs font-bold text-red-800 mb-2">รายละเอียด:</p>
                    {analysis.shortDays.map((entry, idx) => (
                      <p key={idx} className="text-xs text-red-700 py-1">• {entry.name} ({entry.date}): {formatTimeToThai(entry.hoursWorked || '')}</p>
                    ))}
                  </div>
                )}
              </div>
            </div>
          )}

          {/* Data Table - Thai Government Style */}
          {data.length > 0 && (
            <div className="mb-8 overflow-x-auto border-2 border-gray-400">
              <div className="p-4 bg-green-50 border-b border-gray-400">
                <h3 className="text-center font-bold text-lg text-gray-800">
                  {excelHeader.title || 'ระบบการเข้าและออกเวลางาน'}
                </h3>
                <p className="text-center text-sm text-gray-700 mt-1">
                  {excelHeader.subtitle || ''}
                </p>
              </div>
              
              <table className="w-full border-collapse font-bold">
                <thead>
                  <tr className="bg-green-600 text-white">
                    <th className="border border-gray-400 px-3 py-2 text-center text-sm font-bold w-8 text-green-900">ลำดับ</th>
                    <th className="border border-gray-400 px-3 py-2 text-center text-sm font-bold min-w-32 text-green-900">ชื่อ-นามสกุล</th>
                    <th className="border border-gray-400 px-3 py-2 text-center text-sm font-bold min-w-40 text-green-900">วันเวลาเข้างาน(วันที่)</th>
                    <th className="border border-gray-400 px-3 py-2 text-center text-sm font-bold min-w-28 text-green-900">เวลาเข้างาน(เวลา)</th>
                    <th className="border border-gray-400 px-3 py-2 text-center text-sm font-bold min-w-40 text-green-900">วันเวลาออกงาน(วันที่)</th>
                    <th className="border border-gray-400 px-3 py-2 text-center text-sm font-bold min-w-28 text-green-900">เวลาออกงาน(เวลา)</th>
                    <th className="border border-gray-400 px-3 py-2 text-center text-sm font-bold min-w-24 text-green-900">รวมเวลางาน</th>
                    <th className="border border-gray-400 px-3 py-2 text-center text-sm font-bold min-w-32 text-green-900">สถานที่</th>
                    <th className="border border-gray-400 px-3 py-2 text-center text-sm font-bold min-w-24 text-green-900">ละติจูด(เข้า)</th>
                    <th className="border border-gray-400 px-3 py-2 text-center text-sm font-bold min-w-24 text-green-900">ลองจิจูด(เข้า)</th>
                    <th className="border border-gray-400 px-3 py-2 text-center text-sm font-bold min-w-24 text-green-900">ละติจูด(ออก)</th>
                    <th className="border border-gray-400 px-3 py-2 text-center text-sm font-bold min-w-24 text-green-900">ลองจิจูด(ออก)</th>
                    <th className="border border-gray-400 px-3 py-2 text-center text-sm font-bold min-w-24 text-green-900">กลุ่มงาน</th>
                    <th className="border border-gray-400 px-3 py-2 text-center text-sm font-bold min-w-32 text-green-900">บันทึกงาน/หมายเหตุ</th>
                    <th className="border border-gray-400 px-3 py-2 text-center text-sm font-bold min-w-20 text-green-900">แก้ไข</th>
                  </tr>
                </thead>
                <tbody>
                  {data.map((entry, idx) => (
                    <tr key={entry.id} className={`transition ${editingId === entry.id ? 'editing-row' : ''} ${entry.isMissing ? 'missing-row' : 'hover:bg-yellow-100'}`}>
                      {editingId === entry.id ? (
                        <>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50 text-center text-sm text-gray-800 font-medium">{idx + 1}</td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50"><input type="text" value={editValues.name || ''} onChange={(e) => setEditValues({...editValues, name: e.target.value})} className="w-full px-2 py-1 border border-gray-300 rounded text-sm" /></td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50"><input type="text" value={editValues.date || ''} onChange={(e) => setEditValues({...editValues, date: e.target.value})} className="w-full px-2 py-1 border border-gray-300 rounded text-sm" placeholder="วันที่เข้า" /></td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50"><input type="text" inputMode="numeric" value={editValues.startTime || ''} onChange={(e) => setEditValues({...editValues, startTime: e.target.value})} onBlur={(e) => setEditValues({...editValues, startTime: normalizeTimeInput(e.target.value)})} className="w-full px-2 py-1 border border-gray-300 rounded text-sm text-center" placeholder="20:02:24" /></td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50"><input type="text" value={editValues.endDate || ''} onChange={(e) => setEditValues({...editValues, endDate: e.target.value})} className="w-full px-2 py-1 border border-gray-300 rounded text-sm" placeholder="วันที่ออก" /></td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50"><input type="text" inputMode="numeric" value={editValues.endTime || ''} onChange={(e) => setEditValues({...editValues, endTime: e.target.value})} onBlur={(e) => setEditValues({...editValues, endTime: normalizeTimeInput(e.target.value)})} className="w-full px-2 py-1 border border-gray-300 rounded text-sm text-center" placeholder="08:13:16" /></td>
                          <td className="border border-gray-400 px-3 py-2 bg-blue-100 font-bold text-center text-sm text-gray-800">{formatTimeToThai(calculateHours(editValues.startTime || '', editValues.endTime || ''))}</td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50"><input type="text" value={editValues.location || ''} onChange={(e) => setEditValues({...editValues, location: e.target.value})} className="w-full px-2 py-1 border border-gray-300 rounded text-sm" /></td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50"><input type="text" value={editValues.latCheckin || ''} onChange={(e) => setEditValues({...editValues, latCheckin: e.target.value})} className="w-full px-2 py-1 border border-gray-300 rounded text-sm" /></td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50"><input type="text" value={editValues.lonCheckin || ''} onChange={(e) => setEditValues({...editValues, lonCheckin: e.target.value})} className="w-full px-2 py-1 border border-gray-300 rounded text-sm" /></td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50"><input type="text" value={editValues.latCheckout || ''} onChange={(e) => setEditValues({...editValues, latCheckout: e.target.value})} className="w-full px-2 py-1 border border-gray-300 rounded text-sm" /></td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50"><input type="text" value={editValues.lonCheckout || ''} onChange={(e) => setEditValues({...editValues, lonCheckout: e.target.value})} className="w-full px-2 py-1 border border-gray-300 rounded text-sm" /></td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50"><input type="text" value={editValues.workGroup || ''} onChange={(e) => setEditValues({...editValues, workGroup: e.target.value})} className="w-full px-2 py-1 border border-gray-300 rounded text-sm" /></td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50"><input type="text" value={editValues.notes || ''} onChange={(e) => setEditValues({...editValues, notes: e.target.value})} className="w-full px-2 py-1 border border-gray-300 rounded text-sm" /></td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50 text-center"><div className="flex gap-1 justify-center"><button onClick={saveEdit} className="bg-green-600 hover:bg-green-700 text-white px-2 py-1 rounded text-xs flex items-center gap-1 transition"><Save className="w-3 h-3" />บันทึก</button><button onClick={cancelEdit} className="bg-red-500 hover:bg-red-600 text-white px-2 py-1 rounded text-xs flex items-center gap-1 transition"><X className="w-3 h-3" /></button></div></td>
                        </>
                      ) : (
                        <>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50 text-center text-sm text-gray-800 font-medium">{idx + 1}</td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50 text-sm text-gray-800">{entry.name}</td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50 text-center text-sm text-gray-800">{entry.isMissing ? '' : entry.date}</td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50 text-center text-sm text-gray-800">{entry.isMissing ? '' : entry.startTime}</td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50 text-center text-sm text-gray-800">{entry.isMissing ? '' : entry.endDate}</td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50 text-center text-sm text-gray-800">{entry.isMissing ? '' : entry.endTime}</td>
                          <td className="border border-gray-400 px-3 py-2 bg-blue-100 font-bold text-center text-sm text-gray-800">{formatTimeToThai(entry.hoursWorked || '')}</td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50 text-center text-sm text-gray-800">{entry.location}</td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50 text-center text-sm text-gray-800">{entry.latCheckin}</td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50 text-center text-sm text-gray-800">{entry.lonCheckin}</td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50 text-center text-sm text-gray-800">{entry.latCheckout}</td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50 text-center text-sm text-gray-800">{entry.lonCheckout}</td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50 text-center text-sm text-gray-800">{entry.workGroup}</td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50 text-sm text-gray-800">{entry.notes}</td>
                          <td className="border border-gray-400 px-3 py-2 bg-yellow-50 text-center"><button onClick={() => startEdit(entry)} className="bg-blue-600 hover:bg-blue-700 text-white px-3 py-1 rounded text-xs flex items-center gap-1 transition mx-auto"><Edit2 className="w-3 h-3" />แก้ไข</button></td>
                        </>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* Download Button */}
          {data.length > 0 && (
            <div className="p-8 border-t-4 border-gray-300 bg-gradient-to-r from-gray-50 to-white">
              <div className="mb-6 grid grid-cols-1 lg:grid-cols-[minmax(220px,360px)_1fr] gap-4 items-end">
                <div className="flex-1 min-w-[220px]">
                  <label htmlFor="staffSelect" className="block mb-1 text-sm font-bold text-gray-800">
                    ผู้ปฏิบัติงาน
                  </label>
                  <select
                    id="staffSelect"
                    name="staff_name"
                    required
                    value={selectedStaffName}
                    onChange={(event) => setSelectedStaffName(event.target.value)}
                    className="border border-gray-500 rounded w-full px-3 py-2 text-gray-900 font-bold bg-white focus:border-blue-600 focus:ring-2 focus:ring-blue-200 outline-none"
                  >
                    <option value="">-- เลือกผู้ปฏิบัติงาน --</option>
                    {STAFF_LIST.map((staff) => (
                      <option key={staff.name} value={staff.name}>
                        {staff.name}
                      </option>
                    ))}
                  </select>
                </div>

                {selectedOperatorStaff && (
                  <div className="flex items-center gap-4 rounded border border-gray-300 bg-white px-4 py-3">
                    <img
                      src={selectedOperatorStaff.signature}
                      alt={selectedOperatorStaff.name}
                      className="h-10 w-28 object-contain"
                    />
                    <div className="text-sm text-gray-900 font-bold">
                      <div>{selectedOperatorStaff.name}</div>
                      <div>ตำแหน่ง...{selectedOperatorStaff.position}...</div>
                    </div>
                  </div>
                )}
              </div>

              <div className="flex justify-center gap-4">
                <button
                  onClick={downloadExcelWithSignatures}
                  className="bg-gradient-to-r from-green-500 to-green-600 hover:from-green-600 hover:to-green-700 text-white px-10 py-4 rounded-xl flex items-center gap-3 transition-all duration-300 shadow-lg hover:shadow-2xl hover:scale-105 font-bold text-lg"
                >
                  <Download className="w-6 h-6" />
                  ดาวน์โหลด Excel
                </button>
                <button
                  onClick={downloadPDF}
                  className="bg-gradient-to-r from-red-500 to-red-600 hover:from-red-600 hover:to-red-700 text-white px-10 py-4 rounded-xl flex items-center gap-3 transition-all duration-300 shadow-lg hover:shadow-2xl hover:scale-105 font-bold text-lg"
                >
                  <FileText className="w-6 h-6" />
                  ดาวน์โหลด PDF
                </button>
              </div>
            </div>
          )}

          {/* Empty State */}
          {data.length === 0 && !analysis && (
            <div className="text-center py-20 border-t-4 border-gray-300 bg-gradient-to-b from-white to-gray-50">
              <Upload className="w-20 h-20 text-gray-300 mx-auto mb-6 drop-shadow" />
              <p className="text-gray-500 text-xl font-medium">
                ยังไม่มีข้อมูล กรุณาอัพโหลดไฟล์ Excel
              </p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

