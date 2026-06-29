import { NextResponse } from 'next/server';
import * as XLSX from 'xlsx';

export async function GET() {
  try {
    // Create workbook
    const wb = XLSX.utils.book_new();
    
    // Create worksheet data
    const wsData: any[] = [];
    
    // A1: Title
    wsData[0] = ['ระบบการเข้าและออกเวลางาน'];
    
    // A2: Subtitle with date range
    wsData[1] = ['วอร์ดม ศรีใสตั้งแต่ 1 พฤษภาคม 2569 ถึง 31 พฤษภาคม 2569'];
    
    // Empty row
    wsData[2] = [];
    
    // Row 4 (index 3): Headers
    wsData[3] = [
      'ลำดับ',
      'ชื่อ-นามสกุล',
      'วันที่เข้างาน',
      'เวลาเข้างาน',
      'วันที่ออกงาน',
      'เวลาออกงาน',
      'สถานะ',
      'รวมเวลางาน'
    ];
    
    // Sample data starting from row 5 (index 4)
    const sampleData = [
      [1, 'โรงแรม ศรีใส', 'เสาร์ 30 เมษายนคม 2569 20:04:21', '20:04', 'อาทิตย์ 1 พฤษภาคม 2569 08:00:54', '08:00', 'เข้างาน', '11:56'],
      [2, 'โรงแรม ศรีใส', 'หญิงสุดสวย 28 พฤษภาคม 2569 20:03:12', '20:03', 'ศุกร์ 29 พฤษภาคม 2569 08:00:52', '08:00', 'เข้างาน', '11:57'],
      [3, 'โรงแรม ศรีใส', 'บุษ 27 พฤษภาคม 2569 20:01:12', '20:01', 'หญิงสุดสวย 28 พฤษภาคม 2569 08:01:16', '08:01', 'เข้างาน', '12:00'],
      [4, 'โรงแรม ศรีใส', 'อิศรร 26 พฤษภาคม 2569 20:00:48', '20:00', 'เสาร์ 27 พฤษภาคม 2569 08:00:57', '08:00', 'เข้างาน', '12:00'],
      [5, 'โรงแรม ศรีใส', 'ธันร์ 25 พฤษภาคม 2569 20:01:14', '20:01', 'อิศรร 26 พฤษภาคม 2569 08:00:57', '08:00', 'เข้างาน', '11:59'],
      [6, 'โรงแรม ศรีใส', 'อาทิตย์ 24 พฤษภาคม 2569 20:01:39', '20:01', 'ธันร์ 25 พฤษภาคม 2569 08:13:07', '08:13', 'เข้างาน', '12:11'],
      [7, 'โรงแรม ศรีใส', 'เสาร์ 23 พฤษภาคม 2569 20:01:33', '20:01', 'อาทิตย์ 24 พฤษภาคม 2569 09:13:29', '09:13', 'เข้างาน', '13:11'],
      [8, 'โรงแรม ศรีใส', 'ศุกร์ 22 พฤษภาคม 2569 20:00:42', '20:00', 'เสาร์ 23 พฤษภาคม 2569 08:01:39', '08:01', 'เข้างาน', '12:00'],
      [9, 'โรงแรม ศรีใส', 'หญิงสุดสวย 21 พฤษภาคม 2569 20:00:38', '20:00', 'ศุกร์ 22 พฤษภาคม 2569 08:02:58', '08:02', 'เข้างาน', '12:02'],
      [10, 'โรงแรม ศรีใส', 'บุษ 20 พฤษภาคม 2569 20:01:06', '20:01', 'หญิงสุดสวย 21 พฤษภาคม 2569 08:00:38', '08:00', 'เข้างาน', '11:59'],
    ];
    
    // Add sample data
    sampleData.forEach((row, idx) => {
      wsData[4 + idx] = row;
    });
    
    // Create worksheet
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    
    // Set column widths
    ws['!cols'] = [
      { wch: 8 },  // ลำดับ
      { wch: 20 }, // ชื่อ-นามสกุล
      { wch: 30 }, // วันที่เข้างาน
      { wch: 12 }, // เวลาเข้างาน
      { wch: 30 }, // วันที่ออกงาน
      { wch: 12 }, // เวลาออกงาน
      { wch: 12 }, // สถานะ
      { wch: 15 }  // รวมเวลางาน
    ];
    
    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(wb, ws, 'เวลางาน');
    
    // Generate file
    const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
    
    // Return response
    return new NextResponse(excelBuffer, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="sample_timetrack.xlsx"',
      },
    });
  } catch (error) {
    console.error('Error generating Excel:', error);
    return NextResponse.json({ error: 'Failed to generate file' }, { status: 500 });
  }
}
