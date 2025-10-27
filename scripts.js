// ข้อมูลทั้งหมด
let allData = [];
let filteredData = [];

// ข้อมูลจริง - ใส่ SHEET_ID ของ Google Sheets ตรงนี้
const SHEET_ID = '1CJhSu3XwPTC35SxXzGnOwz8a8c_D_1eI';
const SHEET_NAME = 'web app(อยู่ระหว่างการจัดทำ)';

// เริ่มต้นเมื่อโหลดหน้าเสร็จ
$(document).ready(function() {
    console.log('🚀 Starting application...');
    
    // โหลดข้อมูลทันที
    loadData();

    // ปุ่ม Excel
    $('.btn.excel').click(exportToExcel);
    
    // Dropdown: แสดงผลทันทีเมื่อเลือก
    $('#typeFilter, #yearFilter, #deptFilter, #rankFilter').on('change', function() {
        console.log('Dropdown changed:', $(this).attr('id'), '=', $(this).val());
        filterData();
    });
    
    // Input text: ค้นหาแบบ Real-time เมื่อพิมพ์
    $('#keyword, #authorFilter').on('input', function() {
        console.log('Input changed:', $(this).attr('id'), '=', $(this).val());
        filterData();
    });
});

// ฟังก์ชันโหลดข้อมูลจาก Google Sheets
function loadData() {
    console.log('📥 Loading data from Google Sheets...');
    console.log('SHEET_ID:', SHEET_ID);
    console.log('SHEET_NAME:', SHEET_NAME);
    
    if (!SHEET_ID || SHEET_ID === '') {
        alert('กรุณาใส่ SHEET_ID ในไฟล์ script.js');
        $('#dataTable tbody').html('<tr><td colspan="6" class="no-data">กรุณาใส่ SHEET_ID</td></tr>');
        return;
    }
    
    const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(SHEET_NAME)}`;
    console.log('📍 Request URL:', url);
    
    // แสดง Loading
    $('#dataTable tbody').html('<tr><td colspan="6" style="text-align:center; padding:20px;">⏳ กำลังโหลดข้อมูล...</td></tr>');
    
    $.ajax({
        url: url,
        dataType: 'text',
        timeout: 10000,
        success: function(response) {
            console.log('✅ AJAX Success! Response length:', response.length);
            console.log('First 100 chars:', response.substring(0, 100));
            
            try {
                // ลบ prefix และ suffix ของ Google Sheets API
                const jsonString = response.substring(47, response.length - 2);
                const data = JSON.parse(jsonString);
                
                console.log('📊 Parsed data structure:', {
                    cols: data.table.cols.length,
                    rows: data.table.rows.length
                });
                
                // แสดงชื่อคอลัมน์
                const columnNames = data.table.cols.map(col => col.label || 'unnamed');
                console.log('📋 Column names:', columnNames);
                
                // แปลงข้อมูลเป็น Array of Objects
                allData = data.table.rows.map(row => {
                    const obj = {};
                    row.c.forEach((cell, index) => {
                        const header = data.table.cols[index].label || `col${index}`;
                        obj[header] = cell ? (cell.v || '') : '';
                    });
                    return obj;
                });
                
                console.log('✅ Data loaded successfully:', allData.length, 'rows');
                console.log('📄 First row sample:', allData[0]);
                
                // แสดงข้อมูลทั้งหมดทันทีโดยไม่ต้องรอเลือก filter
                filteredData = [...allData];
                populateAuthorsDatalist();
                populateDropdownOptions(); // เพิ่มฟังก์ชันนี้
                displayData(filteredData);
                updateCount();
                
            } catch (error) {
                console.error('❌ Error parsing data:', error);
                console.error('Error stack:', error.stack);
                alert('เกิดข้อผิดพลาดในการแปลงข้อมูล\n\n' + error.message);
                $('#dataTable tbody').html('<tr><td colspan="6" class="no-data">เกิดข้อผิดพลาดในการแปลงข้อมูล<br>' + error.message + '</td></tr>');
            }
        },
        error: function(xhr, status, error) {
            console.error('❌ AJAX Error!');
            console.error('Status:', status);
            console.error('Error:', error);
            console.error('XHR Status:', xhr.status);
            console.error('XHR Response:', xhr.responseText);
            
            let errorMsg = '⚠️ ไม่สามารถโหลดข้อมูลได้\n\n';
            errorMsg += 'กรุณาตรวจสอบ:\n';
            errorMsg += '1. Google Sheets ต้องแชร์เป็น "Anyone with the link"\n';
            errorMsg += '2. ชื่อ Sheet (แท็บ) ต้องถูกต้อง: "' + SHEET_NAME + '"\n';
            errorMsg += '3. SHEET_ID ต้องถูกต้อง: ' + SHEET_ID + '\n\n';
            errorMsg += 'รายละเอียดข้อผิดพลาด:\n';
            errorMsg += 'Status: ' + status + '\n';
            errorMsg += 'Error: ' + error;
            
            alert(errorMsg);
            
            $('#dataTable tbody').html(`
                <tr><td colspan="6" class="no-data">
                    ❌ ไม่สามารถโหลดข้อมูลได้<br><br>
                    <strong>วิธีแก้ไข:</strong><br>
                    1. เปิด Google Sheets<br>
                    2. คลิก "Share" (แชร์) มุมขวาบน<br>
                    3. เลือก "Anyone with the link"<br>
                    4. สิทธิ์: "Viewer"<br>
                    5. คลิก "Done"<br><br>
                    <small>ตรวจสอบ Console (กด F12) สำหรับรายละเอียดเพิ่มเติม</small>
                </td></tr>
            `);
        }
    });
}

// ฟังก์ชันเติมข้อมูลใน Datalist ผู้สร้างสรรค์ (AutoComplete)
function populateAuthorsDatalist() {
    const authors = new Set();
    
    allData.forEach(row => {
        const keys = Object.keys(row);
        
        // คอลัมน์เจ้าของผลงาน (index 2)
        const authorColumn = keys[2];
        if (row[authorColumn] && row[authorColumn].trim() !== '') {
            authors.add(row[authorColumn].trim());
        }
    });

    const $datalist = $('#authorList');
    $datalist.empty();
    
    // เรียงตามตัวอักษร และเพิ่มเข้า datalist
    Array.from(authors).sort().forEach(author => {
        $datalist.append(`<option value="${author}">`);
    });
    
    console.log('👥 Authors loaded:', authors.size);
}

// ฟังก์ชันเติมข้อมูล Dropdown Options จากข้อมูลจริง
function populateDropdownOptions() {
    const years = new Set();
    
    allData.forEach(row => {
        const columns = Object.values(row);
        
        // ปี (คอลัมน์ที่ 1)
        if (columns[1] && columns[1].toString().trim() !== '') {
            const year = columns[1].toString().trim();
            // ดึงเฉพาะตัวเลขปี (กรณีมีข้อความอื่นปนมา)
            const yearMatch = year.match(/\d{4}/);
            if (yearMatch) {
                years.add(yearMatch[0]);
            }
        }
    });

    // เติมข้อมูล ปี
    const $yearFilter = $('#yearFilter');
    $yearFilter.find('option:not(:first)').remove(); // เก็บ option "ทั้งหมด"
    Array.from(years).sort((a, b) => a - b).forEach(year => {
        $yearFilter.append(`<option value="${year}">${year}</option>`);
    });
    
    console.log('📋 ปีที่มีในข้อมูล:', Array.from(years).sort());
}

// ฟังก์ชันกรองข้อมูล
function filterData() {
    const yearType = $('#typeFilter').val().trim();
    const year = $('#yearFilter').val().trim();
    const dept = $('#deptFilter').val().trim();
    const rank = $('#rankFilter').val().trim();
    const author = $('#authorFilter').val().toLowerCase().trim();
    const searchName = $('#keyword').val().toLowerCase().trim();

    console.log('🔍 Filtering with:', {yearType, year, dept, rank, author, searchName});

    // ถ้าไม่มีการกรองเลย แสดงข้อมูลทั้งหมด
    if (!yearType && !year && !dept && !rank && !author && !searchName) {
        filteredData = [...allData];
        displayData(filteredData);
        updateCount();
        return;
    }

    filteredData = allData.filter(row => {
        const columns = Object.values(row);
        
        // ประเภทของปี (คอลัมน์ที่ 0)
        if (yearType && !columns[0].toString().toLowerCase().includes(yearType.toLowerCase())) return false;
        
        // ปี (คอลัมน์ที่ 1)
        if (year && !columns[1].toString().includes(year)) return false;
        
        // เจ้าของผลงาน (คอลัมน์ที่ 2)
        if (author && !columns[2].toString().toLowerCase().includes(author)) return false;
        
        // สังกัดภาควิชา (คอลัมน์ที่ 3)
        if (dept && !columns[3].toString().includes(dept)) return false;
        
        // ชื่อผลงาน (คอลัมน์ที่ 4)
        if (searchName && !columns[4].toString().toLowerCase().includes(searchName)) return false;
        
        // ระดับผลงาน - ค้นหาในทุกคอลัมน์
        if (rank) {
            const allText = columns.join(' ');
            if (!allText.includes(rank)) return false;
        }
        
        return true;
    });

    console.log('✅ Filtered results:', filteredData.length);
    displayData(filteredData);
    updateCount();
}

// ฟังก์ชันแสดงข้อมูลในตาราง
function displayData(data) {
    const $tbody = $('#dataTable tbody');
    $tbody.empty();

    if (data.length === 0) {
        $tbody.html('<tr><td colspan="6" class="no-data">ไม่พบข้อมูล</td></tr>');
        return;
    }

    console.log('📊 Displaying', data.length, 'rows');

    data.forEach((row, index) => {
        const columns = Object.values(row);
        
        // Debug: แสดงข้อมูล 3 แถวแรก
        if (index < 3) {
            console.log(`Row ${index}:`, columns);
        }
        
        // คอลัมน์ที่ 6 คืออ้างอิง (index 5 หรือ 6 ขึ้นอยู่กับข้อมูล)
        const reference = columns[5] || columns[6] || '-';
        
        const rowHTML = `
            <tr>
                <td>${columns[0] || '-'}</td>
                <td>${columns[1] || '-'}</td>
                <td>${columns[2] || '-'}</td>
                <td>${columns[3] || '-'}</td>
                <td>${columns[4] || '-'}</td>
                <td>${reference}</td>
            </tr>
        `;
        $tbody.append(rowHTML);
    });
}

// ฟังก์ชันอัพเดทจำนวนผลลัพธ์
function updateCount() {
    $('#count').text(filteredData.length);
    console.log('📈 Count updated:', filteredData.length);
}

// ฟังก์ชัน Export ไปยัง Excel
function exportToExcel() {
    if (filteredData.length === 0) {
        alert('ไม่มีข้อมูลสำหรับ Export');
        return;
    }

    console.log('📤 Exporting', filteredData.length, 'rows to Excel');

    const wb = XLSX.utils.book_new();
    const headers = Object.keys(filteredData[0]);
    const data = filteredData.map(row => Object.values(row));
    const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
    
    XLSX.utils.book_append_sheet(wb, ws, 'ผลงานตีพิมพ์');
    
    const fileName = `ผลงานตีพิมพ์_${new Date().toLocaleDateString('th-TH')}.xlsx`;
    XLSX.writeFile(wb, fileName);
    
    console.log('✅ Excel file exported:', fileName);
}