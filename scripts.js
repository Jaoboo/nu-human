// ข้อมูลทั้งหมด
let allData = [];
let filteredData = [];

// ฟังก์ชัน normalize ข้อความ - ลบช่องว่างทั้งหมดและแปลงเป็นตัวพิมพ์เล็ก
function normalizeText(text) {
    if (!text) return '';
    return text.toString().replace(/\s+/g, '').toLowerCase();
}

// Default Options สำหรับแต่ละ Dropdown
const defaultOptions = {
    dept: [
        "ทั้งหมด",
        "ภาควิชาภาษาศาสตร์ คติชนวิทยา ปรัชญาและศาสนา",
        "ภาควิชาภาษาตะวันตก",
        "ภาควิชาดนตรี",
        "ผู้ตีพิมพ์มาจากหลายภาควิชา",
        "ภาควิชาภาษาตะวันออก",
        "ภาควิชาภาษาไทย",
        "ภาควิชาภาษาอังกฤษ",
        "สำนักงานเลขานุการคณะ"
    ],
    rank: [
        "ทั้งหมด",
        "วารสารระดับชาติ TCI กลุ่ม 1",
        "วารสารระดับชาติ TCI กลุ่ม 2",
        "Proceedings ระดับนานาชาติ",
        "วารสารระดับนานาชาติ Q1",
        "วารสารระดับนานาชาติ Q2",
        "วารสารระดับนานาชาติ Q3",
        "วารสารระดับนานาชาติ Q4",
    ]
};

// ข้อมูลจริง - ใส่ SHEET_ID ของ Google Sheets ตรงนี้
const SHEET_ID = '1CJhSu3XwPTC35SxXzGnOwz8a8c_D_1eI';
const SHEET_NAME = 'web app(อยู่ระหว่างจัดทำ)';

// เริ่มต้นเมื่อโหลดหน้าเสร็จ
$(document).ready(function() {
    
    // โหลดข้อมูลทันที
    loadData();

    // ปุ่ม Reset
    $('.btn.reset').click(function() {
        // ล้างค่า filters
        $('#typeFilter').val('ปีการศึกษา');
        $('#yearFilter').val('');
        $('#deptFilter').val('');
        $('#rankFilter').val('');
        $('#authorFilter').val('');
        $('#keyword').val('');
        
        // แสดงข้อมูลทั้งหมด
        filteredData = [...allData];
        displayData(filteredData);
        updateCount();
    });

    // ปุ่ม Refresh - โหลดข้อมูลใหม่จาก Google Sheets
    $('.btn.refresh').click(function() {
        const $icon = $(this).find('i');
        $icon.addClass('loading');
        
        // โหลดข้อมูลใหม่
        loadData(true);
    });

    // ปุ่ม Excel - Export เฉพาะข้อมูลที่แสดงในตาราง
    $('.btn.excel, .btn.excel-table').click(function() {
        exportTableToExcel();
    });
    
    // Dropdown: แสดงผลทันทีเมื่อเลือก
    $('#typeFilter').on('change', function() {
        console.log('Type filter changed:', $(this).val());
        // อัพเดท dropdown ปีตามประเภทที่เลือก
        populateDropdownOptions();
        // กรองข้อมูลใหม่
        filterData();
    });
    
    $('#yearFilter, #deptFilter, #rankFilter').on('change', function() {
        console.log('Dropdown changed:', $(this).attr('id'), '=', $(this).val());
        filterData();
    });
    
    // Input text: ค้นหาแบบ Real-time เมื่อพิมพ์
    $('#keyword, #authorFilter').on('input', function() {
        console.log('Input changed:', $(this).attr('id'), '=', $(this).val());
        filterData();
    });
    
    // ปุ่มค้นหา - กรองข้อมูลอีกครั้ง
    $('.btn.search').click(function() {
        filterData();
    });
});

// ฟังก์ชันโหลดข้อมูลจาก Google Sheets
function loadData(isRefresh = false) {
    console.log('📥 Loading data from Google Sheets...');
    console.log('SHEET_ID:', SHEET_ID);
    console.log('SHEET_NAME:', SHEET_NAME);
    
    if (!SHEET_ID || SHEET_ID === '') {
        alert('กรุณาใส่ SHEET_ID ในไฟล์ script.js');
        $('#dataTable tbody').html('<tr><td colspan="8" class="no-data">กรุณาใส่ SHEET_ID</td></tr>');
        return;
    }
    
    // เพิ่ม timestamp เพื่อป้องกัน cache เมื่อ refresh
    const timestamp = isRefresh ? '&_=' + new Date().getTime() : '';
    const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(SHEET_NAME)}${timestamp}`;
    
    // แสดง Loading
    $('#dataTable tbody').html('<tr><td colspan="8" style="text-align:center; padding:20px;">⏳ กำลังโหลดข้อมูล...</td></tr>');
    
    $.ajax({
        url: url,
        dataType: 'text',
        timeout: 10000,
        cache: false, // ป้องกัน cache
        success: function(response) {
            console.log('AJAX Success! Response length:', response.length);
            console.log('First 100 chars:', response.substring(0, 100));
            
            try {
                // ลบ prefix และ suffix ของ Google Sheets API
                const jsonString = response.substring(47, response.length - 2);
                const data = JSON.parse(jsonString);
                
                console.log('Parsed data structure:', {
                    cols: data.table.cols.length,
                    rows: data.table.rows.length
                });
                
                // แสดงชื่อคอลัมน์
                const columnNames = data.table.cols.map(col => col.label || 'unnamed');
                console.log('Column names:', columnNames);
                
                // แปลงข้อมูลเป็น Array of Objects
                allData = data.table.rows.map(row => {
                    const obj = {};
                    row.c.forEach((cell, index) => {
                        const header = data.table.cols[index].label || `col${index}`;
                        obj[header] = cell ? (cell.v || '') : '';
                    });
                    return obj;
                });
                
                console.log('Data loaded successfully:', allData.length, 'rows');
                console.log('First row sample:', allData[0]);
                console.log('All column names in Google Sheets:', Object.keys(allData[0]));
                console.log('Sample data from each column:');
                Object.keys(allData[0]).forEach((key, index) => {
                    console.log(`  Column ${index}: "${key}" = "${allData[0][key]}"`);
                });
                
                // แสดงข้อมูลทั้งหมดทันทีโดยไม่ต้องรอเลือก filter
                filteredData = [...allData];
                populateDropdownOptions();
                displayData(filteredData);
                updateCount();
                
                // ซ่อน loading animation
                $('.btn.refresh i').removeClass('loading');
                
                
            } catch (error) {
                console.error('Error parsing data:', error);
                console.error('Error stack:', error.stack);
                alert('เกิดข้อผิดพลาดในการแปลงข้อมูล\n\n' + error.message);
                $('#dataTable tbody').html('<tr><td colspan="8" class="no-data">เกิดข้อผิดพลาดในการแปลงข้อมูล<br>' + error.message + '</td></tr>');
                $('.btn.refresh i').removeClass('loading');
            }
        },
        error: function(xhr, status, error) {
            console.error('AJAX Error!');
            console.error('Status:', status);
            console.error('Error:', error);
            console.error('XHR Status:', xhr.status);
            console.error('XHR Response:', xhr.responseText);
            
            let errorMsg = 'ไม่สามารถโหลดข้อมูลได้\n\n';
            errorMsg += 'กรุณาตรวจสอบ:\n';
            errorMsg += '1. Google Sheets ต้องแชร์เป็น "Anyone with the link"\n';
            errorMsg += '2. ชื่อ Sheet (แท็บ) ต้องถูกต้อง: "' + SHEET_NAME + '"\n';
            errorMsg += '3. SHEET_ID ต้องถูกต้อง: ' + SHEET_ID + '\n\n';
            errorMsg += 'รายละเอียดข้อผิดพลาด:\n';
            errorMsg += 'Status: ' + status + '\n';
            errorMsg += 'Error: ' + error;
            
            alert(errorMsg);
            $('.btn.refresh i').removeClass('loading');
        }
    });
}

function populateDropdownOptions() {
    const years = new Set();
    const depts = new Set();
    const ranks = new Set();
    
    // ดึงประเภทปีที่เลือก
    const yearType = $('#typeFilter').val().trim() || 'ปีการศึกษา';
    
    console.log('🔍 กำลังประมวลผล allData:', allData.length, 'แถว');
    
    allData.forEach((row, index) => {
        const columns = Object.values(row);
        
        // เลือก index ตามประเภทปี
        let yearIndex = 2; // default ปีการศึกษา
        
        if(yearType === "ปีการศึกษา") {
            yearIndex = 2;
        } else if(yearType === "ปีปฏิทิน") {
            yearIndex = 1;
        } else if(yearType === "ปีงบประมาณ") {
            yearIndex = 0;
        }
        
        // ดึงปี
        const yearValue = columns[yearIndex];
        if (yearValue != null && yearValue.toString().trim() !== '') {
            years.add(yearValue.toString().trim());
        }
        
        // ดึงภาควิชา (index 4)
        const deptValue = columns[4];
        if (deptValue != null && deptValue.toString().trim() !== '') {
            depts.add(deptValue.toString().trim());
        }
        
        // ดึงระดับผลงาน (index 5)
        const rankValue = columns[5];
        if (rankValue != null && rankValue.toString().trim() !== '') {
            ranks.add(rankValue.toString().trim());
        }
    });

    console.log('📊 ข้อมูลที่ดึงได้จาก Google Sheets:');
    console.log('  ปี:', Array.from(years).sort((a, b) => b - a));
    console.log('  ภาควิชา:', Array.from(depts).sort());
    console.log('  ระดับผลงาน:', Array.from(ranks).sort());

    // ============================================
    // เติมข้อมูล ปี (ไม่มี default - แสดงทั้งหมดจาก Sheets)
    // ============================================
    const currentYear = $('#yearFilter').val(); // เก็บค่าที่เลือกไว้
    const $yearFilter = $('#yearFilter');
    $yearFilter.find('option:not(:first)').remove();
    
    const yearArray = Array.from(years).sort((a, b) => b - a);
    yearArray.forEach(year => {
        const selected = year == currentYear ? 'selected' : '';
        $yearFilter.append(`<option value="${year}" ${selected}>${year}</option>`);
    });
    
    console.log('✅ ปีที่แสดงใน dropdown:', yearArray.length, 'รายการ');
    
    // ============================================
    // เติมข้อมูล สังกัดภาควิชา - เพิ่มเฉพาะที่ไม่ซ้ำกับ default
    // ============================================
    const currentDept = $('#deptFilter').val(); // เก็บค่าที่เลือกไว้
    const $deptFilter = $('#deptFilter');
    $deptFilter.find('option:not(:first)').remove();
    
    // สร้าง Map ของ default dept (normalize เป็น key, ข้อความจริงเป็น value)
    const defaultDeptsMap = new Map();
    defaultOptions.dept.filter(d => d !== "ทั้งหมด").forEach(dept => {
        defaultDeptsMap.set(normalizeText(dept), dept);
    });
    
    // เพิ่มเฉพาะข้อมูลที่ไม่ซ้ำกับ default (เปรียบเทียบแบบ normalize)
    const newDepts = Array.from(depts).filter(dept => {
        const normalizedDept = normalizeText(dept);
        return !defaultDeptsMap.has(normalizedDept);
    });
    
    if (newDepts.length > 0) {
        console.log('✨ ภาควิชาใหม่ที่เพิ่มเข้ามา:', newDepts);
    } else {
        console.log('ℹ️ ไม่มีภาควิชาใหม่ (ทุกภาควิชามีใน default อยู่แล้ว)');
    }
    
    // รวม default + ใหม่ แล้วเรียงตัวอักษร
    const allDepts = [
        ...defaultOptions.dept.filter(d => d !== "ทั้งหมด"),
        ...newDepts
    ];
    
    // เรียงตามตัวอักษรภาษาไทย
    allDepts.sort((a, b) => a.localeCompare(b, 'th'));
    
    // เพิ่มทั้งหมดเข้า dropdown
    allDepts.forEach(dept => {
        const selected = dept == currentDept ? 'selected' : '';
        $deptFilter.append(`<option value="${dept}" ${selected}>${dept}</option>`);
    });
    
    // ============================================
    // เติมข้อมูล ระดับผลงาน - เพิ่มเฉพาะที่ไม่ซ้ำกับ default
    // ============================================
    const currentRank = $('#rankFilter').val(); // เก็บค่าที่เลือกไว้
    const $rankFilter = $('#rankFilter');
    $rankFilter.find('option:not(:first)').remove();
    
    // สร้าง Map ของ default rank (normalize เป็น key, ข้อความจริงเป็น value)
    const defaultRanksMap = new Map();
    defaultOptions.rank.filter(r => r !== "ทั้งหมด").forEach(rank => {
        defaultRanksMap.set(normalizeText(rank), rank);
    });
    
    // เพิ่มเฉพาะข้อมูลที่ไม่ซ้ำกับ default (เปรียบเทียบแบบ normalize)
    const newRanks = Array.from(ranks).filter(rank => {
        const normalizedRank = normalizeText(rank);
        return !defaultRanksMap.has(normalizedRank);
    });
    
    if (newRanks.length > 0) {
        console.log('✨ ระดับผลงานใหม่ที่เพิ่มเข้ามา:', newRanks);
    } else {
        console.log('ℹ️ ไม่มีระดับผลงานใหม่ (ทุกระดับมีใน default อยู่แล้ว)');
    }
    
    // รวม default + ใหม่ แล้วเรียงตัวอักษร
    const allRanks = [
        ...defaultOptions.rank.filter(r => r !== "ทั้งหมด"),
        ...newRanks
    ];
    
    // เรียงตามตัวอักษรภาษาไทย
    allRanks.sort((a, b) => a.localeCompare(b, 'th'));
    
    // เพิ่มทั้งหมดเข้า dropdown
    allRanks.forEach(rank => {
        const selected = rank == currentRank ? 'selected' : '';
        $rankFilter.append(`<option value="${rank}" ${selected}>${rank}</option>`);
    });
    
    console.log('✅ ระดับผลงานที่แสดงใน dropdown:', allRanks.length, 'รายการ (เรียงตามตัวอักษร)');
    
    // ============================================
    // สรุปผลลัพธ์
    // ============================================
    console.log('📋 สรุปข้อมูลทั้งหมดใน Dropdown:');
    console.log('  - ปี:', years.size, 'รายการ (ทั้งหมดจาก Google Sheets)');
    console.log('  - ภาควิชา:', defaultDeptsMap.size, 'default +', newDepts.length, 'ใหม่ =', allDepts.length, 'รายการ (เรียงตามตัวอักษร)');
    console.log('  - ระดับผลงาน:', defaultRanksMap.size, 'default +', newRanks.length, 'ใหม่ =', allRanks.length, 'รายการ (เรียงตามตัวอักษร)');
    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
}

// ฟังก์ชันกรองข้อมูล
function filterData() {
    const yearType = $('#typeFilter').val().trim() || 'ปีการศึกษา';
    const year = $('#yearFilter').val().trim();
    const dept = $('#deptFilter').val().trim();
    const rank = $('#rankFilter').val().trim();
    const author = $('#authorFilter').val().toLowerCase().trim();
    const searchName = $('#keyword').val().toLowerCase().trim();

    console.log('🔍 Filtering with:', {yearType, year, dept, rank, author, searchName});

    // ถ้าไม่มีการกรองเลย (เว้นประเภทปี) แสดงข้อมูลทั้งหมด
    if (!year && !dept && !rank && !author && !searchName) {
        filteredData = [...allData];
        displayData(filteredData);
        updateCount();
        return;
    }

    filteredData = allData.filter(row => {
        const columns = Object.values(row);
        
        // ปี - เลือก index ตามประเภทปี
        if (year) {
            let yearIndex = 2; // default ปีการศึกษา
            
            if(yearType === "ปีการศึกษา") {
                yearIndex = 2;
            } else if(yearType === "ปีปฏิทิน") {
                yearIndex = 1;
            } else if(yearType === "ปีงบประมาณ") {
                yearIndex = 0;
            }
            
            const yearValue = columns[yearIndex] || '';
            if (!yearValue.toString().includes(year)) return false;
        }
        
        // เจ้าของผลงาน - อยู่ใน index 3
        if (author) {
            const authorValue = (columns[3] || '').toString().toLowerCase();
            if (!authorValue.includes(author)) return false;
        }
        
        // สังกัดภาควิชา - อยู่ใน index 4
        if (dept) {
            const deptValue = (columns[4] || '').toString();
            if (!deptValue.includes(dept)) return false;
        }
        
        // ชื่อผลงาน (ค้นหาทั้งภาษาไทยและภาษาอังกฤษ)
        // ภาษาไทย - index 7, ภาษาอังกฤษ - index 8
        if (searchName) {
            const thaiName = (columns[7] || '').toString().toLowerCase();
            const engName = (columns[8] || '').toString().toLowerCase();
            if (!thaiName.includes(searchName) && !engName.includes(searchName)) return false;
        }
        
        // ระดับผลงาน - อยู่ใน index 5
        if (rank) {
            const rankValue = (columns[5] || '').toString();
            if (!rankValue.includes(rank)) return false;
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
        $tbody.html('<tr><td colspan="8" class="no-data">ไม่พบข้อมูล</td></tr>');
        return;
    }

    console.log('📊 Displaying', data.length, 'rows');

    // ดึงค่าประเภทปีที่เลือก
    const yearType = $('#typeFilter').val().trim() || 'ปีการศึกษา';

    data.forEach((row, index) => {
        // Debug: แสดงข้อมูล 3 แถวแรก
        if (index < 3) {
            console.log(`Row ${index} data:`, row);
            console.log('All values:', Object.values(row));
        }
        
        // ฟังก์ชันตรวจสอบและแสดงค่า - ถ้าว่างให้ใส่ "-"
        const getValue = (val) => {
            if (val === null || val === undefined || val === '') return '-';
            const strVal = val.toString().trim();
            return strVal === '' ? '-' : strVal;
        };

        const columns = Object.values(row);

        // ดึงค่าปีตามประเภทที่เลือก
        let yearValue = '-';
        let yearIndex = 2; // default ปีการศึกษา

        if(yearType === "ปีการศึกษา") {
            yearIndex = 2;
        } else if(yearType === "ปีปฏิทิน") {
            yearIndex = 1;
        } else if(yearType === "ปีงบประมาณ") {
            yearIndex = 0;
        }

        yearValue = getValue(columns[yearIndex]);
        
        // ชื่อผู้ผลิตผลงาน
        let author = getValue(columns[3]);
        // แทนที่ , ด้วย ,<br> เพื่อขึ้นบรรทัดใหม่
        if (author !== '-') {
            author = author.replace(/,\s*/g, ',<br>');
        }
        
        // สังกัดภาควิชา
        const department = getValue(columns[4]);
        
        // ชื่อผลงานภาษาไทย
        const titleThai = getValue(columns[7]);
        
        // ชื่อผลงานภาษาอังกฤษ
        const titleEng = getValue(columns[8]);
        
        // ระดับผลงาน
        const level = getValue(columns[5]);
        
        // อ้างอิง
        const reference = getValue(columns[10]);

        const rowHTML = `
            <tr>
                <td>${yearType}</td>
                <td>${yearValue}</td>
                <td class="${author === '-' ? 'center-dash' : ''}">${author}</td>
                <td class="${department === '-' ? 'center-dash' : ''}">${department}</td>
                <td>${titleThai}</td>
                <td>${titleEng}</td>
                <td class="${level === '-' ? 'center-dash' : ''}">${level}</td>
                <td class="${reference === '-' ? 'center-dash' : ''}">${reference}</td>
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

// ฟังก์ชัน Export ไปยัง Excel - Export เฉพาะข้อมูลที่แสดงในตาราง (filteredData)
function exportTableToExcel() {
    if (filteredData.length === 0) {
        showNotification('ไม่มีข้อมูลสำหรับ Export', 'warning');
        return;
    }

    console.log('📤 Exporting', filteredData.length, 'rows to Excel');

    // สร้าง workbook
    const wb = XLSX.utils.book_new();
    
    // ดึงค่าประเภทปีที่เลือก
    const yearType = $('#typeFilter').val().trim() || 'ปีการศึกษา';
    
    // สร้างข้อมูลสำหรับ export
    const exportData = filteredData.map(row => {
        const columns = Object.values(row);
        
        // ดึงค่าปีตามประเภทที่เลือก
        let yearIndex = 2; // default ปีการศึกษา
        if(yearType === "ปีปฏิทิน") {
            yearIndex = 1;
        } else if(yearType === "ปีงบประมาณ") {
            yearIndex = 0;
        }
        
        return {
            'ประเภทของปี': yearType,
            'ปี': columns[yearIndex] || '-',
            'ชื่อผู้ผลิตผลงาน': columns[3] || '-',
            'สังกัดภาควิชา': columns[4] || '-',
            'ชื่อผลงาน (ไทย)': columns[7] || '-',
            'ชื่อผลงาน (Eng)': columns[8] || '-',
            'ระดับผลงาน': columns[5] || '-',
            'อ้างอิง': columns[10] || '-'
        };
    });
    
    // สร้าง worksheet จาก JSON
    const ws = XLSX.utils.json_to_sheet(exportData);
    
    // ตั้งค่าความกว้างคอลัมน์
    ws['!cols'] = [
        { wch: 15 }, // ประเภทของปี
        { wch: 10 }, // ปี
        { wch: 30 }, // ชื่อผู้ผลิตผลงาน
        { wch: 25 }, // สังกัดภาควิชา
        { wch: 50 }, // ชื่อผลงาน (ไทย)
        { wch: 50 }, // ชื่อผลงาน (Eng)
        { wch: 35 }, // ระดับผลงาน
        { wch: 50 }  // อ้างอิง
    ];
    
    XLSX.utils.book_append_sheet(wb, ws, 'ผลงานตีพิมพ์');
    
    // สร้างชื่อไฟล์พร้อมวันที่
    const now = new Date();
    const dateStr = now.toLocaleDateString('th-TH', { 
        year: 'numeric', 
        month: '2-digit', 
        day: '2-digit' 
    }).replace(/\//g, '-');
    const fileName = `ผลงานตีพิมพ์_${dateStr}.xlsx`;
    
    // Download ไฟล์
    XLSX.writeFile(wb, fileName);
    
    console.log('✅ Excel file exported:', fileName);
    showNotification(`Export สำเร็จ: ${filteredData.length} แถว`, 'success');
}