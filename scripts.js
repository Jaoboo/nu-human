// ข้อมูลทั้งหมด
let allData = [];
let filteredData = [];

// ฟังก์ชัน normalize ข้อความ - ลบช่องว่างทั้งหมดและแปลงเป็นตัวพิมพ์เล็ก
function normalizeText(text) {
    if (!text) return '';
    return text.toString().replace(/\s+/g, '').toLowerCase();
}

function getYearValue(val) {
    if (val === null || val === undefined || val === '') return '-';
    
    // แสดงค่าตามที่มีเลย
    return val.toString().trim();
}

// ฟังก์ชันแปลงวันที่เป็น mm-yyyy (พ.ศ.) - เฉพาะช่อง Published เท่านั้น
function formatDateToMonthYear(dateValue) {
    if (!dateValue || dateValue === '-' || dateValue === '') return '-';
    
    try {
        let date;
        const str = dateValue.toString().trim();
        
        // ตรวจสอบรูปแบบ Date(yyyy,m,d) จาก Google Sheets
        const dateMatch = str.match(/Date\((\d{4}),(\d{1,2}),(\d{1,2})\)/);
        if (dateMatch) {
            const year = parseInt(dateMatch[1]) + 543; // แปลงเป็น พ.ศ.
            let monthNum = parseInt(dateMatch[2]) + 1; // บวก 1 เพราะ Google Sheets ส่งมาเป็น 0-indexed
            // ถ้าบวกแล้วได้ 13 ให้เป็น 1
            if (monthNum === 13) monthNum = 1;
            const month = monthNum.toString().padStart(2, '0');
            return `${month}-${year}`;
        }
        
        // ลองหารูปแบบ mm/yyyy, yyyy-mm, mm-yyyy
        const patterns = [
            /(\d{1,2})\/(\d{4})/,           // mm/yyyy หรือ m/yyyy
            /(\d{4})-(\d{1,2})/,            // yyyy-mm
            /(\d{1,2})-(\d{4})/,            // mm-yyyy
        ];
        
        for (const pattern of patterns) {
            const match = str.match(pattern);
            if (match) {
                let month, year;
                if (pattern.source.startsWith('(\\d{4})')) {
                    // yyyy-mm format
                    year = parseInt(match[1]) + 543; // แปลงเป็น พ.ศ.
                    month = match[2].padStart(2, '0');
                } else {
                    // mm/yyyy or mm-yyyy format
                    month = match[1].padStart(2, '0');
                    year = parseInt(match[2]) + 543; // แปลงเป็น พ.ศ.
                }
                return `${month}-${year}`;
            }
        }
        
        // ถ้าเป็นตัวเลข (Excel date serial number)
        if (typeof dateValue === 'number') {
            // แปลง Excel serial number เป็น Date และบวก offset timezone ไทย (UTC+7)
            const utcDate = new Date((dateValue - 25569) * 86400 * 1000);
            date = new Date(utcDate.getTime() + (7 * 60 * 60 * 1000)); // บวก 7 ชั่วโมง
            
            // ตรวจสอบว่า date ถูกต้องหรือไม่
            if (isNaN(date.getTime())) {
                return dateValue.toString();
            }
            
            // แปลงเป็น mm-yyyy (พ.ศ.)
            const month = String(date.getUTCMonth() + 1).padStart(2, '0');
            const year = date.getUTCFullYear() + 543;
            return `${month}-${year}`;
        }
        
        // ถ้าไม่ match pattern ใดเลย ให้คืนค่าเดิม (กรณีเป็นปีอย่างเดียวหรือข้อความอื่นๆ)
        return str;
        
    } catch (error) {
        console.error('Error formatting date:', dateValue, error);
        return dateValue.toString();
    }
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
        "วารสารระดับนานาชาติ Q 1",
        "วารสารระดับนานาชาติ Q 2",
        "วารสารระดับนานาชาติ Q 3",
        "วารสารระดับนานาชาติ Q 4",
    ]
};

// ข้อมูลจริง - ใส่ SHEET_ID ของ Google Sheets ตรงนี้
const SHEET_ID = '1CJhSu3XwPTC35SxXzGnOwz8a8c_D_1eI';
const SHEET_NAME = 'web app(อยู่ระหว่างจัดทำ)';

// ฟังก์ชันดึงค่าจาก URL Parameters
function getURLParameters() {
    const params = new URLSearchParams(window.location.search);
    return {
        type: params.get('type') || 'ปีปฏิทิน',
        year: params.get('year') || '',
        dept: params.get('dept') || '',
        rank: params.get('rank') || '',
        author: params.get('author') || '',
        keyword: params.get('keyword') || ''
    };
}

// ฟังก์ชันอัพเดท URL Parameters
function updateURLParameters() {
    const params = new URLSearchParams();
    
    const type = $('#typeFilter').val();
    const year = $('#yearFilter').val();
    const dept = $('#deptFilter').val();
    const rank = $('#rankFilter').val();
    const author = $('#authorFilter').val();
    const keyword = $('#keyword').val();
    
    // เพิ่มเฉพาะค่าที่ไม่ว่าง
    if (type && type !== 'ปีปฏิทิน') params.set('type', type);
    if (year) params.set('year', year);
    if (dept) params.set('dept', dept);
    if (rank) params.set('rank', rank);
    if (author) params.set('author', author);
    if (keyword) params.set('keyword', keyword);
    
    // อัพเดท URL โดยไม่ reload หน้า
    const newURL = params.toString() ? 
        `${window.location.pathname}?${params.toString()}` : 
        window.location.pathname;
    
    window.history.pushState({}, '', newURL);
}

// ฟังก์ชันโหลดค่า filter จาก URL
function loadFiltersFromURL() {
    const params = getURLParameters();
    
    $('#typeFilter').val(params.type);
    
    // รอให้ dropdown โหลดเสร็จก่อนตั้งค่า
    setTimeout(() => {
        if (params.year) $('#yearFilter').val(params.year);
        if (params.dept) $('#deptFilter').val(params.dept);
        if (params.rank) $('#rankFilter').val(params.rank);
        if (params.author) $('#authorFilter').val(params.author);
        if (params.keyword) $('#keyword').val(params.keyword);
        
        // กรองข้อมูลตาม URL parameters
        if (params.year || params.dept || params.rank || params.author || params.keyword) {
            filterData();
        }
    }, 500);
}

// เริ่มต้นเมื่อโหลดหน้าเสร็จ
$(document).ready(function() {
    
    // โหลดข้อมูลทันที
    loadData();

    // ปุ่ม Reset
    $('.btn.reset').click(function() {
        // ล้างค่า filters
        $('#typeFilter').val('ปีปฏิทิน');
        $('#yearFilter').val('');
        $('#deptFilter').val('');
        $('#rankFilter').val('');
        $('#authorFilter').val('');
        $('#keyword').val('');
        
        // ล้าง URL parameters
        window.history.pushState({}, '', window.location.pathname);
        
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
        // อัพเดท URL
        updateURLParameters();
    });
    
    $('#yearFilter, #deptFilter, #rankFilter').on('change', function() {
        console.log('Dropdown changed:', $(this).attr('id'), '=', $(this).val());
        filterData();
        updateURLParameters();
    });
    
    // Input text: ค้นหาแบบ Real-time เมื่อพิมพ์
    $('#keyword, #authorFilter').on('input', function() {
        console.log('Input changed:', $(this).attr('id'), '=', $(this).val());
        filterData();
        updateURLParameters();
    });
    
    // ปุ่มค้นหา - กรองข้อมูลอีกครั้ง
    $('.btn.search').click(function() {
        filterData();
        updateURLParameters();
    });
});

// ฟังก์ชันโหลดข้อมูลจาก Google Sheets
function loadData(isRefresh = false) {
    console.log('Loading data from Google Sheets...');
    console.log('SHEET_ID:', SHEET_ID);
    console.log('SHEET_NAME:', SHEET_NAME);
    
    if (!SHEET_ID || SHEET_ID === '') {
        alert('กรุณาใส่ SHEET_ID ในไฟล์ script.js');
        $('#dataTable tbody').html('<tr><td colspan="9" class="no-data">กรุณาใส่ SHEET_ID</td></tr>');
        return;
    }
    
    // เพิ่ม timestamp เพื่อป้องกัน cache เมื่อ refresh
    const timestamp = isRefresh ? '&_=' + new Date().getTime() : '';
    const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(SHEET_NAME)}${timestamp}`;
    
    // แสดง Loading
    $('#dataTable tbody').html('<tr><td colspan="9" style="text-align:center; padding:20px;">⏳ กำลังโหลดข้อมูล...</td></tr>');
    
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
                
                // โหลดค่า filter จาก URL parameters
                loadFiltersFromURL();
                
                // ซ่อน loading animation
                $('.btn.refresh i').removeClass('loading');
                
                
            } catch (error) {
                console.error('Error parsing data:', error);
                console.error('Error stack:', error.stack);
                alert('Error parsing data\n\n' + error.message);
                $('#dataTable tbody').html('<tr><td colspan="9" class="no-data">Error parsing data<br>' + error.message + '</td></tr>');
                $('.btn.refresh i').removeClass('loading');
            }
        },
        error: function(xhr, status, error) {
            console.error('AJAX Error!');
            console.error('Status:', status);
            console.error('Error:', error);
            console.error('XHR Status:', xhr.status);
            console.error('XHR Response:', xhr.responseText);
            
            let errorMsg = 'Unable to load data\n\n';
            errorMsg += 'Please check:\n';
            errorMsg += '1. Your internet connection\n';
            errorMsg += '2. Unable to display data from spreadsheet';
            
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
    const yearType = $('#typeFilter').val().trim() || 'ปีปฏิทิน';
    
    allData.forEach((row, index) => {
        const columns = Object.values(row);
        
        // ดึงปีจากทั้ง 3 คอลัมน์ (index 0, 1, 2)
        // ปีงบประมาณ - index 0
        const budgetYear = columns[0];
        if (budgetYear != null && budgetYear.toString().trim() !== '') {
            years.add(budgetYear.toString().trim());
        }
        
        // ปีปฏิทิน - index 1
        const calendarYear = columns[1];
        if (calendarYear != null && calendarYear.toString().trim() !== '') {
            years.add(calendarYear.toString().trim());
        }
        
        // ปีการศึกษา - index 2
        const academicYear = columns[2];
        if (academicYear != null && academicYear.toString().trim() !== '') {
            years.add(academicYear.toString().trim());
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

    // ============================================
    // เติมข้อมูล ปี (จากทั้ง 3 คอลัมน์)
    // ============================================
    const currentYear = $('#yearFilter').val(); // เก็บค่าที่เลือกไว้
    const $yearFilter = $('#yearFilter');
    $yearFilter.find('option:not(:first)').remove();
    
    const yearArray = Array.from(years).sort((a, b) => b - a);
    yearArray.forEach(year => {
        const selected = year == currentYear ? 'selected' : '';
        $yearFilter.append(`<option value="${year}" ${selected}>${year}</option>`);
    });
    
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
}

// ฟังก์ชันกรองข้อมูล
function filterData() {
    const yearType = $('#typeFilter').val().trim() || 'ปีปฏิทิน';
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
            let yearIndex = 1; // default ปีปฏิทิน
            
            if(yearType === "ปีปฏิทิน") {
                yearIndex = 1;
            } else if(yearType === "ปีการศึกษา") {
                yearIndex = 2;
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

    displayData(filteredData);
    updateCount();
}

// เพิ่ม debug ในฟังก์ชัน displayData
function displayData(data) {
    const $tbody = $('#dataTable tbody');
    $tbody.empty();

    if (data.length === 0) {
        $tbody.html('<tr><td colspan="9" class="no-data">ไม่พบข้อมูล</td></tr>');
        return;
    }

    const yearType = $('#typeFilter').val().trim() || 'ปีปฏิทิน';

    data.forEach((row, index) => {
        const columns = Object.values(row);
        
        // Debug: แสดงข้อมูลทุกคอลัมน์ของแถวแรก
        if (index === 0) {
            console.log('=== DEBUG ROW 0 ===');
            console.log('Total columns:', columns.length);
            columns.forEach((col, i) => {
                console.log(`Column ${i}:`, col, `(Type: ${typeof col})`);
            });
        }
        
        const getValue = (val) => {
            if (val === null || val === undefined || val === '') return '-';
            const strVal = val.toString().trim();
            return strVal === '' ? '-' : strVal;
        };

        let yearValue = '-';
        let yearIndex = 1; // default ปีปฏิทิน

        if(yearType === "ปีปฏิทิน") {
            yearIndex = 1;
        } else if(yearType === "ปีการศึกษา") {
            yearIndex = 2;
        } else if(yearType === "ปีงบประมาณ") {
            yearIndex = 0;
        }

        // Debug: ดูค่าที่ดึงมา
        if (index === 0) {
            console.log(`Year Type: ${yearType}`);
            console.log(`Year Index: ${yearIndex}`);
            console.log(`Raw Year Value:`, columns[yearIndex]);
            console.log(`After getYearValue:`, getYearValue(columns[yearIndex]));
        }

        yearValue = getYearValue(columns[yearIndex]);
        
        let author = getValue(columns[3]);
        if (author !== '-') {
            author = author.replace(/,\s*/g, ',<br>');
        }
        
        const department = getValue(columns[4]);
        const titleThai = getValue(columns[7]);
        const titleEng = getValue(columns[8]);
        const level = getValue(columns[5]);
        const reference = getValue(columns[10]);
        const publishedRaw = getValue(columns[14]);
        const publisher = getValue(columns[9]);
        const published = formatDateToMonthYear(publishedRaw);
        

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
                <td class="${publisher === '-' ? 'center-dash' : ''}">${publisher}</td>
                <td class="${published === '-' ? 'center-dash' : ''}">${published}</td>
                
            </tr>
        `;
        $tbody.append(rowHTML);
    });
}

// ฟังก์ชันอัพเดทจำนวนผลลัพธ์
function updateCount() {
    $('#count').text(filteredData.length);
}

// ฟังก์ชัน Export ไปยัง Excel - Export เฉพาะข้อมูลที่แสดงในตาราง (filteredData)
function exportTableToExcel() {
    if (filteredData.length === 0) {
        alert('ไม่มีข้อมูลสำหรับ Export');
        return;
    }

    // สร้าง workbook
    const wb = XLSX.utils.book_new();
    
    // ดึงค่าประเภทปีที่เลือก
    const yearType = $('#typeFilter').val().trim() || 'ปีปฏิทิน';
    
    // สร้างข้อมูลสำหรับ export
    const exportData = filteredData.map(row => {
        const columns = Object.values(row);
        
        // ดึงค่าปีตามประเภทที่เลือก - ใช้ getYearValue
        let yearIndex = 1; // default ปีปฏิทิน
        if(yearType === "ปีการศึกษา") {
            yearIndex = 2;
        } else if(yearType === "ปีงบประมาณ") {
            yearIndex = 0;
        }
        
        const yearValue = getYearValue(columns[yearIndex]); // ใช้ getYearValue แทน
        
        // Published - แปลงเป็น mm-yyyy
        const publishedRaw = columns[14] || '-';
        const published = formatDateToMonthYear(publishedRaw);
        
        return {
            'ประเภทปี': yearType,
            'ปี': yearValue,
            'ชื่อผู้ผลิตผลงาน': columns[3] || '-',
            'สังกัดภาควิชา': columns[4] || '-',
            'ชื่อผลงาน (ไทย)': columns[7] || '-',
            'ชื่อผลงาน (Eng)': columns[8] || '-',
            'ระดับผลงาน': columns[5] || '-',
            'อ้างอิง': columns[10] || '-',
            'Publisher': columns[9] || '-',
            'Published': published
        };
    });
    
    // สร้าง worksheet จาก JSON
    const ws = XLSX.utils.json_to_sheet(exportData);
    
    // ตั้งค่าความกว้างคอลัมน์
    ws['!cols'] = [
        { wch: 10 }, // ประเภทของปี
        { wch: 10 }, // ปี
        { wch: 30 }, // ชื่อผู้ผลิตผลงาน
        { wch: 25 }, // สังกัดภาควิชา
        { wch: 50 }, // ชื่อผลงาน (ไทย)
        { wch: 50 }, // ชื่อผลงาน (Eng)
        { wch: 15 }, // ระดับผลงาน
        { wch: 50 }, // อ้างอิง
        { wch: 30 }, // Publisher
        { wch: 10 }  // Published (mm-yyyy)
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
}