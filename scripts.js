// ข้อมูลทั้งหมด
let allData = [];
let filteredData = [];

// Google Sheets URL
const SHEET_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vT2KK5zhWgJf9t21MO2qxykVWctw0sfJ8M7SwERKm0Nb7hxamhqfLEsa5vhqz5XHg/pub?gid=1841845518&single=true&output=csv';

// ข้อมูลทดสอบ (ลบออกเมื่อ Google Sheets ใช้งานได้แล้ว)
const TEST_DATA = [
    {
        'ปี': '2567',
        'เจ้าของผลงาน': 'สถิตาภรณ์ ศรีหิรัญ',
        'สังกัดภาควิชา': 'ภาษาอังกฤษ',
        'ชื่อผลงาน': 'การศึกษาภาษาอังกฤษเพื่อการสื่อสาร',
        'อ้างอิง': 'วารสารมนุษยศาสตร์และสังคมศาสตร์ ปีที่ 15 ฉบับที่ 1'
    },
    {
        'ปี': '2566',
        'เจ้าของผลงาน': 'ดร.สมชาย ใจดี',
        'สังกัดภาควิชา': 'ภาษาไทย',
        'ชื่อผลงาน': 'วรรณกรรมไทยร่วมสมัย',
        'อ้างอิง': 'วารสาร TCI กลุ่ม 1'
    },
    {
        'ปี': '2566',
        'เจ้าของผลงาน': 'ผศ.ดร.วิภา จันทร์งาม',
        'สังกัดภาควิชา': 'ภาษาตะวันออก',
        'ชื่อผลงาน': 'Chinese Language Teaching Methods',
        'อ้างอิง': 'International Journal Q2'
    }
];

// เริ่มต้นเมื่อโหลดหน้าเสร็จ
$(document).ready(function() {
    // โหลดข้อมูลจาก Google Sheets
    loadDataFromGoogleSheets();

    // เพิ่ม Event Listeners
    $('.btn.search').click(filterData);
    $('.btn.excel').click(exportToExcel);
    $('.btn.reset').click(resetForm);
    
    // ค้นหาแบบ Real-time เมื่อพิมพ์
    $('#keyword, #authorFilter').on('input', filterData);
    $('#typeFilter, #yearFilter, #deptFilter, #rankFilter').on('change', filterData);
});

// ฟังก์ชันโหลดข้อมูลจาก Google Sheets
function loadDataFromGoogleSheets() {
    $('#dataTable tbody').html('<tr><td colspan="5" class="loading">กำลังโหลดข้อมูล...</td></tr>');

    // ลองใช้ข้อมูลทดสอบก่อน
    console.log('Loading data...');
    
    fetch(SHEET_URL)
        .then(response => {
            if (!response.ok) {
                throw new Error('ไม่สามารถเข้าถึง Google Sheets');
            }
            return response.text();
        })
        .then(csvData => {
            console.log('CSV Data loaded:', csvData.substring(0, 200));
            allData = parseCSV(csvData);
            
            // ถ้าไม่มีข้อมูลจาก Google Sheets ให้ใช้ TEST_DATA
            if (allData.length === 0) {
                console.log('No data from Google Sheets, using test data');
                allData = TEST_DATA;
            }
            
            filteredData = [...allData];
            populateAuthorsDatalist();
            displayData(filteredData);
            updateCount();
        })
        .catch(error => {
            console.error('Error loading from Google Sheets:', error);
            console.log('Using test data instead');
            
            // ใช้ข้อมูลทดสอบแทน
            allData = TEST_DATA;
            filteredData = [...allData];
            populateAuthorsDatalist();
            displayData(filteredData);
            updateCount();
        });
}

// ฟังก์ชันแปลง CSV เป็น Array
function parseCSV(csv) {
    const lines = csv.split('\n');
    const headers = lines[0].split(',').map(h => h.trim());
    const data = [];

    for (let i = 1; i < lines.length; i++) {
        if (lines[i].trim() === '') continue;
        
        const values = lines[i].split(',');
        const row = {};
        
        headers.forEach((header, index) => {
            row[header] = values[index] ? values[index].trim() : '';
        });
        
        data.push(row);
    }

    return data;
}

// ฟังก์ชันเติมข้อมูลใน Datalist ผู้สร้างสรรค์ (AutoComplete)
function populateAuthorsDatalist() {
    const authors = new Set();
    
    allData.forEach(row => {
        // คอลัมน์เจ้าของผลงาน (index 1 = คอลัมน์ที่ 2)
        const authorColumn = Object.keys(row)[1];
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
}

// ฟังก์ชันกรองข้อมูล
function filterData() {
    const yearType = $('#typeFilter').val();
    const year = $('#yearFilter').val();
    const dept = $('#deptFilter').val();
    const rank = $('#rankFilter').val();
    const author = $('#authorFilter').val().toLowerCase().trim();
    const searchName = $('#keyword').val().toLowerCase().trim();

    filteredData = allData.filter(row => {
        const rowValues = Object.values(row).join(' ').toLowerCase();
        
        if (year && !rowValues.includes(year)) return false;
        if (dept && !rowValues.includes(dept.toLowerCase())) return false;
        if (rank && !rowValues.includes(rank.toLowerCase())) return false;
        if (author && !rowValues.includes(author)) return false;
        if (searchName && !rowValues.includes(searchName)) return false;
        
        return true;
    });

    displayData(filteredData);
    updateCount();
}

// ฟังก์ชันแสดงข้อมูลในตาราง
function displayData(data) {
    const $tbody = $('#dataTable tbody');
    $tbody.empty();

    if (data.length === 0) {
        $tbody.html('<tr><td colspan="5" class="no-data">ไม่พบข้อมูล</td></tr>');
        return;
    }

    data.forEach((row) => {
        const columns = Object.values(row);
        const rowHTML = `
            <tr>
                <td>${columns[0] || '-'}</td>
                <td>${columns[1] || '-'}</td>
                <td>${columns[2] || '-'}</td>
                <td>${columns[3] || '-'}</td>
                <td>${columns[4] || '-'}</td>
            </tr>
        `;
        $tbody.append(rowHTML);
    });
}

// ฟังก์ชันอัพเดทจำนวนผลลัพธ์
function updateCount() {
    $('#count').text(filteredData.length);
}

// ฟังก์ชัน Export ไปยัง Excel
function exportToExcel() {
    if (filteredData.length === 0) {
        alert('ไม่มีข้อมูลสำหรับ Export');
        return;
    }

    const wb = XLSX.utils.book_new();
    const headers = Object.keys(filteredData[0]);
    const data = filteredData.map(row => Object.values(row));
    const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
    
    XLSX.utils.book_append_sheet(wb, ws, 'ผลงานตีพิมพ์');
    
    const fileName = `ผลงานตีพิมพ์_${new Date().toLocaleDateString('th-TH')}.xlsx`;
    XLSX.writeFile(wb, fileName);
}

// ฟังก์ชันรีเซ็ตฟอร์ม
function resetForm() {
    $('#typeFilter').val('');
    $('#yearFilter').val('');
    $('#deptFilter').val('');
    $('#rankFilter').val('');
    $('#authorFilter').val('');
    $('#keyword').val('');
    
    filteredData = [...allData];
    displayData(filteredData);
    updateCount();
}