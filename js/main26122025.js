// js/main.js - แก้ไขให้รองรับทั้ง 3 ฟอร์ม (Dashboard ทำงานแยกอิสระ) + Mobile Test Page + Reports Page
// ====================================================
// ⭐ GLOBAL CONFIGURATION
// ====================================================

const GAS_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbz6pqVa04Xt0k_bDTy2jHbhShT4IUnZoJXTkOV6MNwWhNTCQMsmjp7y72c-sBfUFw4J/exec'; 

// ⭐ ตัวแปร global สำหรับทั้งระบบ
window.API_URL = GAS_WEB_APP_URL;
window.GAS_WEB_APP_URL = GAS_WEB_APP_URL;
// ✅ เพิ่มตัวแปรนี้สำหรับ dashboard.js (ดึงค่า API URL)
window.GAS_API_URL = GAS_WEB_APP_URL;  

// ⭐ เก็บข้อมูลรายชื่อ (cache) - ป้องกันการเรียก API ซ้ำ
window.personnelData = {
    personnelList: [],
    teacherList: [],
    lastUpdated: null,
    isLoaded: false,
    retryCount: 0,
    maxRetries: 2,
    loadingAttempts: 0,
    hasRealData: false
};

// ⭐ ตัวแปรสถานะระบบ
let isMobileMenuOpen = false;
let currentPage = 'homepage.html';

// ====================================================
// 🚀 API SERVICE FUNCTIONS - จัดการ API จากที่เดียว
// ====================================================

/**
 * ⭐ SERVICE: เรียก API Google Apps Script (ใช้ XMLHttpRequest แก้ปัญหา CORS)
 */
async function callGASAPI(action, payload = {}) {
    console.log(`📡 API CALL: ${action}`, { action, payloadSize: JSON.stringify(payload).length });
    
    // ✅ ใช้ค่า API URL จากตัวแปร global (ให้ dashboard.js เข้าถึงได้)
    const url = window.GAS_API_URL || window.API_URL || window.GAS_WEB_APP_URL || GAS_WEB_APP_URL;
    
    if (!url) {
        console.error('❌ API URL is not defined');
        return {
            status: 'ERROR',
            error: 'API URL ไม่ได้ถูกกำหนดค่า',
            suggestion: 'กรุณาตรวจสอบการตั้งค่า'
        };
    }
    
    return new Promise((resolve) => {
        try {
            // ใช้ XMLHttpRequest แทน fetch (แก้ปัญหา CORS)
            const xhr = new XMLHttpRequest();
            
            // สร้าง request data
            const requestData = {
                action: action,
                ...payload,
                timestamp: new Date().toISOString()
            };
            
            const params = new URLSearchParams();
            params.append('jsonPayload', JSON.stringify(requestData));
            
            xhr.open('POST', url, true);
            xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded;charset=UTF-8');
            
            // ✅ ตั้งค่า timeout 30 วินาที (นานขึ้น)
            xhr.timeout = 30000;
            
            xhr.onload = function() {
                console.log(`📥 Response status: ${xhr.status}`);
                
                if (xhr.status >= 200 && xhr.status < 300) {
                    try {
                        const result = JSON.parse(xhr.responseText);
                        console.log(`✅ API ${action} success`);
                        resolve({
                            status: 'SUCCESS',
                            ...result
                        });
                    } catch (parseError) {
                        console.error('❌ Parse error:', parseError);
                        
                        // ถ้า response ไม่ใช่ JSON แต่เป็นข้อความธรรมดา
                        if (xhr.responseText.includes('SUCCESS') || xhr.responseText.includes('บันทึก')) {
                            resolve({
                                status: 'SUCCESS',
                                message: 'บันทึกข้อมูลสำเร็จ',
                                rawResponse: xhr.responseText.substring(0, 200)
                            });
                        } else {
                            resolve({
                                status: 'ERROR',
                                error: 'Invalid JSON response from server',
                                rawResponse: xhr.responseText.substring(0, 200)
                            });
                        }
                    }
                } else {
                    console.error(`❌ HTTP error: ${xhr.status} ${xhr.statusText}`);
                    resolve({
                        status: 'ERROR',
                        error: `HTTP ${xhr.status}: ${xhr.statusText}`,
                        statusCode: xhr.status
                    });
                }
            };
            
            xhr.onerror = function() {
                console.error('❌ Network error - Failed to fetch');
                resolve({
                    status: 'ERROR',
                    error: 'Network error - ไม่สามารถเชื่อมต่อกับเซิร์ฟเวอร์ได้',
                    suggestion: 'ตรวจสอบการเชื่อมต่ออินเทอร์เน็ตหรือลองใหม่อีกครั้ง'
                });
            };
            
            xhr.ontimeout = function() {
                console.log('⏰ Request timeout (ไม่แสดงให้ผู้ใช้เห็น)');
                resolve({
                    status: 'TIMEOUT',
                    error: 'request_timeout',
                    message: 'การเชื่อมต่อใช้เวลานาน'
                });
            };
            
            // ส่ง request
            xhr.send(params.toString());
            
        } catch (error) {
            console.error('❌ Unexpected error in callGASAPI:', error);
            resolve({
                status: 'ERROR',
                error: error.message || 'Unknown error'
            });
        }
    });
}

/**
 * ⭐ SERVICE: ดึงรายชื่อบุคลากร (ไม่ใช้ข้อมูลสำรอง)
 */
async function getPersonnelListService(forceRefresh = false) {
    try {
        // ตรวจสอบ cache และมีข้อมูลจริง
        if (!forceRefresh && window.personnelData.isLoaded && 
            window.personnelData.hasRealData && 
            window.personnelData.personnelList.length > 0) {
            
            const cacheAge = new Date() - new Date(window.personnelData.lastUpdated);
            const cacheValid = cacheAge < 5 * 60 * 1000; // Cache 5 นาที
            
            if (cacheValid) {
                console.log('📦 Using cached personnel data');
                return {
                    status: 'SUCCESS',
                    personnelList: window.personnelData.personnelList,
                    cached: true,
                    count: window.personnelData.personnelList.length,
                    hasRealData: true
                };
            }
        }
        
        console.log('📋 Fetching personnel list from API...');
        
        window.personnelData.loadingAttempts++;
        
        let lastError;
        for (let attempt = 1; attempt <= window.personnelData.maxRetries; attempt++) {
            try {
                console.log(`Attempt ${attempt}/${window.personnelData.maxRetries}`);
                
                const result = await callGASAPI('getPersonnelList');
                
                if (result.status === 'SUCCESS' && result.personnelList && result.personnelList.length > 0) {
                    // ✅ โหลดสำเร็จ - อัปเดต cache
                    window.personnelData.personnelList = result.personnelList;
                    window.personnelData.lastUpdated = new Date();
                    window.personnelData.isLoaded = true;
                    window.personnelData.retryCount = 0;
                    window.personnelData.loadingAttempts = 0;
                    window.personnelData.hasRealData = true;
                    
                    console.log(`✅ Personnel data cached: ${result.personnelList.length} names`);
                    
                    return {
                        status: 'SUCCESS',
                        personnelList: result.personnelList,
                        count: result.personnelList.length,
                        cached: false,
                        attempt: attempt,
                        hasRealData: true
                    };
                } else {
                    lastError = result.error || 'No data received';
                    console.warn(`Attempt ${attempt} failed:`, lastError);
                    
                    if (attempt < window.personnelData.maxRetries) {
                        await new Promise(resolve => setTimeout(resolve, 1000 * attempt));
                    }
                }
            } catch (error) {
                lastError = error.message;
                console.warn(`Attempt ${attempt} error:`, error);
                
                if (attempt < window.personnelData.maxRetries) {
                    await new Promise(resolve => setTimeout(resolve, 1000 * attempt));
                }
            }
        }
        
        // ❌ ล้มเหลวทุกครั้ง
        window.personnelData.retryCount++;
        window.personnelData.hasRealData = false;
        
        console.log('⚠️ Failed after all retries');
        
        return {
            status: 'ERROR',
            error: lastError || 'Failed to fetch personnel list',
            personnelList: [],
            hasRealData: false,
            attempts: window.personnelData.loadingAttempts
        };
        
    } catch (error) {
        console.error('❌ Error in getPersonnelListService:', error);
        return {
            status: 'ERROR',
            error: error.message,
            personnelList: [],
            hasRealData: false,
            attempts: window.personnelData.loadingAttempts
        };
    }
}

/**
 * ⭐ SERVICE: ดึงรายชื่อครู
 */
async function getTeacherListService(forceRefresh = false) {
    try {
        if (!forceRefresh && window.personnelData.teacherList.length > 0) {
            console.log('📦 Using cached teacher data');
            return {
                status: 'SUCCESS',
                teacherList: window.personnelData.teacherList,
                cached: true
            };
        }
        
        console.log('👨‍🏫 Fetching teacher list from API...');
        
        const result = await callGASAPI('getTeacherList');
        
        if (result.status === 'SUCCESS' && result.teacherList) {
            window.personnelData.teacherList = result.teacherList;
            console.log(`✅ Teacher data cached: ${result.teacherList.length} names`);
        }
        
        return result;
        
    } catch (error) {
        console.error('❌ Error in getTeacherListService:', error);
        return {
            status: 'ERROR',
            error: error.message,
            teacherList: []
        };
    }
}

// ====================================================
// 📝 FORM HELPER FUNCTIONS
// ====================================================

/**
 * ⭐ SERVICE: โหลดรายชื่อลงใน dropdown (ไม่แสดงข้อความสถานะ)
 */
async function loadNamesIntoDropdown(selectElementId, type = 'personnel', forceRefresh = false) {
    const selectElement = document.getElementById(selectElementId);
    
    if (!selectElement) {
        console.error(`❌ Select element #${selectElementId} not found`);
        return {
            success: false,
            error: `Element #${selectElementId} not found`,
            attempts: window.personnelData.loadingAttempts || 0
        };
    }
    
    try {
        const savedValue = selectElement.value;
        
        // แสดงสถานะกำลังโหลดแบบเรียบง่าย
        selectElement.innerHTML = '<option value="">-- กำลังโหลด... --</option>';
        selectElement.disabled = true;
        
        let result;
        let names = [];
        
        if (type === 'teacher') {
            result = await getTeacherListService(forceRefresh);
            names = result.teacherList || [];
        } else {
            result = await getPersonnelListService(forceRefresh);
            names = result.personnelList || [];
        }
        
        if (result.status === 'SUCCESS' && names.length > 0) {
            // ✅ มีข้อมูลจริง
            let optionsHTML = '<option value="">-- กรุณาเลือก --</option>';
            
            names.forEach(name => {
                optionsHTML += `<option value="${name}">${name}</option>`;
            });
            
            optionsHTML += '<option value="อื่น ๆ (กรอกเอง)">อื่น ๆ (กรอกเอง)</option>';
            
            selectElement.innerHTML = optionsHTML;
            selectElement.disabled = false;
            
            if (savedValue) {
                selectElement.value = savedValue;
            }
            
            console.log(`✅ Loaded ${names.length} ${type} names`);
            
            // ✅ ซ่อนข้อความสถานะทั้งหมด
            hideAllPersonnelStatusMessages();
            
            return {
                success: true,
                names: names,
                count: names.length,
                cached: result.cached || false,
                hasRealData: true,
                attempts: result.attempts || 1
            };
            
        } else {
            // ❌ ไม่มีข้อมูลจริง
            const attempts = result.attempts || window.personnelData.loadingAttempts || 1;
            
            selectElement.innerHTML = `
                <option value="">-- เลือกชื่อ --</option>
                <option value="อื่น ๆ (กรอกเอง)">อื่น ๆ (กรอกเอง)</option>
            `;
            selectElement.disabled = false;
            
            selectElement.removeAttribute('title');
            
            return {
                success: false,
                error: result.error || 'No data available',
                names: [],
                cached: false,
                hasRealData: false,
                attempts: attempts
            };
        }
        
    } catch (error) {
        console.error(`❌ Error loading ${type} names:`, error);
        
        const selectElement = document.getElementById(selectElementId);
        if (selectElement) {
            selectElement.innerHTML = `
                <option value="">-- เลือกชื่อ --</option>
                <option value="อื่น ๆ (กรอกเอง)">อื่น ๆ (กรอกเอง)</option>
            `;
            selectElement.disabled = false;
            selectElement.removeAttribute('title');
        }
        
        return {
            success: false,
            error: error.message,
            names: [],
            cached: false,
            hasRealData: false,
            attempts: window.personnelData.loadingAttempts || 1
        };
    }
}

/**
 * ⭐ ฟังก์ชันใหม่: ซ่อนข้อความสถานะทั้งหมด
 */
function hideAllPersonnelStatusMessages() {
    // ซ่อนข้อความสถานะหลัก
    const statusElement = document.getElementById('personnel-status');
    if (statusElement) {
        statusElement.style.display = 'none';
        statusElement.innerHTML = '';
    }
    
    // ลบคำเตือนต่างๆ
    const warningElement = document.getElementById('manual-input-warning');
    if (warningElement) {
        warningElement.style.display = 'none';
    }
    
    const instructionElement = document.querySelector('.manual-input-instruction');
    if (instructionElement) {
        instructionElement.remove();
    }
}

// ====================================================
// 🔄 FORM RESET FUNCTIONS
// ====================================================

/**
 * ⭐ ฟังก์ชันใหม่: รีเซ็ตเฉพาะหน้า formtraining.html (ไม่กลับหน้าหลัก)
 */
window.resetTrainingFormOnly = function() {
    console.log('🔄 Resetting only training form (staying on same page)...');
    
    // ✅ 1. รีเซ็ต dropdown
    const fullNameSelect = document.getElementById('fullName');
    const customNameInput = document.getElementById('fullNameCustom');
    
    if (fullNameSelect) {
        fullNameSelect.selectedIndex = 0;
        fullNameSelect.value = '';
    }
    
    if (customNameInput) {
        customNameInput.value = '';
        customNameInput.style.display = 'none';
    }
    
    // ✅ 2. รีเซ็ต courses (ถ้ามี)
    const coursesContainer = document.getElementById('courses-container');
    if (coursesContainer) {
        // รีเซ็ต course count ใน trainingApp ถ้ามี
        if (typeof window.trainingApp !== 'undefined') {
            window.trainingApp.courseCount = 0;
            window.trainingApp.uploadedCertificates = {};
            
            // ล้าง container และเพิ่ม course ใหม่
            coursesContainer.innerHTML = '';
            
            // เพิ่ม course ใหม่ (ถ้าฟังก์ชันมีอยู่)
            if (typeof window.trainingApp.addCourse === 'function') {
                window.trainingApp.addCourse();
            }
        }
    }
    
    // ✅ 3. รีเซ็ต photos (ถ้ามี)
    if (typeof window.trainingApp !== 'undefined') {
        window.trainingApp.uploadedPhotos = Array(4).fill(null);
        
        // รีเซ็ต preview images
        for (let i = 0; i < 4; i++) {
            const preview = document.getElementById(`photo-preview-${i}`);
            const photoItem = document.querySelectorAll('.photo-item')[i];
            
            if (preview) {
                preview.src = '';
                preview.style.display = 'none';
            }
            
            if (photoItem) {
                photoItem.classList.remove('uploaded');
            }
            
            const fileInput = document.querySelectorAll('.file-input')[i];
            if (fileInput) {
                fileInput.value = '';
            }
        }
        
        // อัพเดท photo count
        if (typeof window.trainingApp.updatePhotoCount === 'function') {
            window.trainingApp.updatePhotoCount();
        }
    }
    
    // ✅ 4. ซ่อนข้อความสถานะทั้งหมด
    hideAllPersonnelStatusMessages();
    
    // ✅ 5. ลบข้อความ error/success
    const errorMsg = document.getElementById('error-message');
    if (errorMsg) {
        errorMsg.style.display = 'none';
        errorMsg.textContent = '';
    }
    
    const successMsg = document.getElementById('success-message');
    if (successMsg) {
        successMsg.style.display = 'none';
    }
    
    // ✅ 6. รีเซ็ต tab ไปที่ tab แรก
    const tabData = document.getElementById('tab-data');
    const tabPhotos = document.getElementById('tab-photos');
    const dataTab = document.getElementById('data-tab');
    const photosTab = document.getElementById('photos-tab');
    
    if (tabData && tabPhotos && dataTab && photosTab) {
        tabData.classList.add('active');
        tabPhotos.classList.remove('active');
        dataTab.classList.add('active');
        photosTab.classList.remove('active');
    }
    
    // ✅ 7. Focus กลับไปที่ dropdown
    if (fullNameSelect) {
        setTimeout(() => {
            fullNameSelect.focus();
        }, 100);
    }
    
    console.log('✅ Training form reset complete (stayed on same page)');
    return true;
};

/**
 * ⭐ ฟังก์ชันใหม่: รีเซ็ตเฉพาะหน้า teacheraward.html
 */
window.resetTeacherAwardFormOnly = function() {
    console.log('🔄 Resetting only teacher award form (staying on same page)...');
    
    // ✅ 1. รีเซ็ต dropdown
    const fullNameSelect = document.getElementById('fullName');
    const customNameInput = document.getElementById('fullNameCustom');
    
    if (fullNameSelect) {
        fullNameSelect.selectedIndex = 0;
        fullNameSelect.value = '';
    }
    
    if (customNameInput) {
        customNameInput.value = '';
        customNameInput.style.display = 'none';
    }
    
    // ✅ 2. รีเซ็ต award details
    const awardName = document.getElementById('awardName');
    const awardLevel = document.getElementById('awardLevel');
    const awardYear = document.getElementById('awardYear');
    const awardCertificate = document.getElementById('awardCertificate');
    const awardCertificatePreview = document.getElementById('award-certificate-preview');
    
    if (awardName) awardName.value = '';
    if (awardLevel) awardLevel.value = '';
    if (awardYear) awardYear.value = '';
    if (awardCertificate) awardCertificate.value = '';
    if (awardCertificatePreview) {
        awardCertificatePreview.src = '';
        awardCertificatePreview.style.display = 'none';
    }
    
    // ✅ 3. ซ่อนข้อความสถานะทั้งหมด
    hideAllPersonnelStatusMessages();
    
    // ✅ 4. ลบข้อความ error/success
    const errorMsg = document.getElementById('error-message');
    if (errorMsg) {
        errorMsg.style.display = 'none';
        errorMsg.textContent = '';
    }
    
    const successMsg = document.getElementById('success-message');
    if (successMsg) {
        successMsg.style.display = 'none';
    }
    
    // ✅ 5. Focus กลับไปที่ dropdown
    if (fullNameSelect) {
        setTimeout(() => {
            fullNameSelect.focus();
        }, 100);
    }
    
    console.log('✅ Teacher award form reset complete (stayed on same page)');
    return true;
};

/**
 * ⭐ ฟังก์ชันใหม่: รีเซ็ตเฉพาะหน้า studentwork.html
 */
window.resetStudentWorkFormOnly = function() {
    console.log('🔄 Resetting only student work form (staying on same page)...');
    
    // ✅ 1. รีเซ็ต dropdown
    const fullNameSelect = document.getElementById('fullName');
    const customNameInput = document.getElementById('fullNameCustom');
    
    if (fullNameSelect) {
        fullNameSelect.selectedIndex = 0;
        fullNameSelect.value = '';
    }
    
    if (customNameInput) {
        customNameInput.value = '';
        customNameInput.style.display = 'none';
    }
    
    // ✅ 2. รีเซ็ต student details
    const studentId = document.getElementById('studentId');
    const studentClass = document.getElementById('studentClass');
    const workType = document.getElementById('workType');
    const workTitle = document.getElementById('workTitle');
    const workDescription = document.getElementById('workDescription');
    const workYear = document.getElementById('workYear');
    const workFiles = document.getElementById('workFiles');
    const workFilesPreview = document.getElementById('work-files-preview');
    
    if (studentId) studentId.value = '';
    if (studentClass) studentClass.value = '';
    if (workType) workType.value = '';
    if (workTitle) workTitle.value = '';
    if (workDescription) workDescription.value = '';
    if (workYear) workYear.value = '';
    if (workFiles) workFiles.value = '';
    if (workFilesPreview) {
        workFilesPreview.innerHTML = '';
    }
    
    // ✅ 3. รีเซ็ต photos (ถ้ามี)
    if (typeof window.studentWorkApp !== 'undefined') {
        window.studentWorkApp.uploadedPhotos = Array(3).fill(null);
        
        // รีเซ็ต preview images
        for (let i = 0; i < 3; i++) {
            const preview = document.getElementById(`work-photo-preview-${i}`);
            const photoItem = document.querySelectorAll('.work-photo-item')[i];
            
            if (preview) {
                preview.src = '';
                preview.style.display = 'none';
            }
            
            if (photoItem) {
                photoItem.classList.remove('uploaded');
            }
            
            const fileInput = document.querySelectorAll('.work-file-input')[i];
            if (fileInput) {
                fileInput.value = '';
            }
        }
        
        // อัพเดท photo count
        if (typeof window.studentWorkApp.updatePhotoCount === 'function') {
            window.studentWorkApp.updatePhotoCount();
        }
    }
    
    // ✅ 4. ซ่อนข้อความสถานะทั้งหมด
    hideAllPersonnelStatusMessages();
    
    // ✅ 5. ลบข้อความ error/success
    const errorMsg = document.getElementById('error-message');
    if (errorMsg) {
        errorMsg.style.display = 'none';
        errorMsg.textContent = '';
    }
    
    const successMsg = document.getElementById('success-message');
    if (successMsg) {
        successMsg.style.display = 'none';
    }
    
    // ✅ 6. Focus กลับไปที่ dropdown
    if (fullNameSelect) {
        setTimeout(() => {
            fullNameSelect.focus();
        }, 100);
    }
    
    console.log('✅ Student work form reset complete (stayed on same page)');
    return true;
};

/**
 * ⭐ ฟังก์ชันใหม่: รีเซ็ตเฉพาะหน้า test-mobile-upload.html
 */
window.resetMobileTestFormOnly = function() {
    console.log('🔄 Resetting mobile test form (staying on same page)...');
    
    // ✅ 1. ล้างไฟล์ที่เลือกทั้งหมด
    if (typeof window.mobileTest !== 'undefined') {
        window.mobileTest.selectedFiles = {
            reader: null,
            base64: null,
            upload: null
        };
    }
    
    // ✅ 2. รีเซ็ต input files
    const fileInputs = document.querySelectorAll('.file-input');
    fileInputs.forEach(input => {
        input.value = '';
    });
    
    // ✅ 3. ล้างผลลัพธ์ทั้งหมด
    const resultAreas = [
        'connection-result',
        'reader-result',
        'base64-result',
        'upload-result',
        'largefile-result'
    ];
    
    resultAreas.forEach(id => {
        const element = document.getElementById(id);
        if (element) {
            element.innerHTML = '';
        }
    });
    
    // ✅ 4. ซ่อน progress bars
    const progressBars = [
        'reader-progress',
        'upload-progress'
    ];
    
    progressBars.forEach(id => {
        const element = document.getElementById(id);
        if (element) {
            element.style.display = 'none';
        }
    });
    
    // ✅ 5. รีเซ็ต progress fills
    const progressFills = [
        'reader-progress-fill',
        'upload-progress-fill'
    ];
    
    progressFills.forEach(id => {
        const element = document.getElementById(id);
        if (element) {
            element.style.width = '0%';
        }
    });
    
    // ✅ 6. ซ่อน chunk info
    const chunkInfo = document.getElementById('upload-chunk-info');
    if (chunkInfo) {
        chunkInfo.style.display = 'none';
    }
    
    // ✅ 7. รีเซ็ต checkboxes
    const useCompression = document.getElementById('use-compression');
    if (useCompression) {
        useCompression.checked = false;
    }
    
    const simulateSlow = document.getElementById('simulate-slow');
    if (simulateSlow) {
        simulateSlow.checked = true;
    }
    
    console.log('✅ Mobile test form reset complete (stayed on same page)');
    return true;
};

/**
 * ⭐ ฟังก์ชันเดิม: รีเซ็ตฟอร์มทั้งหมดรวมถึง dropdown (ถ้ายังต้องการ)
 */
window.resetTrainingForm = function() {
    console.log('🔄 Resetting training form...');
    
    // 1. รีเซ็ต dropdown
    const fullNameSelect = document.getElementById('fullName');
    const customNameInput = document.getElementById('fullNameCustom');
    
    if (fullNameSelect) {
        fullNameSelect.selectedIndex = 0;
    }
    
    if (customNameInput) {
        customNameInput.value = '';
        customNameInput.style.display = 'none';
    }
    
    // 2. ซ่อนข้อความสถานะทั้งหมด
    hideAllPersonnelStatusMessages();
    
    // 3. รีเซ็ต cache
    window.personnelData.isLoaded = false;
    window.personnelData.hasRealData = false;
    
    console.log('✅ Form reset complete');
    
    return true;
};

// ====================================================
// 📱 MOBILE MENU FUNCTIONS
// ====================================================

function toggleMobileMenu() {
    const sidebar = document.getElementById('sidebar');
    const overlay = document.getElementById('mobile-overlay');
    
    if (!sidebar) return;
    
    isMobileMenuOpen = !isMobileMenuOpen;
    
    if (isMobileMenuOpen) {
        sidebar.classList.add('open');
        if (overlay) overlay.style.display = 'block';
        document.body.style.overflow = 'hidden';
    } else {
        sidebar.classList.remove('open');
        if (overlay) overlay.style.display = 'none';
        document.body.style.overflow = 'auto';
    }
}

function closeMobileMenu() {
    const sidebar = document.getElementById('sidebar');
    const overlay = document.getElementById('mobile-overlay');
    
    if (sidebar) sidebar.classList.remove('open');
    if (overlay) overlay.style.display = 'none';
    isMobileMenuOpen = false;
    document.body.style.overflow = 'auto';
}

// ====================================================
// 📂 PAGE LOADING FUNCTIONS
// ====================================================

function loadContent(url, targetId) {
    const targetElement = document.getElementById(targetId);
    if (!targetElement) {
        console.error('Target element not found:', targetId);
        return;
    }
    
    // ✅ เรียก cleanupDashboard ก่อนโหลดหน้าใหม่ (ถ้าอยู่ใน dashboard)
    if (currentPage.includes('dashboard.html') && typeof window.cleanupDashboard === 'function') {
        console.log('🧹 Cleaning up dashboard before page change...');
        window.cleanupDashboard();
    }
    
    currentPage = url;
    
    targetElement.innerHTML = `
        <div style="text-align: center; padding: 40px;">
            <div style="display: inline-block; width: 20px; height: 20px; border: 3px solid rgba(0, 123, 255, 0.3); border-radius: 50%; border-top-color: #007bff; animation: spin 1s ease-in-out infinite; margin-right: 10px;"></div>
            <p>กำลังโหลด...</p>
        </div>
    `;
    
    fetch(url)
        .then(response => {
            if (!response.ok) throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            return response.text();
        })
        .then(html => {
            targetElement.innerHTML = html;
            
            executeScriptsFromHTML(html);
            
            if (window.innerWidth <= 768) closeMobileMenu();
            
            // ✅ ส่ง html ไปให้ executePageScripts ด้วย
            executePageScripts(url, html);
            
            console.log(`✅ Loaded: ${url}`);
        })
        .catch(error => {
            console.error('❌ Error loading content:', error);
            targetElement.innerHTML = `
                <div style="padding: 20px; text-align: center; color: #721c24; background-color: #f8d7da; border: 1px solid #f5c6cb; border-radius: 4px;">
                    <h3>เกิดข้อผิดพลาด</h3>
                    <p>ไม่สามารถโหลดหน้าได้: ${error.message}</p>
                    <button onclick="loadContent('pages/homepage.html', 'main-content')" 
                            style="padding: 8px 16px; background-color: #007bff; color: white; border: none; border-radius: 4px; cursor: pointer;">
                        กลับสู่หน้าหลัก
                    </button>
                </div>
            `;
        });
}

function executeScriptsFromHTML(html) {
    console.log('📜 Executing scripts from HTML...');
    
    try {
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = html;
        
        const scripts = tempDiv.querySelectorAll('script');
        
        console.log(`Found ${scripts.length} script(s) in HTML`);
        
        scripts.forEach((script, index) => {
            try {
                if (script.src) {
                    const newScript = document.createElement('script');
                    newScript.src = script.src;
                    if (script.async) newScript.async = script.async;
                    if (script.defer) newScript.defer = script.defer;
                    document.head.appendChild(newScript);
                } else if (script.textContent.trim()) {
                    try {
                        const code = `(function() { ${script.textContent} })()`;
                        const func = new Function(code);
                        func();
                    } catch (funcError) {
                        eval(script.textContent);
                    }
                }
            } catch (scriptError) {
                console.error(`Error processing script ${index + 1}:`, scriptError);
            }
        });
        
        console.log('✅ Scripts executed successfully');
    } catch (error) {
        console.error('❌ Error in executeScriptsFromHTML:', error);
    }
}

// ====================================================
// 📝 FORM FUNCTIONS - สำหรับทั้ง 3 ฟอร์ม + Mobile Test
// ====================================================

function showTrainingForm() {
    console.log('📝 Loading Training Form...');
    loadContent('pages/formtraining.html', 'main-content');
}

function showTeacherAwardForm() {
    console.log('🏆 Loading Teacher Award Form...');
    loadContent('pages/teacheraward.html', 'main-content');
}

function showStudentWorkForm() {
    console.log('👨‍🎓 Loading Student Work Form...');
    loadContent('pages/studentwork.html', 'main-content');
}

// ✅ ฟังก์ชันใหม่: แสดง Mobile Test Page
function showMobileTest() {
    console.log('📱 Loading Mobile Test Page...');
    loadContent('pages/test-mobile-upload.html', 'main-content');
}

// ✅ ฟังก์ชัน: แสดง Dashboard (Dashboard จะทำงานอิสระ)
function showDashboard() {
    console.log('📊 Loading Dashboard...');
    
    // ✅ ตั้งค่าตัวแปร global สำหรับ dashboard.js
    window.GAS_API_URL = window.GAS_WEB_APP_URL || window.API_URL || GAS_WEB_APP_URL;
    console.log('🔗 Setting GAS_API_URL for dashboard:', window.GAS_API_URL);
    
    // ✅ ตั้งค่าตัวแปรบอกว่าโหลดผ่าน main.js
    window.mainJSLoaded = true;
    
    loadContent('pages/dashboard.html', 'main-content');
}

function showReports() {
    console.log('📊 Loading Reports...');
    loadContent('pages/reports.html', 'main-content');
}

function showExport() {
    console.log('📤 Loading Export...');
    loadContent('pages/export.html', 'main-content');
}

function showHomepage() {
    console.log('🏠 Loading Homepage...');
    loadContent('pages/homepage.html', 'main-content');
}

// ✅ Expose functions globally
window.showTrainingForm = showTrainingForm;
window.showTeacherAwardForm = showTeacherAwardForm;
window.showStudentWorkForm = showStudentWorkForm;
window.showMobileTest = showMobileTest; // ✅ เพิ่มฟังก์ชันนี้
window.showDashboard = showDashboard;
window.showReports = showReports;
window.showExport = showExport;
window.showHomepage = showHomepage;

// ====================================================
// 🔧 PAGE SCRIPT EXECUTION - แก้ไขแล้วให้รองรับ Dashboard + Mobile Test + Reports
// ====================================================

function executePageScripts(url, html = null) {
    console.log('🔧 Executing page scripts for:', url);
    
    if (url.includes('formtraining.html')) {
        console.log('📝 Form training page detected');
        setTimeout(() => {
            if (typeof window.initTrainingForm === 'function') {
                console.log('🚀 Calling initTrainingForm...');
                try {
                    window.initTrainingForm();
                } catch (error) {
                    console.error('Error calling initTrainingForm:', error);
                }
            } else if (typeof window.trainingApp !== 'undefined' && 
                      typeof window.trainingApp.init === 'function') {
                console.log('🚀 Calling trainingApp.init() directly...');
                try {
                    window.trainingApp.init();
                } catch (error) {
                    console.error('Error calling trainingApp.init():', error);
                }
            } else {
                console.error('❌ initTrainingForm or trainingApp.init not found');
            }
        }, 100);
    }
    
    if (url.includes('teacheraward.html')) {
        console.log('🏆 Teacher Award page detected');
        setTimeout(() => {
            if (typeof window.initializeTeacherAwardForm === 'function') {
                console.log('🚀 Calling initializeTeacherAwardForm...');
                try {
                    window.initializeTeacherAwardForm();
                } catch (error) {
                    console.error('Error calling initializeTeacherAwardForm:', error);
                }
            }
        }, 100);
    }
    
    if (url.includes('studentwork.html')) {
        console.log('👨‍🎓 Student Work page detected');
        setTimeout(() => {
            if (typeof window.initializeStudentWorkForm === 'function') {
                console.log('🚀 Calling initializeStudentWorkForm...');
                try {
                    window.initializeStudentWorkForm();
                } catch (error) {
                    console.error('Error calling initializeStudentWorkForm:', error);
                }
            }
        }, 100);
    }
    
    // ✅ Mobile Test Page
    if (url.includes('test-mobile-upload.html')) {
        console.log('📱 Mobile Test page detected - Initializing mobile test...');
        
        // รอให้ mobile test script โหลดและ execute เสร็จ
        setTimeout(() => {
            console.log('🚀 Attempting to initialize mobile test...');
            
            // ตรวจสอบว่ามีฟังก์ชัน mobileTest.init หรือไม่
            if (typeof window.mobileTest !== 'undefined' && 
                typeof window.mobileTest.init === 'function') {
                console.log('🎯 Calling mobileTest.init() from main.js...');
                try {
                    window.mobileTest.init();
                    console.log('✅ Mobile Test initialized successfully via main.js');
                } catch (error) {
                    console.error('❌ Error calling mobileTest.init():', error);
                }
            } else {
                console.error('❌ No mobileTest initialization function found');
                console.log('Available functions:', {
                    hasMobileTest: typeof window.mobileTest,
                    hasMobileTestInit: typeof window.mobileTest !== 'undefined' ? 
                                     typeof window.mobileTest.init : 'no mobileTest'
                });
            }
        }, 500); // รอ 500ms ให้ scripts โหลดเสร็จ
    }
    
    // ✅ Dashboard: เรียก initDashboard() เมื่อโหลดเสร็จ
    if (url.includes('dashboard.html')) {
        console.log('📊 Dashboard page detected - Initializing dashboard...');
        
        // รอให้ dashboard script โหลดและ execute เสร็จ
        setTimeout(() => {
            console.log('🚀 Attempting to initialize dashboard...');
            
            // ตรวจสอบว่ามีฟังก์ชัน initDashboard หรือไม่ (จาก dashboard.html)
            if (typeof window.initDashboard === 'function') {
                console.log('🎯 Calling initDashboard() from main.js...');
                try {
                    window.initDashboard();
                    console.log('✅ Dashboard initialized successfully via main.js');
                } catch (error) {
                    console.error('❌ Error calling initDashboard():', error);
                    
                    // Fallback: ลองเรียก dashboardApp.init() โดยตรง
                    if (typeof window.dashboardApp !== 'undefined' && 
                        typeof window.dashboardApp.init === 'function') {
                        console.log('🔄 Fallback: Calling dashboardApp.init() directly...');
                        try {
                            window.dashboardApp.init();
                        } catch (fallbackError) {
                            console.error('❌ Fallback also failed:', fallbackError);
                        }
                    }
                }
            } else if (typeof window.dashboardApp !== 'undefined' && 
                      typeof window.dashboardApp.init === 'function') {
                // ถ้าไม่มี initDashboard แต่มี dashboardApp.init
                console.log('🎯 Calling dashboardApp.init() directly...');
                try {
                    window.dashboardApp.init();
                    console.log('✅ Dashboard initialized via dashboardApp.init()');
                } catch (error) {
                    console.error('❌ Error calling dashboardApp.init():', error);
                }
            } else {
                console.error('❌ No dashboard initialization function found');
                console.log('Available functions:', {
                    hasInitDashboard: typeof window.initDashboard,
                    hasDashboardApp: typeof window.dashboardApp,
                    hasDashboardAppInit: typeof window.dashboardApp !== 'undefined' ? 
                                        typeof window.dashboardApp.init : 'no dashboardApp'
                });
            }
        }, 500); // รอ 500ms ให้ scripts โหลดเสร็จ
    }
    
    // ✅ Reports Page: เรียก initReports() เมื่อโหลดเสร็จ
    if (url.includes('reports.html')) {
        console.log('📊 Reports page detected - Initializing reports...');
        
        // รอให้ reports script โหลดและ execute เสร็จ
        setTimeout(() => {
            console.log('🚀 Attempting to initialize reports...');
            
            // ตรวจสอบว่ามีฟังก์ชัน initReports หรือไม่ (จาก reports.html)
            if (typeof window.initReports === 'function') {
                console.log('🎯 Calling initReports() from main.js...');
                try {
                    window.initReports();
                    console.log('✅ Reports initialized successfully via main.js');
                } catch (error) {
                    console.error('❌ Error calling initReports():', error);
                    
                    // Fallback: ลองเรียก reportApp.init() โดยตรง
                    if (typeof window.reportApp !== 'undefined' && 
                        typeof window.reportApp.init === 'function') {
                        console.log('🔄 Fallback: Calling reportApp.init() directly...');
                        try {
                            window.reportApp.init();
                        } catch (fallbackError) {
                            console.error('❌ Fallback also failed:', fallbackError);
                        }
                    }
                }
            } else if (typeof window.reportApp !== 'undefined' && 
                      typeof window.reportApp.init === 'function') {
                // ถ้าไม่มี initReports แต่มี reportApp.init
                console.log('🎯 Calling reportApp.init() directly...');
                try {
                    window.reportApp.init();
                    console.log('✅ Reports initialized via reportApp.init()');
                } catch (error) {
                    console.error('❌ Error calling reportApp.init():', error);
                }
            } else {
                console.error('❌ No reports initialization function found');
                console.log('Available functions:', {
                    hasInitReports: typeof window.initReports,
                    hasReportApp: typeof window.reportApp,
                    hasReportAppInit: typeof window.reportApp !== 'undefined' ? 
                                     typeof window.reportApp.init : 'no reportApp'
                });
            }
        }, 500); // รอ 500ms ให้ scripts โหลดเสร็จ
    }
}

// ====================================================
// 🚀 INITIALIZATION
// ====================================================

window.onload = function() {
    console.log('🚀 Application starting...');
    
    console.log('🔧 Global variables:');
    console.log('- API_URL:', window.API_URL);
    console.log('- GAS_WEB_APP_URL:', window.GAS_WEB_APP_URL);
    console.log('- GAS_API_URL:', window.GAS_API_URL);
    
    loadContent('pages/menuslidebar.html', 'sidebar');
    
    setTimeout(() => {
        loadContent('pages/homepage.html', 'main-content');
    }, 100);
    
    window.loadContent = loadContent;
    window.toggleMobileMenu = toggleMobileMenu;
    window.closeMobileMenu = closeMobileMenu;
    window.executeScriptsFromHTML = executeScriptsFromHTML;
    
    window.callGASAPI = callGASAPI;
    window.getPersonnelListService = getPersonnelListService;
    window.getTeacherListService = getTeacherListService;
    
    window.loadNamesIntoDropdown = loadNamesIntoDropdown;
    
    // ✅ เพิ่ม cleanup function สำหรับเมื่อปิดหน้า
    window.addEventListener('beforeunload', function() {
        if (typeof window.cleanupDashboard === 'function') {
            window.cleanupDashboard();
        }
    });
    
    console.log('✅ Application initialized (พร้อมรองรับ Dashboard, Mobile Test และ Reports)');
};

// ====================================================
// 📤 SUBMIT FUNCTION - แก้ไขให้รองรับ Mobile Test
// ====================================================

async function submitToGoogleAppsScript(data) {
    console.log('📤 Submitting data to Google Apps Script:', { 
        action: data.action, 
        fullName: data.fullName,
        formType: data.formType || 'unknown'
    });
    
    try {
        const progressOverlay = document.getElementById('submit-progress-overlay');
        const progressFill = document.getElementById('submit-progress-fill');
        const progressText = document.getElementById('submit-progress-text');
        
        if (progressOverlay) {
            progressOverlay.style.display = 'flex';
            progressFill.style.width = '10%';
            progressText.textContent = 'กำลังเตรียมข้อมูล...';
        }
        
        let result;
        try {
            const timeoutPromise = new Promise((_, reject) => {
                setTimeout(() => {
                    reject(new Error('SUBMIT_TIMEOUT'));
                }, 45000);
            });
            
            const apiPromise = callGASAPI(data.action, data);
            result = await Promise.race([apiPromise, timeoutPromise]);
            
        } catch (timeoutError) {
            console.log('⏰ Submit timeout occurred');
            
            if (progressFill && progressText) {
                progressFill.style.width = '70%';
                progressText.textContent = 'กำลังบันทึกข้อมูล...';
            }
            
            await new Promise(resolve => setTimeout(resolve, 10000));
            result = await callGASAPI(data.action, data);
        }
        
        if (progressFill && progressText) {
            if (result.status === 'SUCCESS') {
                progressFill.style.width = '100%';
                progressText.textContent = 'บันทึกข้อมูลสำเร็จ!';
                
                setTimeout(() => {
                    if (progressOverlay) progressOverlay.style.display = 'none';
                    progressFill.style.width = '0%';
                    
                    // ✅ แก้ไข: เรียกฟังก์ชันรีเซ็ตที่เหมาะสมตามฟอร์ม
                    const formType = data.formType || data.action || '';
                    
                    if (formType.includes('training') || data.action === 'submitTraining') {
                        if (typeof window.resetTrainingFormOnly === 'function') {
                            setTimeout(() => {
                                window.resetTrainingFormOnly();
                            }, 500);
                        }
                    } else if (formType.includes('teacher') || data.action === 'submitTeacherAward') {
                        if (typeof window.resetTeacherAwardFormOnly === 'function') {
                            setTimeout(() => {
                                window.resetTeacherAwardFormOnly();
                            }, 500);
                        }
                    } else if (formType.includes('student') || data.action === 'submitStudentWork') {
                        if (typeof window.resetStudentWorkFormOnly === 'function') {
                            setTimeout(() => {
                                window.resetStudentWorkFormOnly();
                            }, 500);
                        }
                    } else if (formType.includes('testMobileUpload') || data.action === 'testMobileUpload') {
                        // ✅ สำหรับ Mobile Test Page
                        if (typeof window.resetMobileTestFormOnly === 'function') {
                            setTimeout(() => {
                                window.resetMobileTestFormOnly();
                            }, 500);
                        }
                    }
                }, 1500);
                
            } else if (result.status === 'TIMEOUT') {
                progressFill.style.width = '100%';
                progressText.textContent = 'กำลังตรวจสอบการบันทึก...';
                
                setTimeout(() => {
                    if (progressOverlay) progressOverlay.style.display = 'none';
                    progressFill.style.width = '0%';
                }, 2000);
                
                return { 
                    success: true,
                    isTimeout: true,
                    message: 'การบันทึกอาจใช้เวลานาน กรุณาตรวจสอบในระบบอีกครั้ง'
                };
            } else {
                progressFill.style.width = '0%';
                progressText.textContent = 'เกิดข้อผิดพลาด';
                setTimeout(() => {
                    if (progressOverlay) progressOverlay.style.display = 'none';
                }, 1000);
            }
        }
        
        if (result.status === 'TIMEOUT') {
            return { 
                success: true,
                isTimeout: true,
                message: 'การบันทึกอาจใช้เวลานาน กรุณาตรวจสอบในระบบอีกครั้ง'
            };
        }
        
        return { 
            success: result.status === 'SUCCESS',
            ...result
        };
        
    } catch (error) {
        console.error('❌ Error submitting to Google Apps Script:', error);
        
        const progressOverlay = document.getElementById('submit-progress-overlay');
        if (progressOverlay) progressOverlay.style.display = 'none';
        
        if (error.message === 'SUBMIT_TIMEOUT') {
            return { 
                success: true,
                isTimeout: true,
                message: 'การบันทึกใช้เวลานาน แต่อาจบันทึกสำเร็จแล้ว กรุณาตรวจสอบในระบบอีกครั้ง'
            };
        }
        
        return { 
            success: false, 
            error: error.message || 'เกิดข้อผิดพลาดในการเชื่อมต่อ'
        };
    }
}

window.submitToGoogleAppsScript = submitToGoogleAppsScript;

// ====================================================
// 🎯 FORM INITIALIZATION FUNCTIONS
// ====================================================

window.initTrainingForm = function() {
    console.log('🎯 initTrainingForm called from main.js');
    
    if (typeof window.trainingApp === 'undefined') {
        console.error('❌ trainingApp is not defined');
        return false;
    }
    
    if (typeof window.trainingApp.init !== 'function') {
        console.error('❌ trainingApp.init is not a function');
        return false;
    }
    
    try {
        console.log('🚀 Initializing training form...');
        window.trainingApp.init();
        return true;
    } catch (error) {
        console.error('❌ Error initializing training form:', error);
        return false;
    }
};

// ✅ ฟังก์ชันใหม่: เรียก init สำหรับ dashboard
window.initDashboard = function() {
    console.log('🎯 initDashboard called from main.js');
    
    if (typeof window.dashboardApp === 'undefined') {
        console.error('❌ dashboardApp is not defined');
        return false;
    }
    
    if (typeof window.dashboardApp.init !== 'function') {
        console.error('❌ dashboardApp.init is not a function');
        return false;
    }
    
    try {
        console.log('🚀 Initializing dashboard...');
        window.dashboardApp.init();
        return true;
    } catch (error) {
        console.error('❌ Error initializing dashboard:', error);
        return false;
    }
};

// ✅ ฟังก์ชันใหม่: เรียก init สำหรับ mobile test
window.initMobileTest = function() {
    console.log('🎯 initMobileTest called from main.js');
    
    if (typeof window.mobileTest === 'undefined') {
        console.error('❌ mobileTest is not defined');
        return false;
    }
    
    if (typeof window.mobileTest.init !== 'function') {
        console.error('❌ mobileTest.init is not a function');
        return false;
    }
    
    try {
        console.log('🚀 Initializing mobile test...');
        window.mobileTest.init();
        return true;
    } catch (error) {
        console.error('❌ Error initializing mobile test:', error);
        return false;
    }
};

// ✅ ฟังก์ชันใหม่: เรียก init สำหรับ reports
// ✅ ฟังก์ชันใหม่: เรียก init สำหรับ reports
window.initReports = function() {
    console.log('🎯 initReports called from main.js');
    
    // ✅ แก้ไข: เรียก initSummary แทน
    if (typeof window.initSummary === 'function') {
        console.log('🚀 Initializing reports via initSummary...');
        window.initSummary();
        return true;
    }
    
    // ✅ หรือเรียก summaryApp.init() โดยตรง
    if (typeof window.summaryApp !== 'undefined' && 
        typeof window.summaryApp.init === 'function') {
        console.log('🎯 Calling summaryApp.init() directly...');
        window.summaryApp.init();
        return true;
    }
    
    console.error('❌ No reports initialization function found');
    return false;
};

window.loadPersonnelData = async function() {
    console.log('👥 loadPersonnelData called from main.js');
    
    const fullNameSelect = document.getElementById('fullName');
    if (!fullNameSelect) {
        console.log('⚠️ Not on formtraining page');
        return false;
    }
    
    try {
        hideAllPersonnelStatusMessages();
        
        fullNameSelect.innerHTML = '<option value="">-- กำลังโหลด... --</option>';
        fullNameSelect.disabled = true;
        
        const result = await loadNamesIntoDropdown('fullName', 'personnel', false);
        
        if (result.success && result.hasRealData) {
            console.log(`✅ Loaded ${result.names.length} real personnel names`);
            
            const customNameInput = document.getElementById('fullNameCustom');
            if (fullNameSelect && customNameInput) {
                fullNameSelect.addEventListener('change', function() {
                    if (this.value === 'อื่น ๆ (กรอกเอง)') {
                        customNameInput.style.display = 'block';
                        customNameInput.focus();
                    } else {
                        customNameInput.style.display = 'none';
                        customNameInput.value = '';
                    }
                });
            }
            
            return result.names;
        } else {
            console.log('⚠️ No real data available');
            fullNameSelect.disabled = false;
            return false;
        }
        
    } catch (error) {
        console.error('❌ Error in loadPersonnelData:', error);
        
        if (fullNameSelect) {
            fullNameSelect.disabled = false;
        }
        
        return false;
    }
};

window.refreshPersonnelList = function(selectElementId = null) {
    console.log('🔄 Refreshing personnel list...');
    
    window.personnelData.isLoaded = false;
    window.personnelData.hasRealData = false;
    window.personnelData.retryCount = 0;
    
    if (selectElementId) {
        loadNamesIntoDropdown(selectElementId, 'personnel', true);
    } else {
        getPersonnelListService(true);
    }
};

// ====================================================
// 🎯 AUTO-LOAD FUNCTIONS
// ====================================================

if (document.getElementById('training-form-container')) {
    console.log('🔄 Auto-initializing formtraining page...');
    
    setTimeout(() => {
        if (typeof window.trainingApp !== 'undefined' && 
            typeof window.trainingApp.init === 'function') {
            window.trainingApp.init();
        }
    }, 300);
}

// ====================================================
// 📱 EVENT LISTENERS
// ====================================================

window.addEventListener('resize', function() {
    if (window.innerWidth > 768) closeMobileMenu();
});

console.log('✅ Main.js loaded successfully (พร้อมรองรับ Dashboard, Mobile Test และ Reports auto-load)');

// ✅ เพิ่ม CSS สำหรับ loading animation และ notifications
const loadingStyle = document.createElement('style');
loadingStyle.textContent = `
    @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }
    
    @keyframes slideIn {
        from { transform: translateX(100%); opacity: 0; }
        to { transform: translateX(0); opacity: 1; }
    }
    
    @keyframes slideOut {
        from { transform: translateX(0); opacity: 1; }
        to { transform: translateX(100%); opacity: 0; }
    }
`;
document.head.appendChild(loadingStyle);