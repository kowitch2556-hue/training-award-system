// ==============================================
// 🎯 CONFIGURATION
// ==============================================

// 🚨 กรุณาเปลี่ยน ID โฟลเดอร์นี้ให้เป็น ID โฟลเดอร์ Google Drive ของคุณ
const FOLDER_ID = '1QLM4C2NBRo9bXTrf4IwL82gRnf67QvMn'; 

// ชื่อ Sheet ต่างๆ
const TRAINING_SHEET_NAME = 'ข้อมูลการอบรม';
const PHOTO_SHEET_NAME = 'รูปกิจกรรม';
const TEACHER_AWARD_SHEET_NAME = 'ข้อมูลผลงานครู';
const STUDENT_WORK_SHEET_NAME = 'ข้อมูลผลงานนักเรียน';
const PERSONNEL_SHEET_NAME = 'รายชื่อบุคลากร';

// 🔄 Configuration สำหรับ Chunk Upload
const CHUNK_SETTINGS = {
  MAX_CHUNK_SIZE: 512 * 1024, // 512KB per chunk (เหมาะสมกับมือถือ)
  CHUNK_EXPIRY_HOURS: 24, // ลบ chunk เก่าทิ้งหลังจาก 24 ชั่วโมง
  MAX_CHUNKS_PER_FILE: 100, // จำกัดจำนวน chunk สูงสุดต่อไฟล์
  CLEANUP_INTERVAL_MINUTES: 60 // ทำความสะอาดทุก 60 นาที
};

// 📁 ชื่อ Properties สำหรับเก็บ chunk
const PROPERTIES_PREFIX = 'CHUNK_';

// ==============================================
// 🛠️ UTILITY FUNCTIONS
// ==============================================

/**
 * ป้องกันการเข้าถึง URL ของ Apps Script โดยตรง
 */
function doGet(e) {
    return ContentService.createTextOutput("API Endpoint - Access Denied.").setMimeType(ContentService.MimeType.TEXT);
}

/**
 * ฟังก์ชันสร้างรหัสเฉพาะสำหรับแต่ละการบันทึก
 */
function generateUniqueId() {
    return 'SUB_' + Date.now() + Math.random().toString(36).substring(2, 6).toUpperCase();
}

/**
 * ฟังก์ชันอัปโหลดไฟล์ไปที่ Google Drive และตั้งค่าการแชร์
 */
function uploadFile(base64Data, filename, mimeType, folder, isChunked = false) {
    try {
        let pureBase64 = base64Data;
        
        // ถ้าเป็น chunked data ที่ส่งมาเป็น string ต่อกันแล้ว
        if (isChunked) {
            // chunked data ควรเป็น base64 ที่ต่อกันแล้ว
            const decoded = Utilities.base64Decode(pureBase64);
            const blob = Utilities.newBlob(decoded, mimeType, filename);
            
            const file = folder.createFile(blob);
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            return file;
        }
        
        // สำหรับ base64 แบบเดิม
        if (base64Data.includes('data:')) {
            const base64Marker = 'base64,';
            const base64Index = base64Data.indexOf(base64Marker);
            if (base64Index > -1) {
                pureBase64 = base64Data.substring(base64Index + base64Marker.length);
            }
        } else if (base64Data.includes(',')) {
            pureBase64 = base64Data.split(',')[1];
        }
        
        // แปลง Base64 เป็น Blob
        const decoded = Utilities.base64Decode(pureBase64);
        const blob = Utilities.newBlob(decoded, mimeType, filename);
        
        // อัปโหลดไฟล์ไปยัง Drive
        const file = folder.createFile(blob);
        
        // ตั้งค่าการแชร์ไฟล์เป็น "Anyone with the link"
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        return file;
    } catch (error) {
        console.log("Upload File Error - Filename: " + filename);
        console.log("Error details: " + error.toString());
        if (base64Data && base64Data.length > 0) {
            console.log("Base64 data sample: " + base64Data.substring(0, 100));
        }
        throw new Error("Failed to upload file: " + error.toString());
    }
}

/**
 * ✅ ฟังก์ชันดึงรายชื่อบุคลากรทั้งหมดจาก Sheet 'รายชื่อบุคลากร'
 */
function getPersonnelNames() {
    try {
        console.log('📋 Fetching personnel names from sheet...');
        
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const personnelSheet = ss.getSheetByName(PERSONNEL_SHEET_NAME);
        
        if (!personnelSheet) {
            console.log('❌ Personnel sheet not found:', PERSONNEL_SHEET_NAME);
            return [];
        }
        
        const lastRow = personnelSheet.getLastRow();
        console.log('📊 Personnel sheet last row:', lastRow);
        
        if (lastRow <= 0) {
            console.log('ℹ️ Personnel sheet is empty');
            return [];
        }
        
        // ดึงข้อมูลจากคอลัมน์ A ทั้งหมด
        const data = personnelSheet.getRange(1, 1, lastRow, 1).getValues();
        console.log('📥 Raw data fetched:', data.length, 'rows');
        
        // กรองเฉพาะแถวที่มีข้อมูล
        const names = data
            .flat()
            .filter(name => name && name.toString().trim() !== '')
            .map(name => name.toString().trim());
        
        console.log('✅ Filtered names:', names.length);
        console.log('📝 Sample names:', names.slice(0, 5));
        
        return names;
        
    } catch (error) {
        console.log('❌ Error in getPersonnelNames:', error.toString());
        console.log('📝 Stack trace:', error.stack);
        return [];
    }
}

// ==============================================
// 🔄 CHUNK UPLOAD FUNCTIONS (ใหม่)
// ==============================================

/**
 * 🔄 สร้าง unique key สำหรับ chunk
 */
function generateChunkKey(recordId, chunkIndex) {
    return `${PROPERTIES_PREFIX}${recordId}_${chunkIndex}`;
}

/**
 * 🔄 บันทึก chunk ลง Cache หรือ Properties
 */
function saveChunkToCache(recordId, chunkIndex, chunkData, metadata) {
    try {
        const cache = CacheService.getScriptCache();
        const properties = PropertiesService.getScriptProperties();
        
        const chunkKey = generateChunkKey(recordId, chunkIndex);
        
        // บันทึกข้อมูล chunk
        cache.put(chunkKey, chunkData, 3600); // เก็บใน cache 1 ชั่วโมง
        
        // บันทึก metadata ใน Properties (เก็บได้นานกว่า)
        const metaKey = `${chunkKey}_META`;
        properties.setProperty(metaKey, JSON.stringify({
            recordId: recordId,
            chunkIndex: chunkIndex,
            timestamp: new Date().toISOString(),
            fileName: metadata.fileName,
            fileType: metadata.fileType,
            totalChunks: metadata.totalChunks,
            fileSize: metadata.fileSize
        }));
        
        console.log(`✅ Saved chunk ${chunkIndex}/${metadata.totalChunks} for ${recordId}`);
        return true;
    } catch (error) {
        console.log(`❌ Error saving chunk: ${error.toString()}`);
        return false;
    }
}

/**
 * 🔄 ดึง chunk จาก Cache
 */
function getChunkFromCache(recordId, chunkIndex) {
    try {
        const cache = CacheService.getScriptCache();
        const chunkKey = generateChunkKey(recordId, chunkIndex);
        const chunkData = cache.get(chunkKey);
        
        if (!chunkData) {
            console.log(`⚠️ Chunk ${chunkIndex} not found for ${recordId}`);
            return null;
        }
        
        return chunkData;
    } catch (error) {
        console.log(`❌ Error getting chunk: ${error.toString()}`);
        return null;
    }
}

/**
 * 🔄 รวม chunk ต่างๆ เป็นไฟล์เดียว
 */
function assembleFileFromChunks(recordId, totalChunks) {
    try {
        console.log(`🔄 Assembling file from ${totalChunks} chunks for ${recordId}`);
        
        let assembledData = '';
        let missingChunks = [];
        
        // เก็บ metadata
        const properties = PropertiesService.getScriptProperties();
        const firstMetaKey = generateChunkKey(recordId, 0) + '_META';
        const metadataStr = properties.getProperty(firstMetaKey);
        
        if (!metadataStr) {
            throw new Error(`Metadata not found for ${recordId}`);
        }
        
        const metadata = JSON.parse(metadataStr);
        console.log(`📄 File metadata: ${metadata.fileName} (${metadata.fileSize} bytes)`);
        
        // รวมทุก chunk
        for (let i = 0; i < totalChunks; i++) {
            const chunkData = getChunkFromCache(recordId, i);
            if (chunkData) {
                assembledData += chunkData;
            } else {
                missingChunks.push(i);
            }
        }
        
        if (missingChunks.length > 0) {
            throw new Error(`Missing chunks: ${missingChunks.join(', ')}`);
        }
        
        console.log(`✅ File assembled: ${assembledData.length} characters total`);
        
        // ส่งคืนข้อมูลและ metadata
        return {
            base64Data: assembledData,
            fileName: metadata.fileName,
            fileType: metadata.fileType,
            fileSize: metadata.fileSize
        };
    } catch (error) {
        console.log(`❌ Error assembling file: ${error.toString()}`);
        throw error;
    }
}

/**
 * 🔄 ลบ chunk เก่าทิ้ง
 */
function cleanupOldChunks() {
    try {
        const properties = PropertiesService.getScriptProperties();
        const cache = CacheService.getScriptCache();
        
        const allProperties = properties.getProperties();
        const now = new Date();
        let cleanedCount = 0;
        
        for (const [key, value] of Object.entries(allProperties)) {
            if (key.startsWith(PROPERTIES_PREFIX) && key.endsWith('_META')) {
                try {
                    const metadata = JSON.parse(value);
                    const chunkTime = new Date(metadata.timestamp);
                    const hoursDiff = (now - chunkTime) / (1000 * 60 * 60);
                    
                    // ลบ chunk ที่เก่ากว่า 24 ชั่วโมง
                    if (hoursDiff > CHUNK_SETTINGS.CHUNK_EXPIRY_HOURS) {
                        // ลบ chunk data จาก cache
                        const chunkKey = key.replace('_META', '');
                        cache.remove(chunkKey);
                        
                        // ลบ metadata จาก properties
                        properties.deleteProperty(key);
                        
                        cleanedCount++;
                        console.log(`🧹 Cleaned old chunk: ${chunkKey}`);
                    }
                } catch (e) {
                    console.log(`⚠️ Error parsing metadata for ${key}: ${e}`);
                }
            }
        }
        
        console.log(`✅ Cleanup completed: ${cleanedCount} chunks removed`);
        return cleanedCount;
    } catch (error) {
        console.log(`❌ Error in cleanup: ${error.toString()}`);
        return 0;
    }
}

/**
 * 🔄 ตรวจสอบสถานะ chunk
 */
function checkChunksStatus(recordId, totalChunks) {
    try {
        const cache = CacheService.getScriptCache();
        const receivedChunks = [];
        const missingChunks = [];
        
        for (let i = 0; i < totalChunks; i++) {
            const chunkKey = generateChunkKey(recordId, i);
            const chunkData = cache.get(chunkKey);
            
            if (chunkData) {
                receivedChunks.push(i);
            } else {
                missingChunks.push(i);
            }
        }
        
        return {
            recordId: recordId,
            totalChunks: totalChunks,
            receivedChunks: receivedChunks.length,
            missingChunks: missingChunks,
            percentage: Math.round((receivedChunks.length / totalChunks) * 100),
            isComplete: receivedChunks.length === totalChunks
        };
    } catch (error) {
        console.log(`❌ Error checking chunks: ${error.toString()}`);
        return {
            recordId: recordId,
            error: error.toString()
        };
    }
}

/**
 * 🔄 ลบ chunk ทั้งหมดของ record
 */
function deleteAllChunks(recordId, totalChunks) {
    try {
        const cache = CacheService.getScriptCache();
        const properties = PropertiesService.getScriptProperties();
        let deletedCount = 0;
        
        for (let i = 0; i < totalChunks; i++) {
            const chunkKey = generateChunkKey(recordId, i);
            const metaKey = `${chunkKey}_META`;
            
            // ลบจาก cache
            cache.remove(chunkKey);
            
            // ลบจาก properties
            properties.deleteProperty(metaKey);
            
            deletedCount++;
        }
        
        console.log(`🗑️ Deleted ${deletedCount} chunks for ${recordId}`);
        return deletedCount;
    } catch (error) {
        console.log(`❌ Error deleting chunks: ${error.toString()}`);
        return 0;
    }
}

// ==============================================
// 🌐 MAIN POST HANDLER
// ==============================================

/**
 * ฟังก์ชันหลักที่รับข้อมูล POST จาก Web App
 */
function doPost(e) {
    console.log('📨 Received POST request');
    
    let result = { status: 'ERROR', error: 'Invalid request' };
    
    try {
        if (!e) {
            throw new Error('Request object (e) is undefined');
        }
        
        let payload = null;
        
        // วิธีที่ 1: มาจาก FormData
        if (e.parameter && e.parameter.jsonPayload) {
            console.log('📝 Parsing from jsonPayload parameter...');
            payload = JSON.parse(e.parameter.jsonPayload);
        }
        // วิธีที่ 2: มาจาก POST body โดยตรง
        else if (e.postData && e.postData.contents) {
            console.log('📝 Parsing from postData contents...');
            try {
                payload = JSON.parse(e.postData.contents);
            } catch (parseError) {
                if (e.postData.type === 'application/x-www-form-urlencoded') {
                    const params = e.postData.contents.split('&');
                    for (const param of params) {
                        if (param.startsWith('jsonPayload=')) {
                            const jsonStr = decodeURIComponent(param.substring(12));
                            payload = JSON.parse(jsonStr);
                            break;
                        }
                    }
                }
            }
        }
        // วิธีที่ 3: มาจาก parameter อื่นๆ
        else if (e.parameter && e.parameter.payload) {
            console.log('📝 Parsing from payload parameter...');
            payload = JSON.parse(e.parameter.payload);
        }
        
        if (!payload) {
            console.log('⚠️ Could not parse payload');
            console.log('e.parameter:', e.parameter);
            console.log('e.postData:', e.postData ? 'exists' : 'null');
            throw new Error('Could not parse request payload');
        }
        
        console.log('✅ Successfully parsed payload');
        console.log('📦 Action:', payload.action);
        console.log('🔍 Payload keys:', Object.keys(payload));
        
        if (payload.action) {
            switch (payload.action) {
                // 📌 Existing actions
                case 'recordTraining':
                    result = recordTraining(payload);
                    break;
                case 'recordTeacherAward': 
                    result = recordTeacherAward(payload);
                    break;
                case 'recordStudentWork': 
                    result = recordStudentWork(payload);
                    break;
                case 'getReportsData':
                    result = getReportsData(payload);
                    break;
                    
                // 📊 NEW REPORT ACTIONS - สำหรับ reports.html
                case 'getTrainingSummary':
                    result = handleGetTrainingSummary(payload);
                    break;
                case 'getStudentSummary':
                    result = handleGetStudentSummary(payload);
                    break;
                case 'getTeacherSummary':
                    result = handleGetTeacherSummary(payload);
                    break;
                    
                // ✅ Personnel Actions
                case 'getPersonnelList':
                    result = handleGetPersonnelList(payload);
                    break;
                case 'getTeacherList':
                    result = handleGetTeacherList(payload);
                    break;
                case 'checkPersonnelSheet':
                    result = handleCheckPersonnelSheet();
                    break;
                case 'searchPersonnel':
                    result = handleSearchPersonnel(payload);
                    break;
                    
                // ✅ Dashboard Actions
                case 'getDashboardData':
                    result = handleGetDashboardData(payload);
                    break;
                case 'getDashboardSummary':
                    result = handleGetDashboardSummary(payload);
                    break;
                case 'getChartData':
                    result = handleGetChartData(payload);
                    break;
                case 'globalSearch':
                    result = handleGlobalSearch(payload);
                    break;
                    
                // ✅ Dashboard Actions (ใหม่)
                case 'getDashboardDataRealtime':
                    result = handleGetDashboardDataRealtime(payload);
                    break;
                case 'getTrendData':
                    result = handleGetTrendData(payload);
                    break;
                case 'getUnprocessedData':
                    result = handleGetUnprocessedData(payload);
                    break;
                    
                // 🔄 NEW: Chunk Upload Actions
                case 'recordTrainingChunk':
                    result = handleTrainingChunkUpload(payload);
                    break;
                    
                case 'recordTeacherAwardChunk':
                    result = handleTeacherAwardChunkUpload(payload);
                    break;
                    
                case 'recordStudentWorkChunk':
                    result = handleStudentWorkChunkUpload(payload);
                    break;
                    
                case 'checkChunksStatus':
                    result = handleCheckChunksStatus(payload);
                    break;
                    
                case 'completeChunkUpload':
                    result = handleCompleteChunkUpload(payload);
                    break;
                    
                case 'cancelChunkUpload':
                    result = handleCancelChunkUpload(payload);
                    break;
                    
                // ✅ Testing Actions
                case 'testTrainingDataCount':
                    result = testTrainingDataCount();
                    break;
                case 'testAllSheetsData':
                    result = testAllSheetsData();
                    break;
                case 'testDashboardDataAccuracy':
                    result = testDashboardDataAccuracy();
                    break;
                case 'quickDataCheck':
                    result = quickDataCheck();
                    break;
                    
                case 'test':
                    result = { 
                        status: 'SUCCESS', 
                        message: 'API is working!', 
                        timestamp: new Date().toISOString(),
                        receivedData: payload 
                    };
                    break;
                case 'testWriteToSheet':
                    result = testWriteToSheet();
                    break;
                case 'testEnvironment':
                    result = testEnvironment();
                    break;
                case 'testStudentWorkRecord':
                    result = testStudentWorkRecord();
                    break;
                case 'testPersonnelFunctions':
                    result = testPersonnelFunctions();
                    break;
                    
                default:
                    result.error = "Invalid action specified: " + payload.action;
                    break;
            }
        } else {
            result.error = "Invalid JSON payload or missing action.";
        }
        
    } catch (error) {
        console.log("❌ Error in doPost: " + error.toString());
        console.log("📝 Stack trace: " + error.stack);
        
        const debugInfo = {
            hasPostData: !!e.postData,
            postDataType: e.postData ? e.postData.type : 'none',
            hasParameters: !!e.parameter,
            parameterKeys: e.parameter ? Object.keys(e.parameter) : []
        };
        
        console.log("🔍 Debug info:", debugInfo);
        
        result.error = `Server Error: ${error.toString()}`;
        result.debug = debugInfo;
    }
    
    console.log("📤 Sending response:", result);
    return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
}

// ==============================================
// 📊 REPORT SUMMARY FUNCTIONS (สำหรับ reports.html)
// ==============================================

/**
 * ✅ Action: ดึงข้อมูลสรุปการอบรม (สำหรับ reports.html)
 */
function handleGetTrainingSummary(payload = {}) {
    try {
        console.log('🎓 Handling getTrainingSummary request...');
        
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const trainingSheet = ss.getSheetByName(TRAINING_SHEET_NAME);
        
        if (!trainingSheet || trainingSheet.getLastRow() <= 1) {
            return {
                status: 'SUCCESS',
                data: {
                    total: 0,
                    totalHours: 0,
                    uniqueParticipants: 0,
                    avgHours: 0,
                    recentData: [],
                    topParticipants: []
                }
            };
        }
        
        const lastRow = trainingSheet.getLastRow();
        const data = trainingSheet.getRange(2, 1, lastRow - 1, 11).getValues();
        
        // คำนวณข้อมูล
        let totalHours = 0;
        const participants = {};
        const recentData = [];
        
        data.forEach(row => {
            // คำนวณชั่วโมง
            const hours = parseFloat(row[7]) || 0;
            totalHours += hours;
            
            // นับจำนวนผู้เข้าร่วม
            const name = row[1];
            if (name) {
                if (!participants[name]) {
                    participants[name] = {
                        count: 0,
                        totalHours: 0
                    };
                }
                participants[name].count++;
                participants[name].totalHours += hours;
            }
            
            // เก็บข้อมูลล่าสุดสำหรับตาราง
            recentData.push({
                date: row[0],
                name: row[1],
                course: row[3],
                type: row[2],
                hours: hours,
                location: row[4]
            });
        });
        
        // เรียงลำดับผู้เข้าร่วมมากที่สุด
        const topParticipants = Object.entries(participants)
            .map(([name, stats]) => ({
                name: name,
                count: stats.count,
                totalHours: stats.totalHours
            }))
            .sort((a, b) => b.count - a.count)
            .slice(0, 10);
        
        return {
            status: 'SUCCESS',
            data: {
                total: data.length,
                totalHours: totalHours,
                uniqueParticipants: Object.keys(participants).length,
                avgHours: data.length > 0 ? (totalHours / Object.keys(participants).length).toFixed(1) : 0,
                recentData: recentData.slice(-10).reverse(), // 10 รายการล่าสุด
                topParticipants: topParticipants
            }
        };
        
    } catch (error) {
        console.log('❌ Error in handleGetTrainingSummary:', error.toString());
        return {
            status: 'ERROR',
            error: error.toString()
        };
    }
}

/**
 * ✅ Action: ดึงข้อมูลสรุปผลงานนักเรียน (สำหรับ reports.html)
 */
function handleGetStudentSummary(payload = {}) {
    try {
        console.log('👨‍🎓 Handling getStudentSummary request...');
        
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const studentSheet = ss.getSheetByName(STUDENT_WORK_SHEET_NAME);
        
        if (!studentSheet || studentSheet.getLastRow() <= 1) {
            return {
                status: 'SUCCESS',
                data: {
                    total: 0,
                    awardCount: 0,
                    uniqueStudents: 0,
                    uniqueAdvisors: 0,
                    recentData: [],
                    topStudents: []
                }
            };
        }
        
        const lastRow = studentSheet.getLastRow();
        const data = studentSheet.getRange(2, 1, lastRow - 1, 8).getValues();
        
        // คำนวณข้อมูล
        let awardCount = 0;
        const students = {};
        const recentData = [];
        
        data.forEach(row => {
            // นับรางวัล
            const awardLevel = row[5];
            if (awardLevel && awardLevel.trim() !== '') {
                awardCount++;
            }
            
            // นับจำนวนนักเรียน
            const studentName = row[1];
            if (studentName) {
                if (!students[studentName]) {
                    students[studentName] = {
                        count: 0,
                        awardCount: 0,
                        bestAward: awardLevel || 'ไม่มีรางวัล'
                    };
                }
                students[studentName].count++;
                if (awardLevel && awardLevel.trim() !== '') {
                    students[studentName].awardCount++;
                }
            }
            
            // เก็บข้อมูลล่าสุดสำหรับตาราง
            recentData.push({
                date: row[0],
                studentName: row[1],
                workName: row[3],
                level: row[5] || 'ไม่มีรางวัล',
                award: row[5] || 'ไม่มีรางวัล',
                advisor: row[4] || 'ไม่มีที่ปรึกษา'
            });
        });
        
        // นับที่ปรึกษาไม่ซ้ำ
        const advisors = new Set();
        data.forEach(row => {
            const advisor = row[4];
            if (advisor && advisor.trim() !== '') {
                advisors.add(advisor.trim());
            }
        });
        
        // เรียงลำดับนักเรียนที่มีผลงานมากที่สุด
        const topStudents = Object.entries(students)
            .map(([name, stats]) => ({
                name: name,
                count: stats.count,
                awardCount: stats.awardCount,
                bestAward: stats.bestAward,
                advisor: data.find(row => row[1] === name)?.[4] || 'ไม่มีที่ปรึกษา'
            }))
            .sort((a, b) => b.awardCount - a.awardCount)
            .slice(0, 10);
        
        return {
            status: 'SUCCESS',
            data: {
                total: data.length,
                awardCount: awardCount,
                uniqueStudents: Object.keys(students).length,
                uniqueAdvisors: advisors.size,
                recentData: recentData.slice(-10).reverse(), // 10 รายการล่าสุด
                topStudents: topStudents
            }
        };
        
    } catch (error) {
        console.log('❌ Error in handleGetStudentSummary:', error.toString());
        return {
            status: 'ERROR',
            error: error.toString()
        };
    }
}

/**
 * ✅ Action: ดึงข้อมูลสรุปรางวัลครู (สำหรับ reports.html)
 */
function handleGetTeacherSummary(payload = {}) {
    try {
        console.log('🏆 Handling getTeacherSummary request...');
        
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const teacherSheet = ss.getSheetByName(TEACHER_AWARD_SHEET_NAME);
        
        if (!teacherSheet || teacherSheet.getLastRow() <= 1) {
            return {
                status: 'SUCCESS',
                data: {
                    total: 0,
                    firstPlace: 0,
                    uniqueTeachers: 0,
                    avgPerTeacher: 0,
                    recentData: [],
                    topTeachers: []
                }
            };
        }
        
        const lastRow = teacherSheet.getLastRow();
        const data = teacherSheet.getRange(2, 1, lastRow - 1, 6).getValues();
        
        // คำนวณข้อมูล
        let firstPlaceCount = 0;
        const teachers = {};
        const recentData = [];
        
        data.forEach(row => {
            // นับรางวัลชนะเลิศ
            const awardLevel = row[3];
            if (awardLevel && (awardLevel.includes('ชนะเลิศ') || awardLevel.includes('เหรียญทอง') || awardLevel.includes('ที่ 1'))) {
                firstPlaceCount++;
            }
            
            // นับจำนวนครู
            const teacherName = row[1];
            if (teacherName) {
                if (!teachers[teacherName]) {
                    teachers[teacherName] = {
                        count: 0,
                        firstPlace: 0
                    };
                }
                teachers[teacherName].count++;
                if (awardLevel && (awardLevel.includes('ชนะเลิศ') || awardLevel.includes('เหรียญทอง') || awardLevel.includes('ที่ 1'))) {
                    teachers[teacherName].firstPlace++;
                }
            }
            
            // เก็บข้อมูลล่าสุดสำหรับตาราง
            recentData.push({
                date: row[0],
                teacherName: row[1],
                awardName: row[2],
                level: row[3] || 'ไม่ระบุ',
                achievement: row[2],
                organization: 'ไม่ระบุ'
            });
        });
        
        // เรียงลำดับครูที่มีรางวัลมากที่สุด
        const topTeachers = Object.entries(teachers)
            .map(([name, stats]) => ({
                name: name,
                count: stats.count,
                firstPlace: stats.firstPlace,
                department: 'ไม่ระบุ'
            }))
            .sort((a, b) => b.count - a.count)
            .slice(0, 10);
        
        return {
            status: 'SUCCESS',
            data: {
                total: data.length,
                firstPlace: firstPlaceCount,
                uniqueTeachers: Object.keys(teachers).length,
                avgPerTeacher: data.length > 0 ? (data.length / Object.keys(teachers).length).toFixed(1) : 0,
                recentData: recentData.slice(-10).reverse(), // 10 รายการล่าสุด
                topTeachers: topTeachers
            }
        };
        
    } catch (error) {
        console.log('❌ Error in handleGetTeacherSummary:', error.toString());
        return {
            status: 'ERROR',
            error: error.toString()
        };
    }
}

// ==============================================
// 🔄 CHUNK UPLOAD HANDLERS (ใหม่)
// ==============================================

/**
 * ✅ Action: บันทึกข้อมูลการอบรมแบบ chunked
 */
function handleTrainingChunkUpload(payload) {
    try {
        console.log('🎓 Handling training chunk upload...');
        
        // ตรวจสอบข้อมูลที่จำเป็น
        if (!payload.recordId) {
            return { status: 'ERROR', error: 'Missing recordId' };
        }
        
        if (payload.chunkData === undefined) {
            return { status: 'ERROR', error: 'Missing chunk data' };
        }
        
        // บันทึก chunk
        const saved = saveChunkToCache(
            payload.recordId,
            payload.chunkIndex,
            payload.chunkData,
            {
                fileName: payload.fileName || 'certificate',
                fileType: payload.fileType || 'application/octet-stream',
                totalChunks: payload.totalChunks,
                fileSize: payload.fileSize
            }
        );
        
        if (!saved) {
            return { status: 'ERROR', error: 'Failed to save chunk' };
        }
        
        // ตรวจสอบสถานะ
        const status = checkChunksStatus(payload.recordId, payload.totalChunks);
        
        return {
            status: 'SUCCESS',
            message: `Chunk ${payload.chunkIndex + 1}/${payload.totalChunks} saved`,
            chunkStatus: status,
            recordId: payload.recordId,
            chunkIndex: payload.chunkIndex,
            totalChunks: payload.totalChunks
        };
        
    } catch (error) {
        console.log('❌ Error in handleTrainingChunkUpload:', error.toString());
        return {
            status: 'ERROR',
            error: error.toString()
        };
    }
}

/**
 * ✅ Action: บันทึกผลงานครูแบบ chunked
 */
function handleTeacherAwardChunkUpload(payload) {
    try {
        console.log('🏆 Handling teacher award chunk upload...');
        
        // ตรวจสอบข้อมูลที่จำเป็น
        if (!payload.recordId) {
            return { status: 'ERROR', error: 'Missing recordId' };
        }
        
        if (payload.chunkData === undefined) {
            return { status: 'ERROR', error: 'Missing chunk data' };
        }
        
        // บันทึก chunk
        const saved = saveChunkToCache(
            payload.recordId,
            payload.chunkIndex,
            payload.chunkData,
            {
                fileName: payload.fileName || 'certificate',
                fileType: payload.fileType || 'application/octet-stream',
                totalChunks: payload.totalChunks,
                fileSize: payload.fileSize
            }
        );
        
        if (!saved) {
            return { status: 'ERROR', error: 'Failed to save chunk' };
        }
        
        // ตรวจสอบสถานะ
        const status = checkChunksStatus(payload.recordId, payload.totalChunks);
        
        return {
            status: 'SUCCESS',
            message: `Chunk ${payload.chunkIndex + 1}/${payload.totalChunks} saved`,
            chunkStatus: status,
            recordId: payload.recordId,
            chunkIndex: payload.chunkIndex,
            totalChunks: payload.totalChunks
        };
        
    } catch (error) {
        console.log('❌ Error in handleTeacherAwardChunkUpload:', error.toString());
        return {
            status: 'ERROR',
            error: error.toString()
        };
    }
}

/**
 * ✅ Action: บันทึกผลงานนักเรียนแบบ chunked
 */
function handleStudentWorkChunkUpload(payload) {
    try {
        console.log('👨‍🎓 Handling student work chunk upload...');
        
        // ตรวจสอบข้อมูลที่จำเป็น
        if (!payload.recordId) {
            return { status: 'ERROR', error: 'Missing recordId' };
        }
        
        if (payload.chunkData === undefined) {
            return { status: 'ERROR', error: 'Missing chunk data' };
        }
        
        // บันทึก chunk
        const saved = saveChunkToCache(
            payload.recordId,
            payload.chunkIndex,
            payload.chunkData,
            {
                fileName: payload.fileName || 'certificate',
                fileType: payload.fileType || 'application/octet-stream',
                totalChunks: payload.totalChunks,
                fileSize: payload.fileSize
            }
        );
        
        if (!saved) {
            return { status: 'ERROR', error: 'Failed to save chunk' };
        }
        
        // ตรวจสอบสถานะ
        const status = checkChunksStatus(payload.recordId, payload.totalChunks);
        
        return {
            status: 'SUCCESS',
            message: `Chunk ${payload.chunkIndex + 1}/${payload.totalChunks} saved`,
            chunkStatus: status,
            recordId: payload.recordId,
            chunkIndex: payload.chunkIndex,
            totalChunks: payload.totalChunks
        };
        
    } catch (error) {
        console.log('❌ Error in handleStudentWorkChunkUpload:', error.toString());
        return {
            status: 'ERROR',
            error: error.toString()
        };
    }
}

/**
 * ✅ Action: ตรวจสอบสถานะ chunk
 */
function handleCheckChunksStatus(payload) {
    try {
        if (!payload.recordId || !payload.totalChunks) {
            return { status: 'ERROR', error: 'Missing recordId or totalChunks' };
        }
        
        const status = checkChunksStatus(payload.recordId, payload.totalChunks);
        
        return {
            status: 'SUCCESS',
            chunkStatus: status
        };
        
    } catch (error) {
        console.log('❌ Error in handleCheckChunksStatus:', error.toString());
        return {
            status: 'ERROR',
            error: error.toString()
        };
    }
}

/**
 * ✅ Action: รวม chunk เป็นไฟล์และบันทึกข้อมูล
 */
function handleCompleteChunkUpload(payload) {
    try {
        console.log('✅ Completing chunk upload...');
        
        const { recordId, totalChunks, formData } = payload;
        
        if (!recordId || !totalChunks) {
            return { status: 'ERROR', error: 'Missing required fields' };
        }
        
        // ตรวจสอบว่าทุก chunk มาแล้วครบ
        const status = checkChunksStatus(recordId, totalChunks);
        
        if (!status.isComplete) {
            return {
                status: 'ERROR',
                error: `Incomplete upload: ${status.missingChunks.length} chunks missing`,
                chunkStatus: status
            };
        }
        
        // รวม chunk เป็นไฟล์
        const fileData = assembleFileFromChunks(recordId, totalChunks);
        
        // เรียกใช้ฟังก์ชันบันทึกข้อมูลเดิม
        let result;
        if (formData.formType === 'training') {
            result = recordTraining({
                ...formData,
                certificate: fileData
            });
        } else if (formData.formType === 'teacherAward') {
            result = recordTeacherAward({
                ...formData,
                certificate: fileData
            });
        } else if (formData.formType === 'studentWork') {
            result = recordStudentWork({
                ...formData,
                certificate: fileData
            });
        } else {
            throw new Error('Invalid form type');
        }
        
        // ลบ chunk หลังจากบันทึกสำเร็จ
        if (result.status === 'SUCCESS') {
            deleteAllChunks(recordId, totalChunks);
        }
        
        return result;
        
    } catch (error) {
        console.log('❌ Error in handleCompleteChunkUpload:', error.toString());
        return {
            status: 'ERROR',
            error: error.toString()
        };
    }
}

/**
 * ✅ Action: ยกเลิก chunk upload
 */
function handleCancelChunkUpload(payload) {
    try {
        const { recordId, totalChunks } = payload;
        
        if (!recordId) {
            return { status: 'ERROR', error: 'Missing recordId' };
        }
        
        const deleted = deleteAllChunks(recordId, totalChunks || 100);
        
        return {
            status: 'SUCCESS',
            message: `Cancelled upload and deleted ${deleted} chunks`,
            recordId: recordId
        };
        
    } catch (error) {
        console.log('❌ Error in handleCancelChunkUpload:', error.toString());
        return {
            status: 'ERROR',
            error: error.toString()
        };
    }
}

// ==============================================
// 📈 DASHBOARD DATA FUNCTIONS (ENHANCED)
// ==============================================

/**
 * ✅ Action: ดึงข้อมูล Dashboard ทั้งหมด
 */
function handleGetDashboardData(payload = {}) {
    try {
        console.log('📈 Handling getDashboardData request...');
        
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const result = {
            status: 'SUCCESS',
            timestamp: new Date().toISOString(),
            data: {
                summary: {},
                charts: {},
                recent: {},
                personnel: {}
            }
        };
        
        // 📊 สรุปข้อมูลแบบละเอียด
        const summary = {
            trainings: { total: 0, byType: {}, byMonth: {} },
            photos: { total: 0, byPerson: {} },
            teacherAwards: { total: 0, byLevel: {} },
            studentWorks: { total: 0, byWorkType: {}, byAwardLevel: {} },
            personnel: { total: 0, byTitle: {} }
        };
        
        // 📋 ดึงข้อมูลบุคลากร
        try {
            const personnelNames = getPersonnelNames();
            summary.personnel.total = personnelNames.length;
            
            personnelNames.forEach(name => {
                if (name.startsWith('นาย')) {
                    summary.personnel.byTitle.นาย = (summary.personnel.byTitle.นาย || 0) + 1;
                } else if (name.startsWith('นางสาว')) {
                    summary.personnel.byTitle.นางสาว = (summary.personnel.byTitle.นางสาว || 0) + 1;
                } else if (name.startsWith('นาง')) {
                    summary.personnel.byTitle.นาง = (summary.personnel.byTitle.นาง || 0) + 1;
                } else {
                    summary.personnel.byTitle.อื่นๆ = (summary.personnel.byTitle.อื่นๆ || 0) + 1;
                }
            });
            
            result.data.personnel = {
                total: summary.personnel.total,
                byTitle: summary.personnel.byTitle,
                list: personnelNames
            };
        } catch (e) {
            console.log('⚠️ Error fetching personnel:', e.toString());
            result.data.personnel = { total: 0, byTitle: {}, list: [] };
        }
        
        // 🎓 ดึงข้อมูลการอบรม
        try {
            const trainingSheet = ss.getSheetByName(TRAINING_SHEET_NAME);
            if (trainingSheet && trainingSheet.getLastRow() > 1) {
                const lastRow = trainingSheet.getLastRow();
                const data = trainingSheet.getRange(2, 1, lastRow - 1, 11).getValues();
                
                summary.trainings.total = data.length;
                
                data.forEach(row => {
                    const type = row[2] || 'ไม่ระบุ';
                    summary.trainings.byType[type] = (summary.trainings.byType[type] || 0) + 1;
                    
                    const timestamp = row[0];
                    if (timestamp instanceof Date) {
                        const monthYear = `${timestamp.getMonth() + 1}/${timestamp.getFullYear() + 543}`;
                        summary.trainings.byMonth[monthYear] = (summary.trainings.byMonth[monthYear] || 0) + 1;
                    }
                });
                
                const recentTrainings = data.map(row => ({
                    timestamp: row[0],
                    name: row[1],
                    type: row[2],
                    course: row[3],
                    location: row[4],
                    startDate: row[5],
                    endDate: row[6],
                    hours: row[7],
                    certificate: row[8]
                }));
                
                result.data.summary.trainings = summary.trainings;
                result.data.recent.trainings = recentTrainings;
            }
        } catch (e) {
            console.log('⚠️ Error fetching trainings:', e.toString());
        }
        
        // 📸 ดึงข้อมูลรูปกิจกรรม
        try {
            const photoSheet = ss.getSheetByName(PHOTO_SHEET_NAME);
            if (photoSheet && photoSheet.getLastRow() > 1) {
                const lastRow = photoSheet.getLastRow();
                const data = photoSheet.getRange(2, 1, lastRow - 1, 4).getValues();
                
                summary.photos.total = data.length;
                
                data.forEach(row => {
                    const person = row[1] || 'ไม่ระบุ';
                    summary.photos.byPerson[person] = (summary.photos.byPerson[person] || 0) + 1;
                });
                
                const recentPhotos = data.map(row => ({
                    timestamp: row[0],
                    person: row[1],
                    photoLink: row[2],
                    description: row[3]
                }));
                
                result.data.summary.photos = summary.photos;
                result.data.recent.photos = recentPhotos;
            }
        } catch (e) {
            console.log('⚠️ Error fetching photos:', e.toString());
        }
        
        // 🏆 ดึงข้อมูลผลงานครู
        try {
            const teacherSheet = ss.getSheetByName(TEACHER_AWARD_SHEET_NAME);
            if (teacherSheet && teacherSheet.getLastRow() > 1) {
                const lastRow = teacherSheet.getLastRow();
                const data = teacherSheet.getRange(2, 1, lastRow - 1, 6).getValues();
                
                summary.teacherAwards.total = data.length;
                
                data.forEach(row => {
                    const level = row[3] || 'ไม่ระบุ';
                    summary.teacherAwards.byLevel[level] = (summary.teacherAwards.byLevel[level] || 0) + 1;
                });
                
                const recentTeacherAwards = data.map(row => ({
                    timestamp: row[0],
                    teacherName: row[1],
                    awardName: row[2],
                    awardLevel: row[3],
                    awardDate: row[4]
                }));
                
                result.data.summary.teacherAwards = summary.teacherAwards;
                result.data.recent.teacherAwards = recentTeacherAwards;
            }
        } catch (e) {
            console.log('⚠️ Error fetching teacher awards:', e.toString());
        }
        
        // 👨‍🎓 ดึงข้อมูลผลงานนักเรียน
        try {
            const studentSheet = ss.getSheetByName(STUDENT_WORK_SHEET_NAME);
            if (studentSheet && studentSheet.getLastRow() > 1) {
                const lastRow = studentSheet.getLastRow();
                const data = studentSheet.getRange(2, 1, lastRow - 1, 8).getValues();
                
                summary.studentWorks.total = data.length;
                
                data.forEach(row => {
                    const workType = row[2] || 'ไม่ระบุ';
                    const awardLevel = row[5] || 'ไม่ระบุ';
                    
                    summary.studentWorks.byWorkType[workType] = (summary.studentWorks.byWorkType[workType] || 0) + 1;
                    summary.studentWorks.byAwardLevel[awardLevel] = (summary.studentWorks.byAwardLevel[awardLevel] || 0) + 1;
                });
                
                const recentStudentWorks = data.map(row => ({
                    timestamp: row[0],
                    studentName: row[1],
                    workType: row[2],
                    projectName: row[3],
                    advisorName: row[4],
                    awardLevel: row[5],
                    awardDate: row[6]
                }));
                
                result.data.summary.studentWorks = summary.studentWorks;
                result.data.recent.studentWorks = recentStudentWorks;
            }
        } catch (e) {
            console.log('⚠️ Error fetching student works:', e.toString());
        }
        
        // 📊 สรุปข้อมูลรวม
        result.data.summary.total = {
            personnel: summary.personnel.total,
            trainings: summary.trainings.total,
            photos: summary.photos.total,
            teacherAwards: summary.teacherAwards.total,
            studentWorks: summary.studentWorks.total,
            totalRecords: summary.trainings.total + summary.photos.total + 
                         summary.teacherAwards.total + summary.studentWorks.total
        };
        
        // 📈 เตรียมข้อมูลสำหรับ Chart
        result.data.charts = {
            trainingsByType: Object.entries(summary.trainings.byType).map(([type, count]) => ({ type, count })),
            teacherAwardsByLevel: Object.entries(summary.teacherAwards.byLevel).map(([level, count]) => ({ level, count })),
            studentWorksByType: Object.entries(summary.studentWorks.byWorkType).map(([type, count]) => ({ type, count })),
            personnelByTitle: Object.entries(summary.personnel.byTitle).map(([title, count]) => ({ title, count }))
        };
        
        console.log('✅ Dashboard data fetched successfully');
        
        return result;
        
    } catch (error) {
        console.log('❌ Error in handleGetDashboardData:', error.toString());
        console.log('📝 Stack trace:', error.stack);
        
        return {
            status: 'ERROR',
            error: 'ไม่สามารถดึงข้อมูล Dashboard ได้: ' + error.toString(),
            timestamp: new Date().toISOString()
        };
    }
}

/**
 * ✅ Action: ดึงข้อมูล Dashboard แบบเรียลไทม์ (ใหม่)
 */
function handleGetDashboardDataRealtime(payload = {}) {
    try {
        console.log('🔄 Handling getDashboardDataRealtime request...');
        
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const result = {
            status: 'SUCCESS',
            timestamp: new Date().toISOString(),
            data: {
                summary: {},
                charts: {},
                recent: {},
                personnel: {},
                latestUpdates: []
            }
        };
        
        // ตั้งค่าช่วงเวลาสำหรับดึงข้อมูลล่าสุด
        const daysBack = payload.daysBack || 30;
        const cutoffDate = new Date();
        cutoffDate.setDate(cutoffDate.getDate() - daysBack);
        
        // 📊 ดึงข้อมูลบุคลากร
        try {
            const personnelNames = getPersonnelNames();
            result.data.personnel = {
                total: personnelNames.length,
                list: personnelNames
            };
        } catch (e) {
            result.data.personnel = { total: 0, list: [] };
        }
        
        // 📈 ฟังก์ชันสำหรับดึงข้อมูลล่าสุดจาก Sheet
        const getLatestData = (sheetName, columns, dateColumnIndex = 0, limit = null) => {
            try {
                const sheet = ss.getSheetByName(sheetName);
                if (!sheet || sheet.getLastRow() <= 1) {
                    return [];
                }
                
                const lastRow = sheet.getLastRow();
                const data = sheet.getRange(2, 1, lastRow - 1, columns).getValues();
                
                // กรองข้อมูลล่าสุดตามช่วงเวลา
                const recentData = data.filter(row => {
                    const rowDate = row[dateColumnIndex];
                    if (!rowDate) return false;
                    
                    const date = rowDate instanceof Date ? rowDate : new Date(rowDate);
                    return date >= cutoffDate;
                });
                
                // เรียงลำดับจากใหม่ไปเก่า
                recentData.sort((a, b) => {
                    const dateA = a[dateColumnIndex] instanceof Date ? a[dateColumnIndex] : new Date(a[dateColumnIndex]);
                    const dateB = b[dateColumnIndex] instanceof Date ? b[dateColumnIndex] : new Date(b[dateColumnIndex]);
                    return dateB - dateA;
                });
                
                return limit ? recentData.slice(0, limit) : recentData;
            } catch (e) {
                console.log(`⚠️ Error fetching latest data from ${sheetName}:`, e.toString());
                return [];
            }
        };
        
        // 🎓 ดึงข้อมูลการอบรมล่าสุด
        const recentTrainings = getLatestData(TRAINING_SHEET_NAME, 11, 0, null);
        result.data.recent.trainings = recentTrainings.map(row => ({
            timestamp: row[0],
            name: row[1],
            type: row[2],
            course: row[3],
            location: row[4],
            startDate: row[5],
            endDate: row[6],
            hours: row[7],
            certificate: row[8]
        }));
        
        // 📸 ดึงข้อมูลรูปกิจกรรมล่าสุด
        const recentPhotos = getLatestData(PHOTO_SHEET_NAME, 4, 0, null);
        result.data.recent.photos = recentPhotos.map(row => ({
            timestamp: row[0],
            person: row[1],
            photoLink: row[2],
            description: row[3]
        }));
        
        // 🏆 ดึงข้อมูลผลงานครูล่าสุด
        const recentTeacherAwards = getLatestData(TEACHER_AWARD_SHEET_NAME, 6, 0, null);
        result.data.recent.teacherAwards = recentTeacherAwards.map(row => ({
            timestamp: row[0],
            teacherName: row[1],
            awardName: row[2],
            awardLevel: row[3],
            awardDate: row[4]
        }));
        
        // 👨‍🎓 ดึงข้อมูลผลงานนักเรียนล่าสุด
        const recentStudentWorks = getLatestData(STUDENT_WORK_SHEET_NAME, 8, 0, null);
        result.data.recent.studentWorks = recentStudentWorks.map(row => ({
            timestamp: row[0],
            studentName: row[1],
            workType: row[2],
            projectName: row[3],
            advisorName: row[4],
            awardLevel: row[5],
            awardDate: row[6]
        }));
        
        // 📊 คำนวณสถิติแบบละเอียด
        result.data.summary = calculateDetailedSummary(ss);
        
        // 📈 เตรียมข้อมูลสำหรับกราฟ
        result.data.charts = prepareChartData(ss);
        
        // 🔔 บันทึกการอัปเดตล่าสุด
        result.data.latestUpdates = getLatestUpdates(ss, 5);
        
        console.log('✅ Realtime dashboard data fetched successfully');
        console.log(`📅 Data from last ${daysBack} days`);
        
        return result;
        
    } catch (error) {
        console.log('❌ Error in handleGetDashboardDataRealtime:', error.toString());
        return {
            status: 'ERROR',
            error: 'ไม่สามารถดึงข้อมูล Dashboard แบบเรียลไทม์ได้: ' + error.toString()
        };
    }
}

/**
 * 📊 คำนวณสถิติแบบละเอียดจากทุก Sheets
 */
function calculateDetailedSummary(ss) {
    const summary = {
        trainings: { total: 0, byType: {}, byMonth: {}, recentCount: 0 },
        photos: { total: 0, byPerson: {}, recentCount: 0 },
        teacherAwards: { total: 0, byLevel: {}, recentCount: 0 },
        studentWorks: { total: 0, byWorkType: {}, byAwardLevel: {}, recentCount: 0 },
        personnel: { total: 0, byTitle: {} }
    };
    
    // คำนวณข้อมูลการอบรม
    const trainingSheet = ss.getSheetByName(TRAINING_SHEET_NAME);
    if (trainingSheet && trainingSheet.getLastRow() > 1) {
        const lastRow = trainingSheet.getLastRow();
        const data = trainingSheet.getRange(2, 1, lastRow - 1, 11).getValues();
        
        summary.trainings.total = data.length;
        
        data.forEach(row => {
            const type = row[2] || 'ไม่ระบุ';
            summary.trainings.byType[type] = (summary.trainings.byType[type] || 0) + 1;
            
            const timestamp = row[0];
            if (timestamp instanceof Date) {
                const monthYear = `${timestamp.getMonth() + 1}/${timestamp.getFullYear() + 543}`;
                summary.trainings.byMonth[monthYear] = (summary.trainings.byMonth[monthYear] || 0) + 1;
            }
        });
    }
    
    // คำนวณข้อมูลรูปกิจกรรม
    const photoSheet = ss.getSheetByName(PHOTO_SHEET_NAME);
    if (photoSheet && photoSheet.getLastRow() > 1) {
        const lastRow = photoSheet.getLastRow();
        const data = photoSheet.getRange(2, 1, lastRow - 1, 4).getValues();
        
        summary.photos.total = data.length;
        
        data.forEach(row => {
            const person = row[1] || 'ไม่ระบุ';
            summary.photos.byPerson[person] = (summary.photos.byPerson[person] || 0) + 1;
        });
    }
    
    // คำนวณข้อมูลผลงานครู
    const teacherSheet = ss.getSheetByName(TEACHER_AWARD_SHEET_NAME);
    if (teacherSheet && teacherSheet.getLastRow() > 1) {
        const lastRow = teacherSheet.getLastRow();
        const data = teacherSheet.getRange(2, 1, lastRow - 1, 6).getValues();
        
        summary.teacherAwards.total = data.length;
        
        data.forEach(row => {
            const level = row[3] || 'ไม่ระบุ';
            summary.teacherAwards.byLevel[level] = (summary.teacherAwards.byLevel[level] || 0) + 1;
        });
    }
    
    // คำนวณข้อมูลผลงานนักเรียน
    const studentSheet = ss.getSheetByName(STUDENT_WORK_SHEET_NAME);
    if (studentSheet && studentSheet.getLastRow() > 1) {
        const lastRow = studentSheet.getLastRow();
        const data = studentSheet.getRange(2, 1, lastRow - 1, 8).getValues();
        
        summary.studentWorks.total = data.length;
        
        data.forEach(row => {
            const workType = row[2] || 'ไม่ระบุ';
            const awardLevel = row[5] || 'ไม่ระบุ';
            
            summary.studentWorks.byWorkType[workType] = (summary.studentWorks.byWorkType[workType] || 0) + 1;
            summary.studentWorks.byAwardLevel[awardLevel] = (summary.studentWorks.byAwardLevel[awardLevel] || 0) + 1;
        });
    }
    
    // คำนวณข้อมูลบุคลากร
    try {
        const personnelNames = getPersonnelNames();
        summary.personnel.total = personnelNames.length;
        
        personnelNames.forEach(name => {
            if (name.startsWith('นาย')) {
                summary.personnel.byTitle.นาย = (summary.personnel.byTitle.นาย || 0) + 1;
            } else if (name.startsWith('นางสาว')) {
                summary.personnel.byTitle.นางสาว = (summary.personnel.byTitle.นางสาว || 0) + 1;
            } else if (name.startsWith('นาง')) {
                summary.personnel.byTitle.นาง = (summary.personnel.byTitle.นาง || 0) + 1;
            } else {
                summary.personnel.byTitle.อื่นๆ = (summary.personnel.byTitle.อื่นๆ || 0) + 1;
            }
        });
    } catch (e) {
        console.log('⚠️ Error calculating personnel summary:', e.toString());
    }
    
    return summary;
}

/**
 * 📈 เตรียมข้อมูลสำหรับกราฟ
 */
function prepareChartData(ss) {
    const charts = {
        trainingsByType: [],
        teacherAwardsByLevel: [],
        studentWorksByType: [],
        personnelByTitle: [],
        monthlyTrend: []
    };
    
    // ข้อมูลการอบรมแยกตามประเภท
    const trainingSheet = ss.getSheetByName(TRAINING_SHEET_NAME);
    if (trainingSheet && trainingSheet.getLastRow() > 1) {
        const lastRow = trainingSheet.getLastRow();
        const data = trainingSheet.getRange(2, 1, lastRow - 1, 11).getValues();
        
        const byType = {};
        const monthlyData = {};
        
        data.forEach(row => {
            const type = row[2] || 'ไม่ระบุ';
            byType[type] = (byType[type] || 0) + 1;
            
            const timestamp = row[0];
            if (timestamp instanceof Date) {
                const month = timestamp.getMonth() + 1;
                const year = timestamp.getFullYear() + 543;
                const key = `${year}-${month.toString().padStart(2, '0')}`;
                monthlyData[key] = (monthlyData[key] || 0) + 1;
            }
        });
        
        charts.trainingsByType = Object.entries(byType).map(([type, count]) => ({ type, count }));
        
        // เตรียมข้อมูลแนวโน้มรายเดือน
        const sortedMonths = Object.keys(monthlyData).sort();
        charts.monthlyTrend = sortedMonths.map(month => {
            const [year, monthNum] = month.split('-');
            return {
                month: `${monthNum}/${year}`,
                count: monthlyData[month]
            };
        });
    }
    
    // ข้อมูลผลงานครูแยกตามระดับรางวัล
    const teacherSheet = ss.getSheetByName(TEACHER_AWARD_SHEET_NAME);
    if (teacherSheet && teacherSheet.getLastRow() > 1) {
        const lastRow = teacherSheet.getLastRow();
        const data = teacherSheet.getRange(2, 1, lastRow - 1, 6).getValues();
        
        const byLevel = {};
        data.forEach(row => {
            const level = row[3] || 'ไม่ระบุ';
            byLevel[level] = (byLevel[level] || 0) + 1;
        });
        
        charts.teacherAwardsByLevel = Object.entries(byLevel).map(([level, count]) => ({ level, count }));
    }
    
    // ข้อมูลผลงานนักเรียน
    const studentSheet = ss.getSheetByName(STUDENT_WORK_SHEET_NAME);
    if (studentSheet && studentSheet.getLastRow() > 1) {
        const lastRow = studentSheet.getLastRow();
        const data = studentSheet.getRange(2, 1, lastRow - 1, 8).getValues();
        
        const byWorkType = {};
        data.forEach(row => {
            const workType = row[2] || 'ไม่ระบุ';
            byWorkType[workType] = (byWorkType[workType] || 0) + 1;
        });
        
        charts.studentWorksByType = Object.entries(byWorkType).map(([type, count]) => ({ type, count }));
    }
    
    // ข้อมูลบุคลากรแยกตามคำนำหน้า
    try {
        const personnelNames = getPersonnelNames();
        const byTitle = {
            'นาย': 0,
            'นาง': 0,
            'นางสาว': 0,
            'อื่นๆ': 0
        };
        
        personnelNames.forEach(name => {
            if (name.startsWith('นาย')) {
                byTitle.นาย++;
            } else if (name.startsWith('นางสาว')) {
                byTitle.นางสาว++;
            } else if (name.startsWith('นาง')) {
                byTitle.นาง++;
            } else {
                byTitle.อื่นๆ++;
            }
        });
        
        charts.personnelByTitle = Object.entries(byTitle)
            .filter(([_, count]) => count > 0)
            .map(([title, count]) => ({ title, count }));
    } catch (e) {
        console.log('⚠️ Error preparing personnel chart data:', e.toString());
    }
    
    return charts;
}

/**
 * 🔔 ดึงข้อมูลการอัปเดตล่าสุด
 */
function getLatestUpdates(ss, limit = 5) {
    const updates = [];
    
    try {
        const sheets = [
            { name: TRAINING_SHEET_NAME, type: 'training', label: 'การอบรม' },
            { name: PHOTO_SHEET_NAME, type: 'photo', label: 'รูปกิจกรรม' },
            { name: TEACHER_AWARD_SHEET_NAME, type: 'teacher_award', label: 'ผลงานครู' },
            { name: STUDENT_WORK_SHEET_NAME, type: 'student_work', label: 'ผลงานนักเรียน' }
        ];
        
        sheets.forEach(sheetConfig => {
            const sheet = ss.getSheetByName(sheetConfig.name);
            if (!sheet || sheet.getLastRow() <= 1) return;
            
            const lastRow = sheet.getLastRow();
            const data = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
            
            if (data && data[0]) {
                updates.push({
                    type: sheetConfig.type,
                    label: sheetConfig.label,
                    timestamp: data[0],
                    details: getUpdateDetails(sheetConfig.type, data)
                });
            }
        });
        
        // เรียงลำดับจากใหม่ไปเก่า
        updates.sort((a, b) => {
            const dateA = a.timestamp instanceof Date ? a.timestamp : new Date(a.timestamp);
            const dateB = b.timestamp instanceof Date ? b.timestamp : new Date(b.timestamp);
            return dateB - dateA;
        });
        
    } catch (error) {
        console.log('⚠️ Error getting latest updates:', error.toString());
    }
    
    return updates.slice(0, limit);
}

/**
 * 📝 แปลงข้อมูลการอัปเดต
 */
function getUpdateDetails(type, data) {
    switch (type) {
        case 'training':
            return `บันทึกการอบรม: ${data[1] || 'ไม่ระบุชื่อ'} - ${data[3] || 'ไม่ระบุหลักสูตร'}`;
        case 'photo':
            return `เพิ่มรูปกิจกรรม: ${data[1] || 'ไม่ระบุผู้เกี่ยวข้อง'}`;
        case 'teacher_award':
            return `บันทึกรางวัลครู: ${data[1] || 'ไม่ระบุชื่อครู'} - ${data[2] || 'ไม่ระบุรางวัล'}`;
        case 'student_work':
            return `บันทึกผลงานนักเรียน: ${data[1] || 'ไม่ระบุชื่อนักเรียน'} - ${data[3] || 'ไม่ระบุผลงาน'}`;
        default:
            return 'อัปเดตข้อมูลใหม่';
    }
}

/**
 * ✅ Action: ดึงข้อมูลสถิติแบบสรุป (แบบย่อ)
 */
function handleGetDashboardSummary(payload = {}) {
    try {
        console.log('📊 Handling getDashboardSummary request...');
        
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const result = {
            status: 'SUCCESS',
            timestamp: new Date().toISOString(),
            data: {}
        };
        
        // ดึงข้อมูลอย่างรวดเร็ว
        const sheets = [
            { name: PERSONNEL_SHEET_NAME, key: 'personnel' },
            { name: TRAINING_SHEET_NAME, key: 'trainings' },
            { name: PHOTO_SHEET_NAME, key: 'photos' },
            { name: TEACHER_AWARD_SHEET_NAME, key: 'teacherAwards' },
            { name: STUDENT_WORK_SHEET_NAME, key: 'studentWorks' }
        ];
        
        sheets.forEach(sheetConfig => {
            const sheet = ss.getSheetByName(sheetConfig.name);
            if (sheet) {
                const lastRow = sheet.getLastRow();
                const count = lastRow > 0 ? lastRow - 1 : 0;
                result.data[sheetConfig.key] = {
                    count: count,
                    hasData: count > 0,
                    lastUpdated: sheet.getRange(1, 1).getValue() || 'N/A'
                };
            } else {
                result.data[sheetConfig.key] = {
                    count: 0,
                    hasData: false,
                    exists: false
                };
            }
        });
        
        // คำนวณรวม
        result.data.totalRecords = Object.values(result.data).reduce((sum, item) => sum + (item.count || 0), 0);
        
        // ดึงข้อมูลบุคลากรเพิ่มเติม
        try {
            const personnelNames = getPersonnelNames();
            result.data.personnel.total = personnelNames.length;
            result.data.personnel.list = personnelNames;
        } catch (e) {
            result.data.personnel.total = 0;
        }
        
        return result;
        
    } catch (error) {
        console.log('❌ Error in handleGetDashboardSummary:', error.toString());
        return {
            status: 'ERROR',
            error: 'ไม่สามารถดึงข้อมูลสรุปได้: ' + error.toString()
        };
    }
}

/**
 * ✅ Action: ดึงข้อมูลสำหรับกราฟและแผนภูมิ
 */
function handleGetChartData(payload = {}) {
    try {
        console.log('📈 Handling getChartData request...');
        
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const result = {
            status: 'SUCCESS',
            timestamp: new Date().toISOString(),
            charts: {}
        };
        
        // 1. ข้อมูลการอบรมแยกตามประเภท
        const trainingSheet = ss.getSheetByName(TRAINING_SHEET_NAME);
        if (trainingSheet && trainingSheet.getLastRow() > 1) {
            const lastRow = trainingSheet.getLastRow();
            const data = trainingSheet.getRange(2, 1, lastRow - 1, 11).getValues();
            
            const byType = {};
            const byMonth = {};
            
            data.forEach(row => {
                const type = row[2] || 'ไม่ระบุ';
                byType[type] = (byType[type] || 0) + 1;
                
                const timestamp = row[0];
                if (timestamp instanceof Date) {
                    const month = timestamp.getMonth() + 1;
                    const year = timestamp.getFullYear() + 543;
                    const monthYear = `${month}/${year}`;
                    byMonth[monthYear] = (byMonth[monthYear] || 0) + 1;
                }
            });
            
            result.charts.trainingsByType = {
                labels: Object.keys(byType),
                data: Object.values(byType),
                total: data.length
            };
            
            result.charts.trainingsByMonth = {
                labels: Object.keys(byMonth),
                data: Object.values(byMonth)
            };
        }
        
        // 2. ข้อมูลผลงานครูแยกตามระดับรางวัล
        const teacherSheet = ss.getSheetByName(TEACHER_AWARD_SHEET_NAME);
        if (teacherSheet && teacherSheet.getLastRow() > 1) {
            const lastRow = teacherSheet.getLastRow();
            const data = teacherSheet.getRange(2, 1, lastRow - 1, 6).getValues();
            
            const byLevel = {};
            data.forEach(row => {
                const level = row[3] || 'ไม่ระบุ';
                byLevel[level] = (byLevel[level] || 0) + 1;
            });
            
            result.charts.teacherAwardsByLevel = {
                labels: Object.keys(byLevel),
                data: Object.values(byLevel),
                total: data.length
            };
        }
        
        // 3. ข้อมูลผลงานนักเรียน
        const studentSheet = ss.getSheetByName(STUDENT_WORK_SHEET_NAME);
        if (studentSheet && studentSheet.getLastRow() > 1) {
            const lastRow = studentSheet.getLastRow();
            const data = studentSheet.getRange(2, 1, lastRow - 1, 8).getValues();
            
            const byWorkType = {};
            const byAwardLevel = {};
            
            data.forEach(row => {
                const workType = row[2] || 'ไม่ระบุ';
                const awardLevel = row[5] || 'ไม่ระบุ';
                
                byWorkType[workType] = (byWorkType[workType] || 0) + 1;
                byAwardLevel[awardLevel] = (byAwardLevel[awardLevel] || 0) + 1;
            });
            
            result.charts.studentWorksByType = {
                labels: Object.keys(byWorkType),
                data: Object.values(byWorkType),
                total: data.length
            };
            
            result.charts.studentAwardsByLevel = {
                labels: Object.keys(byAwardLevel),
                data: Object.values(byAwardLevel)
            };
        }
        
        // 4. ข้อมูลบุคลากรแยกตามคำนำหน้า
        try {
            const personnelNames = getPersonnelNames();
            const byTitle = {
                'นาย': 0,
                'นาง': 0,
                'นางสาว': 0,
                'อื่นๆ': 0
            };
            
            personnelNames.forEach(name => {
                if (name.startsWith('นาย')) {
                    byTitle.นาย++;
                } else if (name.startsWith('นางสาว')) {
                    byTitle.นางสาว++;
                } else if (name.startsWith('นาง')) {
                    byTitle.นาง++;
                } else {
                    byTitle.อื่นๆ++;
                }
            });
            
            result.charts.personnelByTitle = {
                labels: Object.keys(byTitle),
                data: Object.values(byTitle),
                total: personnelNames.length
            };
        } catch (e) {
            console.log('⚠️ Error analyzing personnel:', e.toString());
        }
        
        return result;
        
    } catch (error) {
        console.log('❌ Error in handleGetChartData:', error.toString());
        return {
            status: 'ERROR',
            error: 'ไม่สามารถดึงข้อมูลกราฟได้: ' + error.toString()
        };
    }
}

/**
 * ✅ Action: ดึงข้อมูลสถิติแบบช่วงเวลา (สำหรับกราฟแนวโน้ม)
 */
function handleGetTrendData(payload = {}) {
    try {
        console.log('📈 Handling getTrendData request...');
        
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const monthsBack = payload.monthsBack || 12;
        
        const result = {
            status: 'SUCCESS',
            timestamp: new Date().toISOString(),
            trend: {
                trainings: [],
                teacherAwards: [],
                studentWorks: []
            }
        };
        
        // สร้างช่วงเวลาย้อนหลัง
        const months = [];
        const now = new Date();
        for (let i = monthsBack - 1; i >= 0; i--) {
            const date = new Date(now.getFullYear(), now.getMonth() - i, 1);
            const monthKey = `${date.getFullYear() + 543}-${(date.getMonth() + 1).toString().padStart(2, '0')}`;
            const monthLabel = `${date.getMonth() + 1}/${date.getFullYear() + 543}`;
            months.push({ key: monthKey, label: monthLabel, date: date });
        }
        
        // ดึงข้อมูลการอบรมแบบแยกตามเดือน
        const trainingSheet = ss.getSheetByName(TRAINING_SHEET_NAME);
        if (trainingSheet && trainingSheet.getLastRow() > 1) {
            const lastRow = trainingSheet.getLastRow();
            const data = trainingSheet.getRange(2, 1, lastRow - 1, 11).getValues();
            
            const monthlyCount = {};
            
            data.forEach(row => {
                const timestamp = row[0];
                if (timestamp instanceof Date) {
                    const month = timestamp.getMonth() + 1;
                    const year = timestamp.getFullYear() + 543;
                    const key = `${year}-${month.toString().padStart(2, '0')}`;
                    monthlyCount[key] = (monthlyCount[key] || 0) + 1;
                }
            });
            
            result.trend.trainings = months.map(month => ({
                month: month.label,
                count: monthlyCount[month.key] || 0
            }));
        }
        
        // ดึงข้อมูลผลงานครูแบบแยกตามเดือน
        const teacherSheet = ss.getSheetByName(TEACHER_AWARD_SHEET_NAME);
        if (teacherSheet && teacherSheet.getLastRow() > 1) {
            const lastRow = teacherSheet.getLastRow();
            const data = teacherSheet.getRange(2, 1, lastRow - 1, 6).getValues();
            
            const monthlyCount = {};
            
            data.forEach(row => {
                const timestamp = row[0];
                if (timestamp instanceof Date) {
                    const month = timestamp.getMonth() + 1;
                    const year = timestamp.getFullYear() + 543;
                    const key = `${year}-${month.toString().padStart(2, '0')}`;
                    monthlyCount[key] = (monthlyCount[key] || 0) + 1;
                }
            });
            
            result.trend.teacherAwards = months.map(month => ({
                month: month.label,
                count: monthlyCount[month.key] || 0
            }));
        }
        
        // ดึงข้อมูลผลงานนักเรียนแบบแยกตามเดือน
        const studentSheet = ss.getSheetByName(STUDENT_WORK_SHEET_NAME);
        if (studentSheet && studentSheet.getLastRow() > 1) {
            const lastRow = studentSheet.getLastRow();
            const data = studentSheet.getRange(2, 1, lastRow - 1, 8).getValues();
            
            const monthlyCount = {};
            
            data.forEach(row => {
                const timestamp = row[0];
                if (timestamp instanceof Date) {
                    const month = timestamp.getMonth() + 1;
                    const year = timestamp.getFullYear() + 543;
                    const key = `${year}-${month.toString().padStart(2, '0')}`;
                    monthlyCount[key] = (monthlyCount[key] || 0) + 1;
                }
            });
            
            result.trend.studentWorks = months.map(month => ({
                month: month.label,
                count: monthlyCount[month.key] || 0
            }));
        }
        
        return result;
        
    } catch (error) {
        console.log('❌ Error in handleGetTrendData:', error.toString());
        return {
            status: 'ERROR',
            error: 'ไม่สามารถดึงข้อมูลแนวโน้มได้: ' + error.toString()
        };
    }
}

/**
 * ✅ Action: ตรวจสอบข้อมูลใหม่ที่ยังไม่ได้ประมวลผล
 */
function handleGetUnprocessedData(payload = {}) {
    try {
        console.log('🆕 Handling getUnprocessedData request...');
        
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const hoursThreshold = payload.hoursThreshold || 24;
        
        const cutoffTime = new Date();
        cutoffTime.setHours(cutoffTime.getHours() - hoursThreshold);
        
        const result = {
            status: 'SUCCESS',
            timestamp: new Date().toISOString(),
            unprocessed: {
                trainings: [],
                photos: [],
                teacherAwards: [],
                studentWorks: []
            }
        };
        
        // ตรวจสอบข้อมูลการอบรมใหม่
        const trainingSheet = ss.getSheetByName(TRAINING_SHEET_NAME);
        if (trainingSheet && trainingSheet.getLastRow() > 1) {
            const lastRow = trainingSheet.getLastRow();
            const data = trainingSheet.getRange(2, 1, lastRow - 1, 11).getValues();
            
            data.forEach(row => {
                const timestamp = row[0];
                if (timestamp instanceof Date && timestamp >= cutoffTime) {
                    result.unprocessed.trainings.push({
                        timestamp: timestamp,
                        name: row[1],
                        course: row[3],
                        type: row[2]
                    });
                }
            });
        }
        
        // ตรวจสอบข้อมูลรูปกิจกรรมใหม่
        const photoSheet = ss.getSheetByName(PHOTO_SHEET_NAME);
        if (photoSheet && photoSheet.getLastRow() > 1) {
            const lastRow = photoSheet.getLastRow();
            const data = photoSheet.getRange(2, 1, lastRow - 1, 4).getValues();
            
            data.forEach(row => {
                const timestamp = row[0];
                if (timestamp instanceof Date && timestamp >= cutoffTime) {
                    result.unprocessed.photos.push({
                        timestamp: timestamp,
                        person: row[1],
                        description: row[3]
                    });
                }
            });
        }
        
        // ตรวจสอบข้อมูลผลงานครูใหม่
        const teacherSheet = ss.getSheetByName(TEACHER_AWARD_SHEET_NAME);
        if (teacherSheet && teacherSheet.getLastRow() > 1) {
            const lastRow = teacherSheet.getLastRow();
            const data = teacherSheet.getRange(2, 1, lastRow - 1, 6).getValues();
            
            data.forEach(row => {
                const timestamp = row[0];
                if (timestamp instanceof Date && timestamp >= cutoffTime) {
                    result.unprocessed.teacherAwards.push({
                        timestamp: timestamp,
                        teacherName: row[1],
                        awardName: row[2]
                    });
                }
            });
        }
        
        // ตรวจสอบข้อมูลผลงานนักเรียนใหม่
        const studentSheet = ss.getSheetByName(STUDENT_WORK_SHEET_NAME);
        if (studentSheet && studentSheet.getLastRow() > 1) {
            const lastRow = studentSheet.getLastRow();
            const data = studentSheet.getRange(2, 1, lastRow - 1, 8).getValues();
            
            data.forEach(row => {
                const timestamp = row[0];
                if (timestamp instanceof Date && timestamp >= cutoffTime) {
                    result.unprocessed.studentWorks.push({
                        timestamp: timestamp,
                        studentName: row[1],
                        projectName: row[3]
                    });
                }
            });
        }
        
        // สรุปจำนวน
        result.summary = {
            totalNew: result.unprocessed.trainings.length + 
                     result.unprocessed.photos.length + 
                     result.unprocessed.teacherAwards.length + 
                     result.unprocessed.studentWorks.length,
            period: `ล่าสุด ${hoursThreshold} ชั่วโมง`
        };
        
        return result;
        
    } catch (error) {
        console.log('❌ Error in handleGetUnprocessedData:', error.toString());
        return {
            status: 'ERROR',
            error: 'ไม่สามารถตรวจสอบข้อมูลใหม่ได้: ' + error.toString()
        };
    }
}

/**
 * ✅ Action: ค้นหาข้อมูลทั่วทั้งระบบ
 */
function handleGlobalSearch(payload) {
    try {
        if (!payload || !payload.query) {
            return {
                status: 'ERROR',
                error: 'กรุณาระบุคำค้นหา'
            };
        }
        
        console.log('🔍 Handling global search:', payload.query);
        
        const searchQuery = payload.query.trim().toLowerCase();
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const results = {
            trainings: [],
            teacherAwards: [],
            studentWorks: [],
            photos: [],
            personnel: []
        };
        
        // ค้นหาจากบุคลากร
        const personnelNames = getPersonnelNames();
        results.personnel = personnelNames.filter(name => 
            name.toLowerCase().includes(searchQuery)
        );
        
        // ค้นหาจากการอบรม
        const trainingSheet = ss.getSheetByName(TRAINING_SHEET_NAME);
        if (trainingSheet && trainingSheet.getLastRow() > 1) {
            const lastRow = trainingSheet.getLastRow();
            const data = trainingSheet.getRange(2, 1, lastRow - 1, 11).getValues();
            
            data.forEach((row, index) => {
                const rowText = row.join(' ').toLowerCase();
                if (rowText.includes(searchQuery)) {
                    results.trainings.push({
                        row: index + 2,
                        name: row[1],
                        course: row[3],
                        type: row[2],
                        date: row[0]
                    });
                }
            });
        }
        
        // ค้นหาจากผลงานครู
        const teacherSheet = ss.getSheetByName(TEACHER_AWARD_SHEET_NAME);
        if (teacherSheet && teacherSheet.getLastRow() > 1) {
            const lastRow = teacherSheet.getLastRow();
            const data = teacherSheet.getRange(2, 1, lastRow - 1, 6).getValues();
            
            data.forEach((row, index) => {
                const rowText = row.join(' ').toLowerCase();
                if (rowText.includes(searchQuery)) {
                    results.teacherAwards.push({
                        row: index + 2,
                        teacherName: row[1],
                        awardName: row[2],
                        awardLevel: row[3],
                        date: row[0]
                    });
                }
            });
        }
        
        // ค้นหาจากผลงานนักเรียน
        const studentSheet = ss.getSheetByName(STUDENT_WORK_SHEET_NAME);
        if (studentSheet && studentSheet.getLastRow() > 1) {
            const lastRow = studentSheet.getLastRow();
            const data = studentSheet.getRange(2, 1, lastRow - 1, 8).getValues();
            
            data.forEach((row, index) => {
                const rowText = row.join(' ').toLowerCase();
                if (rowText.includes(searchQuery)) {
                    results.studentWorks.push({
                        row: index + 2,
                        studentName: row[1],
                        projectName: row[3],
                        workType: row[2],
                        date: row[0]
                    });
                }
            });
        }
        
        // ค้นหาจากรูปกิจกรรม
        const photoSheet = ss.getSheetByName(PHOTO_SHEET_NAME);
        if (photoSheet && photoSheet.getLastRow() > 1) {
            const lastRow = photoSheet.getLastRow();
            const data = photoSheet.getRange(2, 1, lastRow - 1, 4).getValues();
            
            data.forEach((row, index) => {
                const rowText = row.join(' ').toLowerCase();
                if (rowText.includes(searchQuery)) {
                    results.photos.push({
                        row: index + 2,
                        person: row[1],
                        description: row[3],
                        date: row[0]
                    });
                }
            });
        }
        
        // สรุปผลลัพธ์
        const totalResults = 
            results.personnel.length + 
            results.trainings.length + 
            results.teacherAwards.length + 
            results.studentWorks.length + 
            results.photos.length;
        
        return {
            status: 'SUCCESS',
            message: `พบ ${totalResults} ผลลัพธ์จากการค้นหา "${payload.query}"`,
            query: payload.query,
            results: results,
            counts: {
                personnel: results.personnel.length,
                trainings: results.trainings.length,
                teacherAwards: results.teacherAwards.length,
                studentWorks: results.studentWorks.length,
                photos: results.photos.length,
                total: totalResults
            },
            timestamp: new Date().toISOString()
        };
        
    } catch (error) {
        console.log('❌ Error in handleGlobalSearch:', error.toString());
        return {
            status: 'ERROR',
            error: 'เกิดข้อผิดพลาดในการค้นหา: ' + error.toString()
        };
    }
}

// ==============================================
// 👥 PERSONNEL FUNCTIONS
// ==============================================

/**
 * ✅ Action: ดึงรายชื่อบุคลากรทั้งหมด
 */
function handleGetPersonnelList(payload = {}) {
    try {
        console.log('📋 Handling getPersonnelList request...');
        
        const names = getPersonnelNames();
        
        if (names.length === 0) {
            return {
                status: 'SUCCESS',
                message: 'ไม่มีรายชื่อบุคลากรในระบบ กรุณาเพิ่มชื่อใน Sheet "' + PERSONNEL_SHEET_NAME + '"',
                personnelList: [],
                count: 0,
                hasData: false,
                sheetExists: true
            };
        }
        
        // ถ้ามีการระบุว่าให้กรอง
        let filteredNames = names;
        if (payload.search && payload.search.trim() !== '') {
            const searchTerm = payload.search.trim().toLowerCase();
            filteredNames = names.filter(name => 
                name.toLowerCase().includes(searchTerm)
            );
        }
        
        // ถ้ามีการระบุจำนวนที่ต้องการ
        let finalNames = filteredNames;
        if (payload.limit && payload.limit > 0) {
            finalNames = filteredNames.slice(0, payload.limit);
        }
        
        // ถ้าต้องการเรียงลำดับ
        if (payload.sort && payload.sort === 'asc') {
            finalNames.sort();
        } else if (payload.sort && payload.sort === 'desc') {
            finalNames.sort().reverse();
        }
        
        return {
            status: 'SUCCESS',
            message: `พบรายชื่อบุคลากร ${filteredNames.length} รายชื่อ (แสดง ${finalNames.length} รายการ)`,
            personnelList: finalNames,
            filteredList: filteredNames,
            allNames: names,
            count: names.length,
            filteredCount: filteredNames.length,
            displayedCount: finalNames.length,
            hasData: true,
            timestamp: new Date().toISOString()
        };
        
    } catch (error) {
        console.log('❌ Error in handleGetPersonnelList:', error.toString());
        return {
            status: 'ERROR',
            error: 'ไม่สามารถดึงรายชื่อบุคลากรได้: ' + error.toString(),
            suggestion: 'ตรวจสอบว่า Sheet "' + PERSONNEL_SHEET_NAME + '" มีอยู่และมีข้อมูล'
        };
    }
}

/**
 * ✅ Action: ดึงรายชื่อครูเฉพาะ
 */
function handleGetTeacherList(payload = {}) {
    try {
        console.log('👨‍🏫 Handling getTeacherList request...');
        
        const allNames = getPersonnelNames();
        
        // กรองเฉพาะครู
        const teacherNames = allNames.filter(name => {
            const trimmedName = name.trim();
            return (
                trimmedName.startsWith('นาย') || 
                trimmedName.startsWith('นาง') || 
                trimmedName.startsWith('นางสาว')
            );
        });
        
        if (teacherNames.length === 0) {
            return {
                status: 'SUCCESS',
                message: 'ไม่มีรายชื่อครูในระบบ กรุณาเพิ่มชื่อครูใน Sheet "' + PERSONNEL_SHEET_NAME + '"',
                teacherList: [],
                count: 0,
                hasData: false
            };
        }
        
        // กรองตามคำค้นหา
        let filteredTeachers = teacherNames;
        if (payload.search && payload.search.trim() !== '') {
            const searchTerm = payload.search.trim().toLowerCase();
            filteredTeachers = teacherNames.filter(name => 
                name.toLowerCase().includes(searchTerm)
            );
        }
        
        return {
            status: 'SUCCESS',
            message: `พบรายชื่อครู ${filteredTeachers.length} รายชื่อ`,
            teacherList: filteredTeachers,
            allTeachers: teacherNames,
            count: teacherNames.length,
            filteredCount: filteredTeachers.length,
            hasData: true
        };
        
    } catch (error) {
        console.log('❌ Error in handleGetTeacherList:', error.toString());
        return {
            status: 'ERROR',
            error: 'ไม่สามารถดึงรายชื่อครูได้: ' + error.toString()
        };
    }
}

/**
 * ✅ Action: ตรวจสอบสถานะ Sheet รายชื่อบุคลากร
 */
function handleCheckPersonnelSheet() {
    try {
        console.log('🔍 Checking personnel sheet...');
        
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const personnelSheet = ss.getSheetByName(PERSONNEL_SHEET_NAME);
        
        if (!personnelSheet) {
            return {
                status: 'WARNING',
                message: 'ไม่พบ Sheet "' + PERSONNEL_SHEET_NAME + '"',
                sheetExists: false,
                suggestion: 'กรุณาสร้าง Sheet "' + PERSONNEL_SHEET_NAME + '" และเพิ่มรายชื่อบุคลากร'
            };
        }
        
        const lastRow = personnelSheet.getLastRow();
        const dataRange = personnelSheet.getDataRange();
        const data = dataRange.getValues();
        
        // วิเคราะห์ข้อมูล
        const validNames = [];
        const invalidRows = [];
        
        for (let i = 0; i < data.length; i++) {
            const cellValue = data[i][0];
            if (cellValue && cellValue.toString().trim() !== '') {
                validNames.push(cellValue.toString().trim());
            } else {
                invalidRows.push(i + 1);
            }
        }
        
        // วิเคราะห์รูปแบบชื่อ
        const nameAnalysis = {
            startsWithนาย: validNames.filter(name => name.startsWith('นาย')).length,
            startsWithนาง: validNames.filter(name => name.startsWith('นาง') && !name.startsWith('นางสาว')).length,
            startsWithนางสาว: validNames.filter(name => name.startsWith('นางสาว')).length,
            otherFormats: validNames.filter(name => 
                !name.startsWith('นาย') && 
                !name.startsWith('นาง') && 
                !name.startsWith('นางสาว')
            ).length
        };
        
        return {
            status: 'SUCCESS',
            message: `Sheet "${PERSONNEL_SHEET_NAME}" พบ ${validNames.length} รายชื่อ`,
            sheetExists: true,
            sheetName: PERSONNEL_SHEET_NAME,
            totalRows: lastRow,
            validRows: validNames.length,
            invalidRows: invalidRows.length,
            validNamesCount: validNames.length,
            sampleNames: validNames.slice(0, 10),
            nameAnalysis: nameAnalysis,
            hasData: validNames.length > 0,
            lastUpdated: personnelSheet.getRange(1, 1).getValue()
        };
        
    } catch (error) {
        console.log('❌ Error in handleCheckPersonnelSheet:', error.toString());
        return {
            status: 'ERROR',
            error: 'ไม่สามารถตรวจสอบ Sheet รายชื่อได้: ' + error.toString()
        };
    }
}

/**
 * ✅ Action: ค้นหารายชื่อบุคลากร
 */
function handleSearchPersonnel(payload) {
    try {
        if (!payload || !payload.query) {
            return {
                status: 'ERROR',
                error: 'กรุณาระบุคำค้นหา'
            };
        }
        
        const searchQuery = payload.query.trim().toLowerCase();
        console.log('🔎 Searching personnel for:', searchQuery);
        
        const allNames = getPersonnelNames();
        
        if (allNames.length === 0) {
            return {
                status: 'SUCCESS',
                message: 'ไม่มีรายชื่อบุคลากรในระบบ',
                results: [],
                count: 0
            };
        }
        
        // ค้นหาชื่อที่ตรงกับคำค้นหา
        const searchResults = allNames.filter(name => 
            name.toLowerCase().includes(searchQuery)
        );
        
        // ค้นหาชื่อที่ขึ้นต้นด้วยคำค้นหา
        const startsWithResults = allNames.filter(name => 
            name.toLowerCase().startsWith(searchQuery)
        );
        
        // รวมผลลัพธ์
        const combinedResults = [...startsWithResults, ...searchResults.filter(name => !startsWithResults.includes(name))];
        
        // จำกัดจำนวนผลลัพธ์
        const limit = payload.limit || null;
        const finalResults = limit ? combinedResults.slice(0, limit) : combinedResults;
        
        return {
            status: 'SUCCESS',
            message: `พบ ${finalResults.length} รายชื่อที่ตรงกับ "${payload.query}"`,
            results: finalResults,
            count: finalResults.length,
            searchQuery: payload.query,
            startsWithCount: startsWithResults.length,
            containsCount: searchResults.length
        };
        
    } catch (error) {
        console.log('❌ Error in handleSearchPersonnel:', error.toString());
        return {
            status: 'ERROR',
            error: 'เกิดข้อผิดพลาดในการค้นหา: ' + error.toString()
        };
    }
}

// ==============================================
// 📸 ACTION LOGIC: แยกบันทึกรูปกิจกรรม
// ==============================================

/**
 * บันทึก Link รูปกิจกรรมลงใน Sheet 'รูปกิจกรรม'
 */
function saveActivityPhotos(fullName, photoLinks, ss, submissionId = null) {
    try {
        const photoSheet = ss.getSheetByName(PHOTO_SHEET_NAME);
        if (!photoSheet) {
            const newSheet = ss.insertSheet(PHOTO_SHEET_NAME);
            newSheet.getRange(1, 1, 1, 4).setValues([['Timestamp', 'ชื่อ-นามสกุล', 'Photo Link', 'Description']]);
            return;
        }

        const photoRows = [];
        const now = new Date();
        const timestamp = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "MM/dd/yyyy HH:mm:ss"); 

        photoLinks.forEach((link, index) => {
            photoRows.push([
                timestamp,
                fullName,
                link,
                `รูปกิจกรรม ${index + 1}`
            ]); 
        });

        // บันทึกทั้งหมดลงใน Sheet 'รูปกิจกรรม'
        if (photoRows.length > 0) {
            const lastRow = photoSheet.getLastRow();
            let startRow = lastRow + 1;
            
            if (lastRow === 0) {
                photoSheet.getRange(1, 1, 1, 4).setValues([[
                    'Timestamp', 
                    'ชื่อ-นามสกุล', 
                    'Photo Link', 
                    'Description'
                ]]);
                startRow = 2;
            }
            
            photoSheet.getRange(startRow, 1, photoRows.length, photoRows[0].length).setValues(photoRows);
        }
    } catch (error) {
        console.log("Error saving activity photos: " + error.toString());
        throw error;
    }
}

// ==============================================
// 🎓 ACTION LOGIC: บันทึกข้อมูลการอบรม
// ==============================================

/**
 * ฟังก์ชัน: บันทึกข้อมูลการอบรมลงใน Google Sheet
 */
function recordTraining(payload) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(TRAINING_SHEET_NAME);
        if (!sheet) {
            return { 
                status: 'ERROR', 
                error: 'Sheet "' + TRAINING_SHEET_NAME + '" not found. Please create it.' 
            };
        }
        
        // สร้าง Submission ID
        const submissionId = generateUniqueId(); 
        
        const now = new Date();
        const timestamp = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "MM/dd/yyyy HH:mm:ss"); 
        
        const folder = DriveApp.getFolderById(FOLDER_ID); 
        const rows = [];
        
        // 1. จัดการรูปกิจกรรม
        const photoLinks = [];
        
        if (!payload.activityPhotos || !Array.isArray(payload.activityPhotos)) {
            return { 
                status: 'ERROR', 
                error: 'No activity photos provided or invalid format' 
            };
        }
        
        payload.activityPhotos.forEach((photo, index) => {
            try {
                if (!photo || !photo.base64Data) {
                    console.log(`Photo ${index + 1}: Missing base64 data`);
                    photoLinks.push(`❌ MISSING DATA`);
                    return;
                }
                
                const file = uploadFile(
                    photo.base64Data, 
                    photo.fileName || `${payload.fullName}_activity_photo_${submissionId}_${index + 1}.png`, 
                    photo.mimeType || 'image/png', 
                    folder
                );
                photoLinks.push(file.getUrl());
                console.log(`Successfully uploaded activity photo ${index + 1}: ${file.getName()}`);
            } catch (error) {
                console.log(`Error uploading photo ${index + 1}: ${error}`);
                photoLinks.push(`❌ UPLOAD ERROR: ${error.toString().substring(0, 50)}`); 
            }
        });

        // 2. บันทึก Link รูปกิจกรรมลงใน Sheet 'รูปกิจกรรม'
        try {
            saveActivityPhotos(payload.fullName, photoLinks, ss, submissionId);
        } catch (e) {
            console.log(`Error saving activity photos to separate sheet: ${e.toString()}`);
            photoLinks.unshift("⚠️ Note: Failed to save to photo sheet");
        }
        
        // 3. วนลูปบันทึกแต่ละหลักสูตร
        if (!payload.courses || !Array.isArray(payload.courses)) {
            return { 
                status: 'ERROR', 
                error: 'No course data provided or invalid format' 
            };
        }
        
        payload.courses.forEach((course, courseIndex) => {
            let certificateValue = 'ไม่มี';
            let linkFormula = '';
            
            // จัดการไฟล์เกียรติบัตร
            if (course.certificateFile && course.certificateFile.base64Data) {
                try {
                    const file = uploadFile(
                        course.certificateFile.base64Data, 
                        course.certificateFile.fileName || `${payload.fullName}_certificate_${submissionId}_${courseIndex + 1}`, 
                        course.certificateFile.mimeType || 'application/octet-stream', 
                        folder
                    );
                    
                    let displayName = course.certificateFile.fileName || 'เกียรติบัตร';
                    if (course.certificateFile.mimeType === 'application/pdf') {
                        displayName = displayName.endsWith('.pdf') ? displayName : displayName + '.pdf';
                    } else if (course.certificateFile.mimeType.includes('image/')) {
                        displayName = displayName.endsWith('.jpg') || displayName.endsWith('.png') || displayName.endsWith('.jpeg') 
                            ? displayName 
                            : displayName + '.jpg';
                    }
                    
                    linkFormula = `=HYPERLINK("${file.getUrl()}", "${displayName}")`;
                    certificateValue = displayName;
                    
                    console.log(`Successfully uploaded certificate for course ${courseIndex + 1}: ${displayName}`);
                } catch (error) {
                    console.log(`Error uploading certificate for course ${courseIndex + 1}: ${error}`);
                    linkFormula = `❌ UPLOAD ERROR`;
                    certificateValue = 'ERROR';
                }
            } else if (course.hasCertificate) {
                certificateValue = 'มี (ไม่มีไฟล์)';
            }

            // ตรวจสอบข้อมูลที่จำเป็น
            if (!course.courseName || !course.startDate || !course.endDate) {
                console.log(`Course ${courseIndex + 1}: Missing required fields`);
                return;
            }

            // 4. เตรียมแถวข้อมูล
            const row = [
                timestamp,
                payload.fullName,
                course.type || 'อบรม',
                course.courseName,
                course.location || '',
                course.startDate,
                course.endDate,
                course.hours || '0',
                certificateValue,
                linkFormula,
                submissionId
            ];
            
            rows.push(row);
        });
        
        // 5. บันทึกข้อมูลหลัก
        if (rows.length > 0) {
            const lastRow = sheet.getLastRow();
            let startRow = lastRow + 1;
            
            if (lastRow === 0) {
                sheet.getRange(1, 1, 1, 11).setValues([[
                    'วัน-เวลาบันทึก', 'ชื่อ-นามสกุล', 'ประเภท', 'ชื่อหลักสูตร/กิจกรรม', 
                    'สถานที่', 'วันที่เริ่ม', 'วันที่สิ้นสุด', 'จำนวนชั่วโมง', 
                    'ชื่อไฟล์เกียรติบัตร', 'Link เกียรติบัตร', 'Submission ID'
                ]]);
                startRow = 2;
            }
            
            sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
            
            return { 
                status: 'SUCCESS', 
                message: `บันทึกข้อมูลการอบรม/กิจกรรม ${rows.length} รายการสำเร็จ`,
                submissionId: submissionId,
                rows: rows.length
            };
        } else {
            return { 
                status: 'ERROR', 
                error: 'No valid course data to save' 
            };
        }
    } catch (error) {
        console.log("Error in recordTraining: " + error.toString());
        return { 
            status: 'ERROR', 
            error: `Server Error: ${error.toString()}` 
        };
    }
}

// ==============================================
// 🏆 ACTION LOGIC: บันทึกผลงานครู
// ==============================================

/**
 * ฟังก์ชัน: บันทึกผลงานครูลงใน Google Sheet
 */
function recordTeacherAward(payload) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(TEACHER_AWARD_SHEET_NAME); 
        if (!sheet) {
            return { 
                status: 'ERROR', 
                error: 'Sheet "' + TEACHER_AWARD_SHEET_NAME + '" not found. Please create it.' 
            };
        }
        
        const now = new Date();
        const timestamp = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "MM/dd/yyyy HH:mm:ss"); 
        
        const folder = DriveApp.getFolderById(FOLDER_ID); 
        let linkFormula = '';

        // 1. จัดการการอัปโหลดไฟล์เกียรติบัตร
        if (payload.certificate && payload.certificate.base64Data) {
            try {
                const file = uploadFile(
                    payload.certificate.base64Data, 
                    payload.certificate.fileName, 
                    payload.certificate.mimeType, 
                    folder
                );
                
                const fileDisplay = payload.certificate.fileName;
                linkFormula = `=HYPERLINK("${file.getUrl()}", "${fileDisplay}")`;
            } catch (error) {
                console.log(`Error uploading teacher award certificate: ${error}`);
                linkFormula = `❌ UPLOAD ERROR`;
            }
        }
        
        // Column Mapping
        const row = [
            timestamp,                   
            payload.teacherName,         
            payload.awardName,           
            payload.awardLevel,          
            payload.awardDate,           
            linkFormula                  
        ];
        
        // 3. บันทึกข้อมูล
        const lastRow = sheet.getLastRow();
        let startRow = lastRow + 1;
        
        if (lastRow === 0) {
            sheet.getRange(1, 1, 1, 6).setValues([[
                'วัน-เวลาบันทึก', 'ชื่อครู', 'ชื่อรางวัล', 'ระดับรางวัล', 'วันที่รับรางวัล', 'Link เกียรติบัตร'
            ]]);
            startRow = 2;
        }
        
        sheet.getRange(startRow, 1, 1, row.length).setValues([row]);
        
        return { 
            status: 'SUCCESS', 
            message: `บันทึกข้อมูลผลงานครู ${payload.teacherName} สำเร็จ`,
            rows: 1
        };
    } catch (error) {
        console.log("Error in recordTeacherAward: " + error.toString());
        return { 
            status: 'ERROR', 
            error: `Error recording Teacher Award: ${error.toString()}` 
        };
    }
}

// ==============================================
// 👨‍🎓 ACTION LOGIC: บันทึกผลงานนักเรียน
// ==============================================

/**
 * ฟังก์ชัน: บันทึกผลงานนักเรียนลงใน Google Sheet
 */
function recordStudentWork(payload) {
    try {
        console.log('🎓 Starting recordStudentWork...');
        
        // ตรวจสอบข้อมูลที่จำเป็น
        if (!payload.studentName || !payload.studentName.trim()) {
            return { 
                status: 'ERROR', 
                error: 'กรุณากรอกชื่อนักเรียน',
                missingField: 'studentName'
            };
        }
        
        if (!payload.workType || !payload.workType.trim()) {
            return { 
                status: 'ERROR', 
                error: 'กรุณากรอกประเภทผลงาน',
                missingField: 'workType'
            };
        }
        
        if (!payload.projectName || !payload.projectName.trim()) {
            return { 
                status: 'ERROR', 
                error: 'กรุณากรอกชื่อผลงาน',
                missingField: 'projectName'
            };
        }
        
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(STUDENT_WORK_SHEET_NAME);
        
        if (!sheet) {
            console.log('❌ Sheet not found:', STUDENT_WORK_SHEET_NAME);
            return { 
                status: 'ERROR', 
                error: 'Sheet "' + STUDENT_WORK_SHEET_NAME + '" not found. Please run initializeSheets() first.' 
            };
        }
        
        const now = new Date();
        const timestamp = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "MM/dd/yyyy HH:mm:ss"); 
        
        console.log('📅 Timestamp:', timestamp);
        console.log('👤 Student:', payload.studentName);
        console.log('📝 Work Type:', payload.workType);
        console.log('🏆 Project:', payload.projectName);
        
        const folder = DriveApp.getFolderById(FOLDER_ID); 
        let linkFormula = '';

        // 1. จัดการการอัปโหลดไฟล์เกียรติบัตร
        if (payload.certificate && payload.certificate.base64Data) {
            try {
                console.log('📄 Processing certificate upload...');
                
                const file = uploadFile(
                    payload.certificate.base64Data, 
                    payload.certificate.fileName, 
                    payload.certificate.mimeType, 
                    folder,
                    payload.isChunked || false
                );
                
                const fileDisplay = payload.certificate.fileName;
                linkFormula = `=HYPERLINK("${file.getUrl()}", "${fileDisplay}")`;
                console.log('✅ Certificate uploaded:', file.getUrl());
                
            } catch (error) {
                console.log(`❌ Error uploading certificate: ${error.toString()}`);
                linkFormula = `❌ UPLOAD ERROR: ${error.toString().substring(0, 50)}`;
            }
        } else {
            console.log('📄 No certificate provided or no base64 data');
        }
        
        // ✅ Column Mapping
        const row = [
            timestamp,
            payload.studentName.trim(),
            payload.workType.trim(),
            payload.projectName.trim(),
            (payload.advisorName || '').trim(),
            (payload.awardLevel || '').trim(),
            (payload.awardDate || '').trim(),
            linkFormula
        ];
        
        console.log('📝 Prepared row:', row);
        
        // 3. บันทึกข้อมูล
        const lastRow = sheet.getLastRow();
        let startRow = lastRow + 1;
        
        if (lastRow === 0) {
            console.log('📋 Creating header row...');
            sheet.getRange(1, 1, 1, 8).setValues([[
                'วัน-เวลาบันทึก', 
                'ชื่อนักเรียน', 
                'ประเภทผลงาน', 
                'ชื่อผลงาน', 
                'ชื่อที่ปรึกษา', 
                'ระดับรางวัล', 
                'วันที่รับรางวัล', 
                'Link เกียรติบัตร'
            ]]);
            
            sheet.getRange(1, 1, 1, 8)
                 .setFontWeight('bold')
                 .setBackground('#f0f8ff')
                 .setBorder(true, true, true, true, true, true);
            
            startRow = 2;
            console.log('✅ Header created');
        }
        
        console.log(`📊 Current last row: ${lastRow}, will write to row: ${startRow}`);
        
        // เขียนข้อมูล
        sheet.getRange(startRow, 1, 1, row.length).setValues([row]);
        
        // Format ข้อมูลใหม่
        sheet.getRange(startRow, 1, 1, 8)
             .setBorder(false, false, true, false, false, false);
        
        console.log(`✅ Data written to row ${startRow}`);
        
        return { 
            status: 'SUCCESS', 
            message: `บันทึกข้อมูลผลงานนักเรียน "${payload.studentName}" สำเร็จ`,
            details: {
                row: startRow,
                workType: payload.workType,
                projectName: payload.projectName,
                awardDate: payload.awardDate
            },
            rows: 1
        };
        
    } catch (error) {
        console.log("❌ Error in recordStudentWork: " + error.toString());
        console.log("📝 Stack trace: " + error.stack);
        
        return { 
            status: 'ERROR', 
            error: `Error recording Student Work: ${error.toString()}`,
            suggestion: 'ตรวจสอบว่า Sheet "ข้อมูลผลงานนักเรียน" มีอยู่และสามารถเขียนได้'
        };
    }
}

// ==============================================
// 📊 REPORT LOGIC
// ==============================================

/**
 * ฟังก์ชัน: ดึงข้อมูลสรุปสำหรับ Dashboard หรือ Reports
 */
function getReportsData(payload) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const result = {
            status: 'SUCCESS',
            data: {
                summary: {},
                recentTrainings: [],
                charts: {},
                personnel: {}
            }
        };
        
        // ✅ 1. ดึงข้อมูลบุคลากร
        try {
            const personnelNames = getPersonnelNames();
            result.data.personnel = {
                total: personnelNames.length,
                sample: personnelNames.slice(0, 10),
                hasData: personnelNames.length > 0
            };
        } catch (e) {
            console.log('⚠️ Could not fetch personnel data:', e.toString());
            result.data.personnel = {
                total: 0,
                sample: [],
                hasData: false,
                error: e.toString()
            };
        }
        
        // 2. ดึงข้อมูลการอบรม
        const trainingSheet = ss.getSheetByName(TRAINING_SHEET_NAME);
        if (trainingSheet && trainingSheet.getLastRow() > 1) {
            const trainingData = trainingSheet.getRange(2, 1, trainingSheet.getLastRow() - 1, 11).getValues();
            result.data.summary.totalTrainings = trainingData.length;
            
            // นับตามประเภท
            const typeCount = {};
            trainingData.forEach(row => {
                const type = row[2];
                typeCount[type] = (typeCount[type] || 0) + 1;
            });
            result.data.summary.byType = typeCount;
            
            // ข้อมูลล่าสุด
            result.data.recentTrainings = trainingData.slice(-5).map(row => ({
                date: row[0],
                name: row[1],
                course: row[3],
                hours: row[7]
            }));
        } else {
            result.data.summary.totalTrainings = 0;
        }
        
        // 3. ดึงข้อมูลรูปกิจกรรม
        const photoSheet = ss.getSheetByName(PHOTO_SHEET_NAME);
        if (photoSheet && photoSheet.getLastRow() > 1) {
            const photoData = photoSheet.getRange(2, 1, photoSheet.getLastRow() - 1, 4).getValues();
            result.data.summary.totalPhotos = photoData.length;
        } else {
            result.data.summary.totalPhotos = 0;
        }
        
        return result;
    } catch (error) {
        console.log("Error in getReportsData: " + error.toString());
        return { 
            status: 'ERROR', 
            error: `Error fetching report data: ${error.toString()}` 
        };
    }
}

// ==============================================
// 🧪 TESTING FUNCTIONS
// ==============================================

/**
 * ฟังก์ชันสำหรับทดสอบการเขียนข้อมูล
 */
function testWriteToSheet() {
    try {
        console.log('🧪 Starting testWriteToSheet...');
        
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        
        let testSheet = ss.getSheetByName('TEST_SHEET');
        if (!testSheet) {
            testSheet = ss.insertSheet('TEST_SHEET');
            testSheet.getRange(1, 1, 1, 2).setValues([['Timestamp', 'Message']]);
            testSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
        }
        
        const now = new Date();
        const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        
        const testData = [
            [timestamp, '✅ Test successful! GAS can write to sheet.']
        ];
        
        const lastRow = testSheet.getLastRow();
        testSheet.getRange(lastRow + 1, 1, testData.length, testData[0].length).setValues(testData);
        
        console.log(`✅ Test data written to row ${lastRow + 1}`);
        
        return {
            status: 'SUCCESS',
            message: 'Test write successful! Check TEST_SHEET in your Google Sheet.',
            timestamp: timestamp,
            rowWritten: lastRow + 1
        };
        
    } catch (error) {
        console.log(`❌ Test write failed: ${error.toString()}`);
        return {
            status: 'ERROR',
            error: error.toString(),
            message: 'Test write failed'
        };
    }
}

/**
 * ฟังก์ชันสำหรับทดสอบการเชื่อมต่อและตรวจสอบสภาพแวดล้อม
 */
function testEnvironment() {
    try {
        const result = {
            status: 'SUCCESS',
            checks: {}
        };
        
        // ตรวจสอบ Folder
        try {
            const folder = DriveApp.getFolderById(FOLDER_ID);
            result.checks.folder = {
                exists: true,
                name: folder.getName(),
                id: folder.getId()
            };
        } catch (e) {
            result.checks.folder = {
                exists: false,
                error: e.toString()
            };
        }
        
        // ตรวจสอบ Sheets
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const requiredSheets = [
            TRAINING_SHEET_NAME, 
            PHOTO_SHEET_NAME, 
            TEACHER_AWARD_SHEET_NAME, 
            STUDENT_WORK_SHEET_NAME,
            PERSONNEL_SHEET_NAME
        ];
        
        result.checks.sheets = {};
        requiredSheets.forEach(sheetName => {
            const sheet = ss.getSheetByName(sheetName);
            const exists = !!sheet;
            const rowCount = sheet ? sheet.getLastRow() : 0;
            
            result.checks.sheets[sheetName] = {
                exists: exists,
                rowCount: rowCount,
                hasData: rowCount > 0
            };
        });
        
        // ตรวจสอบรายชื่อบุคลากร
        try {
            const personnelNames = getPersonnelNames();
            result.checks.personnel = {
                count: personnelNames.length,
                sample: personnelNames.slice(0, 3),
                message: personnelNames.length > 0 ? 
                    `พบ ${personnelNames.length} รายชื่อ` : 
                    'ไม่มีรายชื่อบุคลากร'
            };
        } catch (e) {
            result.checks.personnel = {
                count: 0,
                error: e.toString()
            };
        }
        
        // ตรวจสอบ Web App URL
        result.checks.webAppUrl = ScriptApp.getService().getUrl();
        
        return result;
    } catch (error) {
        return {
            status: 'ERROR',
            error: error.toString()
        };
    }
}

/**
 * ฟังก์ชันสำหรับทดสอบการบันทึกผลงานนักเรียน
 */
function testStudentWorkRecord() {
    try {
        console.log('🧪 Starting testStudentWorkRecord...');
        
        const testPayload = {
            action: 'recordStudentWork',
            studentName: 'ทดสอบ นักเรียน',
            workType: 'ผลงานวิชาการ',
            projectName: 'โครงงานวิทยาศาสตร์ทดสอบ',
            advisorName: 'ครูที่ปรึกษาทดสอบ',
            awardLevel: 'ระดับโรงเรียน',
            awardDate: '2025-01-15'
        };
        
        console.log('📤 Test payload:', testPayload);
        
        const result = recordStudentWork(testPayload);
        
        console.log('📥 Test result:', result);
        
        return {
            status: 'SUCCESS',
            message: 'Test completed',
            testResult: result
        };
        
    } catch (error) {
        console.log('❌ Test failed:', error.toString());
        return {
            status: 'ERROR',
            error: error.toString()
        };
    }
}

/**
 * 🔍 ทดสอบการดึงข้อมูลจริงจาก Sheet "ข้อมูลการอบรม"
 */
function testTrainingDataCount() {
  try {
    console.log('🧪 Starting testTrainingDataCount...');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = 'ข้อมูลการอบรม';
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      console.log('❌ ไม่พบ Sheet:', sheetName);
      return {
        status: 'ERROR',
        message: 'ไม่พบ Sheet "ข้อมูลการอบรม"',
        sheetExists: false
      };
    }
    
    const lastRow = sheet.getLastRow();
    console.log('📊 Last Row in sheet:', lastRow);
    
    if (lastRow <= 1) {
      console.log('ℹ️ มีแต่ header หรือไม่มีข้อมูล');
      return {
        status: 'SUCCESS',
        message: 'Sheet "ข้อมูลการอบรม" มีแต่ header หรือไม่มีข้อมูล',
        sheetExists: true,
        lastRow: lastRow,
        dataCount: 0
      };
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    console.log('📥 Dashboard ดึงข้อมูลมา:', data.length, 'แถว');
    
    let actualDataCount = 0;
    const actualRows = [];
    
    data.forEach((row, index) => {
      const rowNumber = index + 2;
      const timestamp = row[0];
      const hasData = timestamp && 
                     (timestamp instanceof Date || 
                      (typeof timestamp === 'string' && timestamp.trim() !== ''));
      
      if (hasData) {
        actualDataCount++;
        actualRows.push({
          row: rowNumber,
          timestamp: timestamp,
          name: row[1] || 'ไม่มีชื่อ',
          course: row[3] || 'ไม่มีชื่อหลักสูตร'
        });
      }
    });
    
    console.log('🎯 ผลลัพธ์:');
    console.log('- แถวทั้งหมดที่ดึงมา:', data.length);
    console.log('- แถวที่มีข้อมูลจริง (actual):', actualDataCount);
    console.log('- แถวว่าง:', data.length - actualDataCount);
    
    console.log('📝 ตัวอย่างข้อมูล 5 แถวแรก:');
    actualRows.slice(0, 5).forEach(item => {
      console.log(`  แถว ${item.row}: ${item.name} - ${item.course}`);
    });
    
    return {
      status: 'SUCCESS',
      message: `พบข้อมูลใน Sheet "ข้อมูลการอบรม": ${actualDataCount} แถว`,
      sheetExists: true,
      sheetName: sheetName,
      lastRow: lastRow,
      rowsFetchedByDashboard: data.length,
      actualDataCount: actualDataCount,
      emptyRows: data.length - actualDataCount,
      sampleData: actualRows.slice(0, 5),
      explanation: {
        dashboardShows: 'Dashboard จะแสดง: ' + actualDataCount + ' รายการ',
        note: 'ถ้า Dashboard แสดง 45 แต่ที่นับได้ ' + actualDataCount + ' แสดงว่ามีปัญหาการนับ'
      }
    };
    
  } catch (error) {
    console.log('❌ Error:', error.toString());
    return {
      status: 'ERROR',
      error: error.toString(),
      message: 'เกิดข้อผิดพลาดในการทดสอบ'
    };
  }
}

/**
 * 🔍 ทดสอบทุก Sheet ที่เกี่ยวข้อง
 */
function testAllSheetsData() {
  try {
    console.log('🧪 Testing all sheets data...');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const results = {};
    
    const sheetsToTest = [
      { name: 'ข้อมูลการอบรม', key: 'trainings' },
      { name: 'รูปกิจกรรม', key: 'photos' },
      { name: 'ข้อมูลผลงานครู', key: 'teacherAwards' },
      { name: 'ข้อมูลผลงานนักเรียน', key: 'studentWorks' },
      { name: 'รายชื่อบุคลากร', key: 'personnel' }
    ];
    
    sheetsToTest.forEach(sheetConfig => {
      const sheet = ss.getSheetByName(sheetConfig.name);
      
      if (!sheet) {
        results[sheetConfig.key] = {
          exists: false,
          count: 0,
          message: `ไม่พบ Sheet: ${sheetConfig.name}`
        };
        return;
      }
      
      const lastRow = sheet.getLastRow();
      const hasData = lastRow > 1;
      
      if (hasData) {
        const dataRange = sheet.getDataRange();
        const data = dataRange.getValues();
        
        let actualCount = 0;
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const hasAnyData = row.some(cell => {
            if (cell === null || cell === undefined) return false;
            if (typeof cell === 'string' && cell.trim() === '') return false;
            return true;
          });
          if (hasAnyData) actualCount++;
        }
        
        results[sheetConfig.key] = {
          exists: true,
          sheetName: sheetConfig.name,
          lastRow: lastRow,
          totalRows: data.length - 1,
          actualDataRows: actualCount,
          emptyRows: (data.length - 1) - actualCount,
          hasData: actualCount > 0
        };
      } else {
        results[sheetConfig.key] = {
          exists: true,
          sheetName: sheetConfig.name,
          lastRow: lastRow,
          totalRows: 0,
          actualDataRows: 0,
          emptyRows: 0,
          hasData: false
        };
      }
    });
    
    const totalDataRows = Object.values(results).reduce((sum, item) => sum + (item.actualDataRows || 0), 0);
    
    console.log('📊 ผลลัพธ์การทดสอบทุก Sheet:');
    Object.keys(results).forEach(key => {
      const item = results[key];
      console.log(`📋 ${item.sheetName || key}:`);
      console.log(`   - มี Sheet: ${item.exists ? '✅' : '❌'}`);
      console.log(`   - แถวสุดท้าย: ${item.lastRow || 0}`);
      console.log(`   - แถวข้อมูลจริง: ${item.actualDataRows || 0}`);
      console.log(`   - แถวว่าง: ${item.emptyRows || 0}`);
    });
    
    return {
      status: 'SUCCESS',
      message: `ทดสอบทุก Sheet สำเร็จ พบข้อมูลทั้งหมด ${totalDataRows} แถว`,
      results: results,
      summary: {
        totalSheets: sheetsToTest.length,
        sheetsExist: Object.values(results).filter(r => r.exists).length,
        totalDataRows: totalDataRows,
        trainingsCount: results.trainings?.actualDataRows || 0,
        photosCount: results.photos?.actualDataRows || 0,
        teacherAwardsCount: results.teacherAwards?.actualDataRows || 0,
        studentWorksCount: results.studentWorks?.actualDataRows || 0,
        personnelCount: results.personnel?.actualDataRows || 0
      }
    };
    
  } catch (error) {
    console.log('❌ Error:', error.toString());
    return {
      status: 'ERROR',
      error: error.toString()
    };
  }
}

/**
 * 🔍 ทดสอบ Dashboard data โดยตรง
 */
function testDashboardDataAccuracy() {
  try {
    console.log('🧪 Testing Dashboard data accuracy...');
    
    const sheetsTest = testAllSheetsData();
    const dashboardData = handleGetDashboardData();
    
    let dashboardTrainingsCount = 0;
    if (dashboardData.status === 'SUCCESS' && dashboardData.data.summary.trainings) {
      dashboardTrainingsCount = dashboardData.data.summary.trainings.total || 0;
    }
    
    const actualTrainingsCount = sheetsTest.results?.trainings?.actualDataRows || 0;
    
    const comparison = {
      actualDataInSheet: actualTrainingsCount,
      dashboardReported: dashboardTrainingsCount,
      matches: actualTrainingsCount === dashboardTrainingsCount,
      difference: Math.abs(actualTrainingsCount - dashboardTrainingsCount)
    };
    
    console.log('📊 การเปรียบเทียบ:');
    console.log('- ข้อมูลจริงใน Sheet:', actualTrainingsCount);
    console.log('- Dashboard รายงาน:', dashboardTrainingsCount);
    console.log('- ตรงกัน:', comparison.matches ? '✅' : '❌');
    console.log('- ต่างกัน:', comparison.difference, 'รายการ');
    
    if (!comparison.matches) {
      console.log('⚠️ ข้อสังเกต:');
      console.log('  - ถ้า Dashboard แสดงมากกว่า: อาจนับแถวว่างรวมด้วย');
      console.log('  - ถ้า Dashboard แสดงน้อยกว่า: อาจมีปัญหาในการดึงข้อมูล');
    }
    
    return {
      status: 'SUCCESS',
      message: comparison.matches ? 
        `✅ Dashboard แสดงข้อมูลถูกต้อง: ${dashboardTrainingsCount} รายการ` :
        `⚠️ Dashboard แสดงไม่ตรงกับข้อมูลจริง (ต่างกัน ${comparison.difference} รายการ)`,
      comparison: comparison,
      sheetsTest: sheetsTest,
      dashboardStatus: dashboardData.status,
      explanation: {
        possibleReasons: [
          '1. Dashboard อาจนับแถวว่างรวมด้วย',
          '2. อาจมีข้อมูลซ้ำซ้อน',
          '3. อาจมีการดึงข้อมูลผิดคอลัมน์',
          '4. อาจมีหลาย Sheet ที่ชื่อคล้ายกัน'
        ]
      }
    };
    
  } catch (error) {
    console.log('❌ Error:', error.toString());
    return {
      status: 'ERROR',
      error: error.toString()
    };
  }
}

/**
 * 🔍 ฟังก์ชันทดสอบแบบง่ายสำหรับเมนู
 */
function quickDataCheck() {
  try {
    const result = testTrainingDataCount();
    
    let message = '';
    if (result.status === 'SUCCESS') {
      message = `📊 ผลการตรวจสอบข้อมูลการอบรม:\n\n` +
               `✅ Sheet: ${result.sheetName || 'ข้อมูลการอบรม'}\n` +
               `📈 แถวทั้งหมดใน Sheet: ${result.lastRow || 0}\n` +
               `📝 แถวข้อมูลจริง: ${result.actualDataCount || 0}\n` +
               `📦 Dashboard ดึงข้อมูล: ${result.rowsFetchedByDashboard || 0} แถว\n\n`;
      
      if (result.explanation) {
        message += `💡 ${result.explanation.dashboardShows}\n`;
        message += `📌 ${result.explanation.note}\n`;
      }
      
      if (result.sampleData && result.sampleData.length > 0) {
        message += `\n📋 ตัวอย่างข้อมูล (5 รายการแรก):\n`;
        result.sampleData.forEach(item => {
          message += `  • แถว ${item.row}: ${item.name} - ${item.course}\n`;
        });
      }
    } else {
      message = `❌ ${result.message || 'เกิดข้อผิดพลาด'}\n\n${result.error || ''}`;
    }
    
    SpreadsheetApp.getUi().alert('ตรวจสอบข้อมูลการอบรม', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
    return result;
    
  } catch (error) {
    const errorMessage = `❌ เกิดข้อผิดพลาด: ${error.toString()}`;
    SpreadsheetApp.getUi().alert('ข้อผิดพลาด', errorMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    return {
      status: 'ERROR',
      error: error.toString()
    };
  }
}

/**
 * ✅ ฟังก์ชันทดสอบการดึงรายชื่อบุคลากร
 */
function testPersonnelFunctions() {
    try {
        console.log('🧪 Testing personnel functions...');
        
        const results = [];
        
        const allNames = getPersonnelNames();
        results.push({
            test: 'getPersonnelNames',
            status: 'SUCCESS',
            count: allNames.length,
            sample: allNames.slice(0, 3)
        });
        
        const personnelList = handleGetPersonnelList();
        results.push({
            test: 'handleGetPersonnelList',
            status: personnelList.status,
            count: personnelList.count || 0
        });
        
        const sheetCheck = handleCheckPersonnelSheet();
        results.push({
            test: 'handleCheckPersonnelSheet',
            status: sheetCheck.status,
            sheetExists: sheetCheck.sheetExists,
            hasData: sheetCheck.hasData
        });
        
        const searchTest = handleSearchPersonnel({ query: 'ครู' });
        results.push({
            test: 'handleSearchPersonnel',
            status: searchTest.status,
            count: searchTest.count || 0
        });
        
        console.log('✅ Personnel functions test completed');
        
        return {
            status: 'SUCCESS',
            message: 'Personnel functions test completed',
            results: results,
            summary: {
                totalTests: results.length,
                passedTests: results.filter(r => r.status === 'SUCCESS').length,
                totalPersonnel: allNames.length
            }
        };
        
    } catch (error) {
        console.log('❌ Personnel functions test failed:', error.toString());
        return {
            status: 'ERROR',
            error: error.toString(),
            message: 'Personnel functions test failed'
        };
    }
}

// ==============================================
// 🔄 SCHEDULED FUNCTIONS (ใหม่)
// ==============================================

/**
 * 🔄 Scheduled function: ทำความสะอาด chunk เก่าทิ้งทุกชั่วโมง
 */
function scheduledChunkCleanup() {
    try {
        console.log('🧹 Starting scheduled chunk cleanup...');
        const cleaned = cleanupOldChunks();
        console.log(`✅ Cleanup completed: ${cleaned} chunks removed`);
        return cleaned;
    } catch (error) {
        console.log(`❌ Scheduled cleanup failed: ${error.toString()}`);
        return 0;
    }
}

/**
 * 🕒 ตั้งค่า trigger สำหรับ scheduled cleanup
 */
function setupChunkCleanupTrigger() {
    try {
        // ลบ trigger เก่าทิ้งก่อน
        const triggers = ScriptApp.getProjectTriggers();
        triggers.forEach(trigger => {
            if (trigger.getHandlerFunction() === 'scheduledChunkCleanup') {
                ScriptApp.deleteTrigger(trigger);
            }
        });
        
        // สร้าง trigger ใหม่ (ทุกชั่วโมง)
        ScriptApp.newTrigger('scheduledChunkCleanup')
            .timeBased()
            .everyHours(1)
            .create();
        
        console.log('✅ Chunk cleanup trigger set up (every 1 hour)');
        return true;
    } catch (error) {
        console.log(`❌ Error setting up trigger: ${error.toString()}`);
        return false;
    }
}

// ==============================================
// 🚀 INITIALIZATION
// ==============================================

/**
 * ฟังก์ชันเรียกใช้งานครั้งแรกเพื่อสร้าง Sheets หากไม่มี
 */
function initializeSheets() {
    try {
        console.log('🚀 Initializing sheets...');
        
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        
        const sheetsToCreate = [
            {
                name: PERSONNEL_SHEET_NAME,
                headers: ['รายชื่อบุคลากร'],
                description: 'สำหรับเก็บรายชื่อบุคลากรที่ใช้ใน drop down list'
            },
            {
                name: TRAINING_SHEET_NAME,
                headers: [
                    'วัน-เวลาบันทึก', 'ชื่อ-นามสกุล', 'ประเภท', 'ชื่อหลักสูตร/กิจกรรม', 
                    'สถานที่', 'วันที่เริ่ม', 'วันที่สิ้นสุด', 'จำนวนชั่วโมง', 
                    'ชื่อไฟล์เกียรติบัตร', 'Link เกียรติบัตร', 'Submission ID'
                ]
            },
            {
                name: PHOTO_SHEET_NAME,
                headers: ['Timestamp', 'ชื่อ-นามสกุล', 'Photo Link', 'Description']
            },
            {
                name: TEACHER_AWARD_SHEET_NAME,
                headers: ['วัน-เวลาบันทึก', 'ชื่อครู', 'ชื่อรางวัล', 'ระดับรางวัล', 'วันที่รับรางวัล', 'Link เกียรติบัตร']
            },
            {
                name: STUDENT_WORK_SHEET_NAME,
                headers: ['วัน-เวลาบันทึก', 'ชื่อนักเรียน', 'ประเภทผลงาน', 'ชื่อผลงาน', 'ชื่อที่ปรึกษา', 'ระดับรางวัล', 'วันที่รับรางวัล', 'Link เกียรติบัตร']
            }
        ];
        
        sheetsToCreate.forEach(sheetConfig => {
            let sheet = ss.getSheetByName(sheetConfig.name);
            if (!sheet) {
                console.log(`📝 Creating sheet: ${sheetConfig.name}`);
                sheet = ss.insertSheet(sheetConfig.name);
                if (sheetConfig.headers && sheetConfig.headers.length > 0) {
                    sheet.getRange(1, 1, 1, sheetConfig.headers.length).setValues([sheetConfig.headers]);
                    sheet.getRange(1, 1, 1, sheetConfig.headers.length)
                         .setFontWeight('bold')
                         .setBackground('#f0f8ff');
                }
                
                if (sheetConfig.name === PERSONNEL_SHEET_NAME) {
                    console.log('ℹ️ Sheet "รายชื่อบุคลากร" ถูกสร้างแล้ว กรุณาเพิ่มรายชื่อในคอลัมน์ A');
                }
            } else {
                console.log(`✅ Sheet already exists: ${sheetConfig.name} (${sheet.getLastRow()} rows)`);
            }
        });
        
        const personnelNames = getPersonnelNames();
        console.log(`📊 Personnel sheet has ${personnelNames.length} names`);
        if (personnelNames.length > 0) {
            console.log('📝 Sample names:', personnelNames.slice(0, 5));
        }
        
        // ตั้งค่า cleanup trigger
        setupChunkCleanupTrigger();
        
        console.log('✅ All sheets initialized successfully');
        return {
            status: 'SUCCESS',
            message: 'Sheets initialized successfully',
            personnelCount: personnelNames.length
        };
        
    } catch (error) {
        console.log(`❌ Error initializing sheets: ${error.toString()}`);
        return {
            status: 'ERROR',
            error: error.toString()
        };
    }
}

// ==============================================
// 📄 MENU FUNCTIONS
// ==============================================

/**
 * สร้างเมนูใน Google Sheets
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('📊 ระบบบันทึกผลงาน')
        .addItem('🚀 เริ่มต้นระบบ (สร้าง Sheets)', 'initializeSheets')
        .addSeparator()
        .addSubMenu(ui.createMenu('📈 Dashboard Tools')
            .addItem('ทดสอบฟังก์ชัน Dashboard', 'testDashboardFunctions')
            .addItem('ตรวจสอบข้อมูลล่าสุด', 'checkLatestDataMenu')
            .addItem('วิเคราะห์แนวโน้มข้อมูล', 'analyzeTrendsMenu'))
        .addSeparator()
        .addSubMenu(ui.createMenu('🔧 ทดสอบระบบ')
            .addItem('ทดสอบการเขียนข้อมูล', 'testWriteToSheet')
            .addItem('ทดสอบสภาพแวดล้อม', 'testEnvironment')
            .addItem('ทดสอบฟังก์ชันรายชื่อ', 'testPersonnelFunctions')
            .addItem('ทดสอบบันทึกผลงานนักเรียน', 'testStudentWorkRecord')
            .addSeparator()
            .addItem('🔍 ตรวจสอบข้อมูลการอบรม', 'quickDataCheck')
            .addItem('📊 ตรวจสอบทุก Sheet', 'testAllSheetsData')
            .addItem('📈 ทดสอบความถูกต้อง Dashboard', 'testDashboardDataAccuracy'))
        .addSeparator()
        .addSubMenu(ui.createMenu('👥 จัดการรายชื่อบุคลากร')
            .addItem('ตรวจสอบ Sheet รายชื่อ', 'handleCheckPersonnelSheetMenu')
            .addItem('แสดงรายชื่อทั้งหมด', 'showPersonnelListMenu')
            .addItem('นับจำนวนรายชื่อ', 'countPersonnelMenu'))
        .addSeparator()
        .addSubMenu(ui.createMenu('🔄 Chunk Upload Management')
            .addItem('ตั้งค่า Cleanup Trigger', 'setupChunkCleanupTrigger')
            .addItem('ทำความสะอาด Chunk เก่า', 'scheduledChunkCleanup'))
        .addToUi();
}

/**
 * ✅ ทดสอบฟังก์ชัน Dashboard
 */
function testDashboardFunctions() {
    try {
        console.log('🧪 Testing dashboard functions...');
        
        const results = [];
        
        const dashboardData = handleGetDashboardData();
        results.push({
            test: 'handleGetDashboardData',
            status: dashboardData.status,
            hasData: dashboardData.status === 'SUCCESS'
        });
        
        const summaryData = handleGetDashboardSummary();
        results.push({
            test: 'handleGetDashboardSummary',
            status: summaryData.status,
            hasData: summaryData.status === 'SUCCESS'
        });
        
        const chartData = handleGetChartData();
        results.push({
            test: 'handleGetChartData',
            status: chartData.status,
            hasCharts: chartData.status === 'SUCCESS' && Object.keys(chartData.charts).length > 0
        });
        
        const searchData = handleGlobalSearch({ query: 'ครู' });
        results.push({
            test: 'handleGlobalSearch',
            status: searchData.status,
            foundResults: searchData.status === 'SUCCESS' && searchData.counts.total > 0
        });
        
        const message = `📊 ผลการทดสอบ Dashboard Functions:\n\n` +
                       `✅ handleGetDashboardData: ${results[0].status}\n` +
                       `✅ handleGetDashboardSummary: ${results[1].status}\n` +
                       `✅ handleGetChartData: ${results[2].status}\n` +
                       `✅ handleGlobalSearch: ${results[3].status}\n\n` +
                       `🎯 ทุกฟังก์ชันพร้อมใช้งานสำหรับ Dashboard!`;
        
        SpreadsheetApp.getUi().alert('ทดสอบ Dashboard Functions', message, SpreadsheetApp.getUi().ButtonSet.OK);
        
        return {
            status: 'SUCCESS',
            results: results
        };
        
    } catch (error) {
        console.log('❌ Dashboard functions test failed:', error.toString());
        SpreadsheetApp.getUi().alert('ข้อผิดพลาด', 'ไม่สามารถทดสอบ Dashboard ได้: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
        return {
            status: 'ERROR',
            error: error.toString()
        };
    }
}

/**
 * เมนูตรวจสอบข้อมูลล่าสุด
 */
function checkLatestDataMenu() {
    try {
        const result = handleGetUnprocessedData({ hoursThreshold: 24 });
        
        let message = `📊 ข้อมูลล่าสุด (24 ชั่วโมงที่ผ่านมา)\n\n`;
        
        if (result.status === 'SUCCESS') {
            const summary = result.summary;
            message += `🆕 มีข้อมูลใหม่ทั้งหมด: ${summary.totalNew} รายการ\n\n`;
            
            if (result.unprocessed.trainings.length > 0) {
                message += `🎓 การอบรมใหม่: ${result.unprocessed.trainings.length} รายการ\n`;
            }
            if (result.unprocessed.photos.length > 0) {
                message += `📸 รูปกิจกรรมใหม่: ${result.unprocessed.photos.length} รายการ\n`;
            }
            if (result.unprocessed.teacherAwards.length > 0) {
                message += `🏆 ผลงานครูใหม่: ${result.unprocessed.teacherAwards.length} รายการ\n`;
            }
            if (result.unprocessed.studentWorks.length > 0) {
                message += `👨‍🎓 ผลงานนักเรียนใหม่: ${result.unprocessed.studentWorks.length} รายการ\n`;
            }
            
            if (summary.totalNew === 0) {
                message += `\nℹ️ ไม่มีข้อมูลใหม่ในช่วง 24 ชั่วโมงที่ผ่านมา`;
            }
        } else {
            message = `❌ ไม่สามารถตรวจสอบข้อมูลล่าสุดได้\n${result.error}`;
        }
        
        SpreadsheetApp.getUi().alert('ตรวจสอบข้อมูลล่าสุด', message, SpreadsheetApp.getUi().ButtonSet.OK);
        
    } catch (error) {
        SpreadsheetApp.getUi().alert('ข้อผิดพลาด', 'เกิดข้อผิดพลาด: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    }
}

/**
 * เมนูวิเคราะห์แนวโน้ม
 */
function analyzeTrendsMenu() {
    try {
        const result = handleGetTrendData({ monthsBack: 6 });
        
        let message = `📈 วิเคราะห์แนวโน้มข้อมูล (6 เดือนย้อนหลัง)\n\n`;
        
        if (result.status === 'SUCCESS') {
            const trend = result.trend;
            
            if (trend.trainings.length > 0) {
                const trainingCounts = trend.trainings.map(t => t.count);
                const totalTrainings = trainingCounts.reduce((a, b) => a + b, 0);
                const avgTrainings = totalTrainings / trend.trainings.length;
                
                message += `🎓 การอบรม:\n`;
                message += `   - รวม: ${totalTrainings} รายการ\n`;
                message += `   - เฉลี่ยต่อเดือน: ${avgTrainings.toFixed(1)} รายการ\n`;
                
                const maxMonth = trend.trainings.reduce((max, curr) => 
                    curr.count > max.count ? curr : max, trend.trainings[0]);
                message += `   - เดือนที่มีมากที่สุด: ${maxMonth.month} (${maxMonth.count} รายการ)\n\n`;
            }
            
            if (trend.teacherAwards.length > 0) {
                const awardCounts = trend.teacherAwards.map(t => t.count);
                const totalAwards = awardCounts.reduce((a, b) => a + b, 0);
                
                message += `🏆 ผลงานครู:\n`;
                message += `   - รวม: ${totalAwards} รายการ\n\n`;
            }
            
            if (trend.studentWorks.length > 0) {
                const workCounts = trend.studentWorks.map(t => t.count);
                const totalWorks = workCounts.reduce((a, b) => a + b, 0);
                
                message += `👨‍🎓 ผลงานนักเรียน:\n`;
                message += `   - รวม: ${totalWorks} รายการ\n`;
            }
            
            message += `\n📅 ข้อมูลล่าสุด: ${new Date().toLocaleDateString('th-TH')}`;
        } else {
            message = `❌ ไม่สามารถวิเคราะห์แนวโน้มได้\n${result.error}`;
        }
        
        SpreadsheetApp.getUi().alert('วิเคราะห์แนวโน้มข้อมูล', message, SpreadsheetApp.getUi().ButtonSet.OK);
        
    } catch (error) {
        SpreadsheetApp.getUi().alert('ข้อผิดพลาด', 'เกิดข้อผิดพลาด: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    }
}

/**
 * แสดงรายชื่อบุคลากรใน UI
 */
function showPersonnelListMenu() {
    try {
        const names = getPersonnelNames();
        
        if (names.length === 0) {
            SpreadsheetApp.getUi().alert('รายชื่อบุคลากร', 'ไม่มีรายชื่อบุคลากรในระบบ', SpreadsheetApp.getUi().ButtonSet.OK);
            return;
        }
        
        const message = `พบรายชื่อบุคลากร ${names.length} รายชื่อ:\n\n` +
                       names.slice(0, 20).join('\n') +
                       (names.length > 20 ? `\n\n...และอีก ${names.length - 20} รายชื่อ` : '');
        
        SpreadsheetApp.getUi().alert('รายชื่อบุคลากร', message, SpreadsheetApp.getUi().ButtonSet.OK);
        
    } catch (error) {
        SpreadsheetApp.getUi().alert('ข้อผิดพลาด', 'ไม่สามารถดึงรายชื่อได้: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    }
}

/**
 * แสดงจำนวนรายชื่อบุคลากร
 */
function countPersonnelMenu() {
    try {
        const names = getPersonnelNames();
        
        const analysis = {
            นาย: names.filter(name => name.startsWith('นาย')).length,
            นาง: names.filter(name => name.startsWith('นาง') && !name.startsWith('นางสาว')).length,
            นางสาว: names.filter(name => name.startsWith('นางสาว')).length,
            อื่นๆ: names.filter(name => 
                !name.startsWith('นาย') && 
                !name.startsWith('นาง') && 
                !name.startsWith('นางสาว')
            ).length
        };
        
        const message = `📊 สรุปจำนวนรายชื่อบุคลากร\n\n` +
                       `รวมทั้งหมด: ${names.length} รายชื่อ\n` +
                       `👨 นาย: ${analysis.นาย} รายชื่อ\n` +
                       `👩 นาง: ${analysis.นาง} รายชื่อ\n` +
                       `👧 นางสาว: ${analysis.นางสาว} รายชื่อ\n` +
                       `📝 อื่นๆ: ${analysis.อื่นๆ} รายชื่อ`;
        
        SpreadsheetApp.getUi().alert('สรุปรายชื่อบุคลากร', message, SpreadsheetApp.getUi().ButtonSet.OK);
        
    } catch (error) {
        SpreadsheetApp.getUi().alert('ข้อผิดพลาด', 'ไม่สามารถนับรายชื่อได้: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    }
}

/**
 * ตรวจสอบ Sheet รายชื่อใน UI
 */
function handleCheckPersonnelSheetMenu() {
    try {
        const result = handleCheckPersonnelSheet();
        
        let message = '';
        if (result.status === 'SUCCESS') {
            message = `✅ ${result.message}\n\n` +
                     `📁 Sheet: ${result.sheetName}\n` +
                     `📊 แถวทั้งหมด: ${result.totalRows} แถว\n` +
                     `👤 รายชื่อที่ถูกต้อง: ${result.validNamesCount} รายชื่อ\n` +
                     `📝 ตัวอย่าง: ${result.sampleNames.join(', ')}`;
        } else {
            message = `❌ ${result.message}\n\n${result.suggestion || ''}`;
        }
        
        SpreadsheetApp.getUi().alert('ตรวจสอบ Sheet รายชื่อ', message, SpreadsheetApp.getUi().ButtonSet.OK);
        
    } catch (error) {
        SpreadsheetApp.getUi().alert('ข้อผิดพลาด', 'ไม่สามารถตรวจสอบได้: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    }
}