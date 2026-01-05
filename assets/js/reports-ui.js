// assets/js/reports-ui.js
class ReportsUI {
    constructor() {
        this.core = new ReportsCore();
        this.isLoading = false;
        this.currentTab = 'training';
    }

    /**
     * เริ่มต้น reports page
     */
    async init() {
        console.log('🎬 Initializing Reports UI...');
        
        // ตั้งค่า event listeners
        this.setupEventListeners();
        
        // โหลดข้อมูลครั้งแรก
        await this.loadData();
        
        // แสดงข้อมูลเริ่มต้น
        this.showTrainingReport();
        
        console.log('✅ Reports UI initialized');
    }

    /**
     * ตั้งค่า event listeners
     */
    setupEventListeners() {
        // Tab switching
        document.querySelectorAll('.report-tab').forEach(tab => {
            tab.addEventListener('click', (e) => {
                const tabType = e.target.dataset.tab;
                if (tabType) {
                    this.switchTab(tabType);
                }
            });
        });

        // Refresh button
        const refreshBtn = document.getElementById('refresh-reports');
        if (refreshBtn) {
            refreshBtn.addEventListener('click', () => this.loadData(true));
        }

        // Export buttons
        document.querySelectorAll('.export-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const type = e.target.dataset.type;
                this.exportReport(type);
            });
        });

        console.log('✅ Event listeners setup complete');
    }

    /**
     * โหลดข้อมูลทั้งหมด
     */
    async loadData(forceRefresh = false) {
        if (this.isLoading) return;
        
        console.log('📥 Loading reports data...');
        this.showLoading(true);
        
        try {
            const data = await this.core.fetchAllReportsData();
            
            // บันทึกข้อมูลในตัวแปร global
            window.reportsData = data;
            
            // อัพเดท UI
            this.updateAllReports(data);
            
            // แสดงสถานะ
            this.showStatus('success', 'โหลดข้อมูลสำเร็จ', 3000);
            
            console.log('✅ Data loaded and UI updated');
            
        } catch (error) {
            console.error('❌ Error loading data:', error);
            this.showStatus('error', 'ไม่สามารถโหลดข้อมูลได้', 5000);
        } finally {
            this.showLoading(false);
        }
    }

    /**
     * เปลี่ยน tab
     */
    switchTab(tabType) {
        if (this.currentTab === tabType) return;
        
        console.log(`🔄 Switching to ${tabType} tab`);
        
        // อัพเดท active tab
        document.querySelectorAll('.report-tab').forEach(tab => {
            tab.classList.remove('active');
            if (tab.dataset.tab === tabType) {
                tab.classList.add('active');
            }
        });
        
        // ซ่อนรายงานทั้งหมด
        document.querySelectorAll('.report-section').forEach(section => {
            section.classList.remove('active');
        });
        
        // แสดงรายงานที่เลือก
        const targetSection = document.getElementById(`${tabType}-report`);
        if (targetSection) {
            targetSection.classList.add('active');
        }
        
        this.currentTab = tabType;
        
        // โหลดข้อมูลถ้ายังไม่มี
        if (window.reportsData && !window.reportsData[tabType]) {
            this.loadTabData(tabType);
        }
    }

    /**
     * โหลดข้อมูลสำหรับ tab เฉพาะ
     */
    async loadTabData(tabType) {
        console.log(`📊 Loading ${tabType} data...`);
        
        let result;
        switch (tabType) {
            case 'training':
                result = await this.core.fetchTrainingData();
                break;
            case 'student':
                result = await this.core.fetchStudentData();
                break;
            case 'teacher':
                result = await this.core.fetchTeacherData();
                break;
        }
        
        if (result && result.success) {
            this.updateReport(tabType, result.data);
        }
    }

    /**
     * อัพเดทรายงานทั้งหมด
     */
    updateAllReports(data) {
        if (data.training) {
            this.updateReport('training', data.training.data);
        }
        if (data.student) {
            this.updateReport('student', data.student.data);
        }
        if (data.teacher) {
            this.updateReport('teacher', data.teacher.data);
        }
        if (data.dashboard) {
            this.updateDashboard(data.dashboard.data);
        }
    }

    /**
     * อัพเดทรายงานเฉพาะประเภท
     */
    updateReport(type, data) {
        console.log(`🔄 Updating ${type} report with`, data);
        
        switch (type) {
            case 'training':
                this.updateTrainingReport(data);
                break;
            case 'student':
                this.updateStudentReport(data);
                break;
            case 'teacher':
                this.updateTeacherReport(data);
                break;
        }
    }

    /**
     * อัพเดทรายงานการอบรม
     */
    updateTrainingReport(data) {
        // สถิติหลัก
        this.updateElement('training-total', data.total);
        this.updateElement('training-hours', data.totalHours);
        this.updateElement('training-participants', data.uniqueParticipants);
        this.updateElement('training-avg-hours', data.avgHours);

        // ตารางล่าสุด
        this.updateRecentTable('training-recent-table', data.recentData, [
            'date', 'name', 'course', 'type', 'hours', 'location'
        ]);

        // Top participants
        this.updateTopList('training-top-list', data.topParticipants, 'name', 'count');
    }

    /**
     * อัพเดทรายงานผลงานนักเรียน
     */
    updateStudentReport(data) {
        // สถิติหลัก
        this.updateElement('student-total', data.total);
        this.updateElement('student-awards', data.awardCount);
        this.updateElement('student-unique', data.uniqueStudents);
        this.updateElement('student-advisors', data.uniqueAdvisors);

        // ตารางล่าสุด
        this.updateRecentTable('student-recent-table', data.recentData, [
            'date', 'studentName', 'workName', 'level', 'award', 'advisor'
        ]);

        // Top students
        this.updateTopList('student-top-list', data.topStudents, 'name', 'awardCount');
    }

    /**
     * อัพเดทรายงานผลงานครู
     */
    updateTeacherReport(data) {
        // สถิติหลัก
        this.updateElement('teacher-total', data.total);
        this.updateElement('teacher-first', data.firstPlace);
        this.updateElement('teacher-unique', data.uniqueTeachers);
        this.updateElement('teacher-avg', data.avgPerTeacher);

        // ตารางล่าสุด
        this.updateRecentTable('teacher-recent-table', data.recentData, [
            'date', 'teacherName', 'awardName', 'level', 'achievement', 'organization'
        ]);

        // Top teachers
        this.updateTopList('teacher-top-list', data.topTeachers, 'name', 'count');
    }

    /**
     * อัพเดท dashboard
     */
    updateDashboard(data) {
        // คุณสามารถเพิ่ม dashboard specific updates ที่นี่
        console.log('📊 Dashboard data available:', data);
    }

    /**
     * แสดง loading state
     */
    showLoading(show) {
        this.isLoading = show;
        
        const loadingOverlay = document.getElementById('loading-overlay');
        const content = document.getElementById('reports-content');
        
        if (loadingOverlay) {
            loadingOverlay.style.display = show ? 'flex' : 'none';
        }
        
        if (content) {
            content.style.opacity = show ? '0.5' : '1';
            content.style.pointerEvents = show ? 'none' : 'auto';
        }
    }

    /**
     * แสดงสถานะ
     */
    showStatus(type, message, duration = 3000) {
        const statusContainer = document.getElementById('status-container');
        if (!statusContainer) return;
        
        const status = document.createElement('div');
        status.className = `status-message status-${type}`;
        status.innerHTML = `
            <span>${message}</span>
            <button class="status-close">&times;</button>
        `;
        
        statusContainer.appendChild(status);
        
        // ปิดอัตโนมัติ
        setTimeout(() => {
            status.classList.add('fade-out');
            setTimeout(() => status.remove(), 300);
        }, duration);
        
        // ปิดด้วยปุ่ม
        status.querySelector('.status-close').addEventListener('click', () => {
            status.classList.add('fade-out');
            setTimeout(() => status.remove(), 300);
        });
    }

    /**
     * Export รายงาน
     */
    exportReport(type) {
        console.log(`📤 Exporting ${type} report...`);
        
        if (!window.reportsData || !window.reportsData[type]) {
            this.showStatus('error', 'ไม่มีข้อมูลสำหรับส่งออก', 3000);
            return;
        }
        
        const data = window.reportsData[type].data;
        const csv = this.convertToCSV(data);
        
        this.downloadCSV(csv, `${type}-report-${new Date().toISOString().split('T')[0]}.csv`);
        
        this.showStatus('success', 'ส่งออกรายงานสำเร็จ', 3000);
    }

    /**
     * แปลงข้อมูลเป็น CSV
     */
    convertToCSV(data) {
        // สำหรับตัวอย่าง จะส่งออกข้อมูล recentData
        const items = data.recentData || [];
        if (items.length === 0) return '';
        
        const headers = Object.keys(items[0]);
        const rows = items.map(item => 
            headers.map(header => `"${item[header] || ''}"`).join(',')
        );
        
        return [headers.join(','), ...rows].join('\n');
    }

    /**
     * ดาวน์โหลด CSV
     */
    downloadCSV(content, filename) {
        const blob = new Blob([content], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);
        
        link.setAttribute('href', url);
        link.setAttribute('download', filename);
        link.style.visibility = 'hidden';
        
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    /**
     * Helper: อัพเดท element ด้วยข้อมูล
     */
    updateElement(id, value) {
        const element = document.getElementById(id);
        if (element) {
            element.textContent = value;
        }
    }

    /**
     * Helper: อัพเดทตารางล่าสุด
     */
    updateRecentTable(tableId, data, columns) {
        const table = document.getElementById(tableId);
        if (!table || !data) return;
        
        const tbody = table.querySelector('tbody');
        if (!tbody) return;
        
        tbody.innerHTML = '';
        
        data.slice(0, 10).forEach(item => {
            const row = document.createElement('tr');
            columns.forEach(col => {
                const cell = document.createElement('td');
                cell.textContent = item[col] || '';
                row.appendChild(cell);
            });
            tbody.appendChild(row);
        });
    }

    /**
     * Helper: อัพเดทรายการยอดนิยม
     */
    updateTopList(listId, data, nameKey, countKey) {
        const list = document.getElementById(listId);
        if (!list || !data) return;
        
        list.innerHTML = '';
        
        data.slice(0, 10).forEach(item => {
            const li = document.createElement('li');
            li.className = 'top-item';
            li.innerHTML = `
                <span class="top-name">${item[nameKey]}</span>
                <span class="top-count">${item[countKey]} รายการ</span>
            `;
            list.appendChild(li);
        });
    }

    /**
     * แสดงรายงานการอบรม (initial)
     */
    showTrainingReport() {
        this.switchTab('training');
    }
}

// Export globally
window.ReportsUI = ReportsUI;