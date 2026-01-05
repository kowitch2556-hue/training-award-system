// assets/js/reports-core.js
class ReportsCore {
    constructor() {
        this.API_URL = window.GAS_API_URL || window.API_URL || window.GAS_WEB_APP_URL;
        console.log('📊 ReportsCore initialized with API:', this.API_URL);
    }

    /**
     * ดึงข้อมูลการอบรมจาก API
     */
    async fetchTrainingData() {
        try {
            console.log('🎓 Fetching training data...');
            
            const result = await window.callGASAPI('getTrainingSummary');
            
            if (result.status === 'SUCCESS') {
                console.log('✅ Training data fetched:', result.data.total, 'records');
                return {
                    success: true,
                    data: result.data || {
                        total: 0,
                        totalHours: 0,
                        uniqueParticipants: 0,
                        avgHours: 0,
                        recentData: [],
                        topParticipants: []
                    }
                };
            } else {
                console.error('❌ Failed to fetch training data:', result.error);
                return {
                    success: false,
                    error: result.error,
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
        } catch (error) {
            console.error('❌ Error fetching training data:', error);
            return {
                success: false,
                error: error.message,
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
    }

    /**
     * ดึงข้อมูลผลงานนักเรียนจาก API
     */
    async fetchStudentData() {
        try {
            console.log('👨‍🎓 Fetching student data...');
            
            const result = await window.callGASAPI('getStudentSummary');
            
            if (result.status === 'SUCCESS') {
                console.log('✅ Student data fetched:', result.data.total, 'records');
                return {
                    success: true,
                    data: result.data || {
                        total: 0,
                        awardCount: 0,
                        uniqueStudents: 0,
                        uniqueAdvisors: 0,
                        recentData: [],
                        topStudents: []
                    }
                };
            } else {
                console.error('❌ Failed to fetch student data:', result.error);
                return {
                    success: false,
                    error: result.error,
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
        } catch (error) {
            console.error('❌ Error fetching student data:', error);
            return {
                success: false,
                error: error.message,
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
    }

    /**
     * ดึงข้อมูลผลงานครูจาก API
     */
    async fetchTeacherData() {
        try {
            console.log('🏆 Fetching teacher data...');
            
            const result = await window.callGASAPI('getTeacherSummary');
            
            if (result.status === 'SUCCESS') {
                console.log('✅ Teacher data fetched:', result.data.total, 'records');
                return {
                    success: true,
                    data: result.data || {
                        total: 0,
                        firstPlace: 0,
                        uniqueTeachers: 0,
                        avgPerTeacher: 0,
                        recentData: [],
                        topTeachers: []
                    }
                };
            } else {
                console.error('❌ Failed to fetch teacher data:', result.error);
                return {
                    success: false,
                    error: result.error,
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
        } catch (error) {
            console.error('❌ Error fetching teacher data:', error);
            return {
                success: false,
                error: error.message,
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
    }

    /**
     * ดึงข้อมูล dashboard สำหรับ reports
     */
    async fetchDashboardData() {
        try {
            console.log('📈 Fetching dashboard data for reports...');
            
            const result = await window.callGASAPI('getDashboardDataRealtime', {
                daysBack: 30
            });
            
            if (result.status === 'SUCCESS') {
                console.log('✅ Dashboard data fetched successfully');
                return {
                    success: true,
                    data: result.data || {
                        summary: {},
                        charts: {},
                        recent: {},
                        personnel: {}
                    }
                };
            } else {
                console.error('❌ Failed to fetch dashboard data:', result.error);
                return {
                    success: false,
                    error: result.error,
                    data: {
                        summary: {},
                        charts: {},
                        recent: {},
                        personnel: {}
                    }
                };
            }
        } catch (error) {
            console.error('❌ Error fetching dashboard data:', error);
            return {
                success: false,
                error: error.message,
                data: {
                    summary: {},
                    charts: {},
                    recent: {},
                    personnel: {}
                }
            };
        }
    }

    /**
     * ดึงข้อมูลทั้งหมดพร้อมกัน
     */
    async fetchAllReportsData() {
        console.log('🚀 Fetching all reports data...');
        
        try {
            const [training, student, teacher, dashboard] = await Promise.allSettled([
                this.fetchTrainingData(),
                this.fetchStudentData(),
                this.fetchTeacherData(),
                this.fetchDashboardData()
            ]);

            const result = {
                training: training.status === 'fulfilled' ? training.value : null,
                student: student.status === 'fulfilled' ? student.value : null,
                teacher: teacher.status === 'fulfilled' ? teacher.value : null,
                dashboard: dashboard.status === 'fulfilled' ? dashboard.value : null,
                timestamp: new Date().toISOString()
            };

            console.log('✅ All reports data fetched');
            return result;
        } catch (error) {
            console.error('❌ Error fetching all reports:', error);
            return {
                training: null,
                student: null,
                teacher: null,
                dashboard: null,
                error: error.message,
                timestamp: new Date().toISOString()
            };
        }
    }
}

// Export globally
window.ReportsCore = ReportsCore;