<template>
    <div class="download-container">
        <div class="filter-section">
            <el-date-picker v-model="dateRange" type="daterange" range-separator="至" start-placeholder="开始日期"
                end-placeholder="结束日期" :shortcuts="dateShortcuts" value-format="YYYY-MM-DD" @change="handleDateChange"
                class="date-picker" />
        </div>
        <div class="format-selector">
            <label>选择文件格式：</label>
            <select v-model="selectedFormat" class="format-select">
                <option value="xlsx">.xlsx (Excel 2007+)</option>
                <option value="xls">.xls (Excel 97-2003)</option>
            </select>
        </div>
        <button class="download-button" @click="handleDownload" :disabled="isDownloading">
            <span v-if="isDownloading">
                <i class="loading-icon"></i>
                下载中 {{ downloadProgress }}%
            </span>
            <span v-else>下载Excel</span>
        </button>
        <div v-if="downloadStatus" class="download-status" :class="downloadStatus.type">
            {{ downloadStatus.message }}
        </div>
    </div>
</template>

<script>
import ExcelJS from 'exceljs';
import axios from 'axios';
import { ElDatePicker } from 'element-plus';
import 'element-plus/dist/index.css';


export default {
    name: 'DownloadButton',
    components: {
        ElDatePicker
    },
    data() {
        return {
            testData: [
                {

                    "service_id": 165,
                    "id": 287,
                    "customer_id": "u7ev",
                    "customer": "test",
                    "dateTime": "2025-05-23",
                    "daiyaTime": "2025-05-23",
                    "doctor": "test",
                    "proxy": "test",
                    "porcelain": "爱尔创 (国产全瓷)",
                    "tiepianColor": "red",
                    "CAD": "csa",
                    "checi": "dwsqcx",
                    "frontPhoto": "https://yayi-1325314533.cos.ap-shanghai.myqcloud.com/uploads/7KWUHjdiqmwmd87292cf83ea8c4f7aa13010027ba931.jpg",
                    "adviceContent": "dewscwcfedswcxdew",
                    "leftFv": "https://yayi-1325314533.cos.ap-shanghai.myqcloud.com/uploads/EVpTOxCwJ40Tb39ebf1c7b297c5edce6088e18d2f8dd.png",
                    "rightFv": "",
                    "front": "",
                    "leftFvEdge": "",
                    "rightFvEdge": "",
                    "intentImg": "undefined",
                    "designAdvice": "cedijn",
                    "designList": "[\"https://yayi-1325314533.cos.ap-shanghai.myqcloud.com/uploads/jjmmMTSZxkUL0bd7b8f194b4a42973c1383805eb926c.jpg\",\"https://yayi-1325314533.cos.ap-shanghai.myqcloud.com/uploads/XYhTb6eInfrcd8254f61f5cece9166890c868763a7b7.jpg\"]",
                    "bianyuanOpen": "false",
                    "bianyuanValue": "0",
                    "roundOpen": "false",
                    "roundValue": "0",
                    "luochaOpen": "false",
                    "luochaValue": "0",
                    "angleOpen": "false",
                    "angleValue": "0",
                    "jiandunOpen": "false",
                    "jiandunValue": "0",
                    "qieduanOpen": "false",
                    "qieduanValue": "0",
                    "textureOpen": "false",
                    "textureValue": "0",
                    "dotOpen": "false",
                    "dotValue": "0",
                    "touliangOpen": "false",
                    "touliangValue": "0",
                    "qieduanLinearsOpen": "false",
                    "qieduanLinearsValue": "0",
                    "thicknessOpen": "false",
                    "thicknessValue": "0",
                    "createtime": "2025-05-23T15:53:11.000Z",
                    "isPrivacy": 0,
                    "problem": "dsxwxssw",
                    "shangci": null,
                    "shangyou": null,
                    "checiImg": "[\"https://yayi-1325314533.cos.ap-shanghai.myqcloud.com/uploads/sn24spZEkL9673d7e1d051714cc4f54aa738e1617f8d.jpg\"]",
                    "shangyouImg": "[\"https://yayi-1325314533.cos.ap-shanghai.myqcloud.com/uploads/xdcZ7iuXnd5nb39ebf1c7b297c5edce6088e18d2f8dd.png\"]",
                    "shangciImg": "[\"https://yayi-1325314533.cos.ap-shanghai.myqcloud.com/uploads/iNlRSIHx622Ab39ebf1c7b297c5edce6088e18d2f8dd.png\"]",
                    "CADImg": "[\"https://yayi-1325314533.cos.ap-shanghai.myqcloud.com/uploads/1Xikz19mHXfGd8254f61f5cece9166890c868763a7b7.jpg\",\"https://yayi-1325314533.cos.ap-shanghai.myqcloud.com/uploads/xH2WHsORl3em0bd7b8f194b4a42973c1383805eb926c.jpg\"]",
                    "adjust": 0,
                    "shangciRemark": "xsaq",
                    "CADRemark": "dqdx",
                    "checiRemark": "dxaxsq",
                    "shangyouRemark": "xsqxsq",
                    "daodianImg": "[\"https://yayi-1325314533.cos.ap-shanghai.myqcloud.com/uploads/oOCQPkyf5Pwsd8254f61f5cece9166890c868763a7b7.jpg\"]",
                    "nurse": "test",
                    "toothSensitivity": 0,
                    "tryInfo": `[{"tryImg":["https://gdcasa.cn:3010/img/images/cec790f75bd47b077a35528f97c421ea.tryImg1.jpg"],"remark":"牙齿掉落 形态不满意/解决方案 拿回去重新制作"}]`,
                    "recoverInfo": `[{"recoverImg":["https://gdcasa.cn:3010/img/images/82d28274dab90403704fad588af5065e.recoverImg1.jpg"],"remark":"又再次脱落 粘接不行"}]`,
                    "imgList": "[\"https://yayi-1325314533.cos.ap-shanghai.myqcloud.com/uploads/BMJwrNRm21ysb5f115f0689e514072e3415e1689861a.jpg\"]",
                    "docterSummary": "dxwseedxwqs"
                }
            ],
            isDownloading: false,
            downloadProgress: 0,
            downloadStatus: null,
            selectedFormat: 'xlsx',
            dateRange: [],
            dateShortcuts: [
                {
                    text: '最近一周',
                    value: () => {
                        const end = new Date();
                        const start = new Date();
                        start.setTime(start.getTime() - 3600 * 1000 * 24 * 7);
                        return [start, end];
                    },
                },
                {
                    text: '最近一个月',
                    value: () => {
                        const end = new Date();
                        const start = new Date();
                        start.setTime(start.getTime() - 3600 * 1000 * 24 * 30);
                        return [start, end];
                    },
                },
                {
                    text: '最近三个月',
                    value: () => {
                        const end = new Date();
                        const start = new Date();
                        start.setTime(start.getTime() - 3600 * 1000 * 24 * 90);
                        return [start, end];
                    },
                }
            ],
            // 模拟的图片URL数组，每个员工有多张照片
            mockImages: [
                [
                    'http://e.hiphotos.baidu.com/image/pic/item/a1ec08fa513d2697e542494057fbb2fb4316d81e.jpg',
                    'http://c.hiphotos.baidu.com/image/pic/item/30adcbef76094b36de8a2fe5a1cc7cd98d109d99.jpg'
                ],
                [
                    'http://h.hiphotos.baidu.com/image/pic/item/7c1ed21b0ef41bd5f2c2a9e953da81cb39db3d1d.jpg',
                    'http://e.hiphotos.baidu.com/image/pic/item/a1ec08fa513d2697e542494057fbb2fb4316d81e.jpg'
                ],
                [
                    'http://c.hiphotos.baidu.com/image/pic/item/30adcbef76094b36de8a2fe5a1cc7cd98d109d99.jpg',
                    'http://h.hiphotos.baidu.com/image/pic/item/7c1ed21b0ef41bd5f2c2a9e953da81cb39db3d1d.jpg'
                ],
                [
                    'http://e.hiphotos.baidu.com/image/pic/item/a1ec08fa513d2697e542494057fbb2fb4316d81e.jpg',
                    'http://c.hiphotos.baidu.com/image/pic/item/30adcbef76094b36de8a2fe5a1cc7cd98d109d99.jpg',
                    'http://h.hiphotos.baidu.com/image/pic/item/7c1ed21b0ef41bd5f2c2a9e953da81cb39db3d1d.jpg'
                ]
            ]
        }
    },
    methods: {
        // 处理日期变化
        handleDateChange(val) {
            console.log('选择的日期范围:', val);
        },

        // 模拟获取员工数据
        async fetchEmployeeData() {
            try {
                // 这里替换为实际的API地址
                const response = await axios.get('https://gdcasa.cn/api/download/customer', {
                    data: {
                        startDate: this.dateRange[0],
                        endDate: this.dateRange[1]
                    }
                });
                console.log('response', response)
                return response.data;
            } catch (error) {
                console.error('获取员工数据失败:', error);
                // 如果API调用失败，返回模拟数据
                return this.getMockEmployeeData();
            }
        },

        // 模拟获取图片数据
        async fetchEmployeeImages() {
            try {
                // 这里替换为实际的API地址
                const response = await axios.get('https://gdcasa.cn/api/download/', {
                    params: {
                        employeeIds: [1, 2, 3, 4],
                        startDate: this.dateRange[0],
                        endDate: this.dateRange[1]
                    }
                });
                return response.data;
            } catch (error) {
                console.error('获取员工图片失败:', error);
                // 如果API调用失败，返回模拟图片数据
                return []
                // return this.mockImages;
            }
        },

        // 获取模拟的员工数据
        getMockEmployeeData() {
            return [
                { name: '张三', age: '28', department: '技术部', position: '工程师', joinDate: '2023-01-15' },
                { name: '李四', age: '32', department: '市场部', position: '经理', joinDate: '2022-06-01' },
                { name: '王五', age: '25', department: '人事部', position: '专员', joinDate: '2023-03-20' },
                { name: '赵六', age: '35', department: '财务部', position: '主管', joinDate: '2021-12-10' }
            ];
        },

        // 将图片URL转换为base64
        async imageToBase64(url) {
            try {
                const response = await fetch(url);
                const blob = await response.blob();
                return new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onloadend = () => resolve(reader.result);
                    reader.onerror = reject;
                    reader.readAsDataURL(blob);
                });
            } catch (error) {
                console.error('图片转换失败:', error);
                return null;
            }
        },

        // 处理单张图片
        async handleSingleImage(url, worksheet, row, col, workbook) {
            if (!url || url === 'undefined') return;

            try {
                const base64 = await this.imageToBase64(url);
                if (base64) {
                    const imageId = workbook.addImage({
                        base64: base64.split(',')[1],
                        extension: 'jpeg',
                    });

                    worksheet.addImage(imageId, {
                        tl: { col: col, row: row },
                        br: { col: col + 1, row: row + 1 }
                    });
                }
            } catch (error) {
                console.error('处理单张图片失败:', error);
            }
        },

        // 处理多张图片
        async handleMultipleImages(jsonString, worksheet, row, col, workbook) {
            if (!jsonString) return;

            try {
                const images = JSON.parse(jsonString);
                // 过滤掉视频文件
                const imageUrls = images.filter(url => {
                    const lowerUrl = url.toLowerCase();
                    return lowerUrl.endsWith('.jpg') ||
                        lowerUrl.endsWith('.jpeg') ||
                        lowerUrl.endsWith('.png') ||
                        lowerUrl.endsWith('.gif');
                });

                for (let i = 0; i < imageUrls.length; i++) {
                    const base64 = await this.imageToBase64(imageUrls[i]);
                    if (base64) {
                        const imageId = workbook.addImage({
                            base64: base64.split(',')[1],
                            extension: 'jpeg',
                        });

                        // 计算图片位置，每张图片占据1/3的单元格宽度
                        const colWidth = 1 / imageUrls.length;
                        worksheet.addImage(imageId, {
                            tl: { col: col + (i * colWidth), row: row },
                            br: { col: col + ((i + 1) * colWidth), row: row + 1 }
                        });
                    }
                }
            } catch (error) {
                console.error('处理多张图片失败:', error);
            }
        },

        // 生成模拟的Excel数据
        async generateMockExcelData() {
            try {
                const employeeData = await this.testData;
                const workbook = new ExcelJS.Workbook();
                const worksheet = workbook.addWorksheet('员工信息');

                // 设置列宽
                worksheet.columns = [
                    { header: '客户姓名', key: 'customer', width: 15 },
                    { header: '日期', key: 'dateTime', width: 15 },
                    { header: '戴牙日期', key: 'daiyaTime', width: 15 },
                    { header: '面诊医生', key: 'doctor', width: 15 },
                    { header: '代理人', key: 'proxy', width: 15 },
                    { header: '贴片颜色', key: 'tiepianColor', width: 15 },
                    { header: '护士姓名', key: 'nurse', width: 15 },
                    { header: 'CAD设计师', key: 'CAD', width: 15 },
                    { header: '车瓷设计师', key: 'checi', width: 15 },
                    { header: '上瓷', key: 'shangci', width: 15 },
                    { header: '上釉', key: 'shangyou', width: 15 },
                    { header: 'CAD留言', key: 'CADRemark', width: 20 },
                    { header: '车瓷留言', key: 'checiRemark', width: 20 },
                    { header: '上瓷留言', key: 'shangciRemark', width: 20 },
                    { header: '上釉留言', key: 'shangyouRemark', width: 20 },
                    { header: '面诊医生建议', key: 'adviceContent', width: 30 },
                    { header: '设计师建议', key: 'designAdvice', width: 30 },
                    { header: '正面微笑照', key: 'frontPhoto', width: 20 },
                    { header: '左45度', key: 'leftFv', width: 20 },
                    { header: '右45度', key: 'rightFv', width: 20 },
                    { header: '正面扩口', key: 'front', width: 20 },
                    { header: '左45度扩口', key: 'leftFvEdge', width: 20 },
                    { header: '右45度扩口', key: 'rightFvEdge', width: 20 },
                    { header: '客户意向照', key: 'intentImg', width: 20 },
                    { header: 'CAD设计图', key: 'CADImg', width: 20 },
                    { header: '车瓷设计图', key: 'checiImg', width: 20 },
                    { header: '上瓷设计图', key: 'shangciImg', width: 20 },
                    { header: '到店设计图', key: 'daodianImg', width: 20 },
                    { header: '上釉设计图', key: 'shangyouImg', width: 20 },
                    { header: '设计师图片列表', key: 'designList', width: 30 }
                ];

                // 设置表头样式
                worksheet.getRow(1).font = { bold: true };
                worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
                worksheet.getRow(1).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFCCCCCC' }
                };

                // 添加数据行并处理图片
                for (let i = 0; i < employeeData.length; i++) {
                    const row = employeeData[i];
                    const dataRow = worksheet.addRow(row);
                    dataRow.height = 100;
                    dataRow.alignment = { vertical: 'middle', horizontal: 'center' };
                    // for (let i = 0; i < employeeImages.length; i++) {
                    //     const images = employeeImages[i];
                    //     for (let j = 0; j < images.length; j++) {
                    //         const base64 = await this.imageToBase64(images[j]);
                    //         if (base64) {
                    //             const imageId = workbook.addImage({
                    //                 base64: base64.split(',')[1],
                    //                 extension: 'jpeg',
                    //             });

                    //             // 计算图片位置，每张图片占据1/3的单元格宽度
                    //             const colWidth = 1 / images.length;
                    //             worksheet.addImage(imageId, {
                    //                 tl: { col: 5 + (j * colWidth), row: i + 1 },
                    //                 br: { col: 5 + ((j + 1) * colWidth), row: i + 2 }
                    //             });
                    //         }
                    //     }
                    // }

                    // 处理单张图片
                    await this.handleSingleImage(row.frontPhoto, worksheet, i + 2, 17, workbook);
                    await this.handleSingleImage(row.leftFv, worksheet, i + 2, 18, workbook);
                    await this.handleSingleImage(row.rightFv, worksheet, i + 2, 19, workbook);
                    await this.handleSingleImage(row.front, worksheet, i + 2, 20, workbook);
                    await this.handleSingleImage(row.leftFvEdge, worksheet, i + 2, 21, workbook);
                    await this.handleSingleImage(row.rightFvEdge, worksheet, i + 2, 22, workbook);
                    await this.handleSingleImage(row.intentImg, worksheet, i + 2, 23, workbook);

                    // 处理多张图片
                    await this.handleMultipleImages(row.CADImg, worksheet, i + 2, 24, workbook);
                    await this.handleMultipleImages(row.checiImg, worksheet, i + 2, 25, workbook);
                    await this.handleMultipleImages(row.shangciImg, worksheet, i + 2, 26, workbook);
                    await this.handleMultipleImages(row.daodianImg, worksheet, i + 2, 27, workbook);
                    await this.handleMultipleImages(row.shangyouImg, worksheet, i + 2, 28, workbook);
                    await this.handleMultipleImages(row.designList, worksheet, i + 2, 29, workbook);
                }

                // 生成Excel文件
                const buffer = await workbook.xlsx.writeBuffer();
                return new Blob([buffer], {
                    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                });
            } catch (error) {
                console.error('生成Excel文件失败:', error);
                throw error;
            }
        },

        // 模拟下载进度
        simulateProgress() {
            return new Promise((resolve) => {
                let progress = 0;
                const interval = setInterval(() => {
                    progress += Math.random() * 10;
                    if (progress >= 100) {
                        progress = 100;
                        clearInterval(interval);
                        resolve();
                    }
                    this.downloadProgress = Math.floor(progress);
                }, 200);
            });
        },

        async handleDownload() {
            try {
                if (!this.dateRange || this.dateRange.length !== 2) {
                    this.downloadStatus = { type: 'error', message: '请选择日期范围' };
                    return;
                }

                this.isDownloading = true;
                this.downloadProgress = 0;
                this.downloadStatus = { type: 'info', message: '准备下载...' };

                // 模拟网络请求延迟
                await new Promise(resolve => setTimeout(resolve, 1000));

                // 模拟下载进度
                await this.simulateProgress();

                // 生成Excel文件
                const blob = await this.generateMockExcelData();

                // 创建下载链接
                const url = window.URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = `客户信息表_${new Date().toLocaleDateString()}.${this.selectedFormat}`;

                // 触发下载
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                window.URL.revokeObjectURL(url);

                this.downloadStatus = { type: 'success', message: '下载完成！' };
            } catch (error) {
                console.error('下载失败:', error);
                this.downloadStatus = { type: 'error', message: '下载失败，请稍后重试' };
            } finally {
                this.isDownloading = false;
            }
        }
    }
}
</script>

<style scoped>
.download-container {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 10px;
}

.filter-section {
    margin-bottom: 20px;
    /* width: 30%; */
    width: 360px;
    display: flex;
    justify-content: center;
}

.date-picker {
    width: 400px;
}

.format-selector {
    display: flex;
    align-items: center;
    gap: 8px;
    margin-bottom: 10px;
}

.format-select {
    padding: 6px 12px;
    border: 1px solid #ddd;
    border-radius: 4px;
    font-size: 14px;
    background-color: white;
}

.download-button {
    padding: 10px 20px;
    background-color: #4CAF50;
    color: white;
    border: none;
    border-radius: 4px;
    cursor: pointer;
    font-size: 14px;
    transition: all 0.3s;
    min-width: 120px;
    display: flex;
    align-items: center;
    justify-content: center;
}

.download-button:hover {
    background-color: #45a049;
}

.download-button:disabled {
    background-color: #cccccc;
    cursor: not-allowed;
}

.loading-icon {
    display: inline-block;
    width: 16px;
    height: 16px;
    margin-right: 8px;
    border: 2px solid #ffffff;
    border-radius: 50%;
    border-top-color: transparent;
    animation: spin 1s linear infinite;
}

@keyframes spin {
    to {
        transform: rotate(360deg);
    }
}

.download-status {
    font-size: 14px;
    padding: 8px 12px;
    border-radius: 4px;
    margin-top: 8px;
}

.download-status.info {
    background-color: #e3f2fd;
    color: #1976d2;
}

.download-status.success {
    background-color: #e8f5e9;
    color: #2e7d32;
}

.download-status.error {
    background-color: #ffebee;
    color: #c62828;
}
</style>