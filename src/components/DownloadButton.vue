<template>
    <div>
        <div v-if="!isPasswordVerified" class="password-container">
            <div class="password-input">
                <el-input v-model="password" type="password" placeholder="请输入密码" @keyup.enter="verifyPassword" />
                <el-button type="primary" @click="verifyPassword">登录</el-button>
            </div>
            <div v-if="passwordError" class="password-error">
                密码错误，请重试
            </div>
        </div>

        <div v-else class="download-container">
            <div class="filter-section">
                <el-date-picker v-model="dateRange" type="daterange" range-separator="至" start-placeholder="开始日期"
                    end-placeholder="结束日期" :shortcuts="dateShortcuts" value-format="YYYY-MM-DD"
                    @change="handleDateChange" class="date-picker" />
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
    </div>
</template>

<script>
import ExcelJS from 'exceljs';
import axios from 'axios';
import { ElDatePicker, ElInput, ElButton } from 'element-plus';
import 'element-plus/dist/index.css';


export default {
    name: 'DownloadButton',
    components: {
        ElDatePicker,
        ElInput,
        ElButton
    },
    data() {
        return {
            isPasswordVerified: false,
            password: '',
            passwordError: false,
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

        }
    },
    methods: {
        verifyPassword() {
            if (this.password === '111520888') {
                this.isPasswordVerified = true;
                this.passwordError = false;
            } else {
                this.passwordError = true;
                this.password = '';
            }
        },
        // 处理日期变化
        handleDateChange(val) {
            console.log('选择的日期范围:', val);
        },

        // 模拟获取员工数据
        async fetchData() {
            try {
                // 这里替换为实际的API地址
                const response = await axios.post('https://gdcasa.cn/api/download/customer',
                    {
                        startDate: this.dateRange[0],
                        endDate: this.dateRange[1]
                    });
                return response.data.re;
            } catch (error) {
                console.error('获取数据失败:', error);
                // 如果API调用失败，返回模拟数据
            }
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
        // 图片链接需要特殊处理，因为它们是直接下载链接 ，图片链接需要特殊处理，因为它们是直接下载链接
        // async imageToBase64(url) {
        //     if (!url || url === 'undefined') return null;

        //     try {
        //         // 使用axios获取图片数据
        //         const response = await axios.get(url, {
        //             responseType: 'arraybuffer',
        //             headers: {
        //                 'Accept': 'image/jpeg,image/png,image/gif'
        //             }
        //         });

        //         // 将arraybuffer转换为base64
        //         const base64 = Buffer.from(response.data).toString('base64');
        //         // 根据图片类型设置正确的MIME类型
        //         const contentType = response.headers['content-type'];
        //         return `data:${contentType};base64,${base64}`;
        //     } catch (error) {
        //         console.error('图片转换失败:', error);
        //         return null;
        //     }
        // },

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

                    // 设置图片位置和大小（行号-1，0-based）
                    worksheet.addImage(imageId, {
                        tl: { col: col, row: row - 1 },
                        br: { col: col + 1, row: row }
                    });

                    worksheet.getRow(row).height = 100; // 100px
                    worksheet.getColumn(col).width = 14.3; // 100px ≈ 14.3 width
                }
            } catch (error) {
                console.error('处理单张图片失败:', error);
            }
        },

        // 处理多张图片
        async handleMultipleImages(jsonString, worksheet, row, col, workbook) {
            if (!jsonString) return;

            try {
                const images = typeof jsonString === 'string' ? JSON.parse(jsonString) : jsonString;
                // 过滤掉视频文件
                const imageUrls = images.filter(url => {
                    const lowerUrl = url.toLowerCase();
                    return lowerUrl.endsWith('.jpg') ||
                        lowerUrl.endsWith('.jpeg') ||
                        lowerUrl.endsWith('.png') ||
                        lowerUrl.endsWith('.gif');
                });

                if (imageUrls.length === 0) return;
                // worksheet.getColumn(col).width = imageUrls.length * 14.3; // 100px ≈ 14.3 width
                // worksheet.getRow(row).height = 100;
                worksheet.getColumn(col).width = 14.3; // 固定100px宽
                worksheet.getRow(row).height = 100 * imageUrls.length; // 高度为n*100px

                for (let i = 0; i < imageUrls.length; i++) {
                    const base64 = await this.imageToBase64(imageUrls[i]);
                    if (base64) {
                        const imageId = workbook.addImage({
                            base64: base64.split(',')[1],
                            extension: 'jpeg',
                        });
                        //  tl: { col: col + i, row: row - 1 },
                        // br: { col: col + i + 1, row: row }
                        worksheet.addImage(imageId, {
                            tl: { col: col, row: row - 1 + i / imageUrls.length },
                            br: { col: col + 1, row: row - 1 + (i + 1) / imageUrls.length }
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
                // const employeeData = await this.testData;
                const employeeData = await this.fetchData();
                const workbook = new ExcelJS.Workbook();
                const worksheet = workbook.addWorksheet('客户信息');

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
                    // { header: '设计师建议', key: 'designAdvice', width: 30 },
                    { header: '正面微笑照', key: 'frontPhoto', width: 20 },
                    { header: '左45度', key: 'leftFv', width: 20 },
                    { header: '右45度', key: 'rightFv', width: 20 },
                    { header: '正面扩口', key: 'front', width: 20 },
                    { header: '左45度扩口', key: 'leftFvEdge', width: 20 },
                    { header: '右45度扩口', key: 'rightFvEdge', width: 20 },//21
                    { header: '设计师建议', key: 'designAdvice', width: 30 },
                    { header: '客户意向照', key: 'intentImg', width: 20 },
                    { header: '设计师图片列表', key: 'designList', width: 30 },//24
                    { header: 'CAD设计图', key: 'CADImg', width: 20 },
                    { header: '车瓷设计图', key: 'checiImg', width: 20 },
                    { header: '上瓷设计图', key: 'shangciImg', width: 20 },
                    { header: '到店设计图', key: 'daodianImg', width: 20 },
                    { header: '上釉设计图', key: 'shangyouImg', width: 20 },//29
                    { header: '试戴图片', key: 'tryImg', width: 30 },//30
                    { header: '试戴描述', key: 'tryRemark', width: 30 },
                    { header: '修复图片', key: 'recoverImg', width: 30 },//32
                    { header: '修复描述', key: 'recoverRemark', width: 30 },
                    { header: '医生戴牙调颔总结', key: 'docterSummary', width: 30 },
                    { header: '医生戴牙调颔图片', key: 'imgList', width: 30 },//35
                    { header: '瓷品', key: 'porcelain', width: 30 },
                    { header: '贴片颜色', key: 'tiepianColor', width: 30 },
                    { header: '是否做过矫正', key: 'adjust', width: 30 },
                    { header: '是否牙齿敏感', key: 'toothSensitivity', width: 30 },


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
                    // 只渲染图片，不显示图片URL
                    const imageKeys = [
                        'frontPhoto', 'leftFv', 'rightFv', 'front', 'leftFvEdge', 'rightFvEdge', 'intentImg',
                        'CADImg', 'checiImg', 'shangciImg', 'daodianImg', 'shangyouImg', 'designList', 'tryImg', 'recoverImg', 'imgList'
                    ];

                    const { tryInfo, recoverInfo, adjust, toothSensitivity } = row;

                    const _tryInfo = tryInfo ? JSON.parse(tryInfo) : [];
                    const _recoverInfo = recoverInfo ? JSON.parse(recoverInfo) : [];
                    row['tryRemark'] = _tryInfo.map(item => item.remark).join('\n');
                    row['recoverRemark'] = _recoverInfo.map(item => item.remark).join('\n');
                    const tryImg = _tryInfo?.[0]?.tryImg ?? []
                    const recoverImg = _recoverInfo?.[0]?.recoverImg ?? []

                    row['tryImg'] = JSON.stringify(tryImg)
                    row['recoverImg'] = JSON.stringify(recoverImg)
                    row['adjust'] = adjust ? '是' : '否'
                    row['toothSensitivity'] = toothSensitivity ? '是' : '否'



                    const rowData = { ...row };

                    imageKeys.forEach(key => rowData[key] = '');
                    const dataRow = worksheet.addRow(rowData);
                    dataRow.height = 100;
                    dataRow.alignment = { vertical: 'middle', horizontal: 'center' };
                    const rowNum = dataRow.number;


                    // 处理单张图片
                    await this.handleSingleImage(row.frontPhoto, worksheet, rowNum, 16, workbook);
                    await this.handleSingleImage(row.leftFv, worksheet, rowNum, 17, workbook);
                    await this.handleSingleImage(row.rightFv, worksheet, rowNum, 18, workbook);
                    await this.handleSingleImage(row.front, worksheet, rowNum, 19, workbook);
                    await this.handleSingleImage(row.leftFvEdge, worksheet, rowNum, 20, workbook);
                    await this.handleSingleImage(row.rightFvEdge, worksheet, rowNum, 21, workbook);
                    await this.handleSingleImage(row.intentImg, worksheet, rowNum, 23, workbook);

                    // 处理多张图片
                    await this.handleMultipleImages(row.CADImg, worksheet, rowNum, 25, workbook);
                    await this.handleMultipleImages(row.checiImg, worksheet, rowNum, 26, workbook);
                    await this.handleMultipleImages(row.shangciImg, worksheet, rowNum, 27, workbook);
                    await this.handleMultipleImages(row.daodianImg, worksheet, rowNum, 28, workbook);
                    await this.handleMultipleImages(row.shangyouImg, worksheet, rowNum, 29, workbook);
                    await this.handleMultipleImages(row.designList, worksheet, rowNum, 24, workbook);
                    await this.handleMultipleImages(row.tryImg, worksheet, rowNum, 30, workbook);
                    await this.handleMultipleImages(row.recoverImg, worksheet, rowNum, 32, workbook);
                    await this.handleMultipleImages(row.imgList, worksheet, rowNum, 35, workbook);
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
.password-container {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 10px;
    padding: 20px;
}

.password-input {
    display: flex;
    gap: 10px;
    width: 300px;
}

.password-error {
    color: #f56c6c;
    font-size: 14px;
}
</style>