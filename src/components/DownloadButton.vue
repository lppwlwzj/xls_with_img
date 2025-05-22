<template>
    <div class="download-container">
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

export default {
    name: 'DownloadButton',
    data() {
        return {
            isDownloading: false,
            downloadProgress: 0,
            downloadStatus: null,
            selectedFormat: 'xlsx',
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

        // 生成模拟的Excel数据
        async generateMockExcelData() {
            // 创建工作簿
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('员工信息');

            // 设置列宽
            worksheet.columns = [
                { header: '姓名', key: 'name', width: 15 },
                { header: '年龄', key: 'age', width: 10 },
                { header: '部门', key: 'department', width: 15 },
                { header: '职位', key: 'position', width: 15 },
                { header: '入职日期', key: 'joinDate', width: 15 },
                { header: '照片', key: 'photo', width: 30 } // 增加照片列宽度
            ];

            // 设置表头样式
            worksheet.getRow(1).font = { bold: true };
            worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
            worksheet.getRow(1).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFCCCCCC' }
            };

            // 准备数据
            const data = [
                { name: '张三', age: '28', department: '技术部', position: '工程师', joinDate: '2023-01-15' },
                { name: '李四', age: '32', department: '市场部', position: '经理', joinDate: '2022-06-01' },
                { name: '王五', age: '25', department: '人事部', position: '专员', joinDate: '2023-03-20' },
                { name: '赵六', age: '35', department: '财务部', position: '主管', joinDate: '2021-12-10' }
            ];

            // 添加数据行
            data.forEach((row) => {
                const dataRow = worksheet.addRow(row);
                dataRow.height = 100; // 增加行高以适应多张图片
                dataRow.alignment = { vertical: 'middle', horizontal: 'center' };
            });

            // 添加图片
            for (let i = 0; i < this.mockImages.length; i++) {
                const images = this.mockImages[i];
                for (let j = 0; j < images.length; j++) {
                    const base64 = await this.imageToBase64(images[j]);
                    if (base64) {
                        const imageId = workbook.addImage({
                            base64: base64.split(',')[1],
                            extension: 'jpeg',
                        });

                        // 计算图片位置，每张图片占据1/3的单元格宽度
                        const colWidth = 1 / images.length;
                        worksheet.addImage(imageId, {
                            tl: { col: 5 + (j * colWidth), row: i + 1 },
                            br: { col: 5 + ((j + 1) * colWidth), row: i + 2 }
                        });
                    }
                }
            }

            // 生成Excel文件
            const buffer = await workbook.xlsx.writeBuffer();
            return new Blob([buffer], {
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            });
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
                link.download = `员工信息表_${new Date().toLocaleDateString()}.${this.selectedFormat}`;

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