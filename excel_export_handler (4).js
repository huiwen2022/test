/**
 * Excel 匯出處理器 v2.0
 * 使用 ExcelJS 匯出 Tabulator 表格資料
 * 
 * 功能特色：
 * - 表格從 B2 開始，A1 為標題區域
 * - 支援篩選結果匯出
 * - 部門欄位有對應顏色
 * - Email 欄位粉黃色背景
 * - 標題凍結功能
 * - 自動調整欄寬
 * 
 * @author Your Name
 * @version 2.0
 */

class ExcelExportHandler {
    constructor() {
        this.initializeConfig();
    }

    /**
     * 初始化配置
     */
    initializeConfig() {
        // 部門顏色配置
        this.departmentColors = {
            '資訊部': 'FFE6F2FF',      // 淺藍色
            '行銷部': 'FFE6FFE6',      // 淺綠色
            '人事部': 'FFFFF2E6',      // 淺橘色
            '財務部': 'FFFFE6E6',      // 淺紅色
            '業務部': 'FFF2E6FF',      // 淺紫色
        };

        // 特殊欄位顏色
        this.specialColors = {
            email: 'FFFFFF99',         // 粉黃色
            title_blue: 'FF4472C4',    // 藍色
            title_orange: 'FFFD7E14',  // 橘色
            white: 'FFFFFFFF'          // 白色
        };

        // 欄位順序和寬度配置
        this.columnConfig = [
            { key: 'id', header: 'ID', width: 10 },
            { key: 'name', header: '姓名', width: 15 },
            { key: 'email', header: '電子郵件', width: 25 },
            { key: 'age', header: '年齡', width: 10 },
            { key: 'department', header: '部門', width: 15 },
            { key: 'description', header: '工作描述', width: 30 }
        ];

        // 基本樣式配置
        this.styles = this.createStyles();
    }

    /**
     * 建立樣式配置
     */
    createStyles() {
        return {
            // 標題樣式（藍色）
            headerBlue: {
                fill: {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: this.specialColors.title_blue }
                },
                font: {
                    color: { argb: 'FFFFFFFF' },
                    bold: true,
                    size: 12
                },
                alignment: {
                    horizontal: 'center',
                    vertical: 'middle'
                },
                border: this.createBorder()
            },

            // 標題樣式（粉黃色）
            headerYellow: {
                fill: {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: this.specialColors.email }
                },
                font: {
                    color: { argb: 'FF000000' },
                    bold: true,
                    size: 12
                },
                alignment: {
                    horizontal: 'center',
                    vertical: 'middle'
                },
                border: this.createBorder()
            },

            // 主標題樣式
            mainTitle: {
                font: {
                    color: { argb: 'FFFFFFFF' },
                    bold: true,
                    size: 14
                },
                alignment: {
                    horizontal: 'center',
                    vertical: 'middle'
                },
                border: this.createBorder()
            },

            // 資料行樣式
            dataCell: {
                font: { size: 11 },
                alignment: {
                    horizontal: 'left',
                    vertical: 'middle',
                    wrapText: true
                },
                border: this.createBorder('thin', 'FFD0D0D0')
            }
        };
    }

    /**
     * 建立邊框樣式
     */
    createBorder(style = 'thin', color = 'FF000000') {
        const borderConfig = { style, color: { argb: color } };
        return {
            top: borderConfig,
            left: borderConfig,
            bottom: borderConfig,
            right: borderConfig
        };
    }

    /**
     * 主要匯出方法
     */
    async exportTableToExcel(table, filename = 'table_export', options = {}) {
        try {
            console.log('=== 開始 Excel 匯出 ===');
            
            // 取得篩選後的資料
            const filteredData = this.getFilteredData(table);
            
            if (filteredData.length === 0) {
                alert('沒有資料可以匯出');
                return;
            }

            // 建立工作簿和工作表
            const { workbook, worksheet } = this.createWorkbook();

            // 設定標題區域 (A1, B1)
            this.setTitleArea(worksheet, options.tableName || '員工資料表');

            // 設定表格標題 (B2開始)
            this.setTableHeaders(worksheet);

            // 新增資料 (B3開始)
            this.addTableData(worksheet, filteredData);

            // 最終設定
            this.finalizeWorksheet(worksheet, filteredData);

            // 下載檔案
            await this.downloadWorkbook(workbook, filename);

            this.logExportSuccess(filename, filteredData.length);

        } catch (error) {
            this.handleExportError(error);
        }
    }

    /**
     * 取得篩選後的資料
     */
    getFilteredData(table) {
        const rows = table.getRows('visible');
        return rows.map(row => {
            const data = row.getData();
            // 清理內部欄位
            const cleanData = { ...data };
            delete cleanData._editing;
            delete cleanData.actions;
            return cleanData;
        });
    }

    /**
     * 建立工作簿和工作表
     */
    createWorkbook() {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('資料表', {
            views: [{
                state: 'frozen',
                ySplit: 2,  // 凍結前兩行（標題和表格標題）
                xSplit: 1   // 凍結第一列（A欄）
            }]
        });

        return { workbook, worksheet };
    }

    /**
     * 設定標題區域 (A1合併儲存格)
     */
    setTitleArea(worksheet, tableName) {
        const currentYear = new Date().getFullYear();
        const fullTitle = `${tableName} ${currentYear}年`;

        // 計算需要合併的欄位範圍（根據表格欄位數量）
        const lastColumn = String.fromCharCode(65 + this.columnConfig.length); // A + 欄位數量
        const mergeRange = `A1:${lastColumn}1`;

        // 合併儲存格
        worksheet.mergeCells(mergeRange);

        // 設定合併後的儲存格內容和樣式
        const titleCell = worksheet.getCell('A1');
        titleCell.value = fullTitle;
        
        // 套用樣式
        Object.assign(titleCell, this.styles.mainTitle);
        titleCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: this.specialColors.title_blue }
        };

        // 設定字體大小稍大一些
        titleCell.font = {
            ...titleCell.font,
            size: 16,
            bold: true
        };

        // 設定第一行高度
        worksheet.getRow(1).height = 35;

        console.log(`標題區域已設定: ${mergeRange} - "${fullTitle}"`);
    }

    /**
     * 設定表格標題 (B2開始)
     */
    setTableHeaders(worksheet) {
        const headerRow = worksheet.getRow(2);
        headerRow.height = 25;

        this.columnConfig.forEach((config, index) => {
            const columnLetter = String.fromCharCode(66 + index); // B, C, D...
            const cell = worksheet.getCell(`${columnLetter}2`);
            
            cell.value = config.header;
            
            // 根據欄位類型設定樣式
            if (config.key === 'email') {
                Object.assign(cell, this.styles.headerYellow);
            } else {
                Object.assign(cell, this.styles.headerBlue);
            }

            // 設定欄寬
            worksheet.getColumn(columnLetter).width = config.width;
        });

        // 設定 A 欄寬度（保持適當寬度用於視覺平衡）
        worksheet.getColumn('A').width = 5;
        
        console.log(`表格標題已設定: B2 到 ${String.fromCharCode(65 + this.columnConfig.length)}2`);
    }

    /**
     * 新增表格資料 (B3開始)
     */
    addTableData(worksheet, data) {
        data.forEach((rowData, rowIndex) => {
            const excelRowIndex = rowIndex + 3; // 從第3行開始
            
            this.columnConfig.forEach((config, colIndex) => {
                const columnLetter = String.fromCharCode(66 + colIndex); // B, C, D...
                const cell = worksheet.getCell(`${columnLetter}${excelRowIndex}`);
                
                // 設定儲存格值
                cell.value = rowData[config.key] || '';
                
                // 套用基本樣式
                Object.assign(cell, this.styles.dataCell);
                
                // 設定背景色
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: this.getCellBackgroundColor(config.key, rowData) }
                };

                // 特殊欄位處理
                this.applyCellSpecialFormatting(cell, config.key);
            });

            // 設定行高
            const row = worksheet.getRow(excelRowIndex);
            row.height = this.calculateRowHeight(rowData.description);
        });
    }

    /**
     * 取得儲存格背景色
     */
    getCellBackgroundColor(columnKey, rowData) {
        switch (columnKey) {
            case 'department':
                return this.departmentColors[rowData.department] || this.specialColors.white;
            case 'email':
                return this.specialColors.email;
            default:
                return this.specialColors.white;
        }
    }

    /**
     * 套用儲存格特殊格式
     */
    applyCellSpecialFormatting(cell, columnKey) {
        switch (columnKey) {
            case 'age':
                cell.alignment.horizontal = 'center';
                break;
            case 'description':
                cell.alignment.wrapText = true;
                break;
        }
    }

    /**
     * 計算行高
     */
    calculateRowHeight(description) {
        if (!description) return 20;
        const lines = description.split('\n').length;
        return Math.min(lines * 15 + 10, 100);
    }

    /**
     * 工作表最終設定
     */
    finalizeWorksheet(worksheet, data) {
        // 自動調整欄寬已在 setTableHeaders 中處理
        // 這裡可以添加其他最終設定
    }

    /**
     * 下載工作簿
     */
    async downloadWorkbook(workbook, filename) {
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });

        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `${filename}.xlsx`;

        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

        window.URL.revokeObjectURL(url);
    }

    /**
     * 記錄匯出成功
     */
    logExportSuccess(filename, dataCount) {
        console.log('=== Excel 匯出成功 ===');
        console.log(`檔案名稱: ${filename}.xlsx`);
        console.log(`匯出資料筆數: ${dataCount}`);
        console.log(`表格位置: B2 開始，A1 為合併標題`);
        console.log(`匯出時間: ${new Date().toLocaleString()}`);
        console.log('========================');
    }

    /**
     * 處理匯出錯誤
     */
    handleExportError(error) {
        console.error('=== Excel 匯出失敗 ===');
        console.error('錯誤詳情:', error);
        console.error('====================');
        alert('Excel 匯出失敗，請查看 Console 錯誤訊息');
    }

    // ==================== 公開方法 ====================

    /**
     * 匯出篩選後的資料
     */
    exportVisible(table, filename = 'filtered_data', tableName = '員工資料表') {
        return this.exportTableToExcel(table, filename, {
            onlyVisible: true,
            tableName: tableName
        });
    }

    /**
     * 匯出所有資料
     */
    exportAll(table, filename = 'all_data', tableName = '完整員工資料表') {
        const hasFilters = table.getFilters().length > 0;

        if (hasFilters) {
            table.clearFilter();
        }

        return this.exportTableToExcel(table, filename, {
            onlyVisible: false,
            tableName: tableName
        }).then(() => {
            if (hasFilters) {
                console.log('提示：篩選已被清除，請重新設定篩選條件');
            }
        });
    }

    /**
     * 匯出特定部門的資料
     */
    exportByDepartment(table, department, filename, tableName) {
        table.setFilter('department', '=', department);

        const departmentFilename = filename || `${department}_data`;
        const departmentTableName = tableName || `${department}資料表`;
        
        return this.exportTableToExcel(table, departmentFilename, {
            tableName: departmentTableName
        });
    }

    /**
     * 批次匯出所有部門
     */
    async exportAllDepartments(table) {
        const departments = Object.keys(this.departmentColors);
        
        for (const department of departments) {
            await this.exportByDepartment(
                table, 
                department, 
                `${department}_${new Date().getFullYear()}`,
                `${department}員工名單`
            );
            
            // 等待一下避免瀏覽器阻擋多個下載
            await new Promise(resolve => setTimeout(resolve, 500));
        }
        
        console.log('所有部門資料匯出完成');
    }

    /**
     * 取得可用的部門列表
     */
    getAvailableDepartments() {
        return Object.keys(this.departmentColors);
    }

    /**
     * 更新部門顏色配置
     */
    updateDepartmentColors(newColors) {
        this.departmentColors = { ...this.departmentColors, ...newColors };
    }

    /**
     * 更新欄位配置
     */
    updateColumnConfig(newConfig) {
        this.columnConfig = newConfig;
    }
}

// ==================== 使用範例 ====================

/*
// 基本使用
const excelExporter = new ExcelExportHandler();

// 匯出篩選資料
excelExporter.exportVisible(table, 'filtered_data', '篩選後員工資料');

// 匯出所有資料  
excelExporter.exportAll(table, 'all_data', '完整員工名單');

// 匯出特定部門
excelExporter.exportByDepartment(table, '資訊部', 'IT_dept', '資訊部員工清單');

// 批次匯出所有部門
excelExporter.exportAllDepartments(table);

// 取得可用部門
const departments = excelExporter.getAvailableDepartments();

// 更新部門顏色
excelExporter.updateDepartmentColors({
    '新部門': 'FFFF00FF'  // 紫色
});
*/

// 瀏覽器環境註冊
if (typeof window !== 'undefined') {
    window.ExcelExportHandler = ExcelExportHandler;
}

// Node.js 環境匯出
if (typeof module !== 'undefined' && module.exports) {
    module.exports = ExcelExportHandler;
}