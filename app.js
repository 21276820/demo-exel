class DocumentEditor {
    constructor() {
        this.currentFile = null;
        this.currentFileType = null;
        this.currentData = null;
        this.documents = this.loadDocuments();
        this.init();
    }

    init() {
        this.bindEvents();
        this.renderDocumentsList();
        this.initTheme();
    }

    initTheme() {
        // 检查本地存储的主题偏好
        const savedTheme = localStorage.getItem('theme');
        const isDark = savedTheme === 'dark' || (!savedTheme && window.matchMedia('(prefers-color-scheme: dark)').matches);
        
        if (isDark) {
            document.body.classList.add('dark-theme');
            this.updateThemeButton('☀️');
        } else {
            this.updateThemeButton('🌙');
        }
        
        // 绑定主题切换按钮事件
        const themeToggle = document.getElementById('themeToggle');
        if (themeToggle) {
            themeToggle.addEventListener('click', () => this.toggleTheme());
        }
    }

    toggleTheme() {
        const isDark = document.body.classList.toggle('dark-theme');
        
        // 保存主题偏好
        localStorage.setItem('theme', isDark ? 'dark' : 'light');
        
        // 更新按钮图标
        this.updateThemeButton(isDark ? '☀️' : '🌙');
    }

    updateThemeButton(icon) {
        const themeToggle = document.getElementById('themeToggle');
        if (themeToggle) {
            themeToggle.textContent = icon;
        }
    }

    loadDocuments() {
        try {
            const stored = localStorage.getItem('documents');
            return stored ? JSON.parse(stored) : [];
        } catch (error) {
            console.error('加载文档失败:', error);
            return [];
        }
    }

    saveDocuments() {
        try {
            localStorage.setItem('documents', JSON.stringify(this.documents));
        } catch (error) {
            console.error('保存文档列表失败:', error);
        }
    }

    renderDocumentsList() {
        const documentsList = document.getElementById('documentsList');
        
        if (this.documents.length === 0) {
            documentsList.innerHTML = '<p class="no-documents">暂无已上传的文档</p>';
            return;
        }
        
        const html = this.documents.map((doc, index) => `
            <div class="document-item">
                <div class="document-item-header">
                    <div class="document-icon">
                        ${doc.type === 'docx' ? '📄' : '📊'}
                    </div>
                    <div class="document-info">
                        <div class="document-name">${doc.name}</div>
                        <div class="document-meta">
                            <span>${doc.type.toUpperCase()}</span>
                            <span>${doc.lastModified}</span>
                        </div>
                    </div>
                </div>
                <div class="document-actions">
                    <button class="open-btn" onclick="documentEditor.openDocument(${index})">打开</button>
                    <button class="delete-btn" onclick="documentEditor.deleteDocument(${index})">删除</button>
                </div>
            </div>
        `).join('');
        
        documentsList.innerHTML = html;
    }

    bindEvents() {
        const fileInput = document.getElementById('fileInput');
        const uploadBtn = document.getElementById('uploadBtn');
        const uploadArea = document.getElementById('uploadArea');
        const saveBtn = document.getElementById('saveBtn');
        const downloadBtn = document.getElementById('downloadBtn');
        const backBtn = document.getElementById('backBtn');

        uploadBtn.addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', (e) => this.handleFileUpload(e));
        
        // 拖拽上传
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });

        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });

        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                this.handleFile(files[0]);
            }
        });

        saveBtn.addEventListener('click', () => this.saveDocument());
        downloadBtn.addEventListener('click', () => this.downloadDocument());
        backBtn.addEventListener('click', () => this.goBack());
    }

    handleFileUpload(e) {
        const file = e.target.files[0];
        if (file) {
            this.handleFile(file);
        }
    }

    handleFile(file) {
        this.currentFile = file;
        this.currentFileType = file.name.split('.').pop().toLowerCase();
        
        if (!['docx', 'xlsx', 'xls'].includes(this.currentFileType)) {
            alert('不支持的文件格式，请上传 .docx, .xlsx 或 .xls 文件');
            return;
        }

        this.readFile(file);
    }

    readFile(file) {
        const reader = new FileReader();
        
        // 添加读取开始提示
        console.log('开始读取文件:', file.name);
        
        reader.onerror = (e) => {
            console.error('文件读取失败:', e);
            alert('文件读取失败，请重试');
        };
        
        if (this.currentFileType === 'docx') {
            reader.onload = (e) => {
                console.log('Word文件读取完成，开始渲染');
                const arrayBuffer = e.target.result;
                this.renderWordDocument(arrayBuffer);
            };
            reader.readAsArrayBuffer(file);
        } else {
            reader.onload = (e) => {
                console.log('Excel文件读取完成，开始渲染');
                const data = new Uint8Array(e.target.result);
                this.renderExcelDocument(data);
            };
            reader.readAsArrayBuffer(file);
        }
    }

    renderWordDocument(arrayBuffer) {
        const viewer = document.getElementById('documentViewer');
        const documentTitle = document.getElementById('documentTitle');
        
        try {
            console.log('开始处理Word文档:', this.currentFile.name);
            
            // 改进Word文档预览体验
            viewer.innerHTML = `
                <div style="padding: 20px;">
                    <h3>📄 Word文档预览</h3>
                    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin: 20px 0;">
                        <div style="padding: 15px; background-color: #f8f9fa; border-radius: 8px; border: 1px solid #e9ecef;">
                            <h4>文档信息</h4>
                            <p><strong>名称:</strong> ${this.currentFile.name}</p>
                            <p><strong>格式:</strong> DOCX</p>
                            <p><strong>大小:</strong> ${(this.currentFile.size / 1024).toFixed(2)} KB</p>
                            <p><strong>修改时间:</strong> ${this.currentFile.lastModifiedDate ? this.currentFile.lastModifiedDate.toLocaleString() : '未知'}</p>
                        </div>
                        <div style="padding: 15px; background-color: #e8f5e8; border-radius: 8px; border: 1px solid #c8e6c9;">
                            <h4>可用操作</h4>
                            <ul style="margin: 10px 0; padding-left: 20px;">
                                <li>查看文档基本信息</li>
                                <li>将文档添加到文档列表</li>
                                <li>下载原始文档</li>
                            </ul>
                        </div>
                    </div>
                    
                    <div style="padding: 20px; background-color: #fff3cd; border: 1px solid #ffeeba; border-radius: 8px; margin: 20px 0;">
                        <h4 style="color: #856404;">📋 预览说明</h4>
                        <p style="color: #856404; margin: 10px 0;">当前版本支持Excel文档的完整编辑功能。我们正在积极开发Word文档的完整预览和编辑功能，敬请期待！</p>
                        <p style="color: #856404;">您可以先使用Excel文档体验我们的编辑功能，或者直接下载Word文档查看原始内容。</p>
                    </div>
                    
                    <div style="margin-top: 20px;">
                        <button onclick="documentEditor.downloadOriginalDocument()" style="
                            background-color: #28a745;
                            color: white;
                            border: none;
                            padding: 12px 24px;
                            font-size: 1rem;
                            border-radius: 4px;
                            cursor: pointer;
                            margin-right: 10px;
                            transition: background-color 0.3s ease;
                        ">📥 下载原始文档</button>
                        <button onclick="documentEditor.goBack()" style="
                            background-color: #6c757d;
                            color: white;
                            border: none;
                            padding: 12px 24px;
                            font-size: 1rem;
                            border-radius: 4px;
                            cursor: pointer;
                            transition: background-color 0.3s ease;
                        ">🔙 返回文档列表</button>
                    </div>
                </div>
            `;
            
            // 将Word文档添加到文档列表
            this.saveDocumentToLocal();
            
            // 显示文档区域
            const uploadSection = document.querySelector('.upload-section');
            const documentsListSection = document.querySelector('.documents-list-section');
            const documentSection = document.getElementById('documentSection');
            
            uploadSection.style.display = 'none';
            documentsListSection.style.display = 'none';
            documentSection.style.display = 'block';
            documentTitle.textContent = this.currentFile.name;
            
            console.log('Word文档处理完成');
        } catch (error) {
            console.error('处理Word文档失败:', error);
            alert('处理Word文档失败，请重试');
        }
    }

    renderExcelDocument(data) {
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        this.currentData = workbook;
        
        // 将工作表转换为数组
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        // 手动生成简洁的HTML表格
        let html = '<table id="excelSheet" style="border-collapse: collapse; width: 100%; margin: 0; padding: 0; border: none;">';
        
        if (jsonData.length > 0) {
            // 生成表头
            html += '<thead><tr>';
            jsonData[0].forEach(cell => {
                html += `<th style="border: 1px solid var(--border-color); padding: 8px; background-color: var(--light-bg); font-weight: bold;">${cell || ''}</th>`;
            });
            html += '</tr></thead>';
            
            // 生成表格内容
            html += '<tbody>';
            for (let i = 1; i < jsonData.length; i++) {
                html += '<tr>';
                jsonData[i].forEach(cell => {
                    html += `<td style="border: 1px solid var(--border-color); padding: 8px; min-width: 100px;">${cell || ''}</td>`;
                });
                html += '</tr>';
            }
            html += '</tbody>';
        }
        
        html += '</table>';
        
        const viewer = document.getElementById('documentViewer');
        // 清空viewer内容，避免累积
        viewer.innerHTML = '';
        viewer.innerHTML = html;
        
        // 转换为可编辑表格
        this.makeExcelEditable();
        this.showDocumentSection();
        
        // 检查并处理空白div
        this.hideEmptyDivs();
    }
    
    hideEmptyDivs() {
        // 专门处理documentEditor div
        const documentEditor = document.getElementById('documentEditor');
        if (documentEditor) {
            documentEditor.style.display = 'none';
        }
        
        // 全局检查所有div，隐藏或移除空白div
        const allDivs = document.querySelectorAll('div');
        
        allDivs.forEach(div => {
            // 检查div是否为空
            const isEmpty = !div.textContent.trim() && 
                           !div.querySelector('img') &&
                           !div.querySelector('svg') &&
                           !div.querySelector('canvas') &&
                           !div.querySelector('table');
            
            if (isEmpty) {
                // 先尝试隐藏
                div.style.display = 'none';
                div.style.height = '0px';
                div.style.margin = '0px';
                div.style.padding = '0px';
                
                // 如果隐藏后仍然有问题，直接移除
                setTimeout(() => {
                    if (div.offsetHeight === 0) {
                        div.remove();
                    }
                }, 100);
            }
        });
        
        // 特别检查documentViewer内部的div
        const viewer = document.getElementById('documentViewer');
        const viewerDivs = viewer.querySelectorAll('div');
        
        viewerDivs.forEach(div => {
            if (!div.textContent.trim() && div.children.length === 0) {
                div.remove();
            }
        });
    }

    makeExcelEditable() {
        const table = document.getElementById('excelSheet');
        if (!table) return;
        
        const cells = table.querySelectorAll('td');
        cells.forEach(cell => {
            const value = cell.textContent;
            cell.innerHTML = `<input type="text" value="${value}" oninput="documentEditor.updateExcelData(this)">`;
        });
    }

    updateExcelData(input) {
        const cell = input.parentElement;
        const row = cell.parentElement;
        const table = row.parentElement.parentElement;
        
        const rowIndex = Array.from(table.querySelectorAll('tr')).indexOf(row);
        const colIndex = Array.from(row.querySelectorAll('td')).indexOf(cell);
        
        // 这里可以添加数据更新逻辑
        console.log(`更新单元格 (${rowIndex}, ${colIndex}): ${input.value}`);
    }

    showDocumentSection() {
        const uploadSection = document.querySelector('.upload-section');
        const documentsListSection = document.querySelector('.documents-list-section');
        const documentSection = document.getElementById('documentSection');
        const documentTitle = document.getElementById('documentTitle');
        
        uploadSection.style.display = 'none';
        documentsListSection.style.display = 'none';
        documentSection.style.display = 'block';
        documentTitle.textContent = this.currentFile.name;
    }

    saveDocument() {
        if (!this.currentFile) return;
        
        if (this.currentFileType === 'docx') {
            // Word文档保存逻辑
            alert('Word文档保存功能正在开发中');
        } else {
            // Excel文档保存逻辑
            try {
                // 获取表格元素
                const table = document.getElementById('excelSheet');
                if (!table) {
                    alert('未找到文档内容，无法保存');
                    return;
                }
                
                // 获取表格数据
                const rows = table.querySelectorAll('tr');
                const data = [];
                
                rows.forEach(row => {
                    const rowData = [];
                    const cells = row.querySelectorAll('td, th');
                    
                    cells.forEach(cell => {
                        const input = cell.querySelector('input');
                        const value = input ? input.value : cell.textContent;
                        rowData.push(value);
                    });
                    
                    data.push(rowData);
                });
                
                // 创建新的工作表
                const worksheet = XLSX.utils.aoa_to_sheet(data);
                
                // 更新工作簿
                const newWorkbook = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(newWorkbook, worksheet, this.currentData.SheetNames[0]);
                
                // 更新当前数据
                this.currentData = newWorkbook;
                
                // 保存到localStorage
                this.saveDocumentToLocal();
                
                alert('文档保存成功！');
            } catch (error) {
                console.error('保存文档失败:', error);
                alert('保存文档失败，请重试');
            }
        }
    }

    saveDocumentToLocal() {
        if (!this.currentFile) return;
        
        try {
            let data = null;
            
            if (this.currentFileType === 'xlsx' || this.currentFileType === 'xls') {
                // 将Excel数据转换为base64格式保存
                const wbout = XLSX.write(this.currentData, { bookType: this.currentFileType, type: 'base64' });
                data = wbout;
            } else if (this.currentFileType === 'docx') {
                // 对于Word文档，我们需要保存原始文件数据
                // 这里可以根据实际需求调整保存方式
                data = 'docx-content'; // 临时占位符
            }
            
            // 创建文档对象
            const document = {
                name: this.currentFile.name,
                type: this.currentFileType,
                data: data,
                lastModified: new Date().toLocaleString()
            };
            
            // 检查是否已存在同名文档
            const existingIndex = this.documents.findIndex(doc => doc.name === this.currentFile.name);
            if (existingIndex !== -1) {
                // 更新现有文档
                this.documents[existingIndex] = document;
            } else {
                // 添加新文档
                this.documents.push(document);
            }
            
            // 保存到localStorage
            localStorage.setItem('documents', JSON.stringify(this.documents));
            
            // 重新渲染文档列表
            this.renderDocumentsList();
        } catch (error) {
            console.error('保存到本地失败:', error);
        }
    }

    downloadDocument() {
        if (!this.currentFile) return;
        
        if (this.currentFileType === 'docx') {
            // Word文档下载
            this.downloadOriginalDocument();
        } else {
            // Excel文档下载
            const worksheet = this.currentData.Sheets[this.currentData.SheetNames[0]];
            const newWorkbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWorkbook, worksheet, this.currentData.SheetNames[0]);
            
            XLSX.writeFile(newWorkbook, this.currentFile.name);
        }
    }

    downloadOriginalDocument() {
        if (!this.currentFile) return;
        
        try {
            // 创建下载链接
            const a = document.createElement('a');
            a.href = URL.createObjectURL(this.currentFile);
            a.download = this.currentFile.name;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            
            console.log('文档下载成功:', this.currentFile.name);
        } catch (error) {
            console.error('下载文档失败:', error);
            alert('下载文档失败，请重试');
        }
    }

    goBack() {
        const uploadSection = document.querySelector('.upload-section');
        const documentSection = document.getElementById('documentSection');
        const documentsListSection = document.querySelector('.documents-list-section');
        
        uploadSection.style.display = 'block';
        documentsListSection.style.display = 'block';
        documentSection.style.display = 'none';
        
        // 重置文件输入
        document.getElementById('fileInput').value = '';
        this.currentFile = null;
        this.currentFileType = null;
        this.currentData = null;
    }

    openDocument(index) {
        const doc = this.documents[index];
        if (!doc) return;
        
        try {
            // 模拟打开文档
            this.currentFile = { name: doc.name };
            this.currentFileType = doc.type;
            
            // 如果是Excel文档，需要处理数据
            if (doc.type === 'xlsx' || doc.type === 'xls') {
                const workbook = XLSX.read(doc.data, { type: 'base64' });
                this.currentData = workbook;
                
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const html = XLSX.utils.sheet_to_html(worksheet, { id: 'excelSheet' });
                
                const viewer = document.getElementById('documentViewer');
                viewer.innerHTML = html;
                this.makeExcelEditable();
            }
            
            this.showDocumentSection();
        } catch (error) {
            console.error('打开文档失败:', error);
            alert('打开文档失败，请重试');
        }
    }

    deleteDocument(index) {
        if (confirm('确定要删除此文档吗？')) {
            this.documents.splice(index, 1);
            this.saveDocuments();
            this.renderDocumentsList();
        }
    }
}

// 初始化应用
const documentEditor = new DocumentEditor();
