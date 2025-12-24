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
        
        // 强制隐藏documentEditor DIV
        const documentEditor = document.getElementById('documentEditor');
        if (documentEditor) {
            documentEditor.style.display = 'none';
            documentEditor.style.height = '0px';
            documentEditor.style.margin = '0px';
            documentEditor.style.padding = '0px';
            documentEditor.innerHTML = '';
        }
        
        if (this.documents.length === 0) {
            documentsList.innerHTML = '<p class="no-documents">暂无已上传的文档</p>';
            return;
        }
        
        // 按最后修改时间排序，最新的在最前面
        const sortedDocuments = [...this.documents].sort((a, b) => {
            // 将时间字符串转换为Date对象进行比较
            const dateA = new Date(a.lastModified);
            const dateB = new Date(b.lastModified);
            // 倒序排序，最新的在前
            return dateB - dateA;
        });
        
        const html = sortedDocuments.map((doc, index) => {
            // 找到原始索引，用于打开和删除操作
            const originalIndex = this.documents.findIndex(d => d.name === doc.name && d.lastModified === doc.lastModified);
            
            return `
            <div class="document-item">
                <div class="document-item-header">
                    <div class="document-icon">
                            ${doc.type === 'docx' ? '📄' : (doc.type === 'pdf' ? '📑' : (doc.type === 'txt' ? '📝' : '📊'))}
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
                    <button class="open-btn" onclick="documentEditor.openDocument(${originalIndex})">打开</button>
                    <button class="delete-btn" onclick="documentEditor.deleteDocument(${originalIndex})">删除</button>
                </div>
            </div>
        `;
        }).join('');
        
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
        
        if (!['docx', 'xlsx', 'xls', 'pdf', 'txt'].includes(this.currentFileType)) {
            alert('不支持的文件格式，请上传 .docx, .xlsx, .xls, .pdf 或 .txt 文件');
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
        } else if (this.currentFileType === 'pdf') {
            reader.onload = (e) => {
                console.log('PDF文件读取完成，开始渲染');
                const arrayBuffer = e.target.result;
                this.renderPdfDocument(arrayBuffer);
            };
            reader.readAsArrayBuffer(file);
        } else if (this.currentFileType === 'txt') {
            reader.onload = (e) => {
                console.log('TXT文件读取完成，开始渲染');
                const text = e.target.result;
                this.renderTextDocument(text);
            };
            reader.readAsText(file);
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
                        <div style="padding: 15px; background-color: var(--light-bg); border-radius: 8px; border: 1px solid var(--border-color);">
                            <h4>文档信息</h4>
                            <p><strong>名称:</strong> ${this.currentFile.name}</p>
                            <p><strong>格式:</strong> DOCX</p>
                            <p><strong>大小:</strong> ${(this.currentFile.size / 1024).toFixed(2)} KB</p>
                            <p><strong>修改时间:</strong> ${this.currentFile.lastModifiedDate ? this.currentFile.lastModifiedDate.toLocaleString() : '未知'}</p>
                        </div>
                        <div style="padding: 15px; background-color: var(--light-bg); border-radius: 8px; border: 1px solid var(--border-color);">
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
                            background-color: var(--secondary-color);
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
                            background-color: var(--muted-color);
                            color: var(--text-color);
                            border: none;
                            padding: 12px 24px;
                            font-size: 1rem;
                            border-radius: 4px;
                            cursor: pointer;
                            transition: background-color 0.3s ease;
                        ">← 返回文档列表</button>
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
    
    renderTextDocument(text) {
        const viewer = document.getElementById('documentViewer');
        const documentTitle = document.getElementById('documentTitle');
        
        try {
            console.log('开始处理TXT文档:', this.currentFile.name);
            
            // 渲染文本文件内容
            viewer.innerHTML = `
                <div style="padding: 20px;">
                    <h3>📄 文本文件编辑</h3>
                    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin: 20px 0;">
                        <div style="padding: 15px; background-color: var(--light-bg); border-radius: 8px; border: 1px solid var(--border-color);">
                            <h4>文档信息</h4>
                            <p><strong>名称:</strong> ${this.currentFile.name}</p>
                            <p><strong>格式:</strong> TXT</p>
                            <p><strong>大小:</strong> ${this.currentFile.size ? (this.currentFile.size / 1024).toFixed(2) : '未知'} KB</p>
                            <p><strong>修改时间:</strong> ${this.currentFile.lastModifiedDate ? this.currentFile.lastModifiedDate.toLocaleString() : '未知'}</p>
                        </div>
                        <div style="padding: 15px; background-color: var(--light-bg); border-radius: 8px; border: 1px solid var(--border-color);">
                            <h4>可用操作</h4>
                            <ul style="margin: 10px 0; padding-left: 20px;">
                                <li>在线编辑文本内容</li>
                                <li>保存修改后的内容</li>
                                <li>下载原始文档</li>
                            </ul>
                        </div>
                    </div>
                    
                    <div style="margin: 20px 0; padding: 20px; background-color: var(--light-bg); border-radius: 8px; border: 1px solid var(--border-color);">
                        <textarea id="txtEditor" style="width: 100%; height: 400px; font-family: monospace; font-size: 1rem; line-height: 1.5; padding: 10px; border: 1px solid #ddd; border-radius: 4px; resize: vertical; background-color: white; color: #333;">${text}</textarea>
                    </div>
                    
                    <div style="margin-top: 20px;">
                        <button onclick="documentEditor.downloadOriginalDocument()" style="
                            background-color: var(--secondary-color);
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
                            background-color: var(--muted-color);
                            color: var(--text-color);
                            border: none;
                            padding: 12px 24px;
                            font-size: 1rem;
                            border-radius: 4px;
                            cursor: pointer;
                            transition: background-color 0.3s ease;
                        ">← 返回文档列表</button>
                    </div>
                </div>
            `;
            
            // 只有在从handleFile调用时才需要保存到文档列表，从openDocument调用时不需要
            if (this.currentFile.size) {
                this.saveDocumentToLocal();
            }
            
            // 显示文档区域
            const uploadSection = document.querySelector('.upload-section');
            const documentsListSection = document.querySelector('.documents-list-section');
            const documentSection = document.getElementById('documentSection');
            
            uploadSection.style.display = 'none';
            documentsListSection.style.display = 'none';
            documentSection.style.display = 'block';
            documentTitle.textContent = this.currentFile.name;
            
            console.log('TXT文档处理完成');
        } catch (error) {
            console.error('处理TXT文档失败:', error);
            alert('处理TXT文档失败，请重试');
        }
    }
    
    renderPdfDocument(arrayBuffer) {
        const viewer = document.getElementById('documentViewer');
        const documentTitle = document.getElementById('documentTitle');
        
        try {
            console.log('开始处理PDF文档:', this.currentFile.name);
            
            // 创建PDF渲染容器
            viewer.innerHTML = `
                <div style="padding: 20px;">
                    <h3>📄 PDF文档预览</h3>
                    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin: 20px 0;">
                        <div style="padding: 15px; background-color: var(--light-bg); border-radius: 8px; border: 1px solid var(--border-color);">
                            <h4>文档信息</h4>
                            <p><strong>名称:</strong> ${this.currentFile.name}</p>
                            <p><strong>格式:</strong> PDF</p>
                            <p><strong>大小:</strong> ${(this.currentFile.size / 1024).toFixed(2)} KB</p>
                            <p><strong>修改时间:</strong> ${this.currentFile.lastModifiedDate ? this.currentFile.lastModifiedDate.toLocaleString() : '未知'}</p>
                        </div>
                        <div style="padding: 15px; background-color: var(--light-bg); border-radius: 8px; border: 1px solid var(--border-color);">
                            <h4>可用操作</h4>
                            <ul style="margin: 10px 0; padding-left: 20px;">
                                <li>在线预览PDF内容</li>
                                <li>将文档添加到文档列表</li>
                                <li>下载原始文档</li>
                            </ul>
                        </div>
                    </div>
                    
                    <div id="pdfContainer" style="
                        margin: 20px 0;
                        padding: 20px;
                        background-color: #f8f9fa;
                        border-radius: 8px;
                        border: 1px solid #e9ecef;
                        overflow: auto;
                    ">
                        <div style="text-align: center; padding: 40px;">
                            <div style="font-size: 3rem; margin-bottom: 15px;">📄</div>
                            <p>正在加载PDF文档，请稍候...</p>
                        </div>
                    </div>
                    
                    <div style="margin-top: 20px; text-align: center;">
                        <button onclick="documentEditor.downloadOriginalDocument()" style="
                            background-color: var(--secondary-color);
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
                            background-color: var(--muted-color);
                            color: var(--text-color);
                            border: none;
                            padding: 12px 24px;
                            font-size: 1rem;
                            border-radius: 4px;
                            cursor: pointer;
                            transition: background-color 0.3s ease;
                        ">← 返回文档列表</button>
                    </div>
                </div>
            `;
            
            // 渲染PDF内容
            this.renderPdfContent(arrayBuffer);
            
            // 将PDF文档添加到文档列表
            this.saveDocumentToLocal();
            
            // 显示文档区域
            const uploadSection = document.querySelector('.upload-section');
            const documentsListSection = document.querySelector('.documents-list-section');
            const documentSection = document.getElementById('documentSection');
            
            uploadSection.style.display = 'none';
            documentsListSection.style.display = 'none';
            documentSection.style.display = 'block';
            documentTitle.textContent = this.currentFile.name;
            
            console.log('PDF文档处理完成');
        } catch (error) {
            console.error('处理PDF文档失败:', error);
            alert('处理PDF文档失败，请重试');
        }
    }
    
    renderPdfContent(arrayBuffer) {
        const pdfContainer = document.getElementById('pdfContainer');
        
        // 设置PDF.js worker
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        
        // 加载PDF文档
        pdfjsLib.getDocument(arrayBuffer).promise.then(pdf => {
            console.log('PDF加载成功，共', pdf.numPages, '页');
            
            // 清空容器
            pdfContainer.innerHTML = '';
            
            // 创建页面容器
            const pagesContainer = document.createElement('div');
            pagesContainer.style.display = 'flex';
            pagesContainer.style.flexDirection = 'column';
            pagesContainer.style.alignItems = 'center';
            pagesContainer.style.gap = '20px';
            
            // 渲染每一页
            for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
                pdf.getPage(pageNum).then(page => {
                    // 创建canvas元素
                    const canvas = document.createElement('canvas');
                    canvas.style.maxWidth = '100%';
                    canvas.style.border = '1px solid #e0e0e0';
                    canvas.style.borderRadius = '4px';
                    
                    const container = document.createElement('div');
                    container.style.width = '100%';
                    container.style.textAlign = 'center';
                    container.appendChild(canvas);
                    
                    // 设置渲染选项
                    const viewport = page.getViewport({ scale: 1.5 });
                    canvas.height = viewport.height;
                    canvas.width = viewport.width;
                    
                    // 渲染页面
                    const renderContext = {
                        canvasContext: canvas.getContext('2d'),
                        viewport: viewport
                    };
                    
                    page.render(renderContext).promise.then(() => {
                        console.log('PDF页面', pageNum, '渲染成功');
                    }).catch(err => {
                        console.error('渲染PDF页面', pageNum, '失败:', err);
                        container.innerHTML = `<div style="padding: 20px; color: #e74c3c;">页面 ${pageNum} 渲染失败</div>`;
                    });
                    
                    pagesContainer.appendChild(container);
                }).catch(err => {
                    console.error('获取PDF页面失败:', err);
                    const errorDiv = document.createElement('div');
                    errorDiv.style.padding = '20px';
                    errorDiv.style.color = '#e74c3c';
                    errorDiv.textContent = `获取页面 ${pageNum} 失败`;
                    pagesContainer.appendChild(errorDiv);
                });
            }
            
            pdfContainer.appendChild(pagesContainer);
        }).catch(err => {
            console.error('加载PDF文档失败:', err);
            pdfContainer.innerHTML = `
                <div style="text-align: center; padding: 40px; color: #e74c3c;">
                    <div style="font-size: 3rem; margin-bottom: 15px;">❌</div>
                    <p>PDF文档加载失败，请重试</p>
                    <p style="font-size: 0.9rem; margin-top: 10px;">错误信息: ${err.message}</p>
                </div>
            `;
        });
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
        // 专门处理documentEditor div，强制隐藏
        const documentEditor = document.getElementById('documentEditor');
        if (documentEditor) {
            documentEditor.style.display = 'none';
            documentEditor.style.height = '0px';
            documentEditor.style.margin = '0px';
            documentEditor.style.padding = '0px';
            documentEditor.innerHTML = ''; // 清空内容
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
        } else if (this.currentFileType === 'pdf') {
            // PDF文档保存逻辑 - 直接保存成功，因为不可编辑
            alert('文档保存成功！');
        } else if (this.currentFileType === 'txt') {
            // TXT文档保存逻辑
            try {
                // 获取textarea中的内容
                const txtEditor = document.getElementById('txtEditor');
                if (!txtEditor) {
                    alert('未找到文档内容，无法保存');
                    return;
                }
                
                const textContent = txtEditor.value;
                
                // 更新当前文件的修改时间
                this.currentFile.lastModifiedDate = new Date();
                
                // 保存到localStorage
                this.saveDocumentToLocal(textContent);
                
                alert('文档保存成功！');
            } catch (error) {
                console.error('保存文档失败:', error);
                alert('保存文档失败，请重试');
            }
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

    saveDocumentToLocal(textContent = null) {
        if (!this.currentFile) return;
        
        try {
            let data = null;
            
            if (this.currentFileType === 'xlsx' || this.currentFileType === 'xls') {
                // 将Excel数据转换为base64格式保存
                const wbout = XLSX.write(this.currentData, { bookType: this.currentFileType, type: 'base64' });
                data = wbout;
            } else if (this.currentFileType === 'txt') {
                // 对于TXT文档，保存文本内容
                if (textContent) {
                    // 如果提供了修改后的内容，直接使用
                    const document = {
                        name: this.currentFile.name,
                        type: this.currentFileType,
                        data: textContent,
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
                } else {
                    // 如果没有提供修改后的内容，使用文件原始内容
                    const reader = new FileReader();
                    reader.onload = (e) => {
                        const textData = e.target.result; // 获取文本内容
                        
                        // 创建文档对象
                        const document = {
                            name: this.currentFile.name,
                            type: this.currentFileType,
                            data: textData,
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
                    };
                    reader.readAsText(this.currentFile);
                }
                return;
            } else if (this.currentFileType === 'docx' || this.currentFileType === 'pdf') {
                // 对于Word和PDF文档，保存原始文件的base64数据
                const reader = new FileReader();
                reader.onload = (e) => {
                    const base64Data = e.target.result.split(',')[1]; // 获取base64数据部分
                    
                    // 创建文档对象
                    const document = {
                        name: this.currentFile.name,
                        type: this.currentFileType,
                        data: base64Data,
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
                };
                reader.readAsDataURL(this.currentFile);
                return; // 异步操作，直接返回
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
        
        if (this.currentFileType === 'docx' || this.currentFileType === 'pdf' || this.currentFileType === 'txt') {
            // Word、PDF和TXT文档下载
            this.downloadOriginalDocument();
        } else {
            // Excel文档下载
            try {
                if (this.currentData && this.currentData.Sheets) {
                    const worksheet = this.currentData.Sheets[this.currentData.SheetNames[0]];
                    const newWorkbook = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(newWorkbook, worksheet, this.currentData.SheetNames[0]);
                    
                    XLSX.writeFile(newWorkbook, this.currentFile.name);
                } else {
                    // 如果没有currentData，尝试使用原始文档下载
                    this.downloadOriginalDocument();
                }
            } catch (error) {
                console.error('Excel下载失败，尝试使用原始文档下载:', error);
                // 降级处理：使用原始文档下载
                this.downloadOriginalDocument();
            }
        }
    }

    downloadOriginalDocument() {
        if (!this.currentFile) return;
        
        try {
            // 查找当前文档在列表中的索引
            const docIndex = this.documents.findIndex(doc => 
                doc.name === this.currentFile.name && doc.type === this.currentFileType
            );
            
            if (docIndex !== -1) {
                const doc = this.documents[docIndex];
                
                // 如果是Excel文档，使用当前数据下载
                if (doc.type === 'xlsx' || doc.type === 'xls') {
                    if (this.currentData) {
                        const worksheet = this.currentData.Sheets[this.currentData.SheetNames[0]];
                        const newWorkbook = XLSX.utils.book_new();
                        XLSX.utils.book_append_sheet(newWorkbook, worksheet, this.currentData.SheetNames[0]);
                        XLSX.writeFile(newWorkbook, doc.name);
                        return;
                    }
                } else {
                    // 对于不同类型的文档进行不同处理
                    const base64Data = doc.data;
                    if (base64Data && base64Data !== 'docx-content' && base64Data !== 'pdf-content') {
                        // 创建Blob对象
                        let blob;
                        let mimeType;
                        
                        if (doc.type === 'txt') {
                            // 对于TXT文档，直接使用文本数据，不进行Base64解码
                            mimeType = 'text/plain;charset=utf-8';
                            blob = new Blob([base64Data], { type: mimeType });
                        } else {
                            // 对于Word和PDF文档，从base64数据下载
                            if (doc.type === 'docx') {
                                mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
                            } else if (doc.type === 'pdf') {
                                mimeType = 'application/pdf';
                            } else {
                                mimeType = 'application/octet-stream';
                            }
                            const byteCharacters = atob(base64Data);
                            const byteNumbers = new Array(byteCharacters.length);
                            for (let i = 0; i < byteCharacters.length; i++) {
                                byteNumbers[i] = byteCharacters.charCodeAt(i);
                            }
                            const byteArray = new Uint8Array(byteNumbers);
                            blob = new Blob([byteArray], { type: mimeType });
                        }
                        
                        // 创建下载链接
                        const a = document.createElement('a');
                        a.href = URL.createObjectURL(blob);
                        a.download = doc.name;
                        document.body.appendChild(a);
                        a.click();
                        document.body.removeChild(a);
                        
                        console.log('文档下载成功:', doc.name);
                        return;
                    }
                }
            }
            
            // 降级处理：如果上述方法失败，尝试使用当前文件对象
            if (this.currentFile && typeof this.currentFile === 'object' && this.currentFile.name) {
                const a = document.createElement('a');
                a.href = URL.createObjectURL(this.currentFile);
                a.download = this.currentFile.name;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                
                console.log('文档下载成功:', this.currentFile.name);
                return;
            }
            
            throw new Error('无法获取文档数据进行下载');
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
            // 设置当前文件信息
            this.currentFile = { 
                name: doc.name,
                type: doc.type
            };
            this.currentFileType = doc.type;
            const viewer = document.getElementById('documentViewer');
            const documentTitle = document.getElementById('documentTitle');
            
            if (doc.type === 'xlsx' || doc.type === 'xls') {
                // 如果是Excel文档，需要处理数据
                const workbook = XLSX.read(doc.data, { type: 'base64' });
                this.currentData = workbook;
                
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const html = XLSX.utils.sheet_to_html(worksheet, { id: 'excelSheet' });
                
                viewer.innerHTML = html;
                this.makeExcelEditable();
            } else if (doc.type === 'pdf') {
                // 如果是PDF文档，渲染预览
                viewer.innerHTML = `
                    <div style="padding: 20px;">
                        <h3>📄 PDF文档预览</h3>
                        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin: 20px 0;">
                            <div style="padding: 15px; background-color: var(--light-bg); border-radius: 8px; border: 1px solid var(--border-color);">
                                <h4>文档信息</h4>
                                <p><strong>名称:</strong> ${doc.name}</p>
                                <p><strong>格式:</strong> PDF</p>
                                <p><strong>修改时间:</strong> ${doc.lastModified}</p>
                            </div>
                            <div style="padding: 15px; background-color: var(--light-bg); border-radius: 8px; border: 1px solid var(--border-color);">
                                <h4>可用操作</h4>
                                <ul style="margin: 10px 0; padding-left: 20px;">
                                    <li>在线预览PDF内容</li>
                                    <li>下载原始文档</li>
                                </ul>
                            </div>
                        </div>
                        
                        <div id="pdfContainer" style="
                            margin: 20px 0;
                            padding: 20px;
                            background-color: #f8f9fa;
                            border-radius: 8px;
                            border: 1px solid #e9ecef;
                            overflow: auto;
                        ">
                            <div style="text-align: center; padding: 40px;">
                                <div style="font-size: 3rem; margin-bottom: 15px;">📄</div>
                                <p>正在加载PDF文档，请稍候...</p>
                            </div>
                        </div>
                        
                        <div style="margin-top: 20px; text-align: center;">
                            <button onclick="documentEditor.downloadOriginalDocument()" style="
                                background-color: var(--secondary-color);
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
                                background-color: var(--muted-color);
                                color: var(--text-color);
                                border: none;
                                padding: 12px 24px;
                                font-size: 1rem;
                                border-radius: 4px;
                                cursor: pointer;
                                transition: background-color 0.3s ease;
                            ">← 返回文档列表</button>
                        </div>
                    </div>
                `;
                
                // 渲染PDF内容
                const pdfContainer = document.getElementById('pdfContainer');
                // 设置PDF.js worker
                pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
                
                // 将base64数据转换为ArrayBuffer
                const binaryString = atob(doc.data);
                const length = binaryString.length;
                const arrayBuffer = new ArrayBuffer(length);
                const view = new Uint8Array(arrayBuffer);
                for (let i = 0; i < length; i++) {
                    view[i] = binaryString.charCodeAt(i);
                }
                
                // 加载PDF文档
                pdfjsLib.getDocument(arrayBuffer).promise.then(pdf => {
                    console.log('PDF加载成功，共', pdf.numPages, '页');
                    
                    // 清空容器
                    pdfContainer.innerHTML = '';
                    
                    // 创建页面容器
                    const pagesContainer = document.createElement('div');
                    pagesContainer.style.display = 'flex';
                    pagesContainer.style.flexDirection = 'column';
                    pagesContainer.style.alignItems = 'center';
                    pagesContainer.style.gap = '20px';
                    
                    // 渲染每一页
                    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
                        pdf.getPage(pageNum).then(page => {
                            // 创建canvas元素
                            const canvas = document.createElement('canvas');
                            canvas.style.maxWidth = '100%';
                            canvas.style.border = '1px solid #e0e0e0';
                            canvas.style.borderRadius = '4px';
                            
                            const container = document.createElement('div');
                            container.style.width = '100%';
                            container.style.textAlign = 'center';
                            container.appendChild(canvas);
                            
                            // 设置渲染选项
                            const viewport = page.getViewport({ scale: 1.5 });
                            canvas.height = viewport.height;
                            canvas.width = viewport.width;
                            
                            // 渲染页面
                            const renderContext = {
                                canvasContext: canvas.getContext('2d'),
                                viewport: viewport
                            };
                            
                            page.render(renderContext).promise.then(() => {
                                console.log('PDF页面', pageNum, '渲染成功');
                            }).catch(err => {
                                console.error('渲染PDF页面', pageNum, '失败:', err);
                                container.innerHTML = `<div style="padding: 20px; color: #e74c3c;">页面 ${pageNum} 渲染失败</div>`;
                            });
                            
                            pagesContainer.appendChild(container);
                        }).catch(err => {
                            console.error('获取PDF页面失败:', err);
                            const errorDiv = document.createElement('div');
                            errorDiv.style.padding = '20px';
                            errorDiv.style.color = '#e74c3c';
                            errorDiv.textContent = `获取页面 ${pageNum} 失败`;
                            pagesContainer.appendChild(errorDiv);
                        });
                    }
                    
                    pdfContainer.appendChild(pagesContainer);
                }).catch(err => {
                    console.error('加载PDF文档失败:', err);
                    pdfContainer.innerHTML = `
                        <div style="text-align: center; padding: 40px; color: #e74c3c;">
                            <div style="font-size: 3rem; margin-bottom: 15px;">❌</div>
                            <p>PDF文档加载失败,请重试</p>
                            <p style="font-size: 0.9rem; margin-top: 10px;">错误信息: ${err.message}</p>
                        </div>
                    `;
                });
            } else if (doc.type === 'txt') {
                // 如果是TXT文档，使用renderTextDocument方法渲染可编辑内容
                this.renderTextDocument(doc.data);
            } else if (doc.type === 'docx') {
                // 如果是Word文档，显示基本信息和下载选项
                viewer.innerHTML = `
                    <div style="padding: 20px;">
                        <h3>📄 Word文档预览</h3>
                        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin: 20px 0;">
                            <div style="padding: 15px; background-color: var(--light-bg); border-radius: 8px; border: 1px solid var(--border-color);">
                                <h4>文档信息</h4>
                                <p><strong>名称:</strong> ${doc.name}</p>
                                <p><strong>格式:</strong> DOCX</p>
                                <p><strong>修改时间:</strong> ${doc.lastModified}</p>
                            </div>
                            <div style="padding: 15px; background-color: var(--light-bg); border-radius: 8px; border: 1px solid var(--border-color);">
                                <h4>可用操作</h4>
                                <ul style="margin: 10px 0; padding-left: 20px;">
                                    <li>查看文档基本信息</li>
                                    <li>下载原始文档</li>
                                </ul>
                            </div>
                        </div>
                        <div style="padding: 20px; background-color: #fff3cd; border: 1px solid #ffeeba; border-radius: 8px; margin: 20px 0;">
                            <h4 style="color: #856404;">📋 说明</h4>
                            <p style="color: #856404; margin: 10px 0;">文档内容已保存在本地，您可以下载原始文件查看完整内容。</p>
                        </div>
                        <div style="margin-top: 20px; text-align: center;">
                            <button onclick="documentEditor.downloadOriginalDocument()" style="
                                background-color: var(--secondary-color);
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
                                background-color: var(--muted-color);
                                color: var(--text-color);
                                border: none;
                                padding: 12px 24px;
                                font-size: 1rem;
                                border-radius: 4px;
                                cursor: pointer;
                                transition: background-color 0.3s ease;
                            ">← 返回文档列表</button>
                        </div>
                    </div>
                `;
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
