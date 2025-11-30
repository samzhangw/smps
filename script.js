class SeatingChart {
    constructor() {
        this.students = [];
        this.seatingConfig = {
            rows: 4,
            cols: 6
        };
        this.seatingMap = {};
        this.draggedStudent = null;
        this.draggedElement = null;
        this.printFontSize = 16;
        this.className = '課後班';
        this.constraints = []; 

        // ★★★ 請將下方的引號內容替換成你的 Google Apps Script 網頁應用程式網址 ★★★
        this.gasUrl = 'https://script.google.com/macros/s/AKfycbyEa42GnSjfraPFBvPkT0XIfHOvDe5lb2jhHWNoVg9jm2Zo5ufYCF3e7oHk2kPHjvOI/exec'; 
        
        this.initializeEventListeners();
        this.loadFromLocalStorage();
        this.renderSeatingGrid();
        this.updateStudentCount();
    }

    initializeEventListeners() {
        document.getElementById('applySettings').addEventListener('click', () => this.applySettings());
        document.getElementById('addStudent').addEventListener('click', () => this.addStudent());
        document.getElementById('autoArrange').addEventListener('click', () => this.autoArrangeSeats());
        document.getElementById('clearSeats').addEventListener('click', () => this.clearSeats());
        document.getElementById('saveConfig').addEventListener('click', () => this.saveConfiguration());
        document.getElementById('loadConfig').addEventListener('click', () => this.loadConfiguration());
        document.getElementById('printSeating').addEventListener('click', () => this.printSeatingChart());
        document.getElementById('exportWord').addEventListener('click', () => this.exportToWord());
        
        // 檔案輸入事件
        document.getElementById('fileInput').addEventListener('change', (e) => this.handleFileLoad(e));
        
        // 鍵盤事件
        document.getElementById('studentName').addEventListener('keypress', (e) => {
            if (e.key === 'Enter') this.addStudent();
        });

        document.getElementById('addMultipleStudents').addEventListener('click', () => this.addMultipleStudents());
        document.getElementById('clearMultipleInput').addEventListener('click', () => this.clearMultipleInput());
        document.getElementById('loadDefaultStudents').addEventListener('click', () => this.loadDefaultStudents());
        
        // 字體與班級設定
        document.getElementById('fontSize').addEventListener('input', (e) => this.updateFontSize(e.target.value));
        document.getElementById('className').addEventListener('input', (e) => this.updateClassName(e.target.value));

        // 相鄰限制
        const addConstraintBtn = document.getElementById('addConstraint');
        if (addConstraintBtn) {
            addConstraintBtn.addEventListener('click', () => this.addConstraintFromUI());
        }

        // ★★★ 雲端按鈕監聽 ★★★
        const saveCloudBtn = document.getElementById('saveToCloud');
        if (saveCloudBtn) {
            saveCloudBtn.addEventListener('click', () => this.saveToCloud());
        }

        const loadCloudBtn = document.getElementById('loadFromCloud');
        if (loadCloudBtn) {
            loadCloudBtn.addEventListener('click', () => this.loadFromCloud());
        }
    }

    // --- 雲端功能 ---

    saveToCloud() {
        if (!this.gasUrl || this.gasUrl.includes('你的_SCRIPT_ID')) {
            this.showToast('請先在 script.js 中設定 Google Apps Script 網址', 'error');
            return;
        }

        const btn = document.getElementById('saveToCloud');
        const originalText = btn.innerHTML;
        btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 儲存中...';
        btn.disabled = true;

        const data = {
            students: this.students,
            seatingConfig: this.seatingConfig,
            seatingMap: this.seatingMap,
            printFontSize: this.printFontSize,
            className: this.className,
            constraints: this.constraints,
            savedAt: new Date().toISOString()
        };

        // 使用 fetch 發送 POST 請求
        fetch(this.gasUrl, {
            method: 'POST',
            body: JSON.stringify(data),
            // 注意：跨域請求通常使用 no-cors 或是 text/plain，
            // 為了讓 GAS 簡單接收，這裡使用預設的 Content-Type (text/plain)
        })
        .then(response => response.json())
        .then(result => {
            if (result.status === 'success') {
                this.showToast('雲端儲存成功！', 'success');
            } else {
                this.showToast('儲存失敗：' + result.message, 'error');
            }
        })
        .catch(error => {
            console.error('Error:', error);
            this.showToast('連線錯誤，請檢查網路或 URL', 'error');
        })
        .finally(() => {
            btn.innerHTML = originalText;
            btn.disabled = false;
        });
    }

    loadFromCloud() {
        if (!this.gasUrl || this.gasUrl.includes('你的_SCRIPT_ID')) {
            this.showToast('請先在 script.js 中設定 Google Apps Script 網址', 'error');
            return;
        }

        if (!confirm('載入雲端資料將會覆蓋目前的設定，確定要繼續嗎？')) {
            return;
        }

        const btn = document.getElementById('loadFromCloud');
        const originalText = btn.innerHTML;
        btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 載入中...';
        btn.disabled = true;

        fetch(this.gasUrl)
        .then(response => response.json())
        .then(data => {
            // 檢查是否為空物件
            if (data && Object.keys(data).length > 0 && data.students) {
                this.students = data.students || [];
                this.seatingConfig = data.seatingConfig || { rows: 4, cols: 6 };
                this.seatingMap = data.seatingMap || {};
                this.printFontSize = data.printFontSize || 16;
                this.className = data.className || '課後班';
                this.constraints = data.constraints || [];

                // 更新 UI
                document.getElementById('rows').value = this.seatingConfig.rows;
                document.getElementById('cols').value = this.seatingConfig.cols;
                
                const fontSizeInput = document.getElementById('fontSize');
                const fontSizeValue = document.getElementById('fontSizeValue');
                if(fontSizeInput) fontSizeInput.value = this.printFontSize;
                if(fontSizeValue) fontSizeValue.textContent = this.printFontSize + 'px';
                
                const classNameInput = document.getElementById('className');
                const classDisplay = document.getElementById('currentClassDisplay');
                if(classNameInput) classNameInput.value = this.className;
                if(classDisplay) classDisplay.textContent = this.className;
                
                this.updateStudentList();
                this.updateStudentCount();
                this.renderSeatingGrid();
                this.refreshConstraintSelectors();
                this.renderConstraintsList();
                this.saveToLocalStorage(); // 同步更新本地儲存

                this.showToast('雲端資料載入成功！', 'success');
            } else {
                this.showToast('雲端目前沒有資料', 'info');
            }
        })
        .catch(error => {
            console.error('Error:', error);
            this.showToast('讀取錯誤，請檢查網路或 URL', 'error');
        })
        .finally(() => {
            btn.innerHTML = originalText;
            btn.disabled = false;
        });
    }

    // --- 原有功能 ---

    addMultipleStudents() {
        const textarea = document.getElementById('multipleStudents');
        const content = textarea.value.trim();
        if (!content) { this.showToast('請輸入學生姓名', 'error'); return; }
        const names = content.split(/[\n\r,，、\s]+/).filter(name => name.trim());
        if (names.length === 0) { this.showToast('沒有找到有效的學生姓名', 'error'); return; }
        
        let addedCount = 0, duplicateCount = 0, invalidCount = 0;
        
        names.forEach(name => {
            const trimmedName = name.trim();
            if (!trimmedName || trimmedName.length > 20) { invalidCount++; return; }
            const existingStudent = this.students.find(student => student.name === trimmedName);
            if (existingStudent) { duplicateCount++; return; }
            
            let studentName = trimmedName;
            let studentNote = '';
            if (trimmedName.includes('|')) {
                const parts = trimmedName.split('|');
                studentName = parts[0].trim();
                studentNote = parts[1] ? parts[1].trim() : '';
                if (!studentName || studentName.length > 20) { invalidCount++; return; }
                if (this.students.find(student => student.name === studentName)) { duplicateCount++; return; }
            }
            
            this.students.push({
                id: Date.now().toString() + Math.random().toString(36).substr(2, 9),
                name: studentName,
                note: studentNote
            });
            addedCount++;
        });
        
        this.updateStudentList();
        this.updateStudentCount();
        this.saveToLocalStorage();
        textarea.value = '';
        
        let message = `成功新增 ${addedCount} 名學生`;
        if (duplicateCount > 0) message += `，跳過 ${duplicateCount} 個重複姓名`;
        if (invalidCount > 0) message += `，跳過 ${invalidCount} 個無效輸入`;
        this.showToast(message, 'success');
    }

    clearMultipleInput() {
        document.getElementById('multipleStudents').value = '';
        this.showToast('已清空批量輸入框', 'info');
    }

    applySettings() {
        const rows = parseInt(document.getElementById('rows').value);
        const cols = parseInt(document.getElementById('cols').value);
        if (rows < 1 || rows > 10 || cols < 1 || cols > 8) {
            this.showToast('請輸入有效的排數（1-10）和列數（1-8）', 'error');
            return;
        }
        if (rows === this.seatingConfig.rows && cols === this.seatingConfig.cols) {
            this.showToast('座位設定沒有變更', 'info');
            return;
        }
        const hasExistingData = this.students.length > 0 || Object.keys(this.seatingMap).length > 0;
        const isConfigChanged = rows !== this.seatingConfig.rows || cols !== this.seatingConfig.cols;

        if (hasExistingData && isConfigChanged) {
            this.showDataLossWarning(rows, cols);
        } else {
            this.applySettingsDirectly(rows, cols);
        }
    }

    showDataLossWarning(newRows, newCols) {
        const existingWarning = document.querySelector('.data-loss-warning');
        if (existingWarning) existingWarning.remove();

        const warningOverlay = document.createElement('div');
        warningOverlay.className = 'data-loss-warning';
        warningOverlay.innerHTML = `
            <div class="warning-content">
                <div class="warning-header"><i class="fas fa-exclamation-triangle"></i><h3>座位設定修改警告</h3></div>
                <div class="warning-body">
                    <p class="warning-message">修改座位設定可能會遺失現有資料！</p>
                    <div class="setting-comparison">
                        <div class="current-setting"><h4>目前設定</h4><div class="setting-value">${this.seatingConfig.rows}排 × ${this.seatingConfig.cols}列</div></div>
                        <div class="arrow">→</div>
                        <div class="new-setting"><h4>新設定</h4><div class="setting-value">${newRows}排 × ${newCols}列</div></div>
                    </div>
                    <div class="data-impact">
                        <h4>影響項目</h4>
                        <ul>
                            <li><i class="fas fa-users"></i> 學生人數：${this.students.length}人</li>
                            <li><i class="fas fa-chair"></i> 已安排座位：${Object.keys(this.seatingMap).length}人</li>
                        </ul>
                    </div>
                </div>
                <div class="warning-footer">
                    <button class="btn btn-secondary" onclick="seatingChart.cancelSettingsChange()"><i class="fas fa-times"></i> 取消</button>
                    <button class="btn btn-warning" onclick="seatingChart.applySettingsDirectly(${newRows}, ${newCols})"><i class="fas fa-check"></i> 確定修改</button>
                </div>
            </div>`;
        
        document.body.appendChild(warningOverlay);
        warningOverlay.addEventListener('click', (e) => { if (e.target === warningOverlay) this.cancelSettingsChange(); });
    }

    cancelSettingsChange() {
        document.getElementById('rows').value = this.seatingConfig.rows;
        document.getElementById('cols').value = this.seatingConfig.cols;
        const warning = document.querySelector('.data-loss-warning');
        if (warning) warning.remove();
        this.showToast('已取消修改座位設定', 'info');
    }

    applySettingsDirectly(rows, cols) {
        const oldConfig = { ...this.seatingConfig };
        this.seatingConfig = { rows, cols };
        
        if (rows < oldConfig.rows || cols < oldConfig.cols) {
            const removedCount = this.preserveValidSeats(this.seatingConfig);
            if (removedCount === 0) this.showToast('座位設定已更新', 'success');
        } else {
            this.showToast('座位設定已擴展，現有座位安排已保留', 'success');
        }
        
        this.renderSeatingGrid();
        this.saveToLocalStorage();
        const warning = document.querySelector('.data-loss-warning');
        if (warning) warning.remove();
    }

    preserveValidSeats(newConfig) {
        const validSeats = {};
        let removedCount = 0;
        Object.keys(this.seatingMap).forEach(seatKey => {
            const [row, col] = seatKey.split('-').map(Number);
            if (row <= newConfig.rows && col <= newConfig.cols) {
                validSeats[seatKey] = this.seatingMap[seatKey];
            } else {
                removedCount++;
            }
        });
        this.seatingMap = validSeats;
        if (removedCount > 0) this.showToast(`座位設定已更新，${removedCount}個超出範圍的座位安排已移除`, 'warning');
        return removedCount;
    }

    addStudent() {
        const nameInput = document.getElementById('studentName');
        const noteInput = document.getElementById('studentNote');
        const name = nameInput.value.trim();
        const note = noteInput.value.trim();

        if (!name) { this.showToast('請輸入學生姓名', 'error'); return; }
        if (this.students.some(student => student.name === name)) { this.showToast('此姓名已存在', 'error'); return; }

        this.students.push({
            id: Date.now().toString(),
            name: name,
            note: note || '',
            addedAt: new Date().toISOString()
        });

        this.updateStudentList();
        this.updateStudentCount();
        this.refreshConstraintSelectors();
        this.saveToLocalStorage();
        nameInput.value = '';
        noteInput.value = '';
        nameInput.focus();
        this.showToast(`已新增學生：${name}`, 'success');
    }

    removeStudent(studentId) {
        const index = this.students.findIndex(student => student.id === studentId);
        if (index !== -1) {
            const studentName = this.students[index].name;
            this.students.splice(index, 1);
            Object.keys(this.seatingMap).forEach(seatKey => {
                if (this.seatingMap[seatKey] === studentId) delete this.seatingMap[seatKey];
            });
            this.updateStudentList();
            this.updateStudentCount();
            this.renderSeatingGrid();
            this.constraints = this.constraints.filter(c => c.a !== studentId && c.b !== studentId);
            this.refreshConstraintSelectors();
            this.renderConstraintsList();
            this.saveToLocalStorage();
            this.showToast(`已移除學生：${studentName}`, 'info');
        }
    }

    editStudent(studentId) {
        const student = this.students.find(s => s.id === studentId);
        if (!student) return;
        const newName = prompt('請輸入新的姓名：', student.name);
        if (newName && newName.trim() !== '') {
            student.name = newName.trim();
            const newNote = prompt('請輸入備註（可選）：', student.note || '');
            student.note = newNote ? newNote.trim() : '';
            this.updateStudentList();
            this.renderSeatingGrid();
            this.saveToLocalStorage();
            this.showToast(`已更新學生：${newName}`, 'success');
        }
    }

    updateStudentList() {
        const studentsList = document.getElementById('studentsList');
        studentsList.innerHTML = '';
        this.students.forEach(student => {
            const studentItem = document.createElement('div');
            studentItem.className = 'student-item';
            studentItem.draggable = true;
            studentItem.dataset.studentId = student.id;
            
            const hasSeat = Object.values(this.seatingMap).includes(student.id);
            let seatLocation = '';
            if (hasSeat) {
                for (const [seatKey, studentId] of Object.entries(this.seatingMap)) {
                    if (studentId === student.id) { seatLocation = seatKey; break; }
                }
                studentItem.classList.add('has-seat');
            }
            
            studentItem.innerHTML = `
                <div class="student-info">
                    <div class="student-name">${student.name}</div>
                    ${student.note ? `<div class="student-note">${student.note}</div>` : ''}
                    ${hasSeat ? `<div class="seat-status">座位 ${seatLocation}</div>` : '<div class="seat-status">未安排</div>'}
                </div>
                <div class="student-actions">
                    <button class="btn btn-info" onclick="seatingChart.editStudent('${student.id}')"><i class="fas fa-edit"></i></button>
                    <button class="btn btn-danger" onclick="seatingChart.removeStudent('${student.id}')"><i class="fas fa-trash"></i></button>
                </div>`;
            
            this.addDragEvents(studentItem, student);
            studentsList.appendChild(studentItem);
        });
        this.refreshConstraintSelectors();
        this.renderConstraintsList();
    }

    updateStudentCount() {
        document.getElementById('studentCount').textContent = this.students.length;
    }

    renderSeatingGrid() {
        const grid = document.getElementById('seatingGrid');
        grid.innerHTML = '';
        grid.style.gridTemplateColumns = `repeat(${this.seatingConfig.cols}, 1fr)`;

        for (let row = 1; row <= this.seatingConfig.rows; row++) {
            for (let col = 1; col <= this.seatingConfig.cols; col++) {
                const seatKey = `${row}-${col}`;
                const seat = document.createElement('div');
                seat.className = 'seat';
                seat.dataset.row = row;
                seat.dataset.col = col;
                seat.dataset.seatKey = seatKey;

                const seatNumber = document.createElement('div');
                seatNumber.className = 'seat-number';
                seatNumber.textContent = `${row}-${col}`;

                const seatContent = document.createElement('div');
                seatContent.className = 'seat-content';

                if (this.seatingMap[seatKey]) {
                    const student = this.students.find(s => s.id === this.seatingMap[seatKey]);
                    if (student) {
                        seat.classList.add('occupied');
                        seatContent.innerHTML = `
                            <div class="seat-student" draggable="true" data-student-id="${student.id}">
                                <div class="seat-student-name">${student.name}</div>
                                ${student.note ? `<div class="seat-student-note">${student.note}</div>` : ''}
                            </div>`;
                    }
                } else {
                    seat.classList.add('empty-seat');
                    seatContent.innerHTML = '<div class="seat-student">空位</div>';
                }

                seat.appendChild(seatNumber);
                seat.appendChild(seatContent);
                this.addSeatDropEvents(seat, seatKey);
                grid.appendChild(seat);
            }
        }
    }

    getAdjacentSeatKeys(seatKey) {
        const [row, col] = seatKey.split('-').map(Number);
        const deltas = [-1, 0, 1];
        const keys = [];
        for (let dr of deltas) {
            for (let dc of deltas) {
                if (dr === 0 && dc === 0) continue;
                const nr = row + dr;
                const nc = col + dc;
                if (nr >= 1 && nr <= this.seatingConfig.rows && nc >= 1 && nc <= this.seatingConfig.cols) {
                    keys.push(`${nr}-${nc}`);
                }
            }
        }
        return keys;
    }

    violatesConstraints(studentId, seatKey) {
        const neighbors = this.getAdjacentSeatKeys(seatKey);
        const forbiddenSet = new Set();
        this.constraints.forEach(c => {
            if (c.a === studentId) forbiddenSet.add(c.b);
            if (c.b === studentId) forbiddenSet.add(c.a);
        });
        if (forbiddenSet.size === 0) return false;
        for (const nKey of neighbors) {
            const neighborId = this.seatingMap[nKey];
            if (neighborId && forbiddenSet.has(neighborId)) return true;
        }
        return false;
    }

    handleSeatClick(seatKey, seatElement) {
        if (this.seatingMap[seatKey]) {
            const currentStudent = this.students.find(s => s.id === this.seatingMap[seatKey]);
            const action = confirm(`座位 ${seatKey} 目前由 ${currentStudent.name} 佔用。\n\n點擊「確定」移除學生\n點擊「取消」選擇其他學生替換`);
            if (action) {
                delete this.seatingMap[seatKey];
                this.renderSeatingGrid();
                this.updateStudentList();
                this.saveToLocalStorage();
                this.showToast(`已將 ${currentStudent.name} 從座位移除`, 'info');
            } else {
                this.showStudentSelectionDialog(seatKey, true);
            }
        } else {
            this.showStudentSelectionDialog(seatKey);
        }
    }

    showStudentSelectionDialog(seatKey, isReplacement = false) {
        if (this.students.length === 0) { this.showToast('請先新增學生', 'error'); return; }
        let availableStudents;
        if (isReplacement) {
            availableStudents = this.students;
        } else {
            availableStudents = this.students.filter(student => !Object.values(this.seatingMap).includes(student.id));
        }
        if (availableStudents.length === 0) { this.showToast('沒有可選擇的學生', 'info'); return; }
        this.createStudentSelectionModal(seatKey, availableStudents, isReplacement);
    }

    createStudentSelectionModal(seatKey, availableStudents, isReplacement = false) {
        const existingModal = document.querySelector('.student-selection-modal');
        if (existingModal) existingModal.remove();

        const modalOverlay = document.createElement('div');
        modalOverlay.className = 'student-selection-modal';
        const modalContent = document.createElement('div');
        modalContent.className = 'modal-content';
        const modalTitle = isReplacement ? `選擇學生替換座位 ${seatKey} 的學生` : `選擇學生安排到座位 ${seatKey}`;
        
        modalContent.innerHTML = `
            <div class="modal-header"><h3>${modalTitle}</h3><button class="modal-close" onclick="this.closest('.student-selection-modal').remove()"><i class="fas fa-times"></i></button></div>
            <div class="modal-body">
                <div class="student-options">
                    ${availableStudents.map(student => {
                        const hasSeat = Object.values(this.seatingMap).includes(student.id);
                        let seatInfo = '';
                        if (hasSeat) {
                            const currentSeat = Object.keys(this.seatingMap).find(key => this.seatingMap[key] === student.id);
                            seatInfo = `<div class="student-current-seat">目前座位：${currentSeat}</div>`;
                        }
                        return `
                            <div class="student-option" onclick="seatingChart.selectStudentForSeat('${seatKey}', '${student.id}', ${isReplacement})">
                                <div class="student-option-name">${student.name}</div>
                                ${student.note ? `<div class="student-option-note">${student.note}</div>` : ''}
                                ${seatInfo}
                            </div>`;
                    }).join('')}
                </div>
            </div>
            <div class="modal-footer"><button class="btn btn-secondary" onclick="this.closest('.student-selection-modal').remove()">取消</button></div>`;
        
        modalOverlay.appendChild(modalContent);
        document.body.appendChild(modalOverlay);
        modalOverlay.addEventListener('click', (e) => { if (e.target === modalOverlay) modalOverlay.remove(); });
    }

    selectStudentForSeat(seatKey, studentId, isReplacement = false) {
        const student = this.students.find(s => s.id === studentId);
        if (student) {
            if (isReplacement && this.seatingMap[seatKey]) {
                const currentStudent = this.students.find(s => s.id === this.seatingMap[seatKey]);
                if (currentStudent && currentStudent.id !== student.id) {
                    this.swapStudents(student, currentStudent, seatKey);
                } else {
                    this.showToast('選擇了同一個學生', 'info');
                }
            } else {
                if (this.violatesConstraints(studentId, seatKey)) {
                    this.showToast('違反相鄰限制，請選擇其他座位', 'error');
                    return;
                }
                this.seatingMap[seatKey] = studentId;
                this.renderSeatingGrid();
                this.updateStudentList();
                this.saveToLocalStorage();
                this.showToast(`已將 ${student.name} 安排到座位 ${seatKey}`, 'success');
            }
            const modal = document.querySelector('.student-selection-modal');
            if (modal) modal.remove();
        }
    }

    showStudentSelectionForMobile(student) {
        const availableSeats = [];
        for (let row = 1; row <= this.seatingConfig.rows; row++) {
            for (let col = 1; col <= this.seatingConfig.cols; col++) {
                const seatKey = `${row}-${col}`;
                if (!this.seatingMap[seatKey]) availableSeats.push(seatKey);
            }
        }
        if (availableSeats.length === 0) { this.showToast('沒有空座位可選擇', 'error'); return; }
        this.createMobileSeatSelectionModal(student, availableSeats);
    }

    createMobileSeatSelectionModal(student, availableSeats) {
        const existingModal = document.querySelector('.student-selection-modal');
        if (existingModal) existingModal.remove();
        const modalOverlay = document.createElement('div');
        modalOverlay.className = 'student-selection-modal';
        const modalContent = document.createElement('div');
        modalContent.className = 'modal-content';
        
        modalContent.innerHTML = `
            <div class="modal-header"><h3>為 ${student.name} 選擇座位</h3><button class="modal-close" onclick="this.closest('.student-selection-modal').remove()"><i class="fas fa-times"></i></button></div>
            <div class="modal-body"><div class="mobile-seat-grid">${this.generateMobileSeatGrid(availableSeats)}</div></div>
            <div class="modal-footer"><button class="btn btn-secondary" onclick="this.closest('.student-selection-modal').remove()">取消</button></div>`;
        
        modalOverlay.appendChild(modalContent);
        document.body.appendChild(modalOverlay);
        modalOverlay.addEventListener('click', (e) => { if (e.target === modalOverlay) modalOverlay.remove(); });
    }

    generateMobileSeatGrid(availableSeats) {
        let gridHTML = '';
        for (let row = 1; row <= this.seatingConfig.rows; row++) {
            gridHTML += '<div class="mobile-seat-row">';
            for (let col = 1; col <= this.seatingConfig.cols; col++) {
                const seatKey = `${row}-${col}`;
                const isAvailable = availableSeats.includes(seatKey);
                const currentStudent = this.students.find(s => s.id === this.seatingMap[seatKey]);
                
                if (isAvailable) {
                    gridHTML += `
                        <div class="mobile-seat-option available" onclick="seatingChart.selectStudentForSeat('${seatKey}', '${this.draggedStudent ? this.draggedStudent.id : ''}')">
                            <div class="mobile-seat-number">${seatKey}</div>
                            <div class="mobile-seat-status">空位</div>
                        </div>`;
                } else if (currentStudent) {
                    gridHTML += `
                        <div class="mobile-seat-option occupied">
                            <div class="mobile-seat-number">${seatKey}</div>
                            <div class="mobile-seat-student">${currentStudent.name}</div>
                        </div>`;
                } else {
                    gridHTML += `
                        <div class="mobile-seat-option disabled">
                            <div class="mobile-seat-number">${seatKey}</div>
                            <div class="mobile-seat-status">不可用</div>
                        </div>`;
                }
            }
            gridHTML += '</div>';
        }
        return gridHTML;
    }

    autoArrangeSeats() {
        if (this.students.length === 0) { this.showToast('請先新增學生', 'error'); return; }
        this.seatingMap = {};
        const shuffledStudents = [...this.students].sort(() => Math.random() - 0.5);
        const availableSeats = [];
        for (let row = 1; row <= this.seatingConfig.rows; row++) {
            for (let col = 1; col <= this.seatingConfig.cols; col++) {
                availableSeats.push(`${row}-${col}`);
            }
        }
        const shuffledSeats = availableSeats.sort(() => Math.random() - 0.5);
        let notPlaced = 0;
        for (const student of shuffledStudents) {
            let placed = false;
            for (const seatKey of shuffledSeats) {
                if (!this.seatingMap[seatKey] && !this.violatesConstraints(student.id, seatKey)) {
                    this.seatingMap[seatKey] = student.id;
                    placed = true;
                    break;
                }
            }
            if (!placed) notPlaced++;
        }
        this.renderSeatingGrid();
        this.updateStudentList();
        this.saveToLocalStorage();
        if (notPlaced > 0) this.showToast(`排座完成，但有 ${notPlaced} 位因限制未安排`, 'warning');
        else this.showToast('座位已隨機安排完成', 'success');
    }

    clearSeats() {
        if (Object.keys(this.seatingMap).length === 0) { this.showToast('目前沒有安排任何座位', 'info'); return; }
        if (confirm('確定要清空所有座位安排嗎？')) {
            this.seatingMap = {};
            this.renderSeatingGrid();
            this.updateStudentList();
            this.saveToLocalStorage();
            this.showToast('所有座位已清空', 'info');
        }
    }

    saveConfiguration() {
        const config = {
            students: this.students,
            seatingConfig: this.seatingConfig,
            seatingMap: this.seatingMap,
            constraints: this.constraints,
            savedAt: new Date().toISOString()
        };
        const dataBlob = new Blob([JSON.stringify(config, null, 2)], { type: 'application/json' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(dataBlob);
        link.download = `座位配置_${new Date().toLocaleDateString('zh-TW')}.json`;
        link.click();
        this.showToast('配置已儲存到檔案', 'success');
    }

    loadConfiguration() {
        document.getElementById('fileInput').click();
    }

    handleFileLoad(event) {
        const file = event.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const config = JSON.parse(e.target.result);
                if (config.students && config.seatingConfig && config.seatingMap) {
                    this.students = config.students;
                    this.seatingConfig = config.seatingConfig;
                    this.seatingMap = config.seatingMap;
                    this.constraints = config.constraints || [];
                    document.getElementById('rows').value = this.seatingConfig.rows;
                    document.getElementById('cols').value = this.seatingConfig.cols;
                    this.updateStudentList();
                    this.updateStudentCount();
                    this.renderSeatingGrid();
                    this.refreshConstraintSelectors();
                    this.renderConstraintsList();
                    this.saveToLocalStorage();
                    this.showToast('配置已成功載入', 'success');
                } else {
                    this.showToast('檔案格式不正確', 'error');
                }
            } catch (error) {
                this.showToast('載入檔案時發生錯誤', 'error');
                console.error('Error loading file:', error);
            }
        };
        reader.readAsText(file);
        event.target.value = '';
    }

    saveToLocalStorage() {
        const data = {
            students: this.students,
            seatingConfig: this.seatingConfig,
            seatingMap: this.seatingMap,
            printFontSize: this.printFontSize,
            className: this.className,
            constraints: this.constraints
        };
        localStorage.setItem('seatingChartData', JSON.stringify(data));
    }

    loadFromLocalStorage() {
        const saved = localStorage.getItem('seatingChartData');
        if (saved) {
            try {
                const data = JSON.parse(saved);
                this.students = data.students || [];
                this.seatingConfig = data.seatingConfig || { rows: 4, cols: 6 };
                this.seatingMap = data.seatingMap || {};
                this.printFontSize = data.printFontSize || 16;
                this.className = data.className || '課後班';
                this.constraints = data.constraints || [];
                document.getElementById('rows').value = this.seatingConfig.rows;
                document.getElementById('cols').value = this.seatingConfig.cols;
                document.getElementById('fontSize').value = this.printFontSize;
                document.getElementById('fontSizeValue').textContent = this.printFontSize + 'px';
                document.getElementById('className').value = this.className;
                document.getElementById('currentClassDisplay').textContent = this.className;
                this.updateStudentList();
                this.refreshConstraintSelectors();
                this.renderConstraintsList();
            } catch (error) {
                console.error('Error loading from localStorage:', error);
                this.initializeDefaultStudents();
            }
        } else {
            this.initializeDefaultStudents(false);
        }
    }

    initializeDefaultStudents(clearSeating = true) {
        const defaultStudents = ['陳昱允', '彭翊恩', '章彥廷', '曾依凡', '張瑞恩', '張宸晞', '麥惠媛', '陳柳鈴', '吳睿凱', '詹欣容', '李誠恩', '涂毅宏', '曾聿寧', '方語喆', '王勻希', '伊妍欣', '麥庭綺', '彭立宸', '吳睿勳', '連廷駿', '余柏勳', '邱閔昊', '伊妤欣'];
        this.students = [];
        defaultStudents.forEach((name, index) => {
            this.students.push({
                id: `default_${Date.now()}_${index}`,
                name: name,
                note: '',
                addedAt: new Date().toISOString()
            });
        });
        if (clearSeating) this.seatingMap = {};
        this.updateStudentList();
        this.updateStudentCount();
        this.renderSeatingGrid();
        this.saveToLocalStorage();
        this.showToast(`已載入 ${defaultStudents.length} 位預設學生`, 'success');
    }

    refreshConstraintSelectors() {
        const selectA = document.getElementById('constraintStudentA');
        const selectB = document.getElementById('constraintStudentB');
        if (!selectA || !selectB) return;
        const makeOptions = () => this.students.map(s => `<option value="${s.id}">${s.name}</option>`).join('');
        const optionsHTML = makeOptions();
        selectA.innerHTML = `<option value="">選擇學生A</option>` + optionsHTML;
        selectB.innerHTML = `<option value="">選擇學生B</option>` + optionsHTML;
    }

    addConstraintFromUI() {
        const selectA = document.getElementById('constraintStudentA');
        const selectB = document.getElementById('constraintStudentB');
        if (!selectA || !selectB) return;
        const a = selectA.value;
        const b = selectB.value;
        if (!a || !b) { this.showToast('請選擇兩位學生', 'error'); return; }
        if (a === b) { this.showToast('不可選擇同一位學生', 'error'); return; }
        const exists = this.constraints.some(c => (c.a === a && c.b === b) || (c.a === b && c.b === a));
        if (exists) { this.showToast('此限制已存在', 'info'); return; }
        this.constraints.push({ a, b });
        this.saveToLocalStorage();
        this.renderConstraintsList();
        this.showToast('已新增相鄰限制', 'success');
    }

    removeConstraint(index) {
        this.constraints.splice(index, 1);
        this.saveToLocalStorage();
        this.renderConstraintsList();
        this.showToast('已移除相鄰限制', 'info');
    }

    renderConstraintsList() {
        const container = document.getElementById('constraintsList');
        if (!container) return;
        if (!this.constraints || this.constraints.length === 0) {
            container.innerHTML = '<div class="constraints-empty">尚未設定限制</div>';
            return;
        }
        const nameById = (id) => (this.students.find(s => s.id === id)?.name) || '(已移除)';
        container.innerHTML = this.constraints.map((c, idx) => {
            return `<div class="constraint-item"><span class="constraint-pair">${nameById(c.a)} ↔ ${nameById(c.b)}</span><button class="btn btn-danger constraint-remove" onclick="seatingChart.removeConstraint(${idx})"><i class="fas fa-trash"></i></button></div>`;
        }).join('');
    }

    loadDefaultStudents() {
        let message = '確定要載入預設學生嗎？';
        if (this.students.length > 0) message += '\n\n⚠️ 這將會替換現有的學生清單';
        if (Object.keys(this.seatingMap).length > 0) message += '\n\n⚠️ 現有的座位安排將會被清空';
        if (confirm(message)) this.initializeDefaultStudents();
    }

    updateFontSize(size) {
        this.printFontSize = parseInt(size);
        document.getElementById('fontSizeValue').textContent = size + 'px';
        this.saveToLocalStorage();
    }

    updateClassName(name) {
        this.className = name.trim() || '課後班';
        document.getElementById('currentClassDisplay').textContent = this.className;
        this.saveToLocalStorage();
    }

    addDragEvents(studentItem, student) {
        studentItem.addEventListener('dragstart', (e) => {
            this.draggedStudent = student;
            this.draggedElement = studentItem;
            e.dataTransfer.setData('text/plain', student.id);
            e.dataTransfer.effectAllowed = 'move';
            studentItem.classList.add('dragging');
            this.showDragHint(student);
        });

        studentItem.addEventListener('dragend', (e) => {
            studentItem.classList.remove('dragging');
            this.draggedStudent = null;
            this.draggedElement = null;
            this.hideDragHint();
            document.querySelectorAll('.seat').forEach(seat => seat.classList.remove('drag-over', 'swap-target'));
        });

        let touchStartTime = 0, touchStartY = 0, touchStartX = 0, isLongPress = false, longPressTimer = null;

        studentItem.addEventListener('touchstart', (e) => {
            touchStartTime = Date.now();
            touchStartY = e.touches[0].clientY;
            touchStartX = e.touches[0].clientX;
            isLongPress = false;
            longPressTimer = setTimeout(() => {
                isLongPress = true;
                this.draggedStudent = student;
                this.draggedElement = studentItem;
                studentItem.classList.add('dragging');
                this.showToast('長按成功！現在可以拖拽學生', 'info');
            }, 500);
        });

        studentItem.addEventListener('touchmove', (e) => {
            if (isLongPress && this.draggedStudent) {
                e.preventDefault();
                const touch = e.touches[0];
                const deltaY = touch.clientY - touchStartY;
                const deltaX = touch.clientX - touchStartX;
                studentItem.style.transform = `translate(${deltaX}px, ${deltaY}px) rotate(5deg) scale(0.95)`;
            }
        });

        studentItem.addEventListener('touchend', (e) => {
            clearTimeout(longPressTimer);
            if (isLongPress && this.draggedStudent) {
                studentItem.style.transform = '';
                studentItem.classList.remove('dragging');
                this.draggedStudent = null;
                this.draggedElement = null;
                const touch = e.changedTouches[0];
                const elementBelow = document.elementFromPoint(touch.clientX, touch.clientY);
                const seatElement = elementBelow?.closest('.seat');
                if (seatElement) {
                    const seatKey = seatElement.dataset.seatKey;
                    if (seatKey) this.handleStudentDrop(seatKey, seatElement);
                }
            } else {
                if (Date.now() - touchStartTime < 300) this.showStudentSelectionForMobile(student);
            }
        });
    }

    addSeatDropEvents(seat, seatKey) {
        seat.addEventListener('dragover', (e) => {
            e.preventDefault();
            e.dataTransfer.dropEffect = 'move';
            if (this.draggedStudent) {
                if (this.seatingMap[seatKey] && this.seatingMap[seatKey] !== this.draggedStudent.id) {
                    seat.classList.add('swap-target');
                    seat.classList.remove('drag-over');
                } else {
                    seat.classList.add('drag-over');
                    seat.classList.remove('swap-target');
                }
            }
        });

        seat.addEventListener('dragleave', (e) => seat.classList.remove('drag-over', 'swap-target'));

        seat.addEventListener('drop', (e) => {
            e.preventDefault();
            seat.classList.remove('drag-over', 'swap-target');
            if (this.draggedStudent) this.handleStudentDrop(seatKey, seat);
        });

        seat.addEventListener('click', () => this.handleSeatClick(seatKey, seat));
    }

    handleStudentDrop(seatKey, seatElement) {
        if (!this.draggedStudent) return;
        if (this.seatingMap[seatKey]) {
            const currentStudent = this.students.find(s => s.id === this.seatingMap[seatKey]);
            if (currentStudent && currentStudent.id !== this.draggedStudent.id) {
                this.showSwapConfirmation(this.draggedStudent, currentStudent, seatKey);
                return;
            }
        }
        Object.keys(this.seatingMap).forEach(key => {
            if (this.seatingMap[key] === this.draggedStudent.id) delete this.seatingMap[key];
        });
        if (this.violatesConstraints(this.draggedStudent.id, seatKey)) {
            this.showToast('違反相鄰限制，請選擇其他座位', 'error');
            return;
        }
        this.seatingMap[seatKey] = this.draggedStudent.id;
        this.renderSeatingGrid();
        this.updateStudentList();
        this.saveToLocalStorage();
        this.showToast(`已將 ${this.draggedStudent.name} 安排到座位 ${seatKey}`, 'success');
    }

    showSwapConfirmation(student1, student2, targetSeatKey) {
        let student1OriginalSeat = null;
        Object.keys(this.seatingMap).forEach(key => { if (this.seatingMap[key] === student1.id) student1OriginalSeat = key; });
        const message = student1OriginalSeat 
            ? `確定要將 ${student1.name} 和 ${student2.name} 的座位互換嗎？`
            : `確定要將 ${student1.name} 安排到座位 ${targetSeatKey}，並移除 ${student2.name} 嗎？`;
        if (confirm(message)) this.swapStudents(student1, student2, targetSeatKey);
    }

    showDragHint(student) {
        const hint = document.createElement('div');
        hint.id = 'drag-hint-popup';
        hint.className = 'drag-hint-popup';
        hint.innerHTML = `<div class="drag-hint-content"><i class="fas fa-exchange-alt"></i><span>拖拽 ${student.name} 到座位進行更換</span></div>`;
        document.body.appendChild(hint);
        setTimeout(() => this.hideDragHint(), 3000);
    }

    hideDragHint() {
        const hint = document.getElementById('drag-hint-popup');
        if (hint) hint.remove();
    }

    swapStudents(student1, student2, targetSeatKey) {
        let student1OriginalSeat = null;
        Object.keys(this.seatingMap).forEach(key => { if (this.seatingMap[key] === student1.id) student1OriginalSeat = key; });

        if (student1OriginalSeat) {
            const newSeatFor1 = targetSeatKey;
            const newSeatFor2 = student1OriginalSeat;
            if (this.violatesConstraints(student1.id, newSeatFor1) || this.violatesConstraints(student2.id, newSeatFor2)) {
                this.showToast('交換後違反相鄰限制，操作已取消', 'error');
                return;
            }
            this.seatingMap[student1OriginalSeat] = student2.id;
            this.seatingMap[targetSeatKey] = student1.id;
            this.showToast(`已將 ${student1.name} 和 ${student2.name} 的座位互換`, 'success');
        } else {
            if (this.violatesConstraints(student1.id, targetSeatKey)) {
                this.showToast('違反相鄰限制，操作已取消', 'error');
                return;
            }
            this.seatingMap[targetSeatKey] = student1.id;
            this.showToast(`已將 ${student1.name} 安排到座位 ${targetSeatKey}，${student2.name} 的座位已移除`, 'success');
        }
        this.renderSeatingGrid();
        this.updateStudentList();
        this.saveToLocalStorage();
    }

    printSeatingChart() {
        if (this.students.length === 0) { this.showToast('請先新增學生', 'error'); return; }
        const printWindow = window.open('', '_blank');
        const currentDate = new Date().toLocaleDateString('zh-TW');
        
        printWindow.document.write(`
            <!DOCTYPE html>
            <html lang="zh-TW">
            <head>
                <meta charset="UTF-8">
                <title>座位表 - ${currentDate}</title>
                <style>
                    body { font-family: 'Microsoft JhengHei', Arial, sans-serif; margin: 20px; background: white; }
                    .print-header { text-align: center; margin-bottom: 30px; border-bottom: 2px solid #333; padding-bottom: 20px; }
                    .print-header h1 { font-size: 24px; margin: 0 0 10px 0; color: #333; }
                    .print-info { font-size: 14px; color: #666; margin-bottom: 20px; }
                    .print-grid { display: grid; gap: 10px; margin: 20px 0; page-break-inside: avoid; }
                    .print-seat { border: 2px solid #333; padding: 15px; text-align: center; background: #f9f9f9; min-height: 60px; display: flex; flex-direction: column; justify-content: center; align-items: center; }
                    .print-seat.occupied { background: #e3f2fd; border-color: #2196f3; }
                    .print-seat-number { font-size: 12px; color: #666; margin-bottom: 5px; }
                    .print-student-name { font-size: ${this.printFontSize}px; font-weight: bold; color: #333; }
                    .print-student-note { font-size: 12px; color: #666; font-style: italic; margin-top: 3px; }
                    .print-empty { color: #999; font-style: italic; }
                    .print-teacher-desk { text-align: center; margin-bottom: 40px; padding: 25px; background: #2d3748; color: white; border: 3px solid #4a5568; border-radius: 15px; font-weight: 700; font-size: 22px; position: relative; box-shadow: 0 8px 25px rgba(0,0,0,0.3); }
                    .print-teacher-desk:before { content: ''; position: absolute; top: -2px; left: -2px; right: -2px; bottom: -2px; background: linear-gradient(45deg, #667eea, #764ba2, #f093fb, #f5576c); border-radius: 17px; z-index: -1; }
                    .print-teacher-desk-content { display: flex; align-items: center; justify-content: center; gap: 15px; position: relative; z-index: 1; }
                    .print-teacher-desk i { font-size: 28px; color: #ffd700; text-shadow: 2px 2px 4px rgba(0,0,0,0.5); }
                    .print-teacher-desk-text { font-size: 24px; font-weight: 800; text-shadow: 2px 2px 4px rgba(0,0,0,0.5); letter-spacing: 2px; }
                    .print-footer { margin-top: 30px; text-align: center; font-size: 12px; color: #666; border-top: 1px solid #ccc; padding-top: 20px; }
                    @media print {
                        body { margin: 0; }
                        .print-seat { page-break-inside: avoid; }
                        .print-teacher-desk { background: #2d3748 !important; border: 3px solid #000 !important; color: #000 !important; box-shadow: none !important; }
                        .print-teacher-desk:before { display: none !important; }
                        .print-teacher-desk-content, .print-teacher-desk i, .print-teacher-desk-text { color: #000 !important; text-shadow: none !important; }
                    }
                </style>
                <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
            </head>
            <body>
                <div class="print-header">
                    <h1>${this.className}學生座位表</h1>
                    <div class="print-info">
                        班級：${this.className} | 座位配置：${this.seatingConfig.rows}排${this.seatingConfig.cols}列 | 學生人數：${this.students.length}人 | 已安排座位：${Object.keys(this.seatingMap).length}人 | 日期：${currentDate}
                    </div>
                </div>
                <div class="print-teacher-desk"><div class="print-teacher-desk-content"><i class="fas fa-chalkboard"></i><span class="print-teacher-desk-text">講台</span></div></div>
                <div class="print-grid" style="grid-template-columns: repeat(${this.seatingConfig.cols}, 1fr);">
                    ${this.generatePrintSeats()}
                </div>
                <div class="print-footer"><p>座位表由學生座位排程系統生成</p><p>列印時間：${new Date().toLocaleString('zh-TW')}</p></div>
            </body>
            </html>`);
        printWindow.document.close();
        printWindow.focus();
        setTimeout(() => { printWindow.print(); printWindow.close(); }, 500);
    }

    generatePrintSeats() {
        let seatsHTML = '';
        for (let row = 1; row <= this.seatingConfig.rows; row++) {
            for (let col = 1; col <= this.seatingConfig.cols; col++) {
                const seatKey = `${row}-${col}`;
                const student = this.students.find(s => s.id === this.seatingMap[seatKey]);
                if (student) {
                    seatsHTML += `
                        <div class="print-seat occupied">
                            <div class="print-seat-number">座位 ${seatKey}</div>
                            <div class="print-student-name">${student.name}</div>
                            ${student.note ? `<div class="print-student-note">${student.note}</div>` : ''}
                        </div>`;
                } else {
                    seatsHTML += `<div class="print-seat"><div class="print-seat-number">座位 ${seatKey}</div><div class="print-student-name print-empty">空位</div></div>`;
                }
            }
        }
        return seatsHTML;
    }

    exportToWord() {
        if (this.students.length === 0) { this.showToast('請先新增學生', 'error'); return; }
        const currentDate = new Date().toLocaleDateString('zh-TW');
        const currentTime = new Date().toLocaleString('zh-TW');
        const wordContent = `
            <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
            <head><meta charset="UTF-8"><title>學生座位表</title></head>
            <body>
                <div style="text-align:center;margin-bottom:30px;border-bottom:2px solid #333;padding-bottom:20px">
                    <h1 style="font-size:24px;color:#333">${this.className}學生座位表</h1>
                    <div style="font-size:14px;color:#666">班級：${this.className} | 日期：${currentDate}</div>
                </div>
                <div style="text-align:center;margin-bottom:40px;padding:25px;background:#2d3748;color:white;border:3px solid #4a5568;">講台</div>
                <table style="width:100%;border-collapse:collapse;margin:20px 0;">
                    ${this.generateWordTable()}
                </table>
                <div style="margin-top:30px;text-align:center;font-size:12px;color:#666;border-top:1px solid #ccc;padding-top:20px"><p>匯出時間：${currentTime}</p></div>
            </body></html>`;
        const blob = new Blob([wordContent], { type: 'application/msword' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `學生座位表_${currentDate}.doc`;
        link.click();
        this.showToast('Word文件已下載', 'success');
    }

    generateWordTable() {
        let tableHTML = '';
        for (let row = 1; row <= this.seatingConfig.rows; row++) {
            tableHTML += '<tr>';
            for (let col = 1; col <= this.seatingConfig.cols; col++) {
                const seatKey = `${row}-${col}`;
                const student = this.students.find(s => s.id === this.seatingMap[seatKey]);
                const bg = student ? '#e3f2fd' : '#f9f9f9';
                tableHTML += `<td style="border:2px solid #333;padding:15px;text-align:center;background:${bg}">
                    <div style="font-size:12px;color:#666">座位 ${seatKey}</div>
                    <div style="font-size:${this.printFontSize}px;font-weight:bold;color:#333">${student ? student.name : '空位'}</div>
                </td>`;
            }
            tableHTML += '</tr>';
        }
        return tableHTML;
    }

    showToast(message, type = 'info') {
        const existingToast = document.querySelector('.toast');
        if (existingToast) existingToast.remove();
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        toast.textContent = message;
        document.body.appendChild(toast);
        setTimeout(() => toast.classList.add('show'), 100);
        setTimeout(() => { toast.classList.remove('show'); setTimeout(() => { if (toast.parentNode) toast.remove(); }, 300); }, 3000);
    }
}

let seatingChart;
document.addEventListener('DOMContentLoaded', () => {
    seatingChart = new SeatingChart();
});
window.seatingChart = seatingChart;