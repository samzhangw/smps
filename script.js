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
        this.printFontSize = 16; // 預設列印字體大小
        this.className = '課後班'; // 預設班級名稱
        this.constraints = []; // 不能相鄰的學生配對，元素: {a: studentId, b: studentId}
        
        this.initializeEventListeners();
        this.loadFromLocalStorage();
        this.renderSeatingGrid();
        this.updateStudentCount();
    }

    initializeEventListeners() {
        // 設定按鈕
        document.getElementById('applySettings').addEventListener('click', () => {
            this.applySettings();
        });

        // 新增學生按鈕
        document.getElementById('addStudent').addEventListener('click', () => {
            this.addStudent();
        });

        // 自動排座位按鈕
        document.getElementById('autoArrange').addEventListener('click', () => {
            this.autoArrangeSeats();
        });

        // 清空座位按鈕
        document.getElementById('clearSeats').addEventListener('click', () => {
            this.clearSeats();
        });

        // 儲存配置按鈕
        document.getElementById('saveConfig').addEventListener('click', () => {
            this.saveConfiguration();
        });

        // 載入配置按鈕
        document.getElementById('loadConfig').addEventListener('click', () => {
            this.loadConfiguration();
        });

        // 列印座位表按鈕
        document.getElementById('printSeating').addEventListener('click', () => {
            this.printSeatingChart();
        });

        // 匯出Word按鈕
        document.getElementById('exportWord').addEventListener('click', () => {
            this.exportToWord();
        });

        // 檔案輸入事件
        document.getElementById('fileInput').addEventListener('change', (e) => {
            this.handleFileLoad(e);
        });

        // 鍵盤事件
        document.getElementById('studentName').addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                this.addStudent();
            }
        });

        // 批量新增學生按鈕
        document.getElementById('addMultipleStudents').addEventListener('click', () => {
            this.addMultipleStudents();
        });

        // 清空批量輸入按鈕
        document.getElementById('clearMultipleInput').addEventListener('click', () => {
            this.clearMultipleInput();
        });

        // 載入預設學生按鈕
        document.getElementById('loadDefaultStudents').addEventListener('click', () => {
            this.loadDefaultStudents();
        });

        // 字體大小控制
        document.getElementById('fontSize').addEventListener('input', (e) => {
            this.updateFontSize(e.target.value);
        });

        // 班級名稱輸入
        document.getElementById('className').addEventListener('input', (e) => {
            this.updateClassName(e.target.value);
        });

        // 相鄰限制：新增與UI變更
        const addConstraintBtn = document.getElementById('addConstraint');
        if (addConstraintBtn) {
            addConstraintBtn.addEventListener('click', () => {
                this.addConstraintFromUI();
            });
        }
    }

    // 批量新增學生功能
    addMultipleStudents() {
        const textarea = document.getElementById('multipleStudents');
        const content = textarea.value.trim();
        
        if (!content) {
            this.showToast('請輸入學生姓名', 'error');
            return;
        }
        
        // 分割輸入內容，支援多種分隔符號
        const names = content.split(/[\n\r,，、\s]+/).filter(name => name.trim());
        
        if (names.length === 0) {
            this.showToast('沒有找到有效的學生姓名', 'error');
            return;
        }
        
        let addedCount = 0;
        let duplicateCount = 0;
        let invalidCount = 0;
        
        names.forEach(name => {
            const trimmedName = name.trim();
            
            // 驗證姓名
            if (!trimmedName || trimmedName.length > 20) {
                invalidCount++;
                return;
            }
            
            // 檢查是否已存在
            const existingStudent = this.students.find(student => student.name === trimmedName);
            if (existingStudent) {
                duplicateCount++;
                return;
            }
            
            // 解析姓名和備註
            let studentName = trimmedName;
            let studentNote = '';
            
            // 檢查是否包含備註分隔符號 "|"
            if (trimmedName.includes('|')) {
                const parts = trimmedName.split('|');
                studentName = parts[0].trim();
                studentNote = parts[1] ? parts[1].trim() : '';
                
                // 重新驗證姓名
                if (!studentName || studentName.length > 20) {
                    invalidCount++;
                    return;
                }
                
                // 重新檢查是否已存在
                const existingStudentWithName = this.students.find(student => student.name === studentName);
                if (existingStudentWithName) {
                    duplicateCount++;
                    return;
                }
            }
            
            // 新增學生
            const newStudent = {
                id: Date.now().toString() + Math.random().toString(36).substr(2, 9),
                name: studentName,
                note: studentNote
            };
            
            this.students.push(newStudent);
            addedCount++;
        });
        
        // 更新UI
        this.updateStudentList();
        this.updateStudentCount();
        this.saveToLocalStorage();
        
        // 清空輸入框
        textarea.value = '';
        
        // 顯示結果
        let message = `成功新增 ${addedCount} 名學生`;
        if (duplicateCount > 0) {
            message += `，跳過 ${duplicateCount} 個重複姓名`;
        }
        if (invalidCount > 0) {
            message += `，跳過 ${invalidCount} 個無效輸入`;
        }
        
        this.showToast(message, 'success');
    }

    // 清空批量輸入
    clearMultipleInput() {
        const textarea = document.getElementById('multipleStudents');
        textarea.value = '';
        textarea.focus();
        this.showToast('已清空批量輸入框', 'info');
    }

    applySettings() {
        const rows = parseInt(document.getElementById('rows').value);
        const cols = parseInt(document.getElementById('cols').value);

        if (rows < 1 || rows > 10 || cols < 1 || cols > 8) {
            this.showToast('請輸入有效的排數（1-10）和列數（1-8）', 'error');
            return;
        }

        // 檢查設定是否與原本一樣
        const isConfigUnchanged = rows === this.seatingConfig.rows && cols === this.seatingConfig.cols;
        
        if (isConfigUnchanged) {
            this.showToast('座位設定沒有變更，無需更新', 'info');
            return;
        }

        // 檢查是否有現有資料
        const hasExistingData = this.students.length > 0 || Object.keys(this.seatingMap).length > 0;
        const isConfigChanged = rows !== this.seatingConfig.rows || cols !== this.seatingConfig.cols;

        if (hasExistingData && isConfigChanged) {
            // 顯示資料遺失警示
            this.showDataLossWarning(rows, cols);
        } else {
            // 直接套用設定
            this.applySettingsDirectly(rows, cols);
        }
    }

    showDataLossWarning(newRows, newCols) {
        // 移除現有的警告對話框
        const existingWarning = document.querySelector('.data-loss-warning');
        if (existingWarning) {
            existingWarning.remove();
        }

        // 創建警告對話框
        const warningOverlay = document.createElement('div');
        warningOverlay.className = 'data-loss-warning';
        
        warningOverlay.innerHTML = `
            <div class="warning-content">
                <div class="warning-header">
                    <i class="fas fa-exclamation-triangle"></i>
                    <h3>座位設定修改警告</h3>
                </div>
                <div class="warning-body">
                    <p class="warning-message">修改座位設定可能會遺失現有資料！</p>
                    
                    <div class="setting-comparison">
                        <div class="current-setting">
                            <h4>目前設定</h4>
                            <div class="setting-value">${this.seatingConfig.rows}排 × ${this.seatingConfig.cols}列</div>
                        </div>
                        <div class="arrow">→</div>
                        <div class="new-setting">
                            <h4>新設定</h4>
                            <div class="setting-value">${newRows}排 × ${newCols}列</div>
                        </div>
                    </div>
                    
                    <div class="data-impact">
                        <h4>影響項目</h4>
                        <ul>
                            <li><i class="fas fa-users"></i> 學生人數：${this.students.length}人</li>
                            <li><i class="fas fa-chair"></i> 已安排座位：${Object.keys(this.seatingMap).length}人</li>
                            <li><i class="fas fa-trash"></i> 座位安排將被清空</li>
                        </ul>
                    </div>
                </div>
                <div class="warning-footer">
                    <button class="btn btn-secondary" onclick="seatingChart.cancelSettingsChange()">
                        <i class="fas fa-times"></i> 取消
                    </button>
                    <button class="btn btn-warning" onclick="seatingChart.applySettingsDirectly(${newRows}, ${newCols})">
                        <i class="fas fa-check"></i> 確定修改
                    </button>
                </div>
            </div>
        `;
        
        document.body.appendChild(warningOverlay);
        
        // 點擊背景關閉對話框
        warningOverlay.addEventListener('click', (e) => {
            if (e.target === warningOverlay) {
                this.cancelSettingsChange();
            }
        });
        
        // ESC鍵關閉對話框
        document.addEventListener('keydown', function closeWarning(e) {
            if (e.key === 'Escape') {
                seatingChart.cancelSettingsChange();
                document.removeEventListener('keydown', closeWarning);
            }
        });
    }

    cancelSettingsChange() {
        // 恢復原來的設定值
        document.getElementById('rows').value = this.seatingConfig.rows;
        document.getElementById('cols').value = this.seatingConfig.cols;
        
        // 移除警告對話框
        const warning = document.querySelector('.data-loss-warning');
        if (warning) {
            warning.remove();
        }
        
        this.showToast('已取消修改座位設定', 'info');
    }

    applySettingsDirectly(rows, cols) {
        const oldConfig = { ...this.seatingConfig };
        this.seatingConfig = { rows, cols };
        
        // 檢查是否需要清空座位安排
        const needsClearSeats = this.shouldClearSeats(oldConfig, this.seatingConfig);
        
        if (needsClearSeats) {
            // 智能保留有效的座位安排
            const removedCount = this.preserveValidSeats(this.seatingConfig);
            
            if (removedCount === 0) {
                this.showToast('座位設定已更新', 'success');
            }
        } else {
            // 設定擴展，保留所有現有座位安排
            this.showToast('座位設定已擴展，現有座位安排已保留', 'success');
        }
        
        this.renderSeatingGrid();
        this.saveToLocalStorage();
        
        // 移除警告對話框
        const warning = document.querySelector('.data-loss-warning');
        if (warning) {
            warning.remove();
        }
    }

    shouldClearSeats(oldConfig, newConfig) {
        // 如果新設定比舊設定小，需要檢查座位安排
        if (newConfig.rows < oldConfig.rows || newConfig.cols < oldConfig.cols) {
            return true;
        }
        
        // 檢查是否有座位超出新設定的範圍
        const hasOutOfRangeSeats = Object.keys(this.seatingMap).some(seatKey => {
            const [row, col] = seatKey.split('-').map(Number);
            return row > newConfig.rows || col > newConfig.cols;
        });
        
        return hasOutOfRangeSeats;
    }

    // 智能保留座位安排
    preserveValidSeats(newConfig) {
        const validSeats = {};
        let removedCount = 0;
        
        Object.keys(this.seatingMap).forEach(seatKey => {
            const [row, col] = seatKey.split('-').map(Number);
            
            // 檢查座位是否在新設定範圍內
            if (row <= newConfig.rows && col <= newConfig.cols) {
                validSeats[seatKey] = this.seatingMap[seatKey];
            } else {
                removedCount++;
            }
        });
        
        this.seatingMap = validSeats;
        
        if (removedCount > 0) {
            this.showToast(`座位設定已更新，${removedCount}個超出範圍的座位安排已移除`, 'warning');
        }
        
        return removedCount;
    }

    addStudent() {
        const nameInput = document.getElementById('studentName');
        const noteInput = document.getElementById('studentNote');
        const name = nameInput.value.trim();
        const note = noteInput.value.trim();

        if (!name) {
            this.showToast('請輸入學生姓名', 'error');
            return;
        }

        // 檢查姓名是否已存在
        if (this.students.some(student => student.name === name)) {
            this.showToast('此姓名已存在', 'error');
            return;
        }

        const student = {
            id: Date.now().toString(), // 使用時間戳作為唯一ID
            name: name,
            note: note || '', // 添加備註欄位
            addedAt: new Date().toISOString()
        };

        this.students.push(student);
        this.updateStudentList();
        this.updateStudentCount();
        this.refreshConstraintSelectors();
        this.saveToLocalStorage();

        // 清空輸入欄位
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
            
            // 從座位中移除
            Object.keys(this.seatingMap).forEach(seatKey => {
                if (this.seatingMap[seatKey] === studentId) {
                    delete this.seatingMap[seatKey];
                }
            });

            this.updateStudentList();
            this.updateStudentCount();
            this.renderSeatingGrid();
            // 移除與該學生相關的限制
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
            
            // 檢查學生是否已有座位，並找到座位位置
            const hasSeat = Object.values(this.seatingMap).includes(student.id);
            let seatLocation = '';
            if (hasSeat) {
                // 找到學生被安排到哪個座位
                for (const [seatKey, studentId] of Object.entries(this.seatingMap)) {
                    if (studentId === student.id) {
                        seatLocation = seatKey;
                        break;
                    }
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
                    <button class="btn btn-info" onclick="seatingChart.editStudent('${student.id}')">
                        <i class="fas fa-edit"></i>
                    </button>
                    <button class="btn btn-danger" onclick="seatingChart.removeStudent('${student.id}')">
                        <i class="fas fa-trash"></i>
                    </button>
                </div>
            `;
            
            // 添加拖拽事件
            this.addDragEvents(studentItem, student);
            studentsList.appendChild(studentItem);
        });

        // 刷新相鄰限制選單
        this.refreshConstraintSelectors();
        // 重新渲染限制清單（名字可能變更）
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
                            </div>
                        `;
                    }
                } else {
                    seat.classList.add('empty-seat');
                    seatContent.innerHTML = '<div class="seat-student">空位</div>';
                }

                seat.appendChild(seatNumber);
                seat.appendChild(seatContent);

                // 添加拖拽事件
                this.addSeatDropEvents(seat, seatKey);

                grid.appendChild(seat);
            }
        }
    }

    // 取得座位周邊八格的鍵值
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

    // 檢查把 studentId 放到 seatKey 是否違反限制
    violatesConstraints(studentId, seatKey) {
        const neighbors = this.getAdjacentSeatKeys(seatKey);
        // 找與 studentId 有限制的對象集合
        const forbiddenSet = new Set();
        this.constraints.forEach(c => {
            if (c.a === studentId) forbiddenSet.add(c.b);
            if (c.b === studentId) forbiddenSet.add(c.a);
        });
        if (forbiddenSet.size === 0) return false;
        for (const nKey of neighbors) {
            const neighborId = this.seatingMap[nKey];
            if (neighborId && forbiddenSet.has(neighborId)) {
                return true;
            }
        }
        return false;
    }

    handleSeatClick(seatKey, seatElement) {
        if (this.seatingMap[seatKey]) {
            // 座位已有人，顯示選項：移除或更換學生
            const currentStudent = this.students.find(s => s.id === this.seatingMap[seatKey]);
            const action = confirm(`座位 ${seatKey} 目前由 ${currentStudent.name} 佔用。\n\n點擊「確定」移除學生\n點擊「取消」選擇其他學生替換`);
            
            if (action) {
                // 移除學生
                delete this.seatingMap[seatKey];
                this.renderSeatingGrid();
                this.updateStudentList();
                this.saveToLocalStorage();
                this.showToast(`已將 ${currentStudent.name} 從座位移除`, 'info');
            } else {
                // 選擇其他學生替換
                this.showStudentSelectionDialog(seatKey, true);
            }
        } else {
            // 座位是空的，顯示學生選擇對話框
            this.showStudentSelectionDialog(seatKey);
        }
    }

    showStudentSelectionDialog(seatKey, isReplacement = false) {
        if (this.students.length === 0) {
            this.showToast('請先新增學生', 'error');
            return;
        }

        let availableStudents;
        if (isReplacement) {
            // 替換模式：顯示所有學生（包括已有座位的）
            availableStudents = this.students;
        } else {
            // 正常模式：只顯示未安排座位的學生
            availableStudents = this.students.filter(student => 
                !Object.values(this.seatingMap).includes(student.id)
            );
        }

        if (availableStudents.length === 0) {
            this.showToast('沒有可選擇的學生', 'info');
            return;
        }

        // 創建自定義選擇對話框
        this.createStudentSelectionModal(seatKey, availableStudents, isReplacement);
    }

    createStudentSelectionModal(seatKey, availableStudents, isReplacement = false) {
        // 移除現有的modal
        const existingModal = document.querySelector('.student-selection-modal');
        if (existingModal) {
            existingModal.remove();
        }

        // 創建modal背景
        const modalOverlay = document.createElement('div');
        modalOverlay.className = 'student-selection-modal';
        
        // 創建modal內容
        const modalContent = document.createElement('div');
        modalContent.className = 'modal-content';
        
        const modalTitle = isReplacement ? `選擇學生替換座位 ${seatKey} 的學生` : `選擇學生安排到座位 ${seatKey}`;
        
        modalContent.innerHTML = `
            <div class="modal-header">
                <h3>${modalTitle}</h3>
                <button class="modal-close" onclick="this.closest('.student-selection-modal').remove()">
                    <i class="fas fa-times"></i>
                </button>
            </div>
            <div class="modal-body">
                <div class="student-options">
                    ${availableStudents.map(student => {
                        // 檢查學生是否已有座位
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
                            </div>
                        `;
                    }).join('')}
                </div>
            </div>
            <div class="modal-footer">
                <button class="btn btn-secondary" onclick="this.closest('.student-selection-modal').remove()">
                    取消
                </button>
            </div>
        `;
        
        modalOverlay.appendChild(modalContent);
        document.body.appendChild(modalOverlay);
        
        // 點擊背景關閉modal
        modalOverlay.addEventListener('click', (e) => {
            if (e.target === modalOverlay) {
                modalOverlay.remove();
            }
        });
        
        // ESC鍵關閉modal
        document.addEventListener('keydown', function closeModal(e) {
            if (e.key === 'Escape') {
                modalOverlay.remove();
                document.removeEventListener('keydown', closeModal);
            }
        });
    }

    selectStudentForSeat(seatKey, studentId, isReplacement = false) {
        const student = this.students.find(s => s.id === studentId);
        if (student) {
            if (isReplacement && this.seatingMap[seatKey]) {
                // 替換模式：檢查是否需要座位互換
                const currentStudent = this.students.find(s => s.id === this.seatingMap[seatKey]);
                if (currentStudent && currentStudent.id !== student.id) {
                    this.swapStudents(student, currentStudent, seatKey);
                } else {
                    // 選擇了同一個學生，不做任何操作
                    this.showToast('選擇了同一個學生', 'info');
                }
            } else {
                // 正常模式：直接安排座位
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
            
            // 關閉modal
            const modal = document.querySelector('.student-selection-modal');
            if (modal) {
                modal.remove();
            }
        }
    }

    showStudentSelectionForMobile(student) {
        // 手機版專用的學生選擇功能
        const availableSeats = [];
        for (let row = 1; row <= this.seatingConfig.rows; row++) {
            for (let col = 1; col <= this.seatingConfig.cols; col++) {
                const seatKey = `${row}-${col}`;
                if (!this.seatingMap[seatKey]) {
                    availableSeats.push(seatKey);
                }
            }
        }

        if (availableSeats.length === 0) {
            this.showToast('沒有空座位可選擇', 'error');
            return;
        }

        // 創建手機版選擇對話框
        this.createMobileSeatSelectionModal(student, availableSeats);
    }

    createMobileSeatSelectionModal(student, availableSeats) {
        // 移除現有的modal
        const existingModal = document.querySelector('.student-selection-modal');
        if (existingModal) {
            existingModal.remove();
        }

        // 創建modal背景
        const modalOverlay = document.createElement('div');
        modalOverlay.className = 'student-selection-modal';
        
        // 創建modal內容
        const modalContent = document.createElement('div');
        modalContent.className = 'modal-content';
        
        modalContent.innerHTML = `
            <div class="modal-header">
                <h3>為 ${student.name} 選擇座位</h3>
                <button class="modal-close" onclick="this.closest('.student-selection-modal').remove()">
                    <i class="fas fa-times"></i>
                </button>
            </div>
            <div class="modal-body">
                <div class="mobile-seat-grid">
                    ${this.generateMobileSeatGrid(availableSeats)}
                </div>
            </div>
            <div class="modal-footer">
                <button class="btn btn-secondary" onclick="this.closest('.student-selection-modal').remove()">
                    取消
                </button>
            </div>
        `;
        
        modalOverlay.appendChild(modalContent);
        document.body.appendChild(modalOverlay);
        
        // 點擊背景關閉modal
        modalOverlay.addEventListener('click', (e) => {
            if (e.target === modalOverlay) {
                modalOverlay.remove();
            }
        });
        
        // ESC鍵關閉modal
        document.addEventListener('keydown', function closeModal(e) {
            if (e.key === 'Escape') {
                modalOverlay.remove();
                document.removeEventListener('keydown', closeModal);
            }
        });
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
                        </div>
                    `;
                } else if (currentStudent) {
                    gridHTML += `
                        <div class="mobile-seat-option occupied">
                            <div class="mobile-seat-number">${seatKey}</div>
                            <div class="mobile-seat-student">${currentStudent.name}</div>
                        </div>
                    `;
                } else {
                    gridHTML += `
                        <div class="mobile-seat-option disabled">
                            <div class="mobile-seat-number">${seatKey}</div>
                            <div class="mobile-seat-status">不可用</div>
                        </div>
                    `;
                }
            }
            gridHTML += '</div>';
        }
        return gridHTML;
    }

    autoArrangeSeats() {
        if (this.students.length === 0) {
            this.showToast('請先新增學生', 'error');
            return;
        }

        // 清空現有座位
        this.seatingMap = {};

        // 隨機打亂學生順序
        const shuffledStudents = [...this.students].sort(() => Math.random() - 0.5);

        // 創建所有可用座位的陣列
        const availableSeats = [];
        for (let row = 1; row <= this.seatingConfig.rows; row++) {
            for (let col = 1; col <= this.seatingConfig.cols; col++) {
                availableSeats.push(`${row}-${col}`);
            }
        }

        // 隨機打亂座位順序
        const shuffledSeats = availableSeats.sort(() => Math.random() - 0.5);

        // 依序放置，盡量避免違反限制
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
            if (!placed) {
                notPlaced++;
            }
        }

        this.renderSeatingGrid();
        this.updateStudentList(); // 更新學生清單狀態
        this.saveToLocalStorage();
        if (notPlaced > 0) {
            this.showToast(`排座完成，但有 ${notPlaced} 位因限制未安排`, 'warning');
        } else {
            this.showToast('座位已隨機安排完成', 'success');
        }
    }

    clearSeats() {
        if (Object.keys(this.seatingMap).length === 0) {
            this.showToast('目前沒有安排任何座位', 'info');
            return;
        }

        if (confirm('確定要清空所有座位安排嗎？')) {
            this.seatingMap = {};
            this.renderSeatingGrid();
            this.updateStudentList(); // 更新學生清單狀態
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

        const dataStr = JSON.stringify(config, null, 2);
        const dataBlob = new Blob([dataStr], { type: 'application/json' });
        
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

                    // 更新UI
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

        // 清空檔案輸入
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

                // 更新UI
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
            // 如果沒有儲存的資料，初始化預設學生（不清空座位安排）
            this.initializeDefaultStudents(false);
        }
    }

    initializeDefaultStudents(clearSeating = true) {
        // 預設學生名單
        const defaultStudents = [
            '陳昱允', '彭翊恩', '章彥廷', '曾依凡', '張瑞恩',
            '張宸晞', '麥惠媛', '陳柳鈴', '吳睿凱', '詹欣容',
            '李誠恩', '涂毅宏', '曾聿寧', '方語喆', '王勻希',
            '伊妍欣', '麥庭綺', '彭立宸', '吳睿勳', '連廷駿',
            '余柏勳', '邱閔昊', '伊妤欣'
        ];

        // 清空現有學生
        this.students = [];

        // 新增預設學生
        defaultStudents.forEach((name, index) => {
            const student = {
                id: `default_${Date.now()}_${index}`,
                name: name,
                note: '',
                addedAt: new Date().toISOString()
            };
            this.students.push(student);
        });

        // 如果需要清空座位安排
        if (clearSeating) {
            this.seatingMap = {};
        }

        // 更新UI
        this.updateStudentList();
        this.updateStudentCount();
        this.renderSeatingGrid();
        this.saveToLocalStorage();
        
        this.showToast(`已載入 ${defaultStudents.length} 位預設學生`, 'success');
    }

    // ===== 相鄰限制 UI 與資料 =====
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
        // 檢查是否已存在（無序配對）
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
            const aName = nameById(c.a);
            const bName = nameById(c.b);
            return `
                <div class=\"constraint-item\">\n                    <span class=\"constraint-pair\">${aName} ↔ ${bName}</span>\n                    <button class=\"btn btn-danger constraint-remove\" onclick=\"seatingChart.removeConstraint(${idx})\"><i class=\"fas fa-trash\"></i></button>\n                </div>
            `;
        }).join('');
    }

    loadDefaultStudents() {
        // 顯示確認對話框
        const hasExistingStudents = this.students.length > 0;
        const hasSeatingArrangement = Object.keys(this.seatingMap).length > 0;
        
        let message = '確定要載入預設學生嗎？';
        if (hasExistingStudents) {
            message += '\n\n⚠️ 這將會替換現有的學生清單';
        }
        if (hasSeatingArrangement) {
            message += '\n\n⚠️ 現有的座位安排將會被清空';
        }
        
        if (confirm(message)) {
            this.initializeDefaultStudents();
        }
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
        // 桌面版拖拽事件
        studentItem.addEventListener('dragstart', (e) => {
            this.draggedStudent = student;
            this.draggedElement = studentItem;
            e.dataTransfer.setData('text/plain', student.id);
            e.dataTransfer.effectAllowed = 'move';
            studentItem.classList.add('dragging');
            
            // 添加拖拽提示
            this.showDragHint(student);
        });

        studentItem.addEventListener('dragend', (e) => {
            studentItem.classList.remove('dragging');
            this.draggedStudent = null;
            this.draggedElement = null;
            
            // 移除拖拽提示
            this.hideDragHint();
            
            // 移除所有座位的拖拽效果
            document.querySelectorAll('.seat').forEach(seat => {
                seat.classList.remove('drag-over', 'swap-target');
            });
        });

        // 手機版觸控事件
        let touchStartTime = 0;
        let touchStartY = 0;
        let touchStartX = 0;
        let isLongPress = false;
        let longPressTimer = null;

        studentItem.addEventListener('touchstart', (e) => {
            touchStartTime = Date.now();
            touchStartY = e.touches[0].clientY;
            touchStartX = e.touches[0].clientX;
            isLongPress = false;
            
            // 長按觸發拖拽模式
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
                
                // 視覺反饋
                studentItem.style.transform = `translate(${deltaX}px, ${deltaY}px) rotate(5deg) scale(0.95)`;
            }
        });

        studentItem.addEventListener('touchend', (e) => {
            clearTimeout(longPressTimer);
            
            if (isLongPress && this.draggedStudent) {
                // 重置樣式
                studentItem.style.transform = '';
                studentItem.classList.remove('dragging');
                this.draggedStudent = null;
                this.draggedElement = null;
                
                // 檢查是否拖拽到座位上
                const touch = e.changedTouches[0];
                const elementBelow = document.elementFromPoint(touch.clientX, touch.clientY);
                const seatElement = elementBelow?.closest('.seat');
                
                if (seatElement) {
                    const seatKey = seatElement.dataset.seatKey;
                    if (seatKey) {
                        this.handleStudentDrop(seatKey, seatElement);
                    }
                }
            } else {
                // 短按顯示學生選擇對話框
                const touchDuration = Date.now() - touchStartTime;
                if (touchDuration < 300) {
                    this.showStudentSelectionForMobile(student);
                }
            }
        });
    }

    addSeatDropEvents(seat, seatKey) {
        seat.addEventListener('dragover', (e) => {
            e.preventDefault();
            e.dataTransfer.dropEffect = 'move';
            
            if (this.draggedStudent) {
                // 檢查是否會發生互換
                if (this.seatingMap[seatKey] && this.seatingMap[seatKey] !== this.draggedStudent.id) {
                    seat.classList.add('swap-target');
                    seat.classList.remove('drag-over');
                } else {
                    seat.classList.add('drag-over');
                    seat.classList.remove('swap-target');
                }
            }
        });

        seat.addEventListener('dragleave', (e) => {
            seat.classList.remove('drag-over', 'swap-target');
        });

        seat.addEventListener('drop', (e) => {
            e.preventDefault();
            seat.classList.remove('drag-over', 'swap-target');
            
            if (this.draggedStudent) {
                this.handleStudentDrop(seatKey, seat);
            }
        });

        // 保留點擊功能作為備用
        seat.addEventListener('click', () => {
            this.handleSeatClick(seatKey, seat);
        });
    }

    handleStudentDrop(seatKey, seatElement) {
        if (!this.draggedStudent) return;

        // 檢查座位是否已被佔用
        if (this.seatingMap[seatKey]) {
            const currentStudent = this.students.find(s => s.id === this.seatingMap[seatKey]);
            if (currentStudent && currentStudent.id !== this.draggedStudent.id) {
                // 顯示互換確認提示
                this.showSwapConfirmation(this.draggedStudent, currentStudent, seatKey);
                return;
            }
        }

        // 移除學生之前的座位（如果有的話）
        Object.keys(this.seatingMap).forEach(key => {
            if (this.seatingMap[key] === this.draggedStudent.id) {
                delete this.seatingMap[key];
            }
        });

        // 檢查限制
        if (this.violatesConstraints(this.draggedStudent.id, seatKey)) {
            this.showToast('違反相鄰限制，請選擇其他座位', 'error');
            return;
        }

        // 分配新座位
        this.seatingMap[seatKey] = this.draggedStudent.id;
        this.renderSeatingGrid();
        this.updateStudentList();
        this.saveToLocalStorage();
        
        this.showToast(`已將 ${this.draggedStudent.name} 安排到座位 ${seatKey}`, 'success');
    }

    showSwapConfirmation(student1, student2, targetSeatKey) {
        // 找到學生1原本的座位
        let student1OriginalSeat = null;
        Object.keys(this.seatingMap).forEach(key => {
            if (this.seatingMap[key] === student1.id) {
                student1OriginalSeat = key;
            }
        });

        const message = student1OriginalSeat 
            ? `確定要將 ${student1.name} 和 ${student2.name} 的座位互換嗎？\n\n${student1.name}: ${student1OriginalSeat} ↔ ${targetSeatKey}\n${student2.name}: ${targetSeatKey} ↔ ${student1OriginalSeat}`
            : `確定要將 ${student1.name} 安排到座位 ${targetSeatKey}，並移除 ${student2.name} 嗎？`;

        if (confirm(message)) {
            this.swapStudents(student1, student2, targetSeatKey);
        }
    }

    showDragHint(student) {
        // 創建拖拽提示
        const hint = document.createElement('div');
        hint.id = 'drag-hint-popup';
        hint.className = 'drag-hint-popup';
        hint.innerHTML = `
            <div class="drag-hint-content">
                <i class="fas fa-exchange-alt"></i>
                <span>拖拽 ${student.name} 到座位進行更換</span>
            </div>
        `;
        document.body.appendChild(hint);
        
        // 3秒後自動移除
        setTimeout(() => {
            this.hideDragHint();
        }, 3000);
    }

    hideDragHint() {
        const hint = document.getElementById('drag-hint-popup');
        if (hint) {
            hint.remove();
        }
    }

    swapStudents(student1, student2, targetSeatKey) {
        // 找到學生1原本的座位
        let student1OriginalSeat = null;
        Object.keys(this.seatingMap).forEach(key => {
            if (this.seatingMap[key] === student1.id) {
                student1OriginalSeat = key;
            }
        });

        if (student1OriginalSeat) {
            // 執行座位互換前檢查限制
            const originalSeat2 = Object.keys(this.seatingMap).find(key => this.seatingMap[key] === student2.id) || targetSeatKey;
            const newSeatFor1 = targetSeatKey;
            const newSeatFor2 = student1OriginalSeat;
            if (this.violatesConstraints(student1.id, newSeatFor1) || this.violatesConstraints(student2.id, newSeatFor2)) {
                this.showToast('交換後違反相鄰限制，操作已取消', 'error');
                return;
            }
            // 執行座位互換
            this.seatingMap[student1OriginalSeat] = student2.id;
            this.seatingMap[targetSeatKey] = student1.id;
            
            this.renderSeatingGrid();
            this.updateStudentList();
            this.saveToLocalStorage();
            
            this.showToast(`已將 ${student1.name} 和 ${student2.name} 的座位互換`, 'success');
        } else {
            // 如果學生1原本沒有座位，直接替換前檢查限制
            if (this.violatesConstraints(student1.id, targetSeatKey)) {
                this.showToast('違反相鄰限制，操作已取消', 'error');
                return;
            }
            this.seatingMap[targetSeatKey] = student1.id;
            
            this.renderSeatingGrid();
            this.updateStudentList();
            this.saveToLocalStorage();
            
            this.showToast(`已將 ${student1.name} 安排到座位 ${targetSeatKey}，${student2.name} 的座位已移除`, 'success');
        }
    }

    printSeatingChart() {
        if (this.students.length === 0) {
            this.showToast('請先新增學生', 'error');
            return;
        }

        // 創建列印視窗
        const printWindow = window.open('', '_blank');
        const currentDate = new Date().toLocaleDateString('zh-TW');
        
        printWindow.document.write(`
            <!DOCTYPE html>
            <html lang="zh-TW">
            <head>
                <meta charset="UTF-8">
                <title>座位表 - ${currentDate}</title>
                <style>
                    body {
                        font-family: 'Microsoft JhengHei', Arial, sans-serif;
                        margin: 20px;
                        background: white;
                    }
                    .print-header {
                        text-align: center;
                        margin-bottom: 30px;
                        border-bottom: 2px solid #333;
                        padding-bottom: 20px;
                    }
                    .print-header h1 {
                        font-size: 24px;
                        margin: 0 0 10px 0;
                        color: #333;
                    }
                    .print-info {
                        font-size: 14px;
                        color: #666;
                        margin-bottom: 20px;
                    }
                    .print-grid {
                        display: grid;
                        gap: 10px;
                        margin: 20px 0;
                        page-break-inside: avoid;
                    }
                    .print-seat {
                        border: 2px solid #333;
                        padding: 15px;
                        text-align: center;
                        background: #f9f9f9;
                        min-height: 60px;
                        display: flex;
                        flex-direction: column;
                        justify-content: center;
                        align-items: center;
                    }
                    .print-seat.occupied {
                        background: #e3f2fd;
                        border-color: #2196f3;
                    }
                    .print-seat-number {
                        font-size: 12px;
                        color: #666;
                        margin-bottom: 5px;
                    }
                    .print-student-name {
                        font-size: ${this.printFontSize}px;
                        font-weight: bold;
                        color: #333;
                    }
                    .print-student-note {
                        font-size: 12px;
                        color: #666;
                        font-style: italic;
                        margin-top: 3px;
                    }
                    .print-empty {
                        color: #999;
                        font-style: italic;
                    }
                    .print-teacher-desk {
                        text-align: center;
                        margin-bottom: 40px;
                        padding: 25px;
                        background: #2d3748;
                        color: white;
                        border: 3px solid #4a5568;
                        border-radius: 15px;
                        font-weight: 700;
                        font-size: 22px;
                        position: relative;
                        box-shadow: 0 8px 25px rgba(0,0,0,0.3);
                    }
                    .print-teacher-desk:before {
                        content: '';
                        position: absolute;
                        top: -2px;
                        left: -2px;
                        right: -2px;
                        bottom: -2px;
                        background: linear-gradient(45deg, #667eea, #764ba2, #f093fb, #f5576c);
                        border-radius: 17px;
                        z-index: -1;
                    }
                    .print-teacher-desk-content {
                        display: flex;
                        align-items: center;
                        justify-content: center;
                        gap: 15px;
                        position: relative;
                        z-index: 1;
                    }
                    .print-teacher-desk i {
                        font-size: 28px;
                        color: #ffd700;
                        text-shadow: 2px 2px 4px rgba(0,0,0,0.5);
                    }
                    .print-teacher-desk-text {
                        font-size: 24px;
                        font-weight: 800;
                        text-shadow: 2px 2px 4px rgba(0,0,0,0.5);
                        letter-spacing: 2px;
                    }
                    .print-footer {
                        margin-top: 30px;
                        text-align: center;
                        font-size: 12px;
                        color: #666;
                        border-top: 1px solid #ccc;
                        padding-top: 20px;
                    }
                    @media print {
                        body { margin: 0; }
                        .print-seat { page-break-inside: avoid; }
                        .print-teacher-desk {
                            background: #2d3748 !important;
                            border: 3px solid #000 !important;
                            color: #000 !important;
                            box-shadow: none !important;
                        }
                        .print-teacher-desk:before {
                            display: none !important;
                        }
                        .print-teacher-desk-content {
                            color: #000 !important;
                        }
                        .print-teacher-desk i {
                            color: #000 !important;
                            text-shadow: none !important;
                        }
                        .print-teacher-desk-text {
                            color: #000 !important;
                            text-shadow: none !important;
                        }
                    }
                </style>
            </head>
            <body>
                <div class="print-header">
                    <h1>${this.className}學生座位表</h1>
                    <div class="print-info">
                        班級：${this.className} | 
                        座位配置：${this.seatingConfig.rows}排${this.seatingConfig.cols}列 | 
                        學生人數：${this.students.length}人 | 
                        已安排座位：${Object.keys(this.seatingMap).length}人 | 
                        日期：${currentDate}
                    </div>
                </div>
                
                <div class="print-teacher-desk">
                    <div class="print-teacher-desk-content">
                        <i class="fas fa-chalkboard"></i>
                        <span class="print-teacher-desk-text">講台</span>
                    </div>
                </div>
                
                <div class="print-grid" style="grid-template-columns: repeat(${this.seatingConfig.cols}, 1fr);">
                    ${this.generatePrintSeats()}
                </div>
                
                <div class="print-footer">
                    <p>座位表由學生座位排程系統生成</p>
                    <p>列印時間：${new Date().toLocaleString('zh-TW')}</p>
                </div>
            </body>
            </html>
        `);
        
        printWindow.document.close();
        printWindow.focus();
        
        // 等待內容載入後列印
        setTimeout(() => {
            printWindow.print();
            printWindow.close();
        }, 500);
        
        this.showToast('列印視窗已開啟', 'success');
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
                        </div>
                    `;
                } else {
                    seatsHTML += `
                        <div class="print-seat">
                            <div class="print-seat-number">座位 ${seatKey}</div>
                            <div class="print-student-name print-empty">空位</div>
                        </div>
                    `;
                }
            }
        }
        return seatsHTML;
    }

    exportToWord() {
        if (this.students.length === 0) {
            this.showToast('請先新增學生', 'error');
            return;
        }

        const currentDate = new Date().toLocaleDateString('zh-TW');
        const currentTime = new Date().toLocaleString('zh-TW');
        
        // 生成Word文檔的HTML內容
        const wordContent = `
            <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
            <head>
                <meta charset="UTF-8">
                <title>學生座位表</title>
                <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
                <style>
                    body {
                        font-family: 'Microsoft JhengHei', Arial, sans-serif;
                        margin: 20px;
                        line-height: 1.6;
                    }
                    .word-header {
                        text-align: center;
                        margin-bottom: 30px;
                        border-bottom: 2px solid #333;
                        padding-bottom: 20px;
                    }
                    .word-header h1 {
                        font-size: 24px;
                        margin: 0 0 10px 0;
                        color: #333;
                    }
                    .word-info {
                        font-size: 14px;
                        color: #666;
                        margin-bottom: 20px;
                    }
                    .word-table {
                        width: 100%;
                        border-collapse: collapse;
                        margin: 20px 0;
                        page-break-inside: avoid;
                    }
                    .word-table td {
                        border: 2px solid #333;
                        padding: 15px;
                        text-align: center;
                        vertical-align: middle;
                        min-height: 60px;
                        background: #f9f9f9;
                    }
                    .word-table td.occupied {
                        background: #e3f2fd;
                        border-color: #2196f3;
                    }
                    .word-seat-number {
                        font-size: 12px;
                        color: #666;
                        margin-bottom: 5px;
                    }
                    .word-student-name {
                        font-size: ${this.printFontSize}px;
                        font-weight: bold;
                        color: #333;
                    }
                    .word-student-note {
                        font-size: 12px;
                        color: #666;
                        font-style: italic;
                        margin-top: 3px;
                    }
                    .word-empty {
                        color: #999;
                        font-style: italic;
                    }
                    .word-teacher-desk {
                        text-align: center;
                        margin-bottom: 40px;
                        padding: 25px;
                        background: #2d3748;
                        color: white;
                        border: 3px solid #4a5568;
                        border-radius: 15px;
                        font-weight: 700;
                        font-size: 22px;
                        position: relative;
                        box-shadow: 0 8px 25px rgba(0,0,0,0.3);
                    }
                    .word-teacher-desk:before {
                        content: '';
                        position: absolute;
                        top: -2px;
                        left: -2px;
                        right: -2px;
                        bottom: -2px;
                        background: linear-gradient(45deg, #667eea, #764ba2, #f093fb, #f5576c);
                        border-radius: 17px;
                        z-index: -1;
                    }
                    .word-teacher-desk-content {
                        display: flex;
                        align-items: center;
                        justify-content: center;
                        gap: 15px;
                        position: relative;
                        z-index: 1;
                    }
                    .word-teacher-desk i {
                        font-size: 28px;
                        color: #ffd700;
                        text-shadow: 2px 2px 4px rgba(0,0,0,0.5);
                    }
                    .word-teacher-desk-text {
                        font-size: 24px;
                        font-weight: 800;
                        text-shadow: 2px 2px 4px rgba(0,0,0,0.5);
                        letter-spacing: 2px;
                    }
                    .word-footer {
                        margin-top: 30px;
                        text-align: center;
                        font-size: 12px;
                        color: #666;
                        border-top: 1px solid #ccc;
                        padding-top: 20px;
                    }
                </style>
            </head>
            <body>
                <div class="word-header">
                    <h1>${this.className}學生座位表</h1>
                    <div class="word-info">
                        班級：${this.className} | 
                        座位配置：${this.seatingConfig.rows}排${this.seatingConfig.cols}列 | 
                        學生人數：${this.students.length}人 | 
                        已安排座位：${Object.keys(this.seatingMap).length}人 | 
                        日期：${currentDate}
                    </div>
                </div>
                
                <div class="word-teacher-desk">
                    <div class="word-teacher-desk-content">
                        <i class="fas fa-chalkboard"></i>
                        <span class="word-teacher-desk-text">講台</span>
                    </div>
                </div>
                
                <table class="word-table">
                    ${this.generateWordTable()}
                </table>
                
                <div class="word-footer">
                    <p>座位表由學生座位排程系統生成</p>
                    <p>匯出時間：${currentTime}</p>
                </div>
            </body>
            </html>
        `;

        // 創建Blob並下載
        const blob = new Blob([wordContent], { 
            type: 'application/msword' 
        });
        
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
                
                if (student) {
                    tableHTML += `
                        <td class="occupied">
                            <div class="word-seat-number">座位 ${seatKey}</div>
                            <div class="word-student-name">${student.name}</div>
                            ${student.note ? `<div class="word-student-note">${student.note}</div>` : ''}
                        </td>
                    `;
                } else {
                    tableHTML += `
                        <td>
                            <div class="word-seat-number">座位 ${seatKey}</div>
                            <div class="word-student-name word-empty">空位</div>
                        </td>
                    `;
                }
            }
            tableHTML += '</tr>';
        }
        return tableHTML;
    }

    showToast(message, type = 'info') {
        // 移除現有的toast
        const existingToast = document.querySelector('.toast');
        if (existingToast) {
            existingToast.remove();
        }

        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        toast.textContent = message;
        document.body.appendChild(toast);

        // 顯示動畫
        setTimeout(() => {
            toast.classList.add('show');
        }, 100);

        // 自動隱藏
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => {
                if (toast.parentNode) {
                    toast.remove();
                }
            }, 300);
        }, 3000);
    }
}

// 初始化應用程式
let seatingChart;
document.addEventListener('DOMContentLoaded', () => {
    seatingChart = new SeatingChart();
});

// 全域函數供HTML調用
window.seatingChart = seatingChart;
