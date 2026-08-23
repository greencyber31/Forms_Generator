document.addEventListener('DOMContentLoaded', () => {
    // State management
    let currentStep = 1;
    let workspace = {
        excel_file: null,
        template_file: null,
        transmittal_template: null,
        headers: [],
        app_docx_tags: [],
        trans_docx_tags: [],
        mappings: {
            template: {},
            transmittal: {}
        }
    };
    let activePreviewType = 'template'; // 'template' or 'transmittal'
    let eventSource = null;

    // Init App
    fetchWorkspace();
    setupEventListeners();

    function showToast(message, type = 'info') {
        const toastContainer = document.getElementById('toast-container');
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        const icon = type === 'success' ? 'fa-circle-check' : (type === 'error' ? 'fa-circle-xmark' : 'fa-circle-info');
        toast.innerHTML = `<i class="fa-solid ${icon}"></i> <span>${message}</span>`;
        toastContainer.appendChild(toast);
        setTimeout(() => {
            toast.style.opacity = '0';
            setTimeout(() => toast.remove(), 300);
        }, 4000);
    }

    function setStep(stepNum) {
        currentStep = stepNum;
        document.querySelectorAll('.step-item').forEach(item => {
            const step = parseInt(item.getAttribute('data-step'));
            if (step === stepNum) item.classList.add('active');
            else item.classList.remove('active');
        });
        document.querySelectorAll('.wizard-step').forEach((sec, idx) => {
            if (idx + 1 === stepNum) sec.classList.add('active');
            else sec.classList.remove('active');
        });

        if (stepNum === 2) {
            loadTagConnector(activePreviewType);
        } else if (stepNum === 3) {
            populateGroupingSelects();
        }
    }

    // Mode State: 'generator' or 'viewer'
    let currentAppMode = 'generator';

    function setAppMode(mode) {
        currentAppMode = mode;
        const generatorView = document.getElementById('form-generator-view');
        const wizardNav = document.getElementById('wizard-nav-container');
        const viewerView = document.getElementById('pdf-viewer-view');
        const viewerNav = document.getElementById('viewer-mode-nav');
        const modeIcon = document.getElementById('app-mode-icon');
        const modeIconI = document.getElementById('mode-icon-i');
        const brandTitle = document.getElementById('app-brand-title');
        const brandSubtitle = document.getElementById('app-brand-subtitle');
        const headerModeLabel = document.getElementById('btn-header-mode-label');
        const headerModeIcon = document.querySelector('#btn-header-mode-toggle i');

        if (mode === 'viewer') {
            if (generatorView) generatorView.style.display = 'none';
            if (wizardNav) wizardNav.style.display = 'none';
            if (viewerView) viewerView.style.display = 'block';
            if (viewerNav) viewerNav.style.display = 'flex';

            const explorerStage = document.getElementById('output-explorer-stage');
            const viewerStage = document.getElementById('output-viewer-stage');
            if (explorerStage) explorerStage.style.display = 'block';
            if (viewerStage) viewerStage.style.display = 'none';

            if (modeIcon) modeIcon.classList.add('viewer-mode-active');
            if (modeIconI) modeIconI.className = 'fa-solid fa-book-open-reader';
            if (brandTitle) brandTitle.textContent = 'PCIC PDF Viewer';
            if (brandSubtitle) brandSubtitle.textContent = 'Output Bundle Explorer & Interactive Reader';
            if (headerModeLabel) headerModeLabel.textContent = 'Form Studio Mode';
            if (headerModeIcon) headerModeIcon.className = 'fa-solid fa-wand-magic-sparkles';

            loadOutputTree(false);
            window.scrollTo({ top: 0, behavior: 'smooth' });
            showToast("Switched to Output Directory & PDF Library", "info");
        } else {
            if (generatorView) generatorView.style.display = 'block';
            if (wizardNav) wizardNav.style.display = 'flex';
            if (viewerView) viewerView.style.display = 'none';
            if (viewerNav) viewerNav.style.display = 'none';

            if (modeIcon) modeIcon.classList.remove('viewer-mode-active');
            if (modeIconI) modeIconI.className = 'fa-solid fa-file-pdf';
            if (brandTitle) brandTitle.textContent = 'PCIC Form Studio';
            if (brandSubtitle) brandSubtitle.textContent = 'Dynamic Batch PDF & Transmittal Engine';
            if (headerModeLabel) headerModeLabel.textContent = 'PDF Viewer Mode';
            if (headerModeIcon) headerModeIcon.className = 'fa-solid fa-layer-group';

            window.scrollTo({ top: 0, behavior: 'smooth' });
            showToast("Returned to Form Studio", "info");
        }
    }

    function setupEventListeners() {
        // App Mode Switchers (Brand Icon / Header Toggle Button / Batch Jump)
        const brandToggle = document.getElementById('brand-mode-toggle');
        const appModeIcon = document.getElementById('app-mode-icon');
        const btnHeaderToggle = document.getElementById('btn-header-mode-toggle');
        const btnJumpViewer = document.getElementById('btn-jump-to-viewer-mode');

        const toggleMode = () => setAppMode(currentAppMode === 'generator' ? 'viewer' : 'generator');

        if (appModeIcon) appModeIcon.addEventListener('click', (e) => {
            e.stopPropagation();
            toggleMode();
        });
        if (brandToggle) brandToggle.addEventListener('click', toggleMode);
        if (btnHeaderToggle) btnHeaderToggle.addEventListener('click', toggleMode);
        if (btnJumpViewer) btnJumpViewer.addEventListener('click', () => setAppMode('viewer'));

        // Wizard navigation
        document.querySelectorAll('.step-item').forEach(item => {
            item.addEventListener('click', () => {
                setStep(parseInt(item.getAttribute('data-step')));
            });
        });

        // Step buttons
        document.getElementById('btn-to-step-2').addEventListener('click', () => setStep(2));
        document.getElementById('btn-back-to-step-1').addEventListener('click', () => setStep(1));
        document.getElementById('btn-to-step-3').addEventListener('click', () => setStep(3));
        document.getElementById('btn-back-to-step-2').addEventListener('click', () => setStep(2));
        document.getElementById('btn-to-step-4').addEventListener('click', () => setStep(4));
        document.getElementById('btn-back-to-step-3').addEventListener('click', () => setStep(3));

        // Open output folder
        document.getElementById('btn-open-output').addEventListener('click', openOutputFolder);
        document.getElementById('btn-dash-open-output').addEventListener('click', openOutputFolder);

        // Upload zone bindings
        setupDropZone('drop-excel', 'input-excel', 'excel');
        setupDropZone('drop-template', 'input-template', 'template');
        setupDropZone('drop-transmittal', 'input-transmittal', 'transmittal');

        // Tab selectors (Application Form vs Transmittal)
        document.getElementById('tab-app-template').addEventListener('click', (e) => {
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            e.currentTarget.classList.add('active');
            activePreviewType = 'template';
            loadTagConnector(activePreviewType);
        });
        document.getElementById('tab-trans-template').addEventListener('click', (e) => {
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            e.currentTarget.classList.add('active');
            activePreviewType = 'transmittal';
            loadTagConnector(activePreviewType);
        });

        // Tag Connector actions
        document.getElementById('btn-auto-match').addEventListener('click', triggerAutoMatch);
        document.getElementById('btn-reset-match').addEventListener('click', resetMappings);
        document.getElementById('btn-refresh-preview').addEventListener('click', () => reloadSamplePdfPreview(activePreviewType));

        // Process button
        document.getElementById('btn-start-process').addEventListener('click', startBatchStream);
        document.getElementById('btn-clear-logs').addEventListener('click', () => {
            document.getElementById('log-container').innerHTML = '';
        });

        // Restart batch / Back to Step 1 handlers
        const handleRestart = () => {
            const btnRestart = document.getElementById('btn-restart-process');
            const btnFooterRestart = document.getElementById('btn-footer-restart-process');
            const btnJump = document.getElementById('btn-jump-to-viewer-mode');
            if (btnRestart) btnRestart.style.display = 'none';
            if (btnFooterRestart) btnFooterRestart.style.display = 'none';
            if (btnJump) btnJump.style.display = 'none';

            setStep(1);
            window.scrollTo({ top: 0, behavior: 'smooth' });
            showToast("Returned to Step 1. Ready for next batch!", "info");
        };

        const btnRestart = document.getElementById('btn-restart-process');
        const btnFooterRestart = document.getElementById('btn-footer-restart-process');
        if (btnRestart) btnRestart.addEventListener('click', handleRestart);
        if (btnFooterRestart) btnFooterRestart.addEventListener('click', handleRestart);
    }

    function loadTagConnector(type) {
        if (type === 'transmittal') {
            loadTransmittalColumnsConnector();
            return;
        }

        const container = document.getElementById('mapping-rows-container');
        const badge = document.getElementById('tag-count-badge');
        container.innerHTML = '<p style="color:var(--text-muted); padding:1rem;"><i class="fa-solid fa-spinner fa-spin"></i> Scanning document tags & loading field connectors...</p>';

        fetch(`/api/template/mapping?type=${type}`)
        .then(res => res.json())
        .then(data => {
            const tags = data.docx_tags || [];
            const headers = data.excel_headers || [];
            const mapping = data.mapping || {};

            badge.textContent = `${tags.length} Tags Detected`;

            if (tags.length === 0) {
                container.innerHTML = '<p style="color:var(--text-muted); padding:1rem;">No Jinja tags (e.g. {{ Tag }}) found in this template. Upload a template with tags in Step 1.</p>';
                return;
            }

            container.innerHTML = '';
            tags.forEach(tag => {
                const row = document.createElement('div');
                let rawVal = mapping[tag] || '';
                let isCustom = false;
                let isUnderline = false;
                let customVal = '';
                let selectedHeader = rawVal;

                if (rawVal.startsWith('STATIC:')) {
                    isCustom = true;
                    customVal = rawVal.substring(7);
                    selectedHeader = '__CUSTOM__';
                } else if (rawVal === '__BLANK_UNDERLINE__' || rawVal === 'UNDERLINE:') {
                    isUnderline = true;
                    selectedHeader = '__BLANK_UNDERLINE__';
                }

                const isMatched = Boolean(rawVal);

                row.className = `mapping-row-item ${isMatched ? 'matched' : 'unmapped'}`;
                row.setAttribute('data-tag', tag);

                let selectOptions = `<option value="">-- Ignore / Unmapped --</option>`;
                selectOptions += `<option value="__BLANK_UNDERLINE__" ${isUnderline ? 'selected' : ''}>Blank (Underline)</option>`;
                selectOptions += `<option value="__CUSTOM__" ${isCustom ? 'selected' : ''}>✍️ Custom Static Text...</option>`;
                if (headers.length > 0) {
                    selectOptions += `<optgroup label="Excel Headers">`;
                    headers.forEach(h => {
                        const sel = (!isCustom && !isUnderline && h === selectedHeader) ? 'selected' : '';
                        selectOptions += `<option value="${h}" ${sel}>${h}</option>`;
                    });
                    selectOptions += `</optgroup>`;
                }

                row.innerHTML = `
                    <div class="tag-label">
                        <span class="tag-badge">&#123;&#123; ${tag} &#125;&#125;</span>
                    </div>
                    <div class="mapping-input-group">
                        <button type="button" class="btn-unmap-field" data-tag="${tag}" title="Ignore / Unmap Field">
                            <i class="fa-solid fa-xmark"></i>
                        </button>
                        <select class="match-select" data-tag="${tag}">
                            ${selectOptions}
                        </select>
                        <input type="text" class="custom-value-input" data-tag="${tag}" placeholder="Type static text..." value="${customVal}" style="${isCustom ? 'display:inline-block;' : 'display:none;'}" />
                    </div>
                `;

                const selectEl = row.querySelector('.match-select');
                const inputEl = row.querySelector('.custom-value-input');
                const unmapBtn = row.querySelector('.btn-unmap-field');

                unmapBtn.addEventListener('click', (e) => {
                    e.preventDefault();
                    selectEl.value = '';
                    inputEl.value = '';
                    inputEl.style.display = 'none';
                    row.className = 'mapping-row-item unmapped';
                    saveCurrentMappings();
                });

                selectEl.addEventListener('change', () => {
                    if (selectEl.value === '__CUSTOM__') {
                        inputEl.style.display = 'inline-block';
                        inputEl.focus();
                        row.className = 'mapping-row-item matched';
                    } else if (selectEl.value === '__BLANK_UNDERLINE__') {
                        inputEl.style.display = 'none';
                        inputEl.value = '';
                        row.className = 'mapping-row-item matched';
                    } else if (selectEl.value) {
                        inputEl.style.display = 'none';
                        row.className = 'mapping-row-item matched';
                    } else {
                        inputEl.style.display = 'none';
                        row.className = 'mapping-row-item unmapped';
                    }
                    saveCurrentMappings();
                });

                let debounceTimer = null;
                inputEl.addEventListener('input', () => {
                    row.className = 'mapping-row-item matched';
                    clearTimeout(debounceTimer);
                    debounceTimer = setTimeout(() => {
                        saveCurrentMappings();
                    }, 500);
                });

                container.appendChild(row);
            });

            // Reload sample PDF preview
            reloadSamplePdfPreview(type);
        })
        .catch(err => {
            container.innerHTML = `<p style="color:var(--accent-red); padding:1rem;">Failed to load mapping: ${err}</p>`;
        });
    }

    function loadTransmittalColumnsConnector() {
        const container = document.getElementById('mapping-rows-container');
        const badge = document.getElementById('tag-count-badge');
        container.innerHTML = '<p style="color:var(--text-muted); padding:1rem;"><i class="fa-solid fa-spinner fa-spin"></i> Loading Excel columns configuration...</p>';

        fetch('/api/template/transmittal-columns')
        .then(res => res.json())
        .then(data => {
            const columns = data.columns || [];
            const activeCount = columns.filter(c => c.enabled).length;
            badge.textContent = `${activeCount} of ${columns.length} Excel Columns Selected`;

            if (columns.length === 0) {
                container.innerHTML = '<p style="color:var(--text-muted); padding:1rem;">No Excel columns loaded. Please upload an Excel file in Step 1.</p>';
                return;
            }

            container.innerHTML = `
                <div class="column-selector-toolbar">
                    <span class="toolbar-title"><i class="fa-solid fa-table-columns"></i> Transmittal Dynamic Column Selector</span>
                    <div class="toolbar-btns">
                        <button type="button" class="btn btn-xs btn-outline" id="btn-col-defaults"><i class="fa-solid fa-wand-magic-sparkles"></i> Select Standard Defaults</button>
                        <button type="button" class="btn btn-xs btn-outline" id="btn-col-select-all"><i class="fa-solid fa-check-double"></i> Select All</button>
                        <button type="button" class="btn btn-xs btn-outline" id="btn-col-deselect-all"><i class="fa-solid fa-square-xmark"></i> Deselect All</button>
                    </div>
                </div>
                <div class="trans-columns-list" id="trans-columns-list"></div>
            `;

            const listEl = container.querySelector('#trans-columns-list');

            columns.forEach(col => {
                const item = document.createElement('div');
                item.className = `trans-col-item ${col.enabled ? 'selected' : 'disabled'}`;
                item.setAttribute('data-header', col.header);

                item.innerHTML = `
                    <label class="col-check-label">
                        <input type="checkbox" class="trans-col-checkbox" ${col.enabled ? 'checked' : ''} />
                        <span class="excel-col-badge"><i class="fa-solid fa-file-excel"></i> ${col.header}</span>
                    </label>
                    <div class="col-order-box">
                        <span class="order-label">Col Position:</span>
                        <input type="number" min="1" max="99" class="col-order-input" value="${col.order || ''}" placeholder="#" />
                        <div class="col-order-nav">
                            <button type="button" class="btn-order-move btn-order-up" title="Increase Col Position (+1)"><i class="fa-solid fa-circle-arrow-up"></i></button>
                            <button type="button" class="btn-order-move btn-order-down" title="Decrease Col Position (-1)"><i class="fa-solid fa-circle-arrow-down"></i></button>
                        </div>
                    </div>
                `;

                const chk = item.querySelector('.trans-col-checkbox');
                const ordInput = item.querySelector('.col-order-input');
                const btnUp = item.querySelector('.btn-order-up');
                const btnDown = item.querySelector('.btn-order-down');

                btnUp.addEventListener('click', (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    let currentVal = parseInt(ordInput.value, 10);
                    if (isNaN(currentVal)) currentVal = 0;
                    ordInput.value = currentVal + 1;
                    if (!chk.checked) {
                        chk.checked = true;
                        item.className = 'trans-col-item selected';
                    }
                    saveTransmittalColumns();
                });

                btnDown.addEventListener('click', (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    let currentVal = parseInt(ordInput.value, 10);
                    if (isNaN(currentVal) || currentVal <= 1) {
                        ordInput.value = 1;
                    } else {
                        ordInput.value = currentVal - 1;
                    }
                    saveTransmittalColumns();
                });

                chk.addEventListener('change', () => {
                    item.className = `trans-col-item ${chk.checked ? 'selected' : 'disabled'}`;
                    if (chk.checked && !ordInput.value) {
                        let maxOrd = 0;
                        document.querySelectorAll('.col-order-input').forEach(inp => {
                            const val = parseInt(inp.value, 10);
                            if (!isNaN(val) && val > maxOrd) maxOrd = val;
                        });
                        ordInput.value = maxOrd + 1;
                    }
                    saveTransmittalColumns();
                });

                let timer = null;
                ordInput.addEventListener('input', () => {
                    clearTimeout(timer);
                    timer = setTimeout(() => {
                        saveTransmittalColumns();
                    }, 400);
                });

                listEl.appendChild(item);
            });

            container.querySelector('#btn-col-defaults').addEventListener('click', () => {
                fetch('/api/template/transmittal-columns', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ columns: [] })
                })
                .then(() => loadTransmittalColumnsConnector());
            });

            container.querySelector('#btn-col-select-all').addEventListener('click', () => {
                let idx = 1;
                document.querySelectorAll('.trans-col-item').forEach(item => {
                    item.className = 'trans-col-item selected';
                    const chk = item.querySelector('.trans-col-checkbox');
                    const inp = item.querySelector('.col-order-input');
                    chk.checked = true;
                    if (!inp.value) inp.value = idx;
                    idx++;
                });
                saveTransmittalColumns();
            });

            container.querySelector('#btn-col-deselect-all').addEventListener('click', () => {
                document.querySelectorAll('.trans-col-item').forEach(item => {
                    item.className = 'trans-col-item disabled';
                    const chk = item.querySelector('.trans-col-checkbox');
                    chk.checked = false;
                });
                saveTransmittalColumns();
            });

            reloadSamplePdfPreview('transmittal');
        })
        .catch(err => {
            container.innerHTML = `<p style="color:var(--accent-red); padding:1rem;">Failed to load column connector: ${err}</p>`;
        });
    }

    let saveTransColTimer = null;
    function saveTransmittalColumns() {
        clearTimeout(saveTransColTimer);
        saveTransColTimer = setTimeout(() => {
            const columns = [];
            document.querySelectorAll('.trans-col-item').forEach(item => {
                const header = item.getAttribute('data-header');
                const chk = item.querySelector('.trans-col-checkbox');
                const ordInput = item.querySelector('.col-order-input');
                const orderVal = parseInt(ordInput.value, 10);

                columns.push({
                    header: header,
                    enabled: chk ? chk.checked : false,
                    order: isNaN(orderVal) ? 999 : orderVal
                });
            });

            workspace.transmittal_columns = columns;

            fetch('/api/template/transmittal-columns', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ columns: columns })
            })
            .then(res => res.json())
            .then(() => {
                const activeCount = columns.filter(c => c.enabled).length;
                const badge = document.getElementById('tag-count-badge');
                if (badge) badge.textContent = `${activeCount} of ${columns.length} Excel Columns Selected`;
                reloadSamplePdfPreview('transmittal');
            });
        }, 300);
    }

    let saveMappingTimer = null;
    function saveCurrentMappings() {
        clearTimeout(saveMappingTimer);
        saveMappingTimer = setTimeout(() => {
            const mapping = {};
            document.querySelectorAll('.mapping-row-item').forEach(row => {
                const tag = row.getAttribute('data-tag');
                const select = row.querySelector('.match-select');
                const input = row.querySelector('.custom-value-input');

                if (select && select.value === '__CUSTOM__') {
                    const staticText = input ? input.value : '';
                    mapping[tag] = `STATIC:${staticText}`;
                } else if (select && select.value) {
                    mapping[tag] = select.value;
                }
            });

            workspace.mappings[activePreviewType] = mapping;

            fetch('/api/template/mapping', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    doc_type: activePreviewType,
                    mapping: mapping
                })
            })
            .then(res => res.json())
            .then(data => {
                reloadSamplePdfPreview(activePreviewType);
            });
        }, 300);
    }

    function triggerAutoMatch() {
        if (activePreviewType === 'transmittal') {
            showToast("Resetting to standard transmittal column defaults...", "info");
            fetch('/api/template/transmittal-columns', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ columns: [] })
            })
            .then(() => loadTransmittalColumnsConnector());
            return;
        }

        showToast("Auto-matching Word tags with Excel columns...", "info");
        fetch('/api/template/auto-match', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ doc_type: activePreviewType })
        })
        .then(res => res.json())
        .then(data => {
            if (data.error) {
                showToast(data.error, 'error');
            } else {
                showToast(data.message, 'success');
                loadTagConnector(activePreviewType);
            }
        })
        .catch(err => showToast(`Auto-match failed: ${err}`, 'error'));
    }

    function resetMappings() {
        if (activePreviewType === 'transmittal') {
            showToast("Unchecking all transmittal columns...", "info");
            document.querySelectorAll('.trans-col-item').forEach(item => {
                item.className = 'trans-col-item disabled';
                const chk = item.querySelector('.trans-col-checkbox');
                if (chk) chk.checked = false;
            });
            saveTransmittalColumns();
            return;
        }

        showToast("Resetting field mappings...", "info");
        workspace.mappings[activePreviewType] = {};
        saveCurrentMappings();
        loadTagConnector(activePreviewType);
    }

    function reloadSamplePdfPreview(type) {
        const iframe = document.getElementById('pdf-preview-frame');
        const overlay = document.getElementById('pdf-loading-overlay');
        const btnRefreshIcon = document.querySelector('#btn-refresh-preview i');

        if (overlay) overlay.classList.add('active');
        if (btnRefreshIcon) btnRefreshIcon.classList.add('fa-spin');

        const sampleUrl = `/api/template/pdf-preview?type=${type}&t=${Date.now()}`;
        const targetViewerSrc = `/static/pdfjs/web/viewer.html?file=${encodeURIComponent(sampleUrl)}#zoom=page-width`;

        const hideLoading = () => {
            if (overlay) overlay.classList.remove('active');
            if (btnRefreshIcon) btnRefreshIcon.classList.remove('fa-spin');
        };

        // If iframe is already initialized with PDF.js, hot-swap document in-memory without full reload
        try {
            if (iframe && iframe.contentWindow && iframe.contentWindow.PDFViewerApplication && iframe.contentWindow.PDFViewerApplication.open) {
                iframe.contentWindow.PDFViewerApplication.open({ url: sampleUrl }).then(() => {
                    hideLoading();
                }).catch(() => {
                    iframe.src = targetViewerSrc;
                });
                return;
            }
        } catch (e) {}

        iframe.onload = () => {
            hideLoading();
        };

        iframe.src = targetViewerSrc;
    }

    function setupDropZone(dropId, inputId, fileType) {
        const dropZone = document.getElementById(dropId);
        const input = document.getElementById(inputId);

        dropZone.addEventListener('click', () => input.click());

        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.classList.add('dragover');
        });
        dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
            if (e.dataTransfer.files.length > 0) {
                handleFileUpload(e.dataTransfer.files[0], fileType);
            }
        });

        input.addEventListener('change', () => {
            if (input.files.length > 0) {
                handleFileUpload(input.files[0], fileType);
            }
        });
    }

    function handleFileUpload(file, fileType) {
        const formData = new FormData();
        formData.append('file', file);
        formData.append('file_type', fileType);

        showToast(`Uploading ${file.name}...`, 'info');

        fetch('/api/upload', {
            method: 'POST',
            body: formData
        })
        .then(res => res.json())
        .then(data => {
            if (data.error) {
                showToast(data.error, 'error');
            } else {
                showToast(data.message, 'success');
                workspace = data.workspace;
                updateWorkspaceCards();
                if (currentStep === 2) {
                    loadTagConnector(activePreviewType);
                }
            }
        })
        .catch(err => showToast(`Upload failed: ${err}`, 'error'));
    }

    function fetchWorkspace() {
        fetch('/api/workspace')
        .then(res => res.json())
        .then(data => {
            workspace = data;
            updateWorkspaceCards();
        });
    }

    function updateWorkspaceCards() {
        const nameExcel = document.getElementById('name-excel');
        const statusExcel = document.getElementById('status-excel');
        if (workspace.excel_file) {
            nameExcel.textContent = workspace.excel_file;
            statusExcel.classList.add('loaded');
        } else {
            nameExcel.textContent = 'Not loaded';
            statusExcel.classList.remove('loaded');
        }

        const nameTemplate = document.getElementById('name-template');
        const statusTemplate = document.getElementById('status-template');
        if (workspace.template_file) {
            nameTemplate.textContent = workspace.template_file;
            statusTemplate.classList.add('loaded');
        } else {
            nameTemplate.textContent = 'Not loaded';
            statusTemplate.classList.remove('loaded');
        }

        const nameTrans = document.getElementById('name-transmittal');
        const statusTrans = document.getElementById('status-transmittal');
        if (workspace.transmittal_template) {
            nameTrans.textContent = workspace.transmittal_template;
            statusTrans.classList.add('loaded');
        } else {
            nameTrans.textContent = 'Not loaded';
            statusTrans.classList.remove('loaded');
        }

        populateGroupingSelects();
    }

    function populateGroupingSelects() {
        const selectPrimary = document.getElementById('select-primary-group');
        const selectSecondary = document.getElementById('select-secondary-group');
        const selectBundle = document.getElementById('select-bundle-group');

        const headers = workspace.headers || [];

        function buildOptions(defaultVal) {
            let opts = '<option value="">-- None / Default --</option>';
            headers.forEach(h => {
                const selected = (h.original.toLowerCase() === defaultVal.toLowerCase() || h.sanitized.toLowerCase() === defaultVal.toLowerCase()) ? 'selected' : '';
                opts += `<option value="${h.original}" ${selected}>${h.original}</option>`;
            });
            return opts;
        }

        selectPrimary.innerHTML = buildOptions('Province');
        selectSecondary.innerHTML = buildOptions('Municipality');
        selectBundle.innerHTML = buildOptions('Barangay');
    }

    function startBatchStream() {
        if (eventSource) {
            eventSource.close();
        }

        const primary = document.getElementById('select-primary-group').value;
        const secondary = document.getElementById('select-secondary-group').value;
        const bundle = document.getElementById('select-bundle-group').value;
        const maxWorkers = document.getElementById('input-max-workers').value;
        const testLimit = document.getElementById('input-test-limit').value;

        const pulse = document.getElementById('execution-pulse');
        const statusText = document.getElementById('progress-status-text');
        const percentVal = document.getElementById('progress-percent-val');
        const fill = document.getElementById('progress-fill');

        const statCompleted = document.getElementById('stat-completed');
        const statTotal = document.getElementById('stat-total');
        const statFailed = document.getElementById('stat-failed');

        const btnRestart = document.getElementById('btn-restart-process');
        const btnFooterRestart = document.getElementById('btn-footer-restart-process');
        if (btnRestart) btnRestart.style.display = 'none';
        if (btnFooterRestart) btnFooterRestart.style.display = 'none';

        pulse.classList.add('active');
        statusText.textContent = "Processing generation batch...";
        fill.style.width = "0%";
        percentVal.textContent = "0%";

        statCompleted.textContent = "0";
        statTotal.textContent = "0";
        statFailed.textContent = "0";

        document.getElementById('log-container').innerHTML = '';
        appendLog('info', 'Connecting to PDF rendering engine stream...');

        let url = `/api/process/stream?max_workers=${maxWorkers}`;
        if (primary) url += `&primary_group=${encodeURIComponent(primary)}`;
        if (secondary) url += `&secondary_group=${encodeURIComponent(secondary)}`;
        if (bundle) url += `&bundle_group=${encodeURIComponent(bundle)}`;
        if (testLimit) url += `&test_limit=${testLimit}`;

        eventSource = new EventSource(url);

        eventSource.onmessage = (e) => {
            const data = JSON.parse(e.data);

            if (data.type === 'log') {
                appendLog(data.status || 'info', `[${data.timestamp || 'LOG'}] ${data.message}`);
            } else if (data.type === 'progress') {
                fill.style.width = `${data.percent}%`;
                percentVal.textContent = `${data.percent}%`;
                statCompleted.textContent = data.current;
                statTotal.textContent = data.total;
                statFailed.textContent = data.failed;
            } else if (data.type === 'complete') {
                pulse.classList.remove('active');
                statusText.textContent = "Batch Processing Complete!";
                fill.style.width = "100%";
                percentVal.textContent = "100%";
                appendLog('success', `Finished! Merged ${data.summary.total_merged} PDFs into output folder.`);
                showToast(`Batch completed successfully!`, 'success');
                
                // Show restart and jump buttons
                if (btnRestart) btnRestart.style.display = 'inline-flex';
                if (btnFooterRestart) btnFooterRestart.style.display = 'inline-flex';
                const btnJump = document.getElementById('btn-jump-to-viewer-mode');
                if (btnJump) btnJump.style.display = 'inline-flex';

                eventSource.close();
            } else if (data.type === 'error') {
                pulse.classList.remove('active');
                statusText.textContent = "Processing encountered error";
                appendLog('error', `ERROR: ${data.message}`);
                showToast(data.message, 'error');
                eventSource.close();
            }
        };

        eventSource.onerror = () => {
            pulse.classList.remove('active');
            statusText.textContent = "Stream connection closed";
            eventSource.close();
        };
    }

    function appendLog(status, msg) {
        const logContainer = document.getElementById('log-container');
        const line = document.createElement('div');
        line.className = `log-line ${status}`;
        line.textContent = msg;
        logContainer.appendChild(line);
        logContainer.scrollTop = logContainer.scrollHeight;
    }

    function openOutputFolder(folderRelPath = '') {
        fetch('/api/output/open', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ folder: folderRelPath || '' })
        })
        .then(res => res.json())
        .then(data => {
            if (data.error) showToast(data.error, 'error');
            else showToast(data.message, 'success');
        })
        .catch(err => showToast(`Failed to open folder: ${err}`, 'error'));
    }

    // =========================================================================
    // Output Folder Tree Explorer & PDF Viewer Engine
    // =========================================================================
    let currentOutputTree = null;
    let activeSelectedPdfRelPath = null;

    let searchDebounceTimer = null;

    function initOutputExplorer() {
        const btnRefresh = document.getElementById('btn-refresh-output-tree');
        const btnOpenFolder = document.getElementById('btn-open-explorer-folder');
        const searchInput = document.getElementById('input-search-output');
        const btnClearSearch = document.getElementById('btn-clear-search');

        if (btnRefresh) {
            btnRefresh.addEventListener('click', () => {
                showToast("Scanning output folder...", "info");
                loadOutputTree();
            });
        }

        if (btnOpenFolder) {
            btnOpenFolder.addEventListener('click', () => openOutputFolder());
        }

        if (searchInput) {
            searchInput.addEventListener('input', () => {
                const q = searchInput.value.trim();
                if (btnClearSearch) btnClearSearch.style.display = q ? 'block' : 'none';
                
                clearTimeout(searchDebounceTimer);
                if (!q) {
                    if (currentOutputTree && currentOutputTree.folders) {
                        const badge = document.getElementById('output-total-badge');
                        if (badge) badge.textContent = `${currentOutputTree.total_files} Files (${currentOutputTree.total_size_formatted})`;
                        renderOutputTree(currentOutputTree.folders, '');
                    }
                } else {
                    searchDebounceTimer = setTimeout(() => {
                        performDeepPdfSearch(q);
                    }, 280);
                }
            });
        }

        if (btnClearSearch) {
            btnClearSearch.addEventListener('click', () => {
                if (searchInput) searchInput.value = '';
                btnClearSearch.style.display = 'none';
                if (currentOutputTree && currentOutputTree.folders) {
                    const badge = document.getElementById('output-total-badge');
                    if (badge) badge.textContent = `${currentOutputTree.total_files} Files (${currentOutputTree.total_size_formatted})`;
                    renderOutputTree(currentOutputTree.folders, '');
                }
            });
        }
    }

    initOutputExplorer();

    function performDeepPdfSearch(query) {
        const container = document.getElementById('output-tree-container');
        const badge = document.getElementById('output-total-badge');
        if (badge) badge.textContent = `Searching inside PDFs...`;

        fetch(`/api/output/search?q=${encodeURIComponent(query)}`)
        .then(res => res.json())
        .then(data => {
            if (badge) {
                badge.textContent = `${data.total_matching_files} Bundles (${data.total_matches} text matches)`;
            }
            renderSearchResults(data, query);
        })
        .catch(err => {
            if (container) {
                container.innerHTML = `<div class="empty-tree-state"><i class="fa-solid fa-triangle-exclamation" style="color:var(--accent-red);"></i><p>Search error: ${err}</p></div>`;
            }
        });
    }

    function renderSearchResults(data, query) {
        const container = document.getElementById('output-tree-container');
        if (!container) return;

        const results = data.results || [];
        if (results.length === 0) {
            container.innerHTML = `
                <div class="empty-tree-state">
                    <i class="fa-solid fa-magnifying-glass"></i>
                    <p>No PDF files or farmer details matched "<strong>${query}</strong>".<br>Try searching for a different Farmer Name, ID, or Location.</p>
                </div>
            `;
            return;
        }

        // Group search results by folder
        const grouped = {};
        results.forEach(res => {
            const folder = res.folder_name;
            if (!grouped[folder]) {
                grouped[folder] = {
                    name: folder,
                    rel_path: res.folder_rel_path,
                    files: []
                };
            }
            grouped[folder].files.push(res);
        });

        container.innerHTML = '';
        const regex = new RegExp(`(${query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})`, 'gi');

        Object.values(grouped).forEach(folder => {
            const card = document.createElement('div');
            card.className = 'folder-group-card';

            const header = document.createElement('div');
            header.className = 'folder-header';
            header.innerHTML = `
                <div class="folder-title-left">
                    <i class="fa-solid fa-folder folder-icon"></i>
                    <span>${folder.name}</span>
                </div>
                <div style="display:flex; align-items:center; gap:0.5rem;">
                    <span class="folder-file-badge">${folder.files.length} Bundles</span>
                    <i class="fa-solid fa-chevron-down folder-toggle-arrow"></i>
                </div>
            `;

            const fileList = document.createElement('div');
            fileList.className = 'folder-file-list';

            folder.files.forEach(file => {
                const wrapper = document.createElement('div');
                wrapper.className = 'pdf-file-wrapper';

                const item = document.createElement('div');
                const isSelected = (file.rel_path === activeSelectedPdfRelPath);
                item.className = `pdf-file-item ${isSelected ? 'active' : ''}`;
                item.setAttribute('data-rel-path', file.rel_path);

                const matchBadge = file.match_count > 0 
                    ? `<span class="search-match-badge"><i class="fa-solid fa-bullseye"></i> ${file.match_count} match${file.match_count > 1 ? 'es' : ''}</span>`
                    : `<span class="search-match-badge" style="background:rgba(59,130,246,0.2); color:#60a5fa;">Name match</span>`;

                item.innerHTML = `
                    <div class="pdf-file-left">
                        <i class="fa-solid fa-file-pdf"></i>
                        <span>${file.file_name}</span>
                    </div>
                    <div style="display:flex; align-items:center; gap:0.4rem;">
                        ${matchBadge}
                        <span class="pdf-file-size">${file.size_formatted}</span>
                    </div>
                `;

                item.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const firstPage = (file.snippets && file.snippets.length > 0) ? file.snippets[0].page : 1;
                    viewPdfFile(file.rel_path, file.file_name, file.size_formatted, folder.name, folder.rel_path, firstPage);
                });

                wrapper.appendChild(item);

                // Render Snippets if any
                if (file.snippets && file.snippets.length > 0) {
                    const snippetsBox = document.createElement('div');
                    snippetsBox.className = 'search-snippets-box';

                    file.snippets.forEach(s => {
                        const pill = document.createElement('div');
                        pill.className = 'search-snippet-pill';
                        
                        const highlightedSnippet = s.snippet.replace(regex, '<mark>$1</mark>');
                        pill.innerHTML = `<strong>Page ${s.page}:</strong> "${highlightedSnippet}"`;

                        pill.addEventListener('click', (e) => {
                            e.stopPropagation();
                            viewPdfFile(file.rel_path, file.file_name, file.size_formatted, folder.name, folder.rel_path, s.page);
                        });

                        snippetsBox.appendChild(pill);
                    });

                    wrapper.appendChild(snippetsBox);
                }

                fileList.appendChild(wrapper);
            });

            header.addEventListener('click', () => {
                card.classList.toggle('collapsed');
            });

            card.appendChild(header);
            card.appendChild(fileList);
            container.appendChild(card);
        });
    }

    function loadOutputTree(autoSelectFirst = false) {
        const container = document.getElementById('output-tree-container');
        const badge = document.getElementById('output-total-badge');
        const searchInput = document.getElementById('input-search-output');
        const searchTerm = searchInput ? searchInput.value.trim() : '';

        if (searchTerm) {
            performDeepPdfSearch(searchTerm);
            return;
        }

        fetch('/api/output/tree')
        .then(res => res.json())
        .then(data => {
            currentOutputTree = data;
            if (badge) {
                badge.textContent = `${data.total_files} Files (${data.total_size_formatted})`;
            }

            renderOutputTree(data.folders || [], '');

            if (autoSelectFirst && data.folders && data.folders.length > 0) {
                for (const folder of data.folders) {
                    if (folder.files && folder.files.length > 0) {
                        const firstFile = folder.files[0];
                        viewPdfFile(firstFile.rel_path, firstFile.name, firstFile.size_formatted, folder.name, folder.rel_path, 1);
                        break;
                    }
                }
            }
        })
        .catch(err => {
            if (container) {
                container.innerHTML = `<div class="empty-tree-state"><i class="fa-solid fa-triangle-exclamation" style="color:var(--accent-red);"></i><p>Failed to load output files: ${err}</p></div>`;
            }
        });
    }

    function renderOutputTree(folders, searchTerm = '') {
        const container = document.getElementById('output-tree-container');
        if (!container) return;

        if (!folders || folders.length === 0) {
            container.innerHTML = `
                <div class="empty-tree-state">
                    <i class="fa-regular fa-folder-open"></i>
                    <p>No generated PDF files found in <code>output/</code> yet.<br>Execute batch processing in Form Generator to create PDF bundles.</p>
                </div>
            `;
            return;
        }

        const q = searchTerm.toLowerCase();
        let matchedTotal = 0;
        container.innerHTML = '';

        folders.forEach(folder => {
            const folderMatches = folder.name.toLowerCase().includes(q);
            const matchingFiles = folder.files.filter(f => folderMatches || f.name.toLowerCase().includes(q));

            if (matchingFiles.length === 0) return;
            matchedTotal += matchingFiles.length;

            const card = document.createElement('div');
            card.className = 'folder-group-card';

            const header = document.createElement('div');
            header.className = 'folder-header';
            header.innerHTML = `
                <div class="folder-title-left">
                    <i class="fa-solid fa-folder folder-icon"></i>
                    <span>${folder.name}</span>
                </div>
                <div style="display:flex; align-items:center; gap:0.5rem;">
                    <span class="folder-file-badge">${matchingFiles.length} Bundles</span>
                    <i class="fa-solid fa-chevron-down folder-toggle-arrow"></i>
                </div>
            `;

            const fileList = document.createElement('div');
            fileList.className = 'folder-file-list';

            matchingFiles.forEach(file => {
                const item = document.createElement('div');
                const isSelected = (file.rel_path === activeSelectedPdfRelPath);
                item.className = `pdf-file-item ${isSelected ? 'active' : ''}`;
                item.setAttribute('data-rel-path', file.rel_path);

                item.innerHTML = `
                    <div class="pdf-file-left">
                        <i class="fa-solid fa-file-pdf"></i>
                        <span>${file.name}</span>
                    </div>
                    <span class="pdf-file-size">${file.size_formatted}</span>
                `;

                item.addEventListener('click', (e) => {
                    e.stopPropagation();
                    viewPdfFile(file.rel_path, file.name, file.size_formatted, folder.name, folder.rel_path, 1);
                });

                fileList.appendChild(item);
            });

            header.addEventListener('click', () => {
                card.classList.toggle('collapsed');
            });

            card.appendChild(header);
            card.appendChild(fileList);
            container.appendChild(card);
        });

        if (matchedTotal === 0 && q) {
            container.innerHTML = `
                <div class="empty-tree-state">
                    <i class="fa-solid fa-magnifying-glass"></i>
                    <p>No PDF files matched "<strong>${searchTerm}</strong>".<br>Try searching for a different Barangay or Municipality name.</p>
                </div>
            `;
        }
    }

    // =========================================================================
    // In-Document Page Navigation & Viewer Logic
    // =========================================================================
    let currentDocRelPath = null;
    let currentDocFileName = '';
    let currentDocFileSize = '';
    let currentDocFolderName = '';
    let currentDocFolderRelPath = '';

    function jumpToDocPage(pageNum) {
        const metaEl = document.getElementById('viewer-file-meta');
        const iframe = document.getElementById('output-pdf-frame');

        if (metaEl) {
            metaEl.textContent = `Folder: ${currentDocFolderName} | Size: ${currentDocFileSize} (Page ${pageNum})`;
        }

        if (iframe && currentDocRelPath) {
            try {
                if (iframe.contentWindow && iframe.contentWindow.PDFViewerApplication && iframe.contentWindow.PDFViewerApplication.page !== undefined) {
                    iframe.contentWindow.PDFViewerApplication.page = pageNum;
                    return;
                }
            } catch (e) {
                // Fallback to URL navigation if direct property access fails
            }
            const fileUrl = `/api/output/view-file?path=${encodeURIComponent(currentDocRelPath)}`;
            iframe.src = `/static/pdfjs/web/viewer.html?file=${encodeURIComponent(fileUrl)}#page=${pageNum}&zoom=page-width`;
        }
    }

    function viewPdfFile(relPath, fileName, fileSize, folderName, folderRelPath, page = 1) {
        activeSelectedPdfRelPath = relPath;
        currentDocRelPath = relPath;
        currentDocFileName = fileName;
        currentDocFileSize = fileSize;
        currentDocFolderName = folderName;
        currentDocFolderRelPath = folderRelPath;

        const explorerStage = document.getElementById('output-explorer-stage');
        const viewerStage = document.getElementById('output-viewer-stage');
        if (explorerStage) explorerStage.style.display = 'none';
        if (viewerStage) viewerStage.style.display = 'flex';

        const nameEl = document.getElementById('viewer-file-name');
        const metaEl = document.getElementById('viewer-file-meta');
        const actionsEl = document.getElementById('viewer-actions');
        const btnDownload = document.getElementById('btn-download-pdf');
        const btnOpenFolder = document.getElementById('btn-open-file-folder');
        const btnToggleFind = document.getElementById('btn-toggle-pdf-find');
        const iframe = document.getElementById('output-pdf-frame');

        if (nameEl) nameEl.textContent = fileName;
        if (metaEl) metaEl.textContent = `Folder: ${folderName} | Size: ${fileSize}`;
        if (actionsEl) actionsEl.style.display = 'flex';

        if (btnDownload) {
            btnDownload.href = `/api/output/download-file?path=${encodeURIComponent(relPath)}`;
            btnDownload.download = fileName;
        }

        if (btnOpenFolder) {
            btnOpenFolder.onclick = (e) => {
                e.preventDefault();
                openOutputFolder(folderRelPath);
            };
        }

        const btnSavePdf = document.getElementById('btn-save-pdf-changes');
        if (btnSavePdf) {
            btnSavePdf.onclick = (e) => {
                e.preventDefault();
                savePdfChanges();
            };
        }

        if (btnToggleFind) {
            btnToggleFind.onclick = (e) => {
                e.preventDefault();
                if (iframe && iframe.contentWindow && iframe.contentWindow.PDFViewerApplication) {
                    const findBar = iframe.contentWindow.PDFViewerApplication.findBar;
                    if (findBar) {
                        if (findBar.opened) {
                            findBar.close();
                        } else {
                            findBar.open();
                        }
                    }
                }
            };
        }

        if (iframe) {
            iframe.style.display = 'block';
            const fileUrl = `/api/output/view-file?path=${encodeURIComponent(relPath)}`;
            const pageNum = (page && page > 1) ? page : 1;
            iframe.src = `/static/pdfjs/web/viewer.html?file=${encodeURIComponent(fileUrl)}#page=${pageNum}&zoom=page-width`;
        }

        window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    async function savePdfChanges() {
        const btnSave = document.getElementById('btn-save-pdf-changes');
        const iframe = document.getElementById('output-pdf-frame');
        const metaEl = document.getElementById('viewer-file-meta');

        if (!currentDocRelPath) {
            showToast("No active PDF selected to save.", "warning");
            return;
        }

        if (!iframe || !iframe.contentWindow || !iframe.contentWindow.PDFViewerApplication || !iframe.contentWindow.PDFViewerApplication.pdfDocument) {
            showToast("PDF document is not yet ready or loaded.", "warning");
            return;
        }

        const originalHtml = btnSave ? btnSave.innerHTML : '';
        if (btnSave) {
            btnSave.disabled = true;
            btnSave.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Saving...';
        }

        try {
            const pdfDoc = iframe.contentWindow.PDFViewerApplication.pdfDocument;
            const pdfBytes = await pdfDoc.saveDocument();
            const blob = new Blob([pdfBytes], { type: 'application/pdf' });

            const formData = new FormData();
            formData.append('path', currentDocRelPath);
            formData.append('file', blob, currentDocFileName);

            const res = await fetch('/api/output/save-pdf', {
                method: 'POST',
                body: formData
            });

            const data = await res.json();
            if (!res.ok) {
                throw new Error(data.error || 'Failed to save PDF');
            }

            if (data.size_formatted) {
                currentDocFileSize = data.size_formatted;
                if (metaEl) {
                    metaEl.textContent = `Folder: ${currentDocFolderName} | Size: ${currentDocFileSize}`;
                }
            }

            showToast(data.message || `Saved changes directly to ${currentDocFileName}!`, "success");

            if (btnSave) {
                btnSave.innerHTML = '<i class="fa-solid fa-check"></i> Saved!';
                setTimeout(() => {
                    btnSave.innerHTML = originalHtml;
                    btnSave.disabled = false;
                }, 2000);
            }
        } catch (err) {
            console.error("Error saving PDF changes:", err);
            showToast(`Error saving PDF: ${err.message}`, "error");
            if (btnSave) {
                btnSave.innerHTML = originalHtml;
                btnSave.disabled = false;
            }
        }
    }

    const btnBackToDir = document.getElementById('btn-back-to-directory');
    if (btnBackToDir) {
        btnBackToDir.addEventListener('click', () => {
            const explorerStage = document.getElementById('output-explorer-stage');
            const viewerStage = document.getElementById('output-viewer-stage');
            const iframe = document.getElementById('output-pdf-frame');
            if (viewerStage) viewerStage.style.display = 'none';
            if (explorerStage) explorerStage.style.display = 'block';
            if (iframe) iframe.src = 'about:blank';
            window.scrollTo({ top: 0, behavior: 'smooth' });
        });
    }
});
