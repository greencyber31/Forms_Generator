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

    function setupEventListeners() {
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
            if (btnRestart) btnRestart.style.display = 'none';
            if (btnFooterRestart) btnFooterRestart.style.display = 'none';

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
                let customVal = '';
                let selectedHeader = rawVal;

                if (rawVal.startsWith('STATIC:')) {
                    isCustom = true;
                    customVal = rawVal.substring(7);
                    selectedHeader = '__CUSTOM__';
                }

                const isMatched = Boolean(rawVal);

                row.className = `mapping-row-item ${isMatched ? 'matched' : 'unmapped'}`;
                row.setAttribute('data-tag', tag);

                let selectOptions = `<option value="">-- Ignore / Unmapped --</option>`;
                selectOptions += `<option value="__CUSTOM__" ${isCustom ? 'selected' : ''}>✍️ Custom Static Text...</option>`;
                if (headers.length > 0) {
                    selectOptions += `<optgroup label="Excel Headers">`;
                    headers.forEach(h => {
                        const sel = (!isCustom && h === selectedHeader) ? 'selected' : '';
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
                    </div>
                `;

                const chk = item.querySelector('.trans-col-checkbox');
                const ordInput = item.querySelector('.col-order-input');

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
        iframe.src = `/api/template/pdf-preview?type=${type}&t=${Date.now()}`;
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
                
                // Show restart buttons
                if (btnRestart) btnRestart.style.display = 'inline-flex';
                if (btnFooterRestart) btnFooterRestart.style.display = 'inline-flex';

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

    function openOutputFolder() {
        fetch('/api/output/open', { method: 'POST' })
        .then(res => res.json())
        .then(data => {
            if (data.error) showToast(data.error, 'error');
            else showToast(data.message, 'success');
        });
    }
});
