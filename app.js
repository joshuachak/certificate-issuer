document.addEventListener('DOMContentLoaded', () => {
    console.log("CertStudio Initializing...");

    // 1. Initialize Fabric Canvas
    const canvas = new fabric.Canvas('certificate-canvas', {
        preserveObjectStacking: true,
        backgroundColor: '#fff',
        targetFindTolerance: 15
    });
    
    // UI Elements
    const templateUpload = document.getElementById('template-upload');
    const tabBtns = document.querySelectorAll('.tab-btn');
    const tabContents = document.querySelectorAll('.tab-content');
    const dataUpload = document.getElementById('data-upload');
    const namesTextarea = document.getElementById('names-textarea');
    const downloadTemplateBtn = document.getElementById('download-template-btn');
    const btnAddName = document.getElementById('add-name-btn');
    const btnAddDate = document.getElementById('add-date-btn');
    const btnAddId = document.getElementById('add-id-btn');
    const btnDelete = document.getElementById('delete-selected-btn');
    const stylingControls = document.getElementById('styling-controls');
    const alignmentControls = document.getElementById('alignment-controls');
    const fontSizeInput = document.getElementById('font-size');
    const textColorInput = document.getElementById('text-color');
    const textAlignSelect = document.getElementById('text-align');
    const fontFamilySelect = document.getElementById('font-family');
    const btnGenerate = document.getElementById('generate-btn');
    const progressContainer = document.getElementById('progress-container');
    const progressFill = document.getElementById('progress-fill');
    const progressText = document.getElementById('progress-text');
    const editorPanel = document.querySelector('.editor-panel');
    const languageSelect = document.getElementById('language-select');
    const btnAddVGuide = document.getElementById('add-v-guide-btn');
    const btnAddHGuide = document.getElementById('add-h-guide-btn');
    const btnClearGuides = document.getElementById('clear-guides-btn');
    
    // Preview Elements
    const previewSection = document.getElementById('preview-section');
    const prevPreviewBtn = document.getElementById('prev-preview-btn');
    const nextPreviewBtn = document.getElementById('next-preview-btn');
    const previewIndexText = document.getElementById('preview-index');

    // State
    let templateImageWidth = 0;
    let templateImageHeight = 0;
    let templateLoaded = false;
    let templateIsPDF = false;
    let originalPDFBytes = null; // Store raw bytes for overlay
    let currentZoom = 1.0;
    let originalFileName = "certificate";
    let bgDataURL = null;
    let currentLang = localStorage.getItem('certstudio_lang') || 'en';
    let currentRecords = [];
    let currentPreviewIndex = 0;

    // Extend translations with custom field keys if not present
    const extraTranslations = {
        en: {
            custom_fields_title: "Custom Fields",
            custom_field_placeholder: "e.g. Score",
            add_custom_field_btn: "+ Add",
            remove_field_tooltip: "Remove Field",
            guides_title: "Alignment Guides",
            add_v_guide: "+ Vertical Guide",
            add_h_guide: "+ Horizontal Guide",
            clear_guides: "Clear All Guides"
        },
        "zh-TW": {
            custom_fields_title: "自訂欄位",
            custom_field_placeholder: "例如：成績",
            add_custom_field_btn: "+ 新增",
            remove_field_tooltip: "移除欄位",
            guides_title: "對齊輔助線",
            add_v_guide: "+ 垂直輔助線",
            add_h_guide: "+ 水平輔助線",
            clear_guides: "清除所有輔助線"
        },
        "zh-CN": {
            custom_fields_title: "自定义字段",
            custom_field_placeholder: "例如：成绩",
            add_custom_field_btn: "+ 新增",
            remove_field_tooltip: "移除字段",
            guides_title: "对齐辅助线",
            add_v_guide: "+ 垂直辅助线",
            add_h_guide: "+ 水平辅助线",
            clear_guides: "清除所有辅助线"
        }
    };
    
    if (typeof translations !== 'undefined') {
        for (let lang in translations) {
            const extra = extraTranslations[lang] || extraTranslations['en'];
            translations[lang] = { ...extra, ...translations[lang] };
        }
    }

    // --- i18n Logic ---
    function setLanguage(lang) {
        currentLang = lang;
        localStorage.setItem('certstudio_lang', lang);
        if (lang === 'ar') { document.documentElement.dir = 'rtl'; document.body.classList.add('rtl'); } 
        else { document.documentElement.dir = 'ltr'; document.body.classList.remove('rtl'); }
        const dict = translations[lang] || translations['en'];
        document.querySelectorAll('[data-i18n]').forEach(el => {
            const key = el.getAttribute('data-i18n');
            if (dict[key]) {
                const svg = el.querySelector('svg');
                if (svg) { el.innerHTML = ''; el.appendChild(svg); el.appendChild(document.createTextNode(' ' + dict[key])); } 
                else { el.textContent = dict[key]; }
            }
        });
        document.querySelectorAll('[data-i18n-placeholder]').forEach(el => {
            const key = el.getAttribute('data-i18n-placeholder');
            if (dict[key]) el.placeholder = dict[key].replace(/\\n/g, '\n');
        });
        languageSelect.value = lang;
    }
    languageSelect.addEventListener('change', (e) => setLanguage(e.target.value));
    setLanguage(currentLang);

    // State for custom fields
    let customFields = []; // Array of { key: string, label: string }
    let parsedHeaders = ['name', 'email', 'date', 'id'];

    const standardKeys = new Set(['name', 'email', 'date', 'id', 'issuance number', 'mail', '姓名', '日期', '學號', '編號', '郵件', 'attachment_filename']);

    // --- Custom Fields Logic ---
    const customFieldsButtonsContainer = document.getElementById('custom-fields-buttons');
    const customFieldInput = document.getElementById('custom-field-input');
    const btnAddCustomField = document.getElementById('add-custom-field-btn');

    function addCustomField(fieldName) {
        if (!fieldName) return;
        const normalizedKey = fieldName.trim().toLowerCase();
        if (normalizedKey === '') return;

        // Don't add standard keys or duplicate custom keys
        if (standardKeys.has(normalizedKey)) return;
        if (customFields.some(f => f.key === normalizedKey)) return;

        customFields.push({ key: normalizedKey, label: fieldName.trim() });

        // Render the button in UI
        const btnWrapper = document.createElement('div');
        btnWrapper.className = 'custom-field-btn-wrapper';
        btnWrapper.dataset.fieldKey = normalizedKey;
        
        const dict = translations[currentLang] || translations['en'];
        const tooltip = dict.remove_field_tooltip || "Remove Field";
        
        btnWrapper.innerHTML = `
            <button class="add-field-btn" style="flex: 1; margin: 0;">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 5v14M5 12h14"/></svg>
                <span>Add [${fieldName.trim()}]</span>
            </button>
            <button class="remove-custom-field-btn" title="${tooltip}">×</button>
        `;

        // Click event to place it on fabric canvas
        btnWrapper.querySelector('.add-field-btn').addEventListener('click', () => {
            addPlaceholder(normalizedKey, `[${fieldName.trim()}]`);
        });

        // Click event to remove the custom field button
        btnWrapper.querySelector('.remove-custom-field-btn').addEventListener('click', () => {
            customFields = customFields.filter(f => f.key !== normalizedKey);
            btnWrapper.remove();
        });

        customFieldsButtonsContainer.appendChild(btnWrapper);
    }

    if (btnAddCustomField) {
        btnAddCustomField.addEventListener('click', () => {
            const val = customFieldInput.value.trim();
            if (val) {
                addCustomField(val);
                customFieldInput.value = '';
            }
        });
    }

    if (customFieldInput) {
        customFieldInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                const val = customFieldInput.value.trim();
                if (val) {
                    addCustomField(val);
                    customFieldInput.value = '';
                }
            }
        });
    }

    // --- Data Parsing ---
    async function parseRecords() {
        let records = [];
        const isPaste = document.querySelector('.tab-btn[data-tab="tab-text"]').classList.contains('active');
        if (isPaste) {
            const val = namesTextarea.value.trim();
            if (val) {
                const lines = val.split('\n').map(l => l.trim()).filter(l => l.length > 0);
                if (lines.length > 0) {
                    let headers = ['name', 'email'];
                    let startIdx = 0;
                    const firstLineParts = lines[0].split('|').map(s => s.trim());
                    const firstLineLower = firstLineParts.map(s => s.toLowerCase());
                    const hasHeaderHeuristic = firstLineLower.some(p => 
                        ['name', '姓名', 'email', 'mail', '郵件', 'date', '日期', 'id', '編號', '學號'].some(k => p.includes(k))
                    );

                    if (hasHeaderHeuristic && firstLineParts.length > 1) {
                        headers = firstLineParts;
                        startIdx = 1;
                        parsedHeaders = headers;
                    } else {
                        parsedHeaders = ['name', 'email'];
                    }

                    // Dynamically register custom fields from quick paste headers
                    headers.forEach(h => {
                        const norm = h.toLowerCase().trim();
                        if (norm && !standardKeys.has(norm)) {
                            addCustomField(h);
                        }
                    });

                    for (let i = startIdx; i < lines.length; i++) {
                        const parts = lines[i].split('|').map(s => s.trim());
                        const record = {};
                        headers.forEach((h, idx) => {
                            const v = parts[idx] || '';
                            record[h] = v;
                            record[h.toLowerCase()] = v;
                        });
                        // Ensure standard keys
                        record.name = record.name || record['姓名'] || record['name'] || parts[0] || '';
                        record.email = record.email || record['email'] || record['mail'] || record['郵件'] || parts[1] || '';
                        record.date = record.date || record['date'] || record['日期'] || '';
                        record.id = record.id || record['id'] || record['編號'] || record['學號'] || '';
                        records.push(record);
                    }
                }
            }
        } else {
            const file = dataUpload.files[0];
            if (file) {
                await new Promise(res => Papa.parse(file, { header: true, skipEmptyLines: true, complete: r => {
                    if (r.meta && r.meta.fields) {
                        parsedHeaders = r.meta.fields.map(f => f.trim());
                        // Dynamically register custom fields from uploaded CSV headers
                        parsedHeaders.forEach(h => {
                            const norm = h.toLowerCase().trim();
                            if (norm && !standardKeys.has(norm)) {
                                addCustomField(h);
                            }
                        });
                    }
                    r.data.forEach(row => {
                        const record = {};
                        for (let k in row) {
                            const v = row[k] || '';
                            record[k.trim()] = v;
                            record[k.toLowerCase().trim()] = v;
                        }
                        // Ensure standard keys
                        record.name = record.name || record['姓名'] || record['name'] || '';
                        record.email = record.email || record['email'] || record['mail'] || record['郵件'] || '';
                        record.date = record.date || record['date'] || record['日期'] || '';
                        record.id = record.id || record['id'] || record['編號'] || record['學號'] || record['issuance number'] || '';
                        records.push(record);
                    });
                    res();
                }}));
            }
        }
        currentRecords = records;
        if (records.length > 0) {
            previewSection.style.display = 'block';
            currentPreviewIndex = 0;
            updatePreview();
        } else {
            previewSection.style.display = 'none';
        }
        return records;
    }

    namesTextarea.addEventListener('input', parseRecords);
    dataUpload.addEventListener('change', parseRecords);
    
    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            tabBtns.forEach(b => b.classList.remove('active'));
            tabContents.forEach(c => c.style.display = 'none');
            btn.classList.add('active');
            const target = document.getElementById(btn.dataset.tab);
            if (target) target.style.display = 'block';
            setTimeout(parseRecords, 10);
        });
    });

    // --- Preview Logic ---
    function updatePreview() {
        if (!currentRecords.length) return;
        const record = currentRecords[currentPreviewIndex];
        previewIndexText.innerText = `${currentPreviewIndex + 1} / ${currentRecords.length}`;
        canvas.getObjects().forEach(o => {
            const fieldType = o.customFieldType;
            if (fieldType) {
                if (fieldType === 'name') {
                    o.set('text', record.name || '[Name]');
                } else if (fieldType === 'date') {
                    o.set('text', record.date || '[Date]');
                } else if (fieldType === 'id') {
                    o.set('text', record.id || '[ID]');
                } else {
                    // Custom fields lookup (case-insensitive)
                    const val = record[fieldType] || record[fieldType.toLowerCase()] || record[fieldType.toUpperCase()] || `[${fieldType}]`;
                    o.set('text', val);
                }
            }
        });
        canvas.renderAll();
    }

    prevPreviewBtn.addEventListener('click', () => { if (currentPreviewIndex > 0) { currentPreviewIndex--; updatePreview(); } });
    nextPreviewBtn.addEventListener('click', () => { if (currentPreviewIndex < currentRecords.length - 1) { currentPreviewIndex++; updatePreview(); } });

    // --- Zoom Logic ---
    function updateZoom(newZoom) {
        currentZoom = Math.min(Math.max(newZoom, 0.1), 5.0);
        const wrapper = document.getElementById('canvas-wrapper');
        if (wrapper) {
            wrapper.style.transform = `scale(${currentZoom})`;
            if (templateLoaded) {
                editorPanel.style.minWidth = `${(templateImageWidth * currentZoom) + 200}px`;
                editorPanel.style.minHeight = `${(templateImageHeight * currentZoom) + 200}px`;
            }
        }
    }

    editorPanel.addEventListener('wheel', (e) => {
        if (!templateLoaded) return;
        if (e.ctrlKey || e.metaKey) { e.preventDefault(); updateZoom(currentZoom * Math.pow(1.1, -e.deltaY / 100)); }
    }, { passive: false });

    document.getElementById('zoom-in-btn')?.addEventListener('click', () => updateZoom(currentZoom * 1.2));
    document.getElementById('zoom-out-btn')?.addEventListener('click', () => updateZoom(currentZoom / 1.2));
    document.getElementById('zoom-reset-btn')?.addEventListener('click', () => {
        if (!templateLoaded) return;
        updateZoom(Math.min(1, (editorPanel.clientWidth - 100) / templateImageWidth));
    });

    // --- Editor Commands ---
    document.getElementById('center-h-btn')?.addEventListener('click', () => {
        const obj = canvas.getActiveObject();
        if (obj) { obj.set({ left: canvas.width / 2 }); obj.setCoords(); canvas.renderAll(); }
    });
    document.getElementById('center-v-btn')?.addEventListener('click', () => {
        const obj = canvas.getActiveObject();
        if (obj) { obj.set({ top: canvas.height / 2 }); obj.setCoords(); canvas.renderAll(); }
    });

    templateUpload.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        originalFileName = file.name.replace(/\.[^/.]+$/, "");
        const reader = new FileReader();
        if (file.type === 'application/pdf') {
            templateIsPDF = true;
            reader.onload = async function() {
                try {
                    // Clone the ArrayBuffer because pdfjsLib might transfer and detach the original buffer
                    originalPDFBytes = new Uint8Array(this.result.slice(0));
                    
                    const typedarray = new Uint8Array(this.result);
                    const pdf = await pdfjsLib.getDocument(typedarray).promise;
                    const page = await pdf.getPage(1);
                    const viewport = page.getViewport({ scale: 2.0 }); 
                    const tempCanvas = document.createElement('canvas');
                    const ctx = tempCanvas.getContext('2d');
                    tempCanvas.height = viewport.height; tempCanvas.width = viewport.width;
                    await page.render({canvasContext: ctx, viewport: viewport}).promise;
                    bgDataURL = tempCanvas.toDataURL('image/jpeg', 0.95);
                    setCanvasBackground(bgDataURL);
                } catch (err) { alert("PDF Error."); }
            };
            reader.readAsArrayBuffer(file);
        } else {
            templateIsPDF = false;
            originalPDFBytes = null;
            reader.onload = (f) => { bgDataURL = f.target.result; setCanvasBackground(bgDataURL); };
            reader.readAsDataURL(file);
        }
    });
    
    function setCanvasBackground(dataUrl) {
        fabric.util.loadImage(dataUrl, (img) => {
            if (!img) return;
            templateImageWidth = img.width; templateImageHeight = img.height;
            canvas.clear();
            canvas.setWidth(templateImageWidth); canvas.setHeight(templateImageHeight);
            const fabricImg = new fabric.Image(img, { selectable: false, evented: false, originX: 'left', originY: 'top' });
            canvas.setBackgroundImage(fabricImg, () => {
                templateLoaded = true; canvas.renderAll();
                updateZoom(Math.min(1, (editorPanel.clientWidth - 100) / templateImageWidth));
            }, { originX: 'left', originY: 'top' });
        });
    }
    
    function addPlaceholder(field, text) {
        if (!templateLoaded) return alert(translations[currentLang].alert_upload_template);
        const textObj = new fabric.Textbox(text, {
            left: canvas.width / 2, top: canvas.height / 2, width: canvas.width * 0.8,
            fontSize: 60, fontFamily: 'Times New Roman', fill: '#000000', textAlign: 'center',
            originX: 'center', originY: 'center', customFieldType: field 
        });
        canvas.add(textObj); canvas.setActiveObject(textObj);
        updatePreview();
    }
    
    btnAddName.addEventListener('click', () => addPlaceholder('name', '[Name]'));
    btnAddDate.addEventListener('click', () => addPlaceholder('date', '[Date]'));
    btnAddId.addEventListener('click', () => addPlaceholder('id', '[ID]'));
    
    canvas.on('selection:created', onObjectSelected);
    canvas.on('selection:updated', onObjectSelected);
    canvas.on('selection:cleared', onObjectCleared);
    
    function onObjectSelected(e) {
        // Reset all guide lines first to ensure clean state
        canvas.getObjects().forEach(o => {
            if (o.isGuideLine) {
                o.set({ stroke: '#8b5cf6', strokeWidth: 1.5, strokeDashArray: [5, 5] });
            }
        });

        const obj = e.selected[0];
        if (obj) {
            if (obj.type === 'textbox') {
                stylingControls.style.display = 'block';
                alignmentControls.style.display = 'block';
                btnDelete.disabled = false;
                fontSizeInput.value = obj.fontSize;
                textColorInput.value = obj.fill;
                textAlignSelect.value = obj.textAlign;
                fontFamilySelect.value = obj.fontFamily;
            } else if (obj.isGuideLine) {
                // For guide lines, hide styling but allow deletion
                stylingControls.style.display = 'none';
                alignmentControls.style.display = 'none';
                btnDelete.disabled = false;
                // Visual feedback for selected guide line: make solid and thicker
                obj.set({ stroke: '#4f46e5', strokeWidth: 2.5, strokeDashArray: null });
            }
        }
        canvas.renderAll();
    }
    
    function onObjectCleared() {
        stylingControls.style.display = 'none';
        alignmentControls.style.display = 'none';
        btnDelete.disabled = true;
        // Reset all guide lines to default dashed purple
        canvas.getObjects().forEach(o => {
            if (o.isGuideLine) {
                o.set({ stroke: '#8b5cf6', strokeWidth: 1.5, strokeDashArray: [5, 5] });
            }
        });
        canvas.renderAll();
    }

    // --- Alignment Guides & Snapping Logic ---
    function addVerticalGuide() {
        if (!templateLoaded) return alert(translations[currentLang].alert_upload_template);
        const line = new fabric.Line([canvas.width / 2, 0, canvas.width / 2, canvas.height], {
            stroke: '#8b5cf6',
            strokeWidth: 1.5,
            strokeDashArray: [5, 5],
            selectable: true,
            hasControls: false,
            hasBorders: false,
            isGuideLine: true,
            guideType: 'vertical',
            hoverCursor: 'ew-resize',
            lockMovementY: true,
            padding: 15 // Increase click target area by 15px on all sides!
        });
        canvas.add(line);
        canvas.setActiveObject(line);
        canvas.renderAll();
    }

    function addHorizontalGuide() {
        if (!templateLoaded) return alert(translations[currentLang].alert_upload_template);
        const line = new fabric.Line([0, canvas.height / 2, canvas.width, canvas.height / 2], {
            stroke: '#8b5cf6',
            strokeWidth: 1.5,
            strokeDashArray: [5, 5],
            selectable: true,
            hasControls: false,
            hasBorders: false,
            isGuideLine: true,
            guideType: 'horizontal',
            hoverCursor: 'ns-resize',
            lockMovementX: true,
            padding: 15 // Increase click target area by 15px on all sides!
        });
        canvas.add(line);
        canvas.setActiveObject(line);
        canvas.renderAll();
    }

    function clearGuides() {
        const guides = canvas.getObjects().filter(o => o.isGuideLine);
        guides.forEach(g => canvas.remove(g));
        canvas.discardActiveObject().renderAll();
    }

    if (btnAddVGuide) btnAddVGuide.addEventListener('click', addVerticalGuide);
    if (btnAddHGuide) btnAddHGuide.addEventListener('click', addHorizontalGuide);
    if (btnClearGuides) btnClearGuides.addEventListener('click', clearGuides);

    // Magnetic snapping on drag
    canvas.on('object:moving', (e) => {
        const obj = e.target;
        if (!obj || obj.isGuideLine) return;

        const rect = obj.getBoundingRect();
        const centerX = rect.left + rect.width / 2;
        const centerY = rect.top + rect.height / 2;
        const snapThreshold = 15;

        let snappedX = false;
        let snappedY = false;

        // Reset guide line styles
        canvas.getObjects().forEach(o => {
            if (o.isGuideLine) {
                o.set({ stroke: '#8b5cf6', strokeWidth: 1.5 });
            }
        });

        canvas.getObjects().forEach(o => {
            if (!o.isGuideLine) return;

            if (o.guideType === 'vertical') {
                if (snappedX) return;
                const guideX = o.left;
                
                // Snap center, left, or right edges
                if (Math.abs(centerX - guideX) < snapThreshold) {
                    obj.set({ left: obj.left + (guideX - centerX) });
                    o.set({ stroke: '#10b981', strokeWidth: 2.5 });
                    snappedX = true;
                } else if (Math.abs(rect.left - guideX) < snapThreshold) {
                    obj.set({ left: obj.left + (guideX - rect.left) });
                    o.set({ stroke: '#10b981', strokeWidth: 2.5 });
                    snappedX = true;
                } else if (Math.abs((rect.left + rect.width) - guideX) < snapThreshold) {
                    obj.set({ left: obj.left + (guideX - (rect.left + rect.width)) });
                    o.set({ stroke: '#10b981', strokeWidth: 2.5 });
                    snappedX = true;
                }
            } else if (o.guideType === 'horizontal') {
                if (snappedY) return;
                const guideY = o.top;

                // Snap center, top, or bottom edges
                if (Math.abs(centerY - guideY) < snapThreshold) {
                    obj.set({ top: obj.top + (guideY - centerY) });
                    o.set({ stroke: '#10b981', strokeWidth: 2.5 });
                    snappedY = true;
                } else if (Math.abs(rect.top - guideY) < snapThreshold) {
                    obj.set({ top: obj.top + (guideY - rect.top) });
                    o.set({ stroke: '#10b981', strokeWidth: 2.5 });
                    snappedY = true;
                } else if (Math.abs((rect.top + rect.height) - guideY) < snapThreshold) {
                    obj.set({ top: obj.top + (guideY - (rect.top + rect.height)) });
                    o.set({ stroke: '#10b981', strokeWidth: 2.5 });
                    snappedY = true;
                }
            }
        });

        canvas.renderAll();
    });

    // Reset styles on mouse up
    canvas.on('mouse:up', () => {
        const active = canvas.getActiveObject();
        canvas.getObjects().forEach(o => {
            if (o.isGuideLine) {
                if (active === o) {
                    // Keep currently selected guide solid and thicker
                    o.set({ stroke: '#4f46e5', strokeWidth: 2.5, strokeDashArray: null });
                } else {
                    // Reset other guides to thin dashed
                    o.set({ stroke: '#8b5cf6', strokeWidth: 1.5, strokeDashArray: [5, 5] });
                }
            }
        });
        canvas.renderAll();
    });
    
    fontSizeInput.addEventListener('input', (e) => {
        const obj = canvas.getActiveObject();
        if (obj) { obj.set('fontSize', parseInt(e.target.value)); canvas.renderAll(); }
    });
    textColorInput.addEventListener('input', (e) => {
        const obj = canvas.getActiveObject();
        if (obj) { obj.set('fill', e.target.value); canvas.renderAll(); }
    });
    textAlignSelect.addEventListener('change', (e) => {
        const obj = canvas.getActiveObject();
        if (obj) { obj.set('textAlign', e.target.value); canvas.renderAll(); }
    });
    fontFamilySelect.addEventListener('change', (e) => {
        const obj = canvas.getActiveObject();
        if (obj) { obj.set('fontFamily', e.target.value); canvas.renderAll(); }
    });
    btnDelete.addEventListener('click', () => {
        canvas.getActiveObjects().forEach(obj => canvas.remove(obj));
        canvas.discardActiveObject().renderAll();
    });

    document.querySelectorAll('.color-swatch').forEach(swatch => {
        swatch.addEventListener('click', () => {
            const color = swatch.dataset.color;
            const obj = canvas.getActiveObject();
            if (obj) { obj.set('fill', color); canvas.renderAll(); if(textColorInput) textColorInput.value = color; }
        });
    });

    downloadTemplateBtn.addEventListener('click', () => {
        const headers = ['Name', 'Email', 'Date', 'ID'];
        customFields.forEach(f => {
            headers.push(f.label);
        });
        const headerRow = headers.join(',');
        const row1 = ['John Doe', 'john@example.com', '2026-03-23', '001', ...customFields.map(() => 'Value1')].join(',');
        const row2 = ['Jane Smith', 'jane@example.com', '2026-03-23', '002', ...customFields.map(() => 'Value2')].join(',');
        const content = "\uFEFF" + `${headerRow}\n${row1}\n${row2}`;
        const blob = new Blob([content], { type: 'text/csv;charset=utf-8;' });
        saveAs(blob, "template.csv");
    });

    // --- Helper: Convert HEX to RGB for pdf-lib ---
    function hexToRgb(hex) {
        const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
        return result ? {
            r: parseInt(result[1], 16) / 255,
            g: parseInt(result[2], 16) / 255,
            b: parseInt(result[3], 16) / 255
        } : {r:0, g:0, b:0};
    }

    // --- Generation Logic ---
    async function generateCertificateBlob(record, templateData, objectsConfig, outputFormat) {
        // Check if there are non-ASCII characters in any fields to be rendered
        let hasUnicode = false;
        for (const cfg of objectsConfig) {
            let displayText = cfg.originalText;
            const fieldType = cfg.customFieldType;
            if (fieldType) {
                if (fieldType === 'name') displayText = record.name || "";
                else if (fieldType === 'date') displayText = record.date || "";
                else if (fieldType === 'id') displayText = record.id || "";
                else {
                    displayText = record[fieldType] || record[fieldType.toLowerCase()] || record[fieldType.toUpperCase()] || "";
                }
            }
            if (displayText && /[^\x00-\x7F]/.test(displayText)) {
                hasUnicode = true;
                break;
            }
        }

        let effectiveFormat = outputFormat;
        if (hasUnicode && outputFormat === 'text') {
            console.warn("Detected non-ASCII (Unicode) characters. Falling back to Image PDF format for compatibility.");
            effectiveFormat = 'image';
        }

        // Mode A: PDF Overlay using pdf-lib (Preserves original PDF content)
        if (effectiveFormat === 'text' && templateIsPDF && originalPDFBytes) {
            const { PDFDocument, rgb, StandardFonts } = PDFLib;
            const pdfDoc = await PDFDocument.load(originalPDFBytes);
            const pages = pdfDoc.getPages();
            const firstPage = pages[0];
            const { width, height } = firstPage.getSize();

            // Coordinate Mapping: Fabric (px) -> PDF Points
            // Note: Fabric 0,0 is Top-Left. PDF-Lib 0,0 is Bottom-Left.
            const scaleX = width / templateData.width;
            const scaleY = height / templateData.height;

            for (const cfg of objectsConfig) {
                let displayText = cfg.originalText;
                const fieldType = cfg.customFieldType;
                if (fieldType) {
                    if (fieldType === 'name') {
                        displayText = record.name || "";
                    } else if (fieldType === 'date') {
                        displayText = record.date || "";
                    } else if (fieldType === 'id') {
                        displayText = record.id || "";
                    } else {
                        displayText = record[fieldType] || record[fieldType.toLowerCase()] || record[fieldType.toUpperCase()] || "";
                    }
                }

                if (!displayText) continue;

                const color = hexToRgb(cfg.fill);
                let font = await pdfDoc.embedFont(StandardFonts.Helvetica);
                if (cfg.fontFamily.toLowerCase().includes('times')) font = await pdfDoc.embedFont(StandardFonts.TimesRoman);
                else if (cfg.fontFamily.toLowerCase().includes('courier')) font = await pdfDoc.embedFont(StandardFonts.Courier);

                // Calculate positions
                const pdfFontSize = cfg.fontSize * 1.0 * scaleY; 
                const textWidth = font.widthOfTextAtSize(displayText, pdfFontSize);
                
                let x = cfg.left * scaleX;
                let y = height - (cfg.top * scaleY); // Flip Y axis

                // Alignment logic
                if (cfg.textAlign === 'center') x -= textWidth / 2;
                else if (cfg.textAlign === 'right') x -= textWidth;

                firstPage.drawText(displayText, {
                    x: x,
                    y: y - (pdfFontSize / 4), // Middle baseline adjustment
                    size: pdfFontSize,
                    font: font,
                    color: rgb(color.r, color.g, color.b),
                });
            }
            return { name: record.name, id: record.id, data: await pdfDoc.save() };
        } 
        
        // Mode B: jsPDF Fallback (For Images or Flattened output)
        const vCanvas = new fabric.StaticCanvas(null, { width: templateData.width, height: templateData.height });
        return new Promise((resolve) => {
            fabric.util.loadImage(templateData.bg, (imgElement) => {
                const fabricImg = new fabric.Image(imgElement, { originX: 'left', originY: 'top' });
                vCanvas.setBackgroundImage(fabricImg, () => {
                    objectsConfig.forEach(cfg => {
                        let displayText = cfg.originalText;
                        const fieldType = cfg.customFieldType;
                        if (fieldType) {
                            if (fieldType === 'name') {
                                displayText = record.name || "";
                            } else if (fieldType === 'date') {
                                displayText = record.date || "";
                            } else if (fieldType === 'id') {
                                displayText = record.id || "";
                            } else {
                                displayText = record[fieldType] || record[fieldType.toLowerCase()] || record[fieldType.toUpperCase()] || "";
                            }
                        }
                        const { text, ...cleanOptions } = cfg;
                        vCanvas.add(new fabric.Textbox(displayText, { ...cleanOptions }));
                    });
                    vCanvas.renderAll();
                    const { jsPDF } = window.jspdf;
                    const pdf = new jsPDF({ orientation: templateData.width > templateData.height ? 'l' : 'p', unit: 'px', format: [templateData.width, templateData.height], hotfixes: ["px_scaling"] });
                    if (effectiveFormat === 'image') {
                        pdf.addImage(vCanvas.toDataURL({ format: 'jpeg', quality: 0.92 }), 'JPEG', 0, 0, templateData.width, templateData.height);
                    } else {
                        pdf.addImage(templateData.bg, 'JPEG', 0, 0, templateData.width, templateData.height);
                        vCanvas.getObjects().forEach(o => {
                            if (o.type === 'textbox' || o.type === 'text') {
                                let font = o.fontFamily.toLowerCase().includes('times') ? 'times' : (o.fontFamily.toLowerCase().includes('courier') ? 'courier' : 'helvetica');
                                pdf.setFont(font, 'normal').setFontSize(o.fontSize * 1.0).setTextColor(o.fill);
                                pdf.text(o.text, o.left, o.top, { align: o.textAlign, baseline: 'middle', maxWidth: o.getScaledWidth() });
                            }
                        });
                    }
                    vCanvas.dispose();
                    resolve({ name: record.name, id: record.id, data: pdf.output('arraybuffer') });
                });
            });
        });
    }

    if(btnGenerate) btnGenerate.addEventListener('click', async () => {
        const records = await parseRecords();
        if (!templateLoaded) return alert(translations[currentLang].alert_upload_template);
        if (!records.length) return alert(translations[currentLang].alert_provide_data);

        const outputFormat = document.querySelector('input[name="output-format"]:checked').value;
        btnGenerate.disabled = true; progressContainer.style.display = 'block';

        const objectsConfig = canvas.getObjects()
            .filter(o => !o.isGuideLine) // Filter out helper guide lines!
            .map(o => ({
                originalText: o.text, left: o.left, top: o.top, width: o.width, fontSize: o.fontSize,
                fontFamily: o.fontFamily, fill: o.fill, textAlign: o.textAlign, originX: o.originX,
                originY: o.originY, customFieldType: o.customFieldType, scaleX: o.scaleX, scaleY: o.scaleY
            }));

        const zip = new JSZip();
        const batchSize = Math.min(Math.floor((navigator.hardwareConcurrency || 4) * 1.5), Math.floor((navigator.deviceMemory || 4) * 3), 20);
        let completed = 0;
        const dict = translations[currentLang];
        
        const uniqueHeaders = [];
        parsedHeaders.forEach(h => {
            const trimmed = h.trim();
            if (trimmed && !uniqueHeaders.some(uh => uh.toLowerCase() === trimmed.toLowerCase())) {
                uniqueHeaders.push(trimmed);
            }
        });
        if (!uniqueHeaders.some(uh => uh.toLowerCase() === 'name')) uniqueHeaders.unshift('Name');
        if (!uniqueHeaders.some(uh => uh.toLowerCase() === 'email')) uniqueHeaders.push('Email');
        if (!uniqueHeaders.some(uh => uh.toLowerCase() === 'date')) uniqueHeaders.push('Date');
        if (!uniqueHeaders.some(uh => uh.toLowerCase() === 'id')) uniqueHeaders.push('ID');

        const finalHeaders = uniqueHeaders.filter(h => h && h.trim() !== '');
        let mailMergeData = finalHeaders.map(h => `"${h.replace(/"/g, '""')}"`).join(',') + ',"Attachment_Filename"\n';

        try {
            for (let i = 0; i < records.length; i += batchSize) {
                const batch = records.slice(i, i + batchSize);
                const results = await Promise.all(batch.map(r => generateCertificateBlob(r, {bg: bgDataURL, width: templateImageWidth, height: templateImageHeight}, objectsConfig, outputFormat)));
                results.forEach((res, idx) => {
                    const record = batch[idx];
                    let namePart = (res.name || `student_${completed + idx + 1}`).replace(/[^a-z0-9\u4e00-\u9fa5]/gi, '_');
                    let idPart = res.id ? `_${String(res.id).replace(/[^a-z0-9]/gi, '_')}` : "";
                    const filename = `${originalFileName}${idPart}_${namePart}.pdf`;
                    
                    zip.file(filename, res.data);
                    
                    // Add to mail merge CSV
                    const rowValues = finalHeaders.map(h => {
                        const val = record[h] || record[h.toLowerCase()] || record[h.toUpperCase()] || '';
                        return `"${val.replace(/"/g, '""')}"`;
                    });
                    rowValues.push(`"${filename}"`);
                    mailMergeData += rowValues.join(',') + '\n';
                });
                completed += batch.length;
                progressFill.style.width = `${(completed / records.length) * 100}%`;
                progressText.innerText = `${completed} / ${records.length} ${dict.generating}`;
            }
            
            // Download mail merge file separately if there is at least one email provided
            if (records.some(r => r.email)) {
                const blob = new Blob(["\uFEFF" + mailMergeData], { type: 'text/csv;charset=utf-8;' });
                saveAs(blob, "mail_merge.csv");
            }

            saveAs(await zip.generateAsync({type:"blob"}), "certificates.zip");
            btnGenerate.disabled = false; progressText.innerText = dict.complete;
            setTimeout(() => progressContainer.style.display = 'none', 3000);
        } catch (err) {
            console.error("Generation failed: ", err);
            alert("Export failed: " + err.message);
            btnGenerate.disabled = false;
            progressContainer.style.display = 'none';
        }
    });
});