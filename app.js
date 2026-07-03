document.addEventListener('DOMContentLoaded', () => {
    console.log("CertStudio Initializing...");

    // 1. Initialize Fabric Canvas
    const canvas = new fabric.Canvas('certificate-canvas', {
        preserveObjectStacking: true,
        backgroundColor: '#fff',
        targetFindTolerance: 15
    });
    window.canvas = canvas;
    
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
    const posXInput = document.getElementById('pos-x');
    const posYInput = document.getElementById('pos-y');
    const boxWInput = document.getElementById('box-w');
    const boxHInput = document.getElementById('box-h');
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
    let cachedSansFontBytes = null;
    let cachedSerifFontBytes = null;
    
    // Remember last textbox styles for inheriting in new textboxes
    let lastTextSettings = {
        fontSize: 60,
        fontFamily: 'Noto Sans TC',
        fill: '#000000',
        textAlign: 'center'
    };

    const SANS_FONT_URL = 'fonts/NotoSansTC-Regular.ttf';
    const SERIF_FONT_URL = 'fonts/NotoSerifTC-Regular.ttf';

    async function loadSansFont() {
        if (cachedSansFontBytes) return cachedSansFontBytes;
        progressContainer.style.display = 'block';
        progressFill.style.width = '20%';
        progressText.innerText = "Downloading Noto Sans Font (約 11MB, 僅需下載一次)...";
        try {
            const res = await fetch(SANS_FONT_URL);
            if (!res.ok) throw new Error("字型伺服器回應錯誤");
            const buffer = await res.arrayBuffer();
            cachedSansFontBytes = new Uint8Array(buffer);
            return cachedSansFontBytes;
        } catch (err) {
            throw new Error("無法下載 Noto Sans 字型，請檢查您的網路連線：" + err.message);
        }
    }

    async function loadSerifFont() {
        if (cachedSerifFontBytes) return cachedSerifFontBytes;
        progressContainer.style.display = 'block';
        progressFill.style.width = '20%';
        progressText.innerText = "Downloading Noto Serif Font (約 16MB, 僅需下載一次)...";
        try {
            const res = await fetch(SERIF_FONT_URL);
            if (!res.ok) throw new Error("字型伺服器回應錯誤");
            const buffer = await res.arrayBuffer();
            cachedSerifFontBytes = new Uint8Array(buffer);
            return cachedSerifFontBytes;
        } catch (err) {
            throw new Error("無法下載 Noto Serif 字型，請檢查您的網路連線：" + err.message);
        }
    }

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
                if (o.baseFontSize == null) o.baseFontSize = o.fontSize;
                applyAutoFit(o); // shrink this record's text to the box if it's too wide
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
            fontSize: lastTextSettings.fontSize,
            fontFamily: lastTextSettings.fontFamily,
            fill: lastTextSettings.fill,
            textAlign: lastTextSettings.textAlign,
            originX: 'center', originY: 'center', customFieldType: field
        });
        textObj.baseFontSize = textObj.fontSize;
        canvas.add(textObj); canvas.setActiveObject(textObj);
        updatePreview();
    }
    
    btnAddName.addEventListener('click', () => addPlaceholder('name', '[Name]'));
    btnAddDate.addEventListener('click', () => addPlaceholder('date', '[Date]'));
    btnAddId.addEventListener('click', () => addPlaceholder('id', '[ID]'));
    
    canvas.on('selection:created', onObjectSelected);
    canvas.on('selection:updated', onObjectSelected);
    canvas.on('selection:cleared', onObjectCleared);
    
    // --- Auto shrink-to-fit: keep long text on one line by shrinking the font to the box ---
    let autoFitEnabled = true;
    const _measCanvas = document.createElement('canvas');
    const _measCtx = _measCanvas.getContext('2d');
    function measureLongestLineWidth(text, fontSize, fontFamily) {
        _measCtx.font = `${fontSize}px "${fontFamily}", "Noto Sans TC", sans-serif`;
        let max = 0;
        String(text).split('\n').forEach(line => { const w = _measCtx.measureText(line).width; if (w > max) max = w; });
        return max;
    }
    function fittedFontSize(text, baseSize, boxWidth, fontFamily) {
        if (!text || boxWidth <= 0) return baseSize;
        const w = measureLongestLineWidth(text, baseSize, fontFamily);
        if (w <= boxWidth || w === 0) return baseSize;
        return Math.max(6, baseSize * (boxWidth / w));
    }
    // Set a textbox's rendered fontSize to fit its box width, driven by its untouched baseFontSize.
    function applyAutoFit(obj) {
        if (!obj || obj.type !== 'textbox') return;
        const base = obj.baseFontSize != null ? obj.baseFontSize : obj.fontSize;
        const boxW = obj.width * (obj.scaleX || 1);
        const size = autoFitEnabled ? fittedFontSize(obj.text, base, boxW, obj.fontFamily) : base;
        if (Math.abs(obj.fontSize - size) > 0.01) obj.set('fontSize', size);
    }
    // Resolve the text a field config shows for a record (mirrors the export/preview logic).
    function displayTextForRecord(cfg, record) {
        const ft = cfg.customFieldType;
        if (!ft) return cfg.originalText;
        if (ft === 'name') return record.name || '';
        if (ft === 'date') return record.date || '';
        if (ft === 'id') return record.id || '';
        return record[ft] || record[ft.toLowerCase()] || record[ft.toUpperCase()] || '';
    }

    // The two Noto faces are the only fonts actually embedded. Map any family name
    // (incl. legacy layouts using Times/Lora/Cinzel/etc.) to the face it renders as,
    // mirroring the serif detection used in the export paths.
    function effectiveFontFamily(ff) {
        const f = (ff || '').toLowerCase();
        if (ff === 'Noto Serif TC' || f.includes('serif') || f.includes('times') || f.includes('lora') || f.includes('cinzel')) {
            return 'Noto Serif TC';
        }
        return 'Noto Sans TC';
    }

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
                fontSizeInput.value = obj.baseFontSize != null ? obj.baseFontSize : obj.fontSize;
                fontFamilySelect.value = effectiveFontFamily(obj.fontFamily);
                posXInput.value = Math.round(obj.left);
                posYInput.value = Math.round(obj.top);
                boxWInput.value = Math.round(obj.getScaledWidth());
                boxHInput.value = Math.round(obj.getScaledHeight());
                textColorInput.value = obj.fill;
                textAlignSelect.value = obj.textAlign;

                // Store selected textbox styles
                lastTextSettings = {
                    fontSize: obj.fontSize || 60,
                    fontFamily: obj.fontFamily || 'Noto Sans TC',
                    fill: obj.fill || '#000000',
                    textAlign: obj.textAlign || 'center'
                };
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

    function exportLayout() {
        const textboxes = canvas.getObjects()
            .filter(o => o.type === 'textbox' || o.type === 'text')
            .map(o => ({
                left: o.left,
                top: o.top,
                width: o.width,
                height: o.height,
                scaleX: o.scaleX,
                scaleY: o.scaleY,
                fontSize: o.baseFontSize != null ? o.baseFontSize : o.fontSize,
                fontFamily: o.fontFamily,
                fill: o.fill,
                textAlign: o.textAlign,
                originalText: o.text,
                customFieldType: o.customFieldType,
                originX: o.originX,
                originY: o.originY
            }));
            
        const guides = canvas.getObjects()
            .filter(o => o.isGuideLine)
            .map(o => ({
                guideType: o.guideType,
                left: o.left,
                top: o.top
            }));
            
        const layoutData = {
            version: "1.0",
            textboxes,
            guides
        };
        
        const blob = new Blob([JSON.stringify(layoutData, null, 2)], { type: 'application/json' });
        saveAs(blob, `certificate_layout_${new Date().toISOString().slice(0,10)}.json`);
    }

    function importLayout(e) {
        const file = e.target.files[0];
        if (!file) return;
        
        const reader = new FileReader();
        reader.onload = function(evt) {
            try {
                const layoutData = JSON.parse(evt.target.result);
                if (!layoutData || !Array.isArray(layoutData.textboxes)) {
                    throw new Error("Invalid structure");
                }
                
                // 1. Clear existing textboxes and guides
                const toRemove = canvas.getObjects().filter(o => o.type === 'textbox' || o.type === 'text' || o.isGuideLine);
                toRemove.forEach(o => canvas.remove(o));
                
                // 2. Add textboxes
                layoutData.textboxes.forEach(cfg => {
                    const textbox = new fabric.Textbox(cfg.originalText || "[Text]", {
                        left: cfg.left,
                        top: cfg.top,
                        width: cfg.width,
                        height: cfg.height,
                        scaleX: cfg.scaleX || 1,
                        scaleY: cfg.scaleY || 1,
                        fontSize: cfg.fontSize,
                        fontFamily: cfg.fontFamily,
                        fill: cfg.fill,
                        textAlign: cfg.textAlign,
                        customFieldType: cfg.customFieldType,
                        originX: cfg.originX || 'center',
                        originY: cfg.originY || 'center',
                        cornerColor: '#4f46e5',
                        cornerStrokeColor: '#4f46e5',
                        borderColor: '#4f46e5',
                        cornerSize: 8,
                        transparentCorners: false,
                        padding: 5
                    });
                    textbox.baseFontSize = cfg.fontSize; // imported size is the base for auto-fit

                    // Enable/disable standard scaling/controls
                    textbox.setControlsVisibility({
                        mt: false,
                        mb: false,
                        ml: false,
                        mr: false,
                        mtr: true // rotation
                    });
                    
                    canvas.add(textbox);
                });
                
                // 3. Add guides
                if (Array.isArray(layoutData.guides)) {
                    layoutData.guides.forEach(g => {
                        if (g.guideType === 'vertical') {
                            const line = new fabric.Line([g.left, 0, g.left, canvas.height], {
                                stroke: '#a855f7',
                                strokeWidth: 1,
                                strokeDashArray: [5, 5],
                                selectable: true,
                                hasControls: false,
                                hasBorders: false,
                                isGuideLine: true,
                                guideType: 'vertical',
                                hoverCursor: 'ew-resize',
                                moveCursor: 'ew-resize',
                                padding: 15
                            });
                            canvas.add(line);
                        } else if (g.guideType === 'horizontal') {
                            const line = new fabric.Line([0, g.top, canvas.width, g.top], {
                                stroke: '#a855f7',
                                strokeWidth: 1,
                                strokeDashArray: [5, 5],
                                selectable: true,
                                hasControls: false,
                                hasBorders: false,
                                isGuideLine: true,
                                guideType: 'horizontal',
                                hoverCursor: 'ns-resize',
                                moveCursor: 'ns-resize',
                                padding: 15
                            });
                            canvas.add(line);
                        }
                    });
                }
                
                canvas.renderAll();
                updatePreview();
                
                // Clean up file input value so same file can be selected again
                e.target.value = '';
            } catch (err) {
                console.error("Error importing layout:", err);
                const dict = translations[currentLang];
                alert(dict.alert_invalid_layout || "Invalid layout configuration file.");
            }
        };
        reader.readAsText(file);
    }

    if (btnAddVGuide) btnAddVGuide.addEventListener('click', addVerticalGuide);
    if (btnAddHGuide) btnAddHGuide.addEventListener('click', addHorizontalGuide);
    if (btnClearGuides) btnClearGuides.addEventListener('click', clearGuides);

    // --- Layout Backup (Import/Export) ---
    const btnExportLayout = document.getElementById('export-layout-btn');
    const btnImportLayout = document.getElementById('import-layout-btn');
    const layoutFileInput = document.getElementById('layout-file-input');

    if (btnExportLayout) btnExportLayout.addEventListener('click', exportLayout);
    if (btnImportLayout) btnImportLayout.addEventListener('click', () => layoutFileInput.click());
    if (layoutFileInput) layoutFileInput.addEventListener('change', importLayout);

    // Magnetic snapping on drag
    canvas.on('object:moving', (e) => {
        const obj = e.target;
        if (!obj || obj.isGuideLine) return;

        // Calculate unsnapped coordinates using active drag pointer coordinates to prevent jittering
        let unsnappedLeft = obj.left;
        let unsnappedTop = obj.top;
        if (e.transform && e.transform.original && e.pointer) {
            const originalLeft = e.transform.original.left;
            const originalTop = e.transform.original.top;
            const startPointerX = e.transform.ex;
            const startPointerY = e.transform.ey;
            const currentPointerX = e.pointer.x;
            const currentPointerY = e.pointer.y;
            unsnappedLeft = originalLeft + (currentPointerX - startPointerX);
            unsnappedTop = originalTop + (currentPointerY - startPointerY);
        }

        const rect = obj.getBoundingRect();
        const dx = unsnappedLeft - obj.left;
        const dy = unsnappedTop - obj.top;

        const unsnappedRect = {
            left: rect.left + dx,
            top: rect.top + dy,
            width: rect.width,
            height: rect.height
        };

        const centerX = unsnappedRect.left + unsnappedRect.width / 2;
        const centerY = unsnappedRect.top + unsnappedRect.height / 2;
        const snapThreshold = 15;

        let targetLeft = unsnappedLeft;
        let targetTop = unsnappedTop;
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
                    targetLeft = unsnappedLeft + (guideX - centerX);
                    o.set({ stroke: '#10b981', strokeWidth: 2.5 });
                    snappedX = true;
                } else if (Math.abs(unsnappedRect.left - guideX) < snapThreshold) {
                    targetLeft = unsnappedLeft + (guideX - unsnappedRect.left);
                    o.set({ stroke: '#10b981', strokeWidth: 2.5 });
                    snappedX = true;
                } else if (Math.abs((unsnappedRect.left + unsnappedRect.width) - guideX) < snapThreshold) {
                    targetLeft = unsnappedLeft + (guideX - (unsnappedRect.left + unsnappedRect.width));
                    o.set({ stroke: '#10b981', strokeWidth: 2.5 });
                    snappedX = true;
                }
            } else if (o.guideType === 'horizontal') {
                if (snappedY) return;
                const guideY = o.top;

                // Snap center, top, or bottom edges
                if (Math.abs(centerY - guideY) < snapThreshold) {
                    targetTop = unsnappedTop + (guideY - centerY);
                    o.set({ stroke: '#10b981', strokeWidth: 2.5 });
                    snappedY = true;
                } else if (Math.abs(unsnappedRect.top - guideY) < snapThreshold) {
                    targetTop = unsnappedTop + (guideY - unsnappedRect.top);
                    o.set({ stroke: '#10b981', strokeWidth: 2.5 });
                    snappedY = true;
                } else if (Math.abs((unsnappedRect.top + unsnappedRect.height) - guideY) < snapThreshold) {
                    targetTop = unsnappedTop + (guideY - (unsnappedRect.top + unsnappedRect.height));
                    o.set({ stroke: '#10b981', strokeWidth: 2.5 });
                    snappedY = true;
                }
            }
        });

        obj.set({ left: targetLeft, top: targetTop });
        obj.setCoords();
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
        if (obj) {
            const val = parseInt(e.target.value) || 60;
            obj.baseFontSize = val;
            applyAutoFit(obj); // sets rendered fontSize (== val unless the current text overflows)
            if (obj.fontSize === val) obj.set('fontSize', val);
            canvas.renderAll();
            lastTextSettings.fontSize = val;
        }
    });
    // X/Y position inputs: same canvas coordinates as Export Layout (obj.left/top).
    // With originX/Y 'center' these are the textbox's centre point.
    function applyPositionInput(axis, raw) {
        const obj = canvas.getActiveObject();
        if (!obj || obj.type !== 'textbox') return;
        const val = parseFloat(raw);
        if (isNaN(val)) return;
        obj.set(axis, val);
        obj.setCoords();
        canvas.renderAll();
    }
    posXInput.addEventListener('input', (e) => applyPositionInput('left', e.target.value));
    posYInput.addEventListener('input', (e) => applyPositionInput('top', e.target.value));
    // Box width is editable (controls wrapping/centring); height is auto from the font size.
    boxWInput.addEventListener('input', (e) => {
        const obj = canvas.getActiveObject();
        if (!obj || obj.type !== 'textbox') return;
        const val = parseFloat(e.target.value);
        if (isNaN(val) || val < 10) return;
        obj.set('width', val / (obj.scaleX || 1));
        obj.setCoords();
        canvas.renderAll();
        boxHInput.value = Math.round(obj.getScaledHeight());
    });
    // Keep the numbers in sync while dragging / after any transform.
    function syncPositionInputs() {
        const obj = canvas.getActiveObject();
        if (obj && obj.type === 'textbox') {
            posXInput.value = Math.round(obj.left);
            posYInput.value = Math.round(obj.top);
            boxWInput.value = Math.round(obj.getScaledWidth());
            boxHInput.value = Math.round(obj.getScaledHeight());
        }
    }
    canvas.on('object:moving', syncPositionInputs);
    canvas.on('object:modified', syncPositionInputs);
    canvas.on('object:scaling', syncPositionInputs);

    // Arrow-key nudge for the selected textbox (Shift = 10px). Ignore when typing in a field.
    document.addEventListener('keydown', (e) => {
        if (!['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) return;
        const tag = (document.activeElement && document.activeElement.tagName) || '';
        if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return;
        const obj = canvas.getActiveObject();
        if (!obj || obj.isGuideLine) return;
        e.preventDefault();
        const step = e.shiftKey ? 10 : 1;
        if (e.key === 'ArrowUp') obj.set('top', obj.top - step);
        else if (e.key === 'ArrowDown') obj.set('top', obj.top + step);
        else if (e.key === 'ArrowLeft') obj.set('left', obj.left - step);
        else if (e.key === 'ArrowRight') obj.set('left', obj.left + step);
        obj.setCoords();
        canvas.renderAll();
        syncPositionInputs();
        if (typeof saveState === 'function') saveState();
    });
    textColorInput.addEventListener('input', (e) => {
        const obj = canvas.getActiveObject();
        if (obj) {
            obj.set('fill', e.target.value);
            canvas.renderAll();
            lastTextSettings.fill = e.target.value;
        }
    });
    textAlignSelect.addEventListener('change', (e) => {
        const obj = canvas.getActiveObject();
        if (obj) {
            obj.set('textAlign', e.target.value);
            canvas.renderAll();
            lastTextSettings.textAlign = e.target.value;
        }
    });
    fontFamilySelect.addEventListener('change', (e) => {
        const obj = canvas.getActiveObject();
        if (obj) {
            obj.set('fontFamily', e.target.value);
            canvas.renderAll();
            lastTextSettings.fontFamily = e.target.value;
        }
    });
    btnDelete.addEventListener('click', () => {
        canvas.getActiveObjects().forEach(obj => canvas.remove(obj));
        canvas.discardActiveObject().renderAll();
    });

    // --- Undo / Redo (snapshot the editable objects; background is kept separately) ---
    let history = [];
    let histPointer = -1;
    let isRestoring = false;
    const UNDO_PROPS = ['customFieldType', 'isGuideLine', 'guideType', 'originalText',
        'selectable', 'hasControls', 'hasBorders', 'lockMovementX', 'lockMovementY', 'hoverCursor', 'padding'];
    function snapshotState() {
        return JSON.stringify(canvas.getObjects().map(o => o.toObject(UNDO_PROPS)));
    }
    function saveState() {
        if (isRestoring) return;
        history = history.slice(0, histPointer + 1);
        history.push(snapshotState());
        if (history.length > 50) history.shift();
        histPointer = history.length - 1;
    }
    function restoreState(json) {
        isRestoring = true;
        canvas.discardActiveObject();
        canvas.getObjects().slice().forEach(o => canvas.remove(o));
        fabric.util.enlivenObjects(JSON.parse(json), (objs) => {
            objs.forEach(o => canvas.add(o));
            canvas.renderAll();
            isRestoring = false;
            onObjectCleared();
        });
    }
    function undo() { if (histPointer > 0) { histPointer--; restoreState(history[histPointer]); } }
    function redo() { if (histPointer < history.length - 1) { histPointer++; restoreState(history[histPointer]); } }
    canvas.on('object:added', saveState);
    canvas.on('object:modified', saveState);
    canvas.on('object:removed', saveState);
    saveState(); // initial (empty) baseline so the first add is undoable
    document.addEventListener('keydown', (e) => {
        if (!(e.metaKey || e.ctrlKey) || e.key.toLowerCase() !== 'z') return;
        const tag = (document.activeElement && document.activeElement.tagName) || '';
        if (tag === 'INPUT' || tag === 'TEXTAREA') return;
        e.preventDefault();
        if (e.shiftKey) redo(); else undo();
    });
    // Commit fine-tuning (arrow nudge / X-Y-W inputs) into history on release.
    [posXInput, posYInput, boxWInput].forEach(inp => inp.addEventListener('change', saveState));

    // Global auto shrink-to-fit toggle.
    const autoFitToggle = document.getElementById('auto-fit-toggle');
    if (autoFitToggle) autoFitToggle.addEventListener('change', (e) => {
        autoFitEnabled = e.target.checked;
        canvas.getObjects().forEach(o => { if (o.type === 'textbox') applyAutoFit(o); });
        canvas.renderAll();
    });

    document.querySelectorAll('.color-swatch').forEach(swatch => {
        swatch.addEventListener('click', () => {
            const color = swatch.dataset.color;
            const obj = canvas.getActiveObject();
            if (obj) {
                obj.set('fill', color);
                canvas.renderAll();
                if(textColorInput) textColorInput.value = color;
                lastTextSettings.fill = color;
            }
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
    async function generateCertificateBlob(record, templateData, objectsConfig, outputFormat, fonts) {
        // Mode A: PDF Overlay using pdf-lib (Preserves original PDF content)
        if (outputFormat === 'text' && templateIsPDF && originalPDFBytes) {
            const { PDFDocument, rgb, StandardFonts } = PDFLib;
            const pdfDoc = await PDFDocument.load(originalPDFBytes);
            
            if (window.fontkit && (fonts.sansBytes || fonts.serifBytes)) {
                pdfDoc.registerFontkit(window.fontkit);
            }

            const pages = pdfDoc.getPages();
            const firstPage = pages[0];
            const { width, height } = firstPage.getSize();
            // The template page's MediaBox may not start at (0,0). pdf-lib drawText uses
            // absolute user-space coords, so overlays must be offset by the MediaBox origin
            // or they land shifted (pdf.js/preview already accounts for it, hence the mismatch).
            const mediaBox = firstPage.getMediaBox();

            let scaleX = width / (templateData.width || 1);
            let scaleY = height / (templateData.height || 1);
            if (isNaN(scaleX) || !isFinite(scaleX)) scaleX = 1;
            if (isNaN(scaleY) || !isFinite(scaleY)) scaleY = 1;

            let sansFont = null;
            if (fonts.sansBytes) {
                sansFont = await pdfDoc.embedFont(fonts.sansBytes);
            }
            let serifFont = null;
            if (fonts.serifBytes) {
                serifFont = await pdfDoc.embedFont(fonts.serifBytes);
            }

            const helveticaFont = await pdfDoc.embedFont(StandardFonts.Helvetica);
            const timesFont = await pdfDoc.embedFont(StandardFonts.TimesRoman);
            const courierFont = await pdfDoc.embedFont(StandardFonts.Courier);

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
                let font;
                let isSerif = cfg.fontFamily === 'Noto Serif TC' || cfg.fontFamily.toLowerCase().includes('serif') || cfg.fontFamily.toLowerCase().includes('times') || cfg.fontFamily.toLowerCase().includes('lora') || cfg.fontFamily.toLowerCase().includes('cinzel');
                let isCourier = cfg.fontFamily.toLowerCase().includes('courier');

                if (isSerif) {
                    font = serifFont || timesFont;
                } else if (isCourier) {
                    font = courierFont;
                } else {
                    font = sansFont || helveticaFont;
                }

                const rawWidth = typeof cfg.width === 'number' && !isNaN(cfg.width) ? cfg.width : 100;
                const rawHeight = typeof cfg.height === 'number' && !isNaN(cfg.height) ? cfg.height : (cfg.fontSize || 20) * 1.2;
                const rawScaleX = typeof cfg.scaleX === 'number' && !isNaN(cfg.scaleX) ? cfg.scaleX : 1;
                const rawScaleY = typeof cfg.scaleY === 'number' && !isNaN(cfg.scaleY) ? cfg.scaleY : 1;

                const actualWidth = rawWidth * rawScaleX;
                const actualHeight = rawHeight * rawScaleY;

                let left = typeof cfg.left === 'number' && !isNaN(cfg.left) ? cfg.left : 0;
                if (cfg.originX === 'center') {
                    left = left - actualWidth / 2;
                } else if (cfg.originX === 'right') {
                    left = left - actualWidth;
                }

                let top = typeof cfg.top === 'number' && !isNaN(cfg.top) ? cfg.top : 0;
                if (cfg.originY === 'center') {
                    top = top - actualHeight / 2;
                } else if (cfg.originY === 'bottom') {
                    top = top - actualHeight;
                }

                const xMin = mediaBox.x + left * scaleX;
                const yMax = mediaBox.y + height - (top * scaleY);
                const pdfWidth = actualWidth * scaleX;
                const rawFontSize = typeof cfg.fontSize === 'number' && !isNaN(cfg.fontSize) ? cfg.fontSize : 20;
                let pdfFontSize = rawFontSize * rawScaleY * scaleY;

                const lines = displayText.split('\n');
                if (autoFitEnabled) { // shrink so the widest line fits the box (matches editor auto-fit)
                    let widest = 0;
                    for (const l of lines) { const w = font.widthOfTextAtSize(l || "", pdfFontSize); if (w > widest) widest = w; }
                    if (widest > pdfWidth && widest > 0) pdfFontSize = Math.max(6 * scaleY, pdfFontSize * (pdfWidth / widest));
                }
                const lineSpacing = pdfFontSize * 1.31; // fabric lineHeight(1.16) × _fontSizeMult(1.13)

                for (let j = 0; j < lines.length; j++) {
                    const lineText = lines[j];
                    if (!lineText && lines.length > 1) continue;

                    const lineWidth = font.widthOfTextAtSize(lineText || "", pdfFontSize);
                    
                    let xStart = xMin;
                    if (cfg.textAlign === 'center') {
                        xStart = xMin + (pdfWidth - lineWidth) / 2;
                    } else if (cfg.textAlign === 'right') {
                        xStart = xMin + pdfWidth - lineWidth;
                    }

                    // Fabric draws the first line's alphabetic baseline at (box top + ascent),
                    // where ascent ≈ 0.87×fontSize for Noto Sans/Serif. Match it exactly so the
                    // selectable text lands where the WYSIWYG canvas (image mode) shows it.
                    // Height-independent by design; verified within 1px via pdf.js round-trip.
                    const yBaseline = yMax - (pdfFontSize * 0.87) - (j * lineSpacing);

                    // pdf-lib mis-writes advance widths for these full Noto CJK fonts, so one
                    // drawText() spreads glyphs apart ("3 Jul 2026" -> "3  Jul 2 0 2 6").
                    // widthOfTextAtSize() is correct, so draw each character at a manually
                    // accumulated x. Fixes Latin spacing and keeps CJK correct (verified via
                    // pdf.js round-trip). Alignment above still uses the correct full-line width.
                    let charX = xStart;
                    for (const ch of (lineText || "")) {
                        firstPage.drawText(ch, {
                            x: charX,
                            y: yBaseline,
                            size: pdfFontSize,
                            font: font,
                            color: rgb(color.r, color.g, color.b),
                        });
                        charX += font.widthOfTextAtSize(ch, pdfFontSize);
                    }
                }
            }
            return { name: record.name, id: record.id, data: await pdfDoc.save() };
        } 
        
        // Mode B: jsPDF Fallback (For Images or Flattened output)
        const vCanvas = new fabric.StaticCanvas(null, { width: templateData.width, height: templateData.height });
        return new Promise((resolve, reject) => {
            fabric.util.loadImage(templateData.bg, async (imgElement) => {
                const fabricImg = new fabric.Image(imgElement, { originX: 'left', originY: 'top' });
                vCanvas.setBackgroundImage(fabricImg, async () => {
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
                    vCanvas.getObjects().forEach(o => { if (o.type === 'textbox') applyAutoFit(o); });
                    vCanvas.renderAll();

                    try {
                        const { jsPDF } = window.jspdf;
                        const pdf = new jsPDF({ orientation: templateData.width > templateData.height ? 'l' : 'p', unit: 'px', format: [templateData.width, templateData.height], hotfixes: ["px_scaling"] });
                        
                        if (outputFormat === 'image') {
                            pdf.addImage(vCanvas.toDataURL({ format: 'jpeg', quality: 0.92 }), 'JPEG', 0, 0, templateData.width, templateData.height);
                        } else {
                            pdf.addImage(templateData.bg, 'JPEG', 0, 0, templateData.width, templateData.height);
                            
                            if (fonts.base64Sans) {
                                pdf.addFileToVFS('NotoSansTC.ttf', fonts.base64Sans);
                                pdf.addFont('NotoSansTC.ttf', 'NotoSansTC', 'normal');
                            }
                            if (fonts.base64Serif) {
                                pdf.addFileToVFS('NotoSerifTC.ttf', fonts.base64Serif);
                                pdf.addFont('NotoSerifTC.ttf', 'NotoSerifTC', 'normal');
                            }

                            vCanvas.getObjects().forEach(o => {
                                if (o.type === 'textbox' || o.type === 'text') {
                                    let isSerif = o.fontFamily === 'Noto Serif TC' || o.fontFamily.toLowerCase().includes('serif') || o.fontFamily.toLowerCase().includes('times') || o.fontFamily.toLowerCase().includes('lora') || o.fontFamily.toLowerCase().includes('cinzel');
                                    let isCourier = o.fontFamily.toLowerCase().includes('courier');
                                    
                                    if (isSerif) {
                                        if (fonts.base64Serif) pdf.setFont('NotoSerifTC', 'normal');
                                        else pdf.setFont('times', 'normal');
                                    } else if (isCourier) {
                                        pdf.setFont('courier', 'normal');
                                    } else {
                                        if (fonts.base64Sans) pdf.setFont('NotoSansTC', 'normal');
                                        else pdf.setFont('helvetica', 'normal');
                                    }
                                    pdf.setFontSize(o.fontSize * 1.0).setTextColor(o.fill);
                                    
                                    let actualW = o.getScaledWidth();
                                    let actualH = o.getScaledHeight();
                                    
                                    let absLeft = o.left;
                                    if (o.originX === 'center') absLeft = o.left - actualW / 2;
                                    else if (o.originX === 'right') absLeft = o.left - actualW;
                                    
                                    let anchorX = absLeft;
                                    if (o.textAlign === 'center') anchorX = absLeft + actualW / 2;
                                    else if (o.textAlign === 'right') anchorX = absLeft + actualW;
                                    
                                    let absTop = o.top;
                                    if (o.originY === 'center') absTop = o.top - actualH / 2;
                                    else if (o.originY === 'bottom') absTop = o.top - actualH;
                                    let anchorY = absTop + actualH / 2;
                                    
                                    pdf.text(o.text, anchorX, anchorY, { align: o.textAlign, baseline: 'middle', maxWidth: actualW });
                                }
                            });
                        }
                        vCanvas.dispose();
                        resolve({ name: record.name, id: record.id, data: pdf.output('arraybuffer') });
                    } catch (err) {
                        vCanvas.dispose();
                        reject(err);
                    }
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
            .filter(o => !o.isGuideLine) 
            .map(o => ({
                originalText: o.text, left: o.left, top: o.top, width: o.width, height: o.height,
                fontSize: o.baseFontSize != null ? o.baseFontSize : o.fontSize,
                fontFamily: o.fontFamily, fill: o.fill, textAlign: o.textAlign, originX: o.originX,
                originY: o.originY, customFieldType: o.customFieldType, scaleX: o.scaleX, scaleY: o.scaleY
            }));

        // Warn (with examples) about records whose text is wider than its box; auto-fit will shrink them.
        if (autoFitEnabled) {
            const affected = new Set();
            records.forEach(rec => {
                objectsConfig.forEach(cfg => {
                    if (!cfg.customFieldType) return;
                    const text = displayTextForRecord(cfg, rec);
                    if (!text) return;
                    const boxW = (typeof cfg.width === 'number' ? cfg.width : 100) * (cfg.scaleX || 1);
                    if (measureLongestLineWidth(text, cfg.fontSize, cfg.fontFamily) > boxW) {
                        affected.add(rec.name || rec.id || text);
                    }
                });
            });
            if (affected.size) {
                const dict = translations[currentLang] || translations['en'];
                const tmpl = dict.overflow_confirm || translations['en'].overflow_confirm;
                const eg = [...affected].slice(0, 5).join('、') + (affected.size > 5 ? ' …' : '');
                const msg = tmpl.replace('{n}', affected.size).replace('{eg}', eg);
                if (!confirm(msg)) {
                    btnGenerate.disabled = false; progressContainer.style.display = 'none';
                    return;
                }
            }
        }

        const zip = new JSZip();
        const batchSize = Math.min(Math.floor((navigator.hardwareConcurrency || 4) * 1.5), Math.floor((navigator.deviceMemory || 4) * 3), 20);
        let completed = 0;
        const dict = translations[currentLang];

        let needSansFont = false;
        let needSerifFont = false;

        for (const cfg of objectsConfig) {
            let isSerif = cfg.fontFamily === 'Noto Serif TC' || cfg.fontFamily.toLowerCase().includes('serif') || cfg.fontFamily.toLowerCase().includes('times') || cfg.fontFamily.toLowerCase().includes('lora') || cfg.fontFamily.toLowerCase().includes('cinzel');
            let isCourier = cfg.fontFamily.toLowerCase().includes('courier');
            
            if (isSerif) {
                needSerifFont = true;
            } else if (!isCourier) {
                needSansFont = true;
            }
        }

        let sansFontBytes = null;
        let serifFontBytes = null;
        let base64SansFont = null;
        let base64SerifFont = null;
        try {
            if (needSansFont) {
                sansFontBytes = await loadSansFont();
                if (sansFontBytes) {
                    progressText.innerText = "Preparing Noto Sans encoding...";
                    base64SansFont = await new Promise((resolve) => {
                        const reader = new FileReader();
                        reader.onloadend = () => resolve(reader.result.split(',')[1]);
                        reader.readAsDataURL(new Blob([sansFontBytes]));
                    });
                }
            }
            if (needSerifFont) {
                serifFontBytes = await loadSerifFont();
                if (serifFontBytes) {
                    progressText.innerText = "Preparing Noto Serif encoding...";
                    base64SerifFont = await new Promise((resolve) => {
                        const reader = new FileReader();
                        reader.onloadend = () => resolve(reader.result.split(',')[1]);
                        reader.readAsDataURL(new Blob([serifFontBytes]));
                    });
                }
            }

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

            // High-Performance Path: PDF Overlay using pdf-lib (Preserves original PDF content)
            if ((outputFormat === 'text' || outputFormat === 'single-pdf') && templateIsPDF && originalPDFBytes) {
                const { PDFDocument, rgb, StandardFonts } = PDFLib;
                
                const mainPdfDoc = await PDFDocument.create();
                if (window.fontkit && (sansFontBytes || serifFontBytes)) {
                    mainPdfDoc.registerFontkit(window.fontkit);
                }

                let sansFont = null;
                if (sansFontBytes) {
                    sansFont = await mainPdfDoc.embedFont(sansFontBytes);
                }
                let serifFont = null;
                if (serifFontBytes) {
                    serifFont = await mainPdfDoc.embedFont(serifFontBytes);
                }
                
                const helveticaFont = await mainPdfDoc.embedFont(StandardFonts.Helvetica);
                const timesFont = await mainPdfDoc.embedFont(StandardFonts.TimesRoman);
                const courierFont = await mainPdfDoc.embedFont(StandardFonts.Courier);

                const srcDoc = await PDFDocument.load(originalPDFBytes);
                const templatePage = srcDoc.getPages()[0];
                const { width, height } = templatePage.getSize();
                // MediaBox origin may be non-zero; copied pages inherit it. Offset overlays by
                // it so pdf-lib drawText lands where pdf.js/preview renders the page content.
                const mediaBox = templatePage.getMediaBox();
                let scaleX = width / (templateImageWidth || 1);
                let scaleY = height / (templateImageHeight || 1);
                if (isNaN(scaleX) || !isFinite(scaleX)) scaleX = 1;
                if (isNaN(scaleY) || !isFinite(scaleY)) scaleY = 1;

                // Copy all template pages in ONE copyPages call so pdf-lib's object copier
                // shares the template's background image across every page. Copying once per
                // record duplicated the ~6MB background on each page (hundreds of MB for large
                // batches); batching keeps the whole combined file at roughly one template size.
                const copiedPages = await mainPdfDoc.copyPages(srcDoc, new Array(records.length).fill(0));
                copiedPages.forEach(p => mainPdfDoc.addPage(p));

                for (let i = 0; i < records.length; i++) {
                    const record = records[i];
                    const page = mainPdfDoc.getPage(i);

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

                        if (!displayText) continue;

                        const color = hexToRgb(cfg.fill);
                        let font;
                        let isSerif = cfg.fontFamily === 'Noto Serif TC' || cfg.fontFamily.toLowerCase().includes('serif') || cfg.fontFamily.toLowerCase().includes('times') || cfg.fontFamily.toLowerCase().includes('lora') || cfg.fontFamily.toLowerCase().includes('cinzel');
                        let isCourier = cfg.fontFamily.toLowerCase().includes('courier');

                        if (isSerif) {
                            font = serifFont || timesFont;
                        } else if (isCourier) {
                            font = courierFont;
                        } else {
                            font = sansFont || helveticaFont;
                        }

                        const rawWidth = typeof cfg.width === 'number' && !isNaN(cfg.width) ? cfg.width : 100;
                        const rawHeight = typeof cfg.height === 'number' && !isNaN(cfg.height) ? cfg.height : (cfg.fontSize || 20) * 1.2;
                        const rawScaleX = typeof cfg.scaleX === 'number' && !isNaN(cfg.scaleX) ? cfg.scaleX : 1;
                        const rawScaleY = typeof cfg.scaleY === 'number' && !isNaN(cfg.scaleY) ? cfg.scaleY : 1;

                        const actualWidth = rawWidth * rawScaleX;
                        const actualHeight = rawHeight * rawScaleY;

                        let left = typeof cfg.left === 'number' && !isNaN(cfg.left) ? cfg.left : 0;
                        if (cfg.originX === 'center') {
                            left = left - actualWidth / 2;
                        } else if (cfg.originX === 'right') {
                            left = left - actualWidth;
                        }

                        let top = typeof cfg.top === 'number' && !isNaN(cfg.top) ? cfg.top : 0;
                        if (cfg.originY === 'center') {
                            top = top - actualHeight / 2;
                        } else if (cfg.originY === 'bottom') {
                            top = top - actualHeight;
                        }

                        const xMin = mediaBox.x + left * scaleX;
                        const yMax = mediaBox.y + height - (top * scaleY);
                        const pdfWidth = actualWidth * scaleX;
                        const rawFontSize = typeof cfg.fontSize === 'number' && !isNaN(cfg.fontSize) ? cfg.fontSize : 20;
                        let pdfFontSize = rawFontSize * rawScaleY * scaleY;

                        const lines = displayText.split('\n');
                        if (autoFitEnabled) { // shrink so the widest line fits the box (matches editor auto-fit)
                            let widest = 0;
                            for (const l of lines) { const w = font.widthOfTextAtSize(l || "", pdfFontSize); if (w > widest) widest = w; }
                            if (widest > pdfWidth && widest > 0) pdfFontSize = Math.max(6 * scaleY, pdfFontSize * (pdfWidth / widest));
                        }
                        const lineSpacing = pdfFontSize * 1.31; // fabric lineHeight(1.16) × _fontSizeMult(1.13)

                        for (let j = 0; j < lines.length; j++) {
                            const lineText = lines[j];
                            if (!lineText && lines.length > 1) continue;

                            const lineWidth = font.widthOfTextAtSize(lineText || "", pdfFontSize);
                            
                            let xStart = xMin;
                            if (cfg.textAlign === 'center') {
                                xStart = xMin + (pdfWidth - lineWidth) / 2;
                            } else if (cfg.textAlign === 'right') {
                                xStart = xMin + pdfWidth - lineWidth;
                            }

                            // Fabric draws the first line's alphabetic baseline at (box top + ascent),
                            // where ascent ≈ 0.87×fontSize for Noto Sans/Serif. Match it exactly so the
                            // selectable text lands where the WYSIWYG canvas (image mode) shows it.
                            // Height-independent by design; verified within 1px via pdf.js round-trip.
                            const yBaseline = yMax - (pdfFontSize * 0.87) - (j * lineSpacing);

                            // pdf-lib mis-writes advance widths for these full Noto CJK fonts, so one
                            // drawText() spreads glyphs apart ("3 Jul 2026" -> "3  Jul 2 0 2 6").
                            // widthOfTextAtSize() is correct, so draw each character at a manually
                            // accumulated x. Fixes Latin spacing and keeps CJK correct (verified via
                            // pdf.js round-trip). Alignment above still uses the correct full-line width.
                            let charX = xStart;
                            for (const ch of (lineText || "")) {
                                page.drawText(ch, {
                                    x: charX,
                                    y: yBaseline,
                                    size: pdfFontSize,
                                    font: font,
                                    color: rgb(color.r, color.g, color.b),
                                });
                                charX += font.widthOfTextAtSize(ch, pdfFontSize);
                            }
                        }
                    }

                    completed++;
                    const renderProgress = (completed / records.length) * 50;
                    progressFill.style.width = `${renderProgress}%`;
                    progressText.innerText = `${completed} / ${records.length} ${dict.generating}`;
                }

                if (outputFormat === 'single-pdf') {
                    progressText.innerText = "Saving combined PDF...";
                    progressFill.style.width = '90%';
                    const allPdfBytes = await mainPdfDoc.save();
                    progressFill.style.width = '100%';
                    
                    const blob = new Blob([allPdfBytes], { type: "application/pdf" });
                    saveAs(blob, `${originalFileName}_all_certificates.pdf`);
                    btnGenerate.disabled = false; progressText.innerText = dict.complete;
                    setTimeout(() => progressContainer.style.display = 'none', 3000);
                    return;
                }

                // Compile and reload the document to trigger font subsetting and keep output size tiny
                progressText.innerText = "Subsetting fonts...";
                const compiledBytes = await mainPdfDoc.save();
                const splitDocSrc = await PDFDocument.load(compiledBytes);

                const savingText = currentLang.startsWith('zh') ? "正在儲存證書" : "Saving certificates";
                for (let i = 0; i < records.length; i++) {
                    const record = records[i];
                    const singleDoc = await PDFDocument.create();
                    const [pageCopy] = await singleDoc.copyPages(splitDocSrc, [i]);
                    singleDoc.addPage(pageCopy);
                    const singlePdfBytes = await singleDoc.save();

                    let namePart = (record.name || `student_${i + 1}`).replace(/[^a-z0-9\u4e00-\u9fa5]/gi, '_');
                    let idPart = record.id ? `_${String(record.id).replace(/[^a-z0-9]/gi, '_')}` : "";
                    const filename = `${originalFileName}${idPart}_${namePart}.pdf`;
                    zip.file(filename, singlePdfBytes);

                    const rowValues = finalHeaders.map(h => {
                        const val = record[h] || record[h.toLowerCase()] || record[h.toUpperCase()] || '';
                        return `"${val.replace(/"/g, '""')}"`;
                    });
                    rowValues.push(`"${filename}"`);
                    mailMergeData += rowValues.join(',') + '\n';

                    const saveProgress = 50 + (((i + 1) / records.length) * 50);
                    progressFill.style.width = `${saveProgress}%`;
                    progressText.innerText = `${savingText}: ${i + 1} / ${records.length}`;
                }
            } else if (outputFormat === 'single-pdf') {
                const { jsPDF } = window.jspdf;
                const pdf = new jsPDF({
                    orientation: templateImageWidth > templateImageHeight ? 'l' : 'p',
                    unit: 'px',
                    format: [templateImageWidth, templateImageHeight],
                    hotfixes: ["px_scaling"]
                });

                if (base64SansFont) {
                    pdf.addFileToVFS('NotoSansTC.ttf', base64SansFont);
                    pdf.addFont('NotoSansTC.ttf', 'NotoSansTC', 'normal');
                }
                if (base64SerifFont) {
                    pdf.addFileToVFS('NotoSerifTC.ttf', base64SerifFont);
                    pdf.addFont('NotoSerifTC.ttf', 'NotoSerifTC', 'normal');
                }

                const vCanvas = new fabric.Canvas(null, { width: templateImageWidth, height: templateImageHeight });
                
                for (let i = 0; i < records.length; i++) {
                    const record = records[i];
                    vCanvas.clear();
                    
                    await new Promise((resolveBg) => {
                        fabric.util.loadImage(bgDataURL, (img) => {
                            const fabricImg = new fabric.Image(img, { selectable: false, evented: false, originX: 'left', originY: 'top' });
                            vCanvas.setBackgroundImage(fabricImg, () => resolveBg());
                        });
                    });

                    objectsConfig.forEach(cfg => {
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
                        const { text, ...cleanOptions } = cfg;
                        vCanvas.add(new fabric.Textbox(displayText, { ...cleanOptions }));
                    });
                    vCanvas.getObjects().forEach(o => { if (o.type === 'textbox') applyAutoFit(o); });
                    vCanvas.renderAll();

                    if (i > 0) {
                        pdf.addPage([templateImageWidth, templateImageHeight], templateImageWidth > templateImageHeight ? 'l' : 'p');
                    }

                    pdf.addImage(bgDataURL, 'JPEG', 0, 0, templateImageWidth, templateImageHeight);

                    vCanvas.getObjects().forEach(o => {
                        if (o.type === 'textbox' || o.type === 'text') {
                            let isSerif = o.fontFamily === 'Noto Serif TC' || o.fontFamily.toLowerCase().includes('serif') || o.fontFamily.toLowerCase().includes('times') || o.fontFamily.toLowerCase().includes('lora') || o.fontFamily.toLowerCase().includes('cinzel');
                            let isCourier = o.fontFamily.toLowerCase().includes('courier');
                            
                            if (isSerif) {
                                if (base64SerifFont) pdf.setFont('NotoSerifTC', 'normal');
                                else pdf.setFont('times', 'normal');
                            } else if (isCourier) {
                                pdf.setFont('courier', 'normal');
                            } else {
                                if (base64SansFont) pdf.setFont('NotoSansTC', 'normal');
                                else pdf.setFont('helvetica', 'normal');
                            }
                            pdf.setFontSize(o.fontSize * 1.0).setTextColor(o.fill);
                            
                            let actualW = o.getScaledWidth();
                            let actualH = o.getScaledHeight();
                            
                            let absLeft = o.left;
                            if (o.originX === 'center') absLeft = o.left - actualW / 2;
                            else if (o.originX === 'right') absLeft = o.left - actualW;
                            
                            let anchorX = absLeft;
                            if (o.textAlign === 'center') anchorX = absLeft + actualW / 2;
                            else if (o.textAlign === 'right') anchorX = absLeft + actualW;
                            
                            let absTop = o.top;
                            if (o.originY === 'center') absTop = o.top - actualH / 2;
                            else if (o.originY === 'bottom') absTop = o.top - actualH;
                            let anchorY = absTop + actualH / 2;
                            
                            pdf.text(o.text, anchorX, anchorY, { align: o.textAlign, baseline: 'middle', maxWidth: actualW });
                        }
                    });

                    completed++;
                    progressFill.style.width = `${(completed / records.length) * 100}%`;
                    progressText.innerText = `${completed} / ${records.length} ${dict.generating}`;
                }
                vCanvas.dispose();

                progressText.innerText = "Saving combined PDF...";
                const allPdfBytes = pdf.output('arraybuffer');
                const blob = new Blob([allPdfBytes], { type: "application/pdf" });
                saveAs(blob, `${originalFileName}_all_certificates.pdf`);
                btnGenerate.disabled = false; progressText.innerText = dict.complete;
                setTimeout(() => progressContainer.style.display = 'none', 3000);
                return;
            } else {
                for (let i = 0; i < records.length; i += batchSize) {
                    const batch = records.slice(i, i + batchSize);
                    const results = await Promise.all(batch.map(r => generateCertificateBlob(r, {bg: bgDataURL, width: templateImageWidth, height: templateImageHeight}, objectsConfig, outputFormat, { sansBytes: sansFontBytes, serifBytes: serifFontBytes, base64Sans: base64SansFont, base64Serif: base64SerifFont })));
                    results.forEach((res, idx) => {
                        const record = batch[idx];
                        let namePart = (res.name || `student_${completed + idx + 1}`).replace(/[^a-z0-9\u4e00-\u9fa5]/gi, '_');
                        let idPart = res.id ? `_${String(res.id).replace(/[^a-z0-9]/gi, '_')}` : "";
                        const filename = `${originalFileName}${idPart}_${namePart}.pdf`;
                        
                        zip.file(filename, res.data);
                        
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
            }
            
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