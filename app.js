document.addEventListener('DOMContentLoaded', () => {
    console.log("CertStudio Initializing...");

    // 1. Initialize Fabric Canvas
    const canvas = new fabric.Canvas('certificate-canvas', {
        preserveObjectStacking: true,
        backgroundColor: '#fff'
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

    // --- Data Parsing ---
    async function parseRecords() {
        let records = [];
        const isPaste = document.querySelector('.tab-btn[data-tab="tab-text"]').classList.contains('active');
        if (isPaste) {
            const val = namesTextarea.value.trim();
            if (val) {
                val.split('\n').forEach(line => {
                    const parts = line.split('|').map(s => s.trim());
                    if (parts[0]) {
                        records.push({
                            name: parts[0],
                            email: parts[1] || '',
                            date: '',
                            id: ''
                        });
                    }
                });            }
        } else {
            const file = dataUpload.files[0];
            if (file) {
                await new Promise(res => Papa.parse(file, { header: true, skipEmptyLines: true, complete: r => {
                    r.data.forEach(row => {
                        const norm = {}; for (let k in row) norm[k.toLowerCase().trim()] = row[k];
                        records.push({
                            name: norm['name']||'', 
                            email: norm['email']||'',
                            date: norm['date']||'', 
                            id: norm['issuance number'] || norm['id'] || ''
                        });
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
            if (o.customFieldType === 'name') o.set('text', record.name || '[Name]');
            else if (o.customFieldType === 'date') o.set('text', record.date || '[Date]');
            else if (o.customFieldType === 'id') o.set('text', record.id || '[ID]');
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
        const obj = e.selected[0];
        if (obj && obj.type === 'textbox') {
            stylingControls.style.display = 'block';
            alignmentControls.style.display = 'block';
            btnDelete.disabled = false;
            fontSizeInput.value = obj.fontSize;
            textColorInput.value = obj.fill;
            textAlignSelect.value = obj.textAlign;
            fontFamilySelect.value = obj.fontFamily;
        }
    }
    
    function onObjectCleared() {
        stylingControls.style.display = 'none';
        alignmentControls.style.display = 'none';
        btnDelete.disabled = true;
    }
    
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
        const content = "Name,Email,Date,ID\nJohn Doe,john@example.com,2026-03-23,001\nJane Smith,jane@example.com,2026-03-23,002";
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
        // Mode A: PDF Overlay using pdf-lib (Preserves original PDF content)
        if (outputFormat === 'text' && templateIsPDF && originalPDFBytes) {
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
                if (cfg.customFieldType === 'name') displayText = record.name || "";
                else if (cfg.customFieldType === 'date') displayText = record.date || "";
                else if (cfg.customFieldType === 'id') displayText = record.id || "";

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
                        if (cfg.customFieldType === 'name') displayText = record.name || "";
                        else if (cfg.customFieldType === 'date') displayText = record.date || "";
                        else if (cfg.customFieldType === 'id') displayText = record.id || "";
                        const { text, ...cleanOptions } = cfg;
                        vCanvas.add(new fabric.Textbox(displayText, { ...cleanOptions }));
                    });
                    vCanvas.renderAll();
                    const { jsPDF } = window.jspdf;
                    const pdf = new jsPDF({ orientation: templateData.width > templateData.height ? 'l' : 'p', unit: 'px', format: [templateData.width, templateData.height], hotfixes: ["px_scaling"] });
                    if (outputFormat === 'image') {
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

        const objectsConfig = canvas.getObjects().map(o => ({
            originalText: o.text, left: o.left, top: o.top, width: o.width, fontSize: o.fontSize,
            fontFamily: o.fontFamily, fill: o.fill, textAlign: o.textAlign, originX: o.originX,
            originY: o.originY, customFieldType: o.customFieldType, scaleX: o.scaleX, scaleY: o.scaleY
        }));

        const zip = new JSZip();
        const batchSize = Math.min(Math.floor((navigator.hardwareConcurrency || 4) * 1.5), Math.floor((navigator.deviceMemory || 4) * 3), 20);
        let completed = 0;
        const dict = translations[currentLang];
        
        let mailMergeData = "Name,Email,Date,ID,Attachment_Filename\n";

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
                const safeName = record.name ? `"${record.name.replace(/"/g, '""')}"` : "";
                const safeEmail = record.email ? `"${record.email.replace(/"/g, '""')}"` : "";
                const safeDate = record.date ? `"${record.date.replace(/"/g, '""')}"` : "";
                const safeId = record.id ? `"${record.id.replace(/"/g, '""')}"` : "";
                mailMergeData += `${safeName},${safeEmail},${safeDate},${safeId},"${filename}"\n`;
            });
            completed += batch.length;
            progressFill.style.width = `${(completed / records.length) * 100}%`;
            progressText.innerText = `${completed} / ${records.length} ${dict.generating}`;
        }
        
        // Download mail merge file separately if there is at least one email provided
        if (records.some(r => r.email)) {
            const blob = new Blob([mailMergeData], { type: 'text/csv;charset=utf-8;' });
            saveAs(blob, "mail_merge.csv");
        }

        saveAs(await zip.generateAsync({type:"blob"}), "certificates.zip");
        btnGenerate.disabled = false; progressText.innerText = dict.complete;
        setTimeout(() => progressContainer.style.display = 'none', 3000);
    });
});