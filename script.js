document.addEventListener('DOMContentLoaded', () => {
    // --- CONSTANTS & VERSION ---
    const VERSION = "1.7.5-GH-DEBUG Step";
    const MAIN_FONT = "Times New Roman";
    const SIZE_TITLE = 32, SIZE_LYRIC = 24, SIZE_CHORD = 14, SIZE_SECTION = 16, SIZE_COPY = 14;
    const PT_TO_PX = 96 / 72; 
    
    // --- APP STATE ---
    let currentPreviewIndex = 0;
    let currentShift = 0;
    let selectedBgPath = "assets/bg-default.png";
    
    const SCALE = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B'];
    const FLAT_MAP = { 'Db': 'C#', 'Eb': 'D#', 'Gb': 'F#', 'Ab': 'G#', 'Bb': 'A#' };

    // Sync UI Badges
    const vBadge = document.getElementById('vBadge');
    if (vBadge) vBadge.innerText = `v${VERSION}`;
    const lyricInput = document.getElementById('valLyrics');
    if (lyricInput) lyricInput.style.fontFamily = "'Consolas', 'Monaco', 'Courier New', monospace";

    // --- BACKGROUND GALLERY DATA ---
    const bgOptions = [
        { name: 'Default', path: 'assets/bg-default.png' },
        { name: 'Holy Spirit', path: 'assets/bg-HolySpirit.png' },
        { name: 'Linen', path: 'assets/bg-linen.png' },
        { name: 'Soft', path: 'assets/bg-soft.png' },
        { name: 'Dark', path: 'assets/bg-dark.png' },
        { name: 'Cloud', path: 'assets/bg-cloud.png' }
    ];

    const bgSelector = document.getElementById('bgSelector');
    bgOptions.forEach((opt, i) => {
        const thumb = document.createElement('div');
        thumb.className = `bg-thumb ${opt.path === selectedBgPath ? 'active' : ''}`;
        if(opt.path) thumb.style.backgroundImage = `url(${opt.path})`;
        thumb.onclick = () => {
            document.querySelectorAll('.bg-thumb').forEach(t => t.classList.remove('active'));
            thumb.classList.add('active');
            selectedBgPath = opt.path;
            updatePreview();
        };
        bgSelector.appendChild(thumb);
    });

    // --- LOGIC 1: DOWNLOAD POWERPOINT (WITH TROUBLESHOOTING) ---
    async function downloadPptx() {
        console.log("DEBUG: Download process started.");
        
        // 1. Check if Library exists
        if (typeof PptxGenJS === 'undefined') {
            alert("STOP: PptxGenJS library not loaded. Check your internet connection or CDN link in index.html.");
            return;
        }

        try {
            const pres = new PptxGenJS();
            pres.layout = 'LAYOUT_16x9';
            
            const songTitle = document.getElementById('valTitle').value.trim() || "Song_Output";
            const align = document.getElementById('slideAlign').value;
            const gapVal = parseInt(document.getElementById('chordGap').value);
            const spacingMult = 0.85 + (gapVal / 100);

            // 2. Metadata Check
            const rawLyrics = lyricInput.value.trim();
            if (!rawLyrics) {
                alert("STOP: No lyrics found in the editor.");
                return;
            }

            const sections = rawLyrics.split(/(?=\[)/).filter(s => s.trim());
            console.log(`DEBUG: Found ${sections.length} sections.`);

            // 3. Define Master Slide
            const masterName = 'LYRIC_MASTER_' + Date.now();
            try {
                pres.defineSlideMaster({
                    title: masterName,
                    background: { path: selectedBgPath },
                    objects: [
                        { placeholder: { type: 'title', name: 'title', x: "5%", y: document.getElementById('yTitle').value + "%", w: "90%", h: "10%", align: align, valign: 'top' } },
                        { placeholder: { type: 'body', name: 'body', x: "5%", y: document.getElementById('yLyrics').value + "%", w: "90%", h: "70%", align: align, valign: 'middle' } },
                        { placeholder: { type: 'footer', name: 'footer', x: "5%", y: document.getElementById('yCopy').value + "%", w: "90%", h: "10%", align: align, valign: 'bottom' } }
                    ]
                });
                console.log("DEBUG: Master slide defined.");
            } catch (e) {
                alert("STOP: Failed to define Master Slide. Error: " + e.message);
                return;
            }

            // 4. Create Slides
            sections.forEach((section, idx) => {
                let slide = pres.addSlide({ masterName: masterName });
                slide.addNotes(section);
                
                slide.addText(songTitle, { placeholder: 'title', fontSize: SIZE_TITLE, fontFace: MAIN_FONT, bold: true });

                const lines = section.replace(/^[\n\r]+|[\n\r]+$/g, '').split('\n');
                let textObjects = [];
                for (let i = 0; i < lines.length; i++) {
                    const line = lines[i];
                    if (line.trim().startsWith('[') && line.trim().endsWith(']')) {
                        textObjects.push({ text: line.trim() + "\n", options: { fontSize: SIZE_SECTION, bold: true } });
                    } else if (isChordLine(line)) {
                        textObjects.push(...createPptxGhostLine(line, lines[i+1] || "", align));
                    } else {
                        textObjects.push({ text: (line || " ") + "\n", options: { fontSize: SIZE_LYRIC } });
                    }
                }
                slide.addText(textObjects, { placeholder: 'body', fontFace: MAIN_FONT, lineSpacing: SIZE_LYRIC * spacingMult });
                slide.addText(document.getElementById('valCopy').value, { placeholder: 'footer', fontSize: SIZE_COPY, fontFace: MAIN_FONT, italic: true });
                console.log(`DEBUG: Slide ${idx + 1} added.`);
            });

            // 5. Final Write
            const safeFileName = songTitle.replace(/[/\\?%*:|"<>]/g, '-') + ".pptx";
            console.log("DEBUG: Attempting to write file: " + safeFileName);
            
            await pres.writeFile({ fileName: safeFileName });
            alert("SUCCESS: PowerPoint generated! Check your downloads.");

        } catch (err) {
            console.error("DEBUG FATAL ERROR:", err);
            alert("FATAL ERROR: The process crashed. \nMessage: " + err.message + "\nCheck Console (F12) for details.");
        }
    }

    // --- LOGIC 2: SMART TRANSPOSE IMPORT ---
    async function handleImport(event) {
        const file = event.target.files[0];
        if (!file) return;
        try {
            const zip = await JSZip.loadAsync(file);
            const parser = new DOMParser();
            let extractedTitle = "", extractedCopy = "", fullLyrics = "";

            const slideFiles = Object.keys(zip.files).filter(name => name.startsWith('ppt/slides/slide'));
            slideFiles.sort((a,b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));

            for (let slidePath of slideFiles) {
                const xmlString = await zip.file(slidePath).async("text");
                const xmlDoc = parser.parseFromString(xmlString, "application/xml");
                const shapes = xmlDoc.getElementsByTagNameNS("*", "sp");
                for (let sp of shapes) {
                    const ph = sp.getElementsByTagNameNS("*", "ph")[0];
                    const phType = ph ? ph.getAttribute("type") : null;
                    let shapeText = "";
                    const tTags = sp.getElementsByTagNameNS("*", "t");
                    for (let t of tTags) shapeText += t.textContent;
                    shapeText = shapeText.trim();
                    if (!shapeText) continue;
                    if ((phType === "title" || phType === "ctrTitle") && !extractedTitle) extractedTitle = shapeText;
                    if (phType === "ftr" && !extractedCopy) extractedCopy = shapeText;
                    if (!extractedCopy && (shapeText.includes("©") || shapeText.toLowerCase().includes("copyright"))) extractedCopy = shapeText;
                }
            }

            const notesFiles = Object.keys(zip.files).filter(name => name.startsWith('ppt/notesSlides/notesSlide'));
            notesFiles.sort((a, b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));
            for (let fileName of notesFiles) {
                const xmlString = await zip.file(fileName).async("text");
                const xmlDoc = parser.parseFromString(xmlString, "application/xml");
                const paragraphs = xmlDoc.getElementsByTagNameNS("*", "p");
                let slideNoteText = "";
                for (let p of paragraphs) {
                    let pText = "";
                    const children = p.getElementsByTagNameNS("*", "*");
                    for (let c of children) {
                        if (c.localName === "t") pText += c.textContent;
                        if (c.localName === "br") pText += "\n"; 
                    }
                    if (pText.trim().length > 0) slideNoteText += pText + "\n";
                    else if (pText === "") slideNoteText += "\n";
                }
                if (slideNoteText.trim()) fullLyrics += slideNoteText.trim() + "\n\n";
            }
            document.getElementById('valTitle').value = extractedTitle || "Imported Song";
            document.getElementById('valCopy').value = extractedCopy || "";
            lyricInput.value = fullLyrics.trim();
            updatePreview();
        } catch (err) { alert("Import Failed."); }
    }

    // --- LOGIC 3: TRANSPOSITION ---
    function transposeChord(chord, steps) {
        return chord.replace(/[A-G][b#]?/g, (match) => {
            let note = FLAT_MAP[match] || match;
            let index = SCALE.indexOf(note);
            if (index === -1) return match;
            let newIndex = (index + steps) % 12;
            while (newIndex < 0) newIndex += 12;
            return SCALE[newIndex];
        });
    }

    function processTranspose(steps) {
        lyricInput.value = lyricInput.value.split('\n').map(line => {
            return isChordLine(line) ? line.replace(/\S+/g, c => transposeChord(c, steps)) : line;
        }).join('\n');
        currentShift += steps;
        document.getElementById('keyShift').innerText = `Shift: ${currentShift}`;
        updatePreview();
    }

    // --- LOGIC 4: LIVE PREVIEW ---
    function updatePreview() {
        const mock = document.getElementById('slideMock');
        if (!mock) return;
        const lyricsRaw = lyricInput.value;
        const sections = lyricsRaw.split(/(?=\[)/).filter(s => s.trim());
        if (currentPreviewIndex >= sections.length) currentPreviewIndex = Math.max(0, sections.length - 1);
        document.getElementById('slideIndicator').innerText = `Slide ${sections.length > 0 ? currentPreviewIndex + 1 : 0} / ${sections.length}`;

        const ratio = mock.offsetWidth / 960; 
        const scale = ratio * PT_TO_PX;
        mock.style.backgroundImage = selectedBgPath ? `url(${selectedBgPath})` : 'none';

        const pt = document.getElementById('prevTitle'), pc = document.getElementById('prevCopy'), pl = document.getElementById('prevLyrics');
        const align = document.getElementById('slideAlign').value;

        [pt, pc, pl].forEach(el => { el.style.textAlign = align; el.style.fontFamily = MAIN_FONT; });

        pt.innerText = document.getElementById('valTitle').value;
        pt.style.top = document.getElementById('yTitle').value + "%";
        pt.style.fontSize = (SIZE_TITLE * scale) + "px"; 
        pt.style.fontWeight = "bold";

        pc.innerText = document.getElementById('valCopy').value;
        pc.style.top = document.getElementById('yCopy').value + "%";
        pc.style.fontSize = (SIZE_COPY * scale) + "px";
        pc.style.fontStyle = "italic";

        pl.style.top = document.getElementById('yLyrics').value + "%";
        pl.style.height = "70%"; pl.style.display = "flex"; pl.style.flexDirection = "column"; pl.style.justifyContent = "center";
        
        const activeSection = (sections[currentPreviewIndex] || "").replace(/^[\n\r]+|[\n\r]+$/g, '');
        pl.innerHTML = ""; 
        const innerContent = document.createElement('div');
        innerContent.style.width = "100%";

        const lines = activeSection.split('\n');
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            const nextLine = lines[i+1] || "";
            const lineDiv = document.createElement('div');
            lineDiv.style.whiteSpace = "pre";

            if (line.trim().startsWith('[') && line.trim().endsWith(']')) {
                lineDiv.style.fontSize = (SIZE_SECTION * scale) + "px"; lineDiv.style.fontWeight = "bold";
                lineDiv.style.marginBottom = (2 * scale) + "px"; lineDiv.innerText = line;
            } else if (isChordLine(line)) {
                lineDiv.style.fontSize = (SIZE_CHORD * scale) + "px"; lineDiv.style.lineHeight = "0.7"; 
                lineDiv.style.marginBottom = (parseInt(document.getElementById('chordGap').value) * scale) + "px";
                lineDiv.innerHTML = createHtmlGhostLine(line, nextLine, scale, align);
            } else {
                lineDiv.style.fontSize = (SIZE_LYRIC * scale) + "px"; lineDiv.style.lineHeight = "1";
                lineDiv.style.marginBottom = (5 * scale) + "px"; lineDiv.innerText = line || " "; 
            }
            innerContent.appendChild(lineDiv);
        }
        pl.appendChild(innerContent);
    }

    function isChordLine(str) {
        if (!str.trim() || (str.trim().startsWith('[') && str.trim().endsWith(']'))) return false;
        const test = str.replace(/[A-G]|[m|maj|min|dim|aug|sus|2|4|5|7|9]|#|b|\s|\/|v|i|\[|\]/gi, "");
        return test.length === 0;
    }

    function createHtmlGhostLine(chords, lyrics, scale, align) {
        let html = "";
        const targetLen = align === 'center' ? Math.max(chords.length, lyrics.length) : chords.length;
        for (let i = 0; i < targetLen; i++) {
            const c = chords[i] || " ", l = lyrics[i] || " ", char = l === " " ? "\u00A0" : l;
            html += `<span style="position:relative; display:inline-block; font-size:${SIZE_LYRIC * scale}px; color: transparent;">`;
            html += char; 
            if (c !== " ") html += `<span style="position:absolute; left:0; bottom:0; font-size:${SIZE_CHORD * scale}px; color:#64748b; visibility:visible; font-family: monospace;">${c}</span>`;
            html += `</span>`;
        }
        return html;
    }

    function createPptxGhostLine(chords, lyrics, align) {
        let result = [];
        const targetLen = align === 'center' ? Math.max(chords.length, lyrics.length) : chords.length;
        for (let i = 0; i < targetLen; i++) {
            const c = chords[i] || " ", l = lyrics[i] || " ";
            if (c !== " ") result.push({ text: c, options: { color: "808080", fontSize: SIZE_CHORD } });
            else result.push({ text: l === "" ? " " : l, options: { transparency: 100, fontSize: SIZE_LYRIC } });
        }
        result.push({ text: "\n" });
        return result;
    }

    const inputs = document.querySelectorAll('input, textarea, select, [type="range"]');
    inputs.forEach(input => input.addEventListener('input', updatePreview));
    document.getElementById('importPptx').addEventListener('change', handleImport);
    document.getElementById('btnUp').onclick = () => processTranspose(1);
    document.getElementById('btnDown').onclick = () => processTranspose(-1);
    document.getElementById('downloadBtn').onclick = downloadPptx;
    document.getElementById('nextSlide').onclick = () => { 
        const count = lyricInput.value.split(/(?=\[)/).filter(s => s.trim()).length;
        if (currentPreviewIndex < count - 1) { currentPreviewIndex++; updatePreview(); }
    };
    document.getElementById('prevSlide').onclick = () => { if (currentPreviewIndex > 0) { currentPreviewIndex--; updatePreview(); } };
    window.addEventListener('resize', updatePreview);
    updatePreview();
});
