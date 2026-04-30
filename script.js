document.addEventListener('DOMContentLoaded', () => {
    // --- CONSTANTS & VERSION ---
    const VERSION = "1.6.1-Debug";
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
        // Logic change: Check if path matches our default instead of just i === 0
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

    // --- LOGIC 2: SMART TRANSPOSE IMPORT (Slides & Notes) ---
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
                    const runs = sp.getElementsByTagNameNS("*", "r");
                    let shapeText = "";
                    let maxFontSize = 0, isBold = false, isItalic = false;
    
                    for (let r of runs) {
                        const t = r.getElementsByTagNameNS("*", "t")[0];
                        if (t) shapeText += t.textContent;
    
                        const rPr = r.getElementsByTagNameNS("*", "rPr")[0];
                        if (rPr) {
                            const sz = rPr.getAttribute("sz");
                            if (sz) maxFontSize = parseInt(sz) / 100;
                            if (rPr.getAttribute("b") === "1") isBold = true;
                            if (rPr.getAttribute("i") === "1") isItalic = true;
                        }
                    }
    
                    const clean = shapeText.trim();
                    if (clean) {
                        console.log(`Found Text: "${clean}" | Size: ${maxFontSize} | Bold: ${isBold} | Italic: ${isItalic}`);
                        
                        if (maxFontSize === 32 && isBold && !extractedTitle) extractedTitle = clean;
                        if (maxFontSize === 14 && isItalic && !extractedCopy) extractedCopy = clean;
                    }
                }
            }

            // 2. Scan Notes for Content
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

            // 3. Feed the UI
            document.getElementById('valTitle').value = extractedTitle || "Imported Song";
            document.getElementById('valCopy').value = extractedCopy || "";

            let consolidated = "";
            if (extractedTitle) consolidated += extractedTitle + "\n";
            if (extractedCopy) consolidated += extractedCopy + "\n\n";
            consolidated += fullLyrics.trim();
            lyricInput.value = consolidated;

            currentShift = 0;
            currentPreviewIndex = 0;
            document.getElementById('keyShift').innerText = "Shift: 0";
            updatePreview();
            alert("Transpose Import Successful!");

        } catch (err) {
            console.error(err);
            alert("Error reading PPTX placeholders.");
        }
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
        const lyricsRaw = lyricInput.value;
        const sections = lyricsRaw.split(/(?=\[)/).filter(s => s.trim());

        if (currentPreviewIndex >= sections.length) currentPreviewIndex = Math.max(0, sections.length - 1);
        document.getElementById('slideIndicator').innerText = `Slide ${sections.length > 0 ? currentPreviewIndex + 1 : 0} / ${sections.length}`;

        const ratio = mock.offsetWidth / 960; 
        const scale = ratio * PT_TO_PX;

        mock.style.backgroundImage = selectedBgPath ? `url(${selectedBgPath})` : 'none';

        const pt = document.getElementById('prevTitle');
        const pc = document.getElementById('prevCopy');
        const pl = document.getElementById('prevLyrics');
        const align = document.getElementById('slideAlign').value;

        [pt, pc, pl].forEach(el => {
            el.style.textAlign = align;
            el.style.fontFamily = MAIN_FONT;
        });

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
                lineDiv.style.fontSize = (SIZE_SECTION * scale) + "px";
                lineDiv.style.fontWeight = "bold";
                lineDiv.style.marginBottom = (2 * scale) + "px";
                lineDiv.innerText = line;
            } else if (isChordLine(line)) {
                lineDiv.style.fontSize = (SIZE_CHORD * scale) + "px";
                lineDiv.style.lineHeight = "0.7"; 
                lineDiv.style.marginBottom = (parseInt(document.getElementById('chordGap').value) * scale) + "px";
                lineDiv.innerHTML = createHtmlGhostLine(line, nextLine, scale, align);
            } else {
                lineDiv.style.fontSize = (SIZE_LYRIC * scale) + "px";
                lineDiv.style.lineHeight = "1";
                lineDiv.style.marginBottom = (5 * scale) + "px"; 
                lineDiv.innerText = line || " "; 
            }
            innerContent.appendChild(lineDiv);
        }
        pl.appendChild(innerContent);
    }

    // --- LOGIC 1: DOWNLOAD POWERPOINT ---
    async function downloadPptx() {
        const pres = new PptxGenJS();
        pres.layout = 'LAYOUT_16x9';
        
        const songTitle = document.getElementById('valTitle').value.trim() || "Song_Studio_Output";
        const align = document.getElementById('slideAlign').value;
        const gapVal = parseInt(document.getElementById('chordGap').value);
        const spacingMult = 0.85 + (gapVal / 100);

        const sections = lyricInput.value.split(/(?=\[)/).filter(s => s.trim());

        sections.forEach(section => {
            let slide = pres.addSlide();
            slide.background = selectedBgPath ? { path: selectedBgPath } : { fill: "FFFFFF" };
            slide.addNotes(section);

            slide.addText(songTitle, {
                x: "5%", y: document.getElementById('yTitle').value + "%", w: "90%",
                fontSize: SIZE_TITLE, fontFace: MAIN_FONT, bold: true, align: align, valign: 'top'
            });

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

            slide.addText(textObjects, {
                x: "5%", y: document.getElementById('yLyrics').value + "%", w: "90%", h: "70%",
                fontFace: MAIN_FONT, valign: 'middle', align: align, lineSpacing: SIZE_LYRIC * spacingMult
            });

            slide.addText(document.getElementById('valCopy').value, {
                x: "5%", y: document.getElementById('yCopy').value + "%", w: "90%",
                fontSize: SIZE_COPY, fontFace: MAIN_FONT, italic: true, align: align, valign: 'top'
            });
        });

        const safeFileName = songTitle.replace(/[/\\?%*:|"<>]/g, '-') + ".pptx";
        await pres.writeFile({ fileName: safeFileName });
    }

    // --- HELPERS ---
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

    // --- LISTENERS ---
    const inputs = document.querySelectorAll('input, textarea, select, [type="range"]');
    inputs.forEach(input => input.addEventListener('input', updatePreview));

    document.getElementById('importPptx').addEventListener('change', handleImport);
    document.getElementById('btnUp').onclick = () => processTranspose(1);
    document.getElementById('btnDown').onclick = () => processTranspose(-1);
    document.getElementById('downloadBtn').onclick = downloadPptx;
    
    document.getElementById('nextSlide').onclick = () => { 
        const sectionsCount = lyricInput.value.split(/(?=\[)/).filter(s => s.trim()).length;
        if (currentPreviewIndex < sectionsCount - 1) { currentPreviewIndex++; updatePreview(); }
    };
    document.getElementById('prevSlide').onclick = () => { 
        if (currentPreviewIndex > 0) { currentPreviewIndex--; updatePreview(); }
    };

    window.addEventListener('resize', updatePreview);
    updatePreview();
});
