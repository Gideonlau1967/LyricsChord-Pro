document.addEventListener('DOMContentLoaded', () => {
    // --- CONSTANTS & VERSION ---
    const VERSION = "1.8.0-NO-BG-STABLE";
    const MAIN_FONT = "Times New Roman";
    const SIZE_TITLE = 32, SIZE_LYRIC = 24, SIZE_CHORD = 14, SIZE_SECTION = 16, SIZE_COPY = 14;
    const PT_TO_PX = 96 / 72; 
    
    // --- APP STATE ---
    let currentPreviewIndex = 0;
    let currentShift = 0;
    let selectedBgPath = "assets/bg-default.png";
    
    const SCALE = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B'];
    const FLAT_MAP = { 'Db': 'C#', 'Eb': 'D#', 'Gb': 'F#', 'Ab': 'G#', 'Bb': 'A#' };

    const vBadge = document.getElementById('vBadge');
    if (vBadge) vBadge.innerText = `v${VERSION}`;
    const lyricInput = document.getElementById('valLyrics');

    // Background selection still updates preview, but is ignored in PPT Download
    const bgSelector = document.getElementById('bgSelector');
    const bgOptions = [
        { name: 'Default', path: 'assets/bg-default.png' },
        { name: 'Holy Spirit', path: 'assets/bg-HolySpirit.png' },
        { name: 'Linen', path: 'assets/bg-linen.png' },
        { name: 'Soft', path: 'assets/bg-soft.png' },
        { name: 'Dark', path: 'assets/bg-dark.png' },
        { name: 'Cloud', path: 'assets/bg-cloud.png' }
    ];

    bgOptions.forEach((opt) => {
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

    // --- LOGIC 1: DOWNLOAD POWERPOINT (SIMPLIFIED & STABLE) ---
    async function downloadPptx() {
        console.log("DEBUG [1/7]: Initialization...");
        
        if (typeof PptxGenJS === 'undefined') {
            alert("DEBUG ERROR: Library PptxGenJS not loaded.");
            return;
        }

        try {
            const pres = new PptxGenJS();
            pres.layout = 'LAYOUT_16x9'; 
            
            const songTitle = document.getElementById('valTitle').value.trim() || "Song_Output";
            const align = document.getElementById('slideAlign').value;
            const gapVal = parseInt(document.getElementById('chordGap').value);
            const spacingMult = 0.85 + (gapVal / 100);

            // Calculation for Y positions (using inches: total height is 5.625)
            const getInchesY = (id) => (parseFloat(document.getElementById(id).value) / 100) * 5.625;

            const rawLyrics = lyricInput.value.trim();
            if (!rawLyrics) return alert("DEBUG ERROR: Editor is empty.");
            const sections = rawLyrics.split(/(?=\[)/).filter(s => s.trim());

            console.log("DEBUG [2/7]: Defining Master Layout...");
            
            // FIX: Removed 'background' and used the most primitive object syntax.
            // This avoids the 'name' undefined error in PptxGenJS 3.12
            pres.defineSlideMaster({
                title: 'SONG_LYRIC_MASTER',
                objects: [
                    { 'title':  { x: 0.5, y: getInchesY('yTitle'), w: 9.0, h: 0.7, align: align, valign: 'top' } },
                    { 'body':   { x: 0.5, y: getInchesY('yLyrics'), w: 9.0, h: 4.2, align: align, valign: 'middle' } },
                    { 'footer': { x: 0.5, y: getInchesY('yCopy'),  w: 9.0, h: 0.4, align: align, valign: 'bottom' } }
                ]
            });
            console.log("DEBUG [3/7]: Master Defined Successfully.");

            sections.forEach((section, idx) => {
                let slide = pres.addSlide({ masterName: 'SONG_LYRIC_MASTER' });
                
                // NO BACKGROUND APPLIED HERE (Per request)

                slide.addNotes(section);
                
                slide.addText(songTitle, { 
                    placeholder: 'title', 
                    fontSize: SIZE_TITLE, fontFace: MAIN_FONT, bold: true 
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
                    placeholder: 'body', 
                    fontFace: MAIN_FONT, 
                    lineSpacing: SIZE_LYRIC * spacingMult 
                });

                slide.addText(document.getElementById('valCopy').value, { 
                    placeholder: 'footer', 
                    fontSize: SIZE_COPY, fontFace: MAIN_FONT, italic: true 
                });
                
                console.log(`DEBUG: Slide ${idx + 1} built.`);
            });

            console.log("DEBUG [4/7]: Writing file...");
            const safeFileName = songTitle.replace(/[/\\?%*:|"<>]/g, '-') + ".pptx";
            await pres.writeFile({ fileName: safeFileName });
            
            console.log("DEBUG [5/7]: Success.");
            alert("SUCCESS: PowerPoint generated (Background disabled for stability).");

        } catch (err) {
            console.error("DEBUG FATAL:", err);
            alert("CRITICAL ERROR: " + err.message);
        }
    }

    // --- LOGIC 2: IMPORT ---
    async function handleImport(event) {
        const file = event.target.files[0];
        if (!file) return;
        try {
            const zip = await JSZip.loadAsync(file);
            const parser = new DOMParser();
            let extractedTitle = "", extractedCopy = "", fullLyrics = "";
            const slideFiles = Object.keys(zip.files).filter(name => name.startsWith('ppt/slides/slide')).sort((a,b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));

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
                    if ((phType === "title" || phType === "ctrTitle") && !extractedTitle) extractedTitle = shapeText.trim();
                    if (phType === "ftr" && !extractedCopy) extractedCopy = shapeText.trim();
                }
            }
            document.getElementById('valTitle').value = extractedTitle || "Imported Song";
            document.getElementById('valCopy').value = extractedCopy || "";
            updatePreview();
        } catch (err) { alert("Import error."); }
    }

    // --- LOGIC 3: TRANSPOSE ---
    function transposeChord(chord, steps) {
        return chord.replace(/[A-G][b#]?/g, (match) => {
            let note = FLAT_MAP[match] || match, index = SCALE.indexOf(note);
            if (index === -1) return match;
            let newIndex = (index + steps) % 12;
            while (newIndex < 0) newIndex += 12;
            return SCALE[newIndex];
        });
    }

    function processTranspose(steps) {
        lyricInput.value = lyricInput.value.split('\n').map(line => isChordLine(line) ? line.replace(/\S+/g, c => transposeChord(c, steps)) : line).join('\n');
        currentShift += steps;
        document.getElementById('keyShift').innerText = `Shift: ${currentShift}`;
        updatePreview();
    }

    // --- LOGIC 4: PREVIEW ---
    function updatePreview() {
        const mock = document.getElementById('slideMock');
        if (!mock) return;
        const lyricsRaw = lyricInput.value;
        const sections = lyricsRaw.split(/(?=\[)/).filter(s => s.trim());
        if (currentPreviewIndex >= sections.length) currentPreviewIndex = Math.max(0, sections.length - 1);
        document.getElementById('slideIndicator').innerText = `Slide ${sections.length > 0 ? currentPreviewIndex + 1 : 0} / ${sections.length}`;
        const scale = (mock.offsetWidth / 960) * PT_TO_PX;
        mock.style.backgroundImage = selectedBgPath ? `url(${selectedBgPath})` : 'none';
        const pt = document.getElementById('prevTitle'), pc = document.getElementById('prevCopy'), pl = document.getElementById('prevLyrics'), align = document.getElementById('slideAlign').value;
        [pt, pc, pl].forEach(el => { el.style.textAlign = align; el.style.fontFamily = MAIN_FONT; });
        pt.innerText = document.getElementById('valTitle').value;
        pt.style.top = document.getElementById('yTitle').value + "%";
        pt.style.fontSize = (SIZE_TITLE * scale) + "px"; pt.style.fontWeight = "bold";
        pc.innerText = document.getElementById('valCopy').value;
        pc.style.top = document.getElementById('yCopy').value + "%";
        pc.style.fontSize = (SIZE_COPY * scale) + "px"; pc.style.fontStyle = "italic";
        pl.style.top = document.getElementById('yLyrics').value + "%";
        pl.style.height = "70%"; pl.style.display = "flex"; pl.style.flexDirection = "column"; pl.style.justifyContent = "center";
        const activeSection = (sections[currentPreviewIndex] || "").replace(/^[\n\r]+|[\n\r]+$/g, '');
        pl.innerHTML = ""; const innerContent = document.createElement('div'); innerContent.style.width = "100%";
        const lines = activeSection.split('\n');
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i], nextLine = lines[i+1] || "", lineDiv = document.createElement('div'); lineDiv.style.whiteSpace = "pre";
            if (line.trim().startsWith('[') && line.trim().endsWith(']')) {
                lineDiv.style.fontSize = (SIZE_SECTION * scale) + "px"; lineDiv.style.fontWeight = "bold"; lineDiv.style.marginBottom = (2 * scale) + "px"; lineDiv.innerText = line;
            } else if (isChordLine(line)) {
                lineDiv.style.fontSize = (SIZE_CHORD * scale) + "px"; lineDiv.style.lineHeight = "0.7"; 
                lineDiv.style.marginBottom = (parseInt(document.getElementById('chordGap').value) * scale) + "px";
                lineDiv.innerHTML = createHtmlGhostLine(line, nextLine, scale, align);
            } else {
                lineDiv.style.fontSize = (SIZE_LYRIC * scale) + "px"; lineDiv.style.lineHeight = "1"; lineDiv.style.marginBottom = (5 * scale) + "px"; lineDiv.innerText = line || " "; 
            }
            innerContent.appendChild(lineDiv);
        }
        pl.appendChild(innerContent);
    }

    function isChordLine(str) {
        if (!str.trim() || (str.trim().startsWith('[') && str.trim().endsWith(']'))) return false;
        return str.replace(/[A-G]|[m|maj|min|dim|aug|sus|2|4|5|7|9]|#|b|\s|\/|v|i|\[|\]/gi, "").length === 0;
    }

    function createHtmlGhostLine(chords, lyrics, scale, align) {
        let html = "";
        const targetLen = align === 'center' ? Math.max(chords.length, lyrics.length) : chords.length;
        for (let i = 0; i < targetLen; i++) {
            const c = chords[i] || " ", l = lyrics[i] || " ", char = l === " " ? "\u00A0" : l;
            html += `<span style="position:relative; display:inline-block; font-size:${SIZE_LYRIC * scale}px; color: transparent;">${char}`; 
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
    document.getElementById('nextSlide').onclick = () => { const count = lyricInput.value.split(/(?=\[)/).filter(s => s.trim()).length; if (currentPreviewIndex < count - 1) { currentPreviewIndex++; updatePreview(); } };
    document.getElementById('prevSlide').onclick = () => { if (currentPreviewIndex > 0) { currentPreviewIndex--; updatePreview(); } };
    window.addEventListener('resize', updatePreview);
    updatePreview();
});
