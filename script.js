document.addEventListener('DOMContentLoaded', () => {
    const VERSION = "1.8.0-PRO";
    const MAIN_FONT = "Times New Roman";
    
    // Default Font Sizes (Points)
    const SIZE_TITLE = 32, SIZE_LYRIC = 24, SIZE_CHORD = 14, SIZE_SECTION = 16, SIZE_COPY = 14;
    const PT_TO_PX = 96 / 72; 

    let currentPreviewIndex = 0;
    let currentShift = 0;
    let selectedBgPath = "assets/bg-default.png"; // Ensure these paths exist in your assets folder

    const SCALE = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B'];
    const FLAT_MAP = { 'Db': 'C#', 'Eb': 'D#', 'Gb': 'F#', 'Ab': 'G#', 'Bb': 'A#' };

    const lyricInput = document.getElementById('valLyrics');
    if (document.getElementById('vBadge')) document.getElementById('vBadge').innerText = `v${VERSION}`;

    // --- 1. BACKGROUND GALLERY ---
    const bgOptions = [
        { name: 'Default', path: 'assets/bg-default.png' },
        { name: 'Holy', path: 'assets/bg-HolySpirit.png' },
        { name: 'Linen', path: 'assets/bg-linen.png' },
        { name: 'Soft', path: 'assets/bg-soft.png' },
        { name: 'Dark', path: 'assets/bg-dark.png' },
        { name: 'Cloud', path: 'assets/bg-cloud.png' }
    ];

    const bgSelector = document.getElementById('bgSelector');
    bgOptions.forEach(opt => {
        const thumb = document.createElement('div');
        thumb.className = `bg-thumb ${opt.path === selectedBgPath ? 'active' : ''}`;
        thumb.style.backgroundImage = `url(${opt.path})`;
        thumb.onclick = () => {
            document.querySelectorAll('.bg-thumb').forEach(t => t.classList.remove('active'));
            thumb.classList.add('active');
            selectedBgPath = opt.path;
            updatePreview();
        };
        bgSelector.appendChild(thumb);
    });

    // --- 2. TRANSPOSE LOGIC ---
    function transposeChord(chord, steps) {
        return chord.replace(/[A-G][b#]?/g, m => {
            let n = FLAT_MAP[m] || m, i = SCALE.indexOf(n);
            if (i === -1) return m;
            let newI = (i + steps) % 12;
            while (newI < 0) newI += 12;
            return SCALE[newI];
        });
    }

    function isChordLine(s) {
        if (!s.trim() || (s.trim().startsWith('[') && s.trim().endsWith(']'))) return false;
        const cleaned = s.replace(/[A-G]|[m|maj|min|dim|aug|sus|2|4|5|7|9]|#|b|\s|\/|v|i|\[|\]/gi, "");
        return cleaned.length === 0;
    }

    // --- 3. PREVIEW ENGINE ---
    function updatePreview() {
        const mock = document.getElementById('slideMock');
        const containerWidth = mock.offsetWidth;
        if (containerWidth === 0) return;

        const scale = (containerWidth / 960) * PT_TO_PX;
        const sections = lyricInput.value.split(/(?=\[)/).filter(s => s.trim());
        const align = document.getElementById('slideAlign').value;
        const chordGap = document.getElementById('chordGap').value;

        if (currentPreviewIndex >= sections.length) currentPreviewIndex = Math.max(0, sections.length - 1);
        document.getElementById('slideIndicator').innerText = `Slide ${sections.length > 0 ? currentPreviewIndex + 1 : 0} / ${sections.length}`;
        
        mock.style.backgroundImage = `url(${selectedBgPath})`;

        // Elements
        const pt = document.getElementById('prevTitle'), pc = document.getElementById('prevCopy'), pl = document.getElementById('prevLyrics');
        const colT = document.getElementById('colTitle').value, colL = document.getElementById('colLyrics').value,
              colC = document.getElementById('colChords').value, colCp = document.getElementById('colCopy').value;

        [pt, pc, pl].forEach(el => { el.style.textAlign = align; el.style.fontFamily = MAIN_FONT; });

        // Title
        pt.innerText = document.getElementById('valTitle').value;
        pt.style.top = document.getElementById('yTitle').value + "%";
        pt.style.fontSize = (SIZE_TITLE * scale) + "px"; pt.style.fontWeight = "bold"; pt.style.color = colT;

        // Copy
        pc.innerText = document.getElementById('valCopy').value;
        pc.style.top = document.getElementById('yCopy').value + "%";
        pc.style.fontSize = (SIZE_COPY * scale) + "px"; pc.style.fontStyle = "italic"; pc.style.color = colCp;

        // Lyrics
        pl.style.top = document.getElementById('yLyrics').value + "%";
        pl.style.height = "70%"; pl.style.display = "flex"; pl.style.flexDirection = "column"; pl.style.justifyContent = "center";
        
        const active = (sections[currentPreviewIndex] || "").trim();
        pl.innerHTML = ""; const inner = document.createElement('div'); inner.style.width = "100%";
        
        const lines = active.split('\n');
        lines.forEach((line, i) => {
            const div = document.createElement('div'); div.style.whiteSpace = "pre";
            if (line.trim().startsWith('[') && line.trim().endsWith(']')) {
                div.style.fontSize = (SIZE_SECTION * scale) + "px"; div.style.fontWeight = "bold"; div.innerText = line; div.style.color = colL;
            } else if (isChordLine(line)) {
                div.style.marginBottom = (chordGap * scale) + "px";
                div.innerHTML = createHtmlLine(line, lines[i+1] || "", scale, align, colC);
            } else {
                div.style.fontSize = (SIZE_LYRIC * scale) + "px"; div.style.lineHeight = "1.2"; div.innerText = line || " "; div.style.color = colL;
            }
            inner.appendChild(div);
        });
        pl.appendChild(inner);
    }

    function createHtmlLine(chords, lyrics, scale, align, cCol) {
        let h = ""; const len = Math.max(chords.length, lyrics.length);
        for (let i = 0; i < len; i++) {
            const c = chords[i] || " ", l = lyrics[i] || " ";
            h += `<span style="position:relative; display:inline-block; font-size:${SIZE_LYRIC*scale}px; color:transparent;">${l===" "?"\u00A0":l}`;
            if (c!==" ") h += `<span style="position:absolute; left:0; bottom:100%; font-size:${SIZE_CHORD*scale}px; color:${cCol}; visibility:visible; font-family:monospace;">${c}</span>`;
            h += `</span>`;
        }
        return h;
    }

    // --- 4. IMPORT PPTX ---
    async function handleImport(event) {
        const file = event.target.files[0]; if (!file) return;
        try {
            const zip = await JSZip.loadAsync(file);
            const parser = new DOMParser();
            let extractedTitle = "", extractedCopy = "", fullLyrics = "";

            const slides = Object.keys(zip.files).filter(n => n.startsWith('ppt/slides/slide')).sort((a,b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));
            for (let path of slides) {
                const xml = await zip.file(path).async("text");
                const doc = parser.parseFromString(xml, "application/xml");
                for (let sp of doc.getElementsByTagNameNS("*", "sp")) {
                    let txt = ""; let isB = false, sz = 0;
                    for (let r of sp.getElementsByTagNameNS("*", "r")) {
                        const t = r.getElementsByTagNameNS("*", "t")[0];
                        if (t) txt += t.textContent;
                        const rPr = r.getElementsByTagNameNS("*", "rPr")[0];
                        if (rPr) {
                            if (rPr.getAttribute("b") === "1") isB = true;
                            if (rPr.getAttribute("sz")) sz = parseInt(rPr.getAttribute("sz")) / 100;
                        }
                    }
                    const clean = txt.trim();
                    if (sz === 32 && isB && !extractedTitle) extractedTitle = clean;
                    else if (sz === 14 && !extractedCopy) extractedCopy = clean;
                }
            }

            const notes = Object.keys(zip.files).filter(n => n.startsWith('ppt/notesSlides/notesSlide')).sort((a,b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));
            for (let path of notes) {
                const xml = await zip.file(path).async("text");
                const paras = parser.parseFromString(xml, "application/xml").getElementsByTagNameNS("*", "p");
                let noteTxt = "";
                for (let p of paras) {
                    let line = "";
                    for (let t of p.getElementsByTagNameNS("*", "t")) line += t.textContent;
                    if (!line.trim() || /^\d+$/.test(line.trim())) continue;
                    noteTxt += line + "\n";
                }
                if (noteTxt.trim()) fullLyrics += noteTxt.trim() + "\n\n";
            }
            document.getElementById('valTitle').value = extractedTitle || "Imported Song";
            document.getElementById('valCopy').value = extractedCopy || "";
            lyricInput.value = fullLyrics.trim();
            currentShift = 0; currentPreviewIndex = 0;
            document.getElementById('keyShift').innerText = "Shift: 0";
            updatePreview();
        } catch (e) { console.error("Import Error:", e); }
    }

    // --- 5. EXPORT PPTX ---
    async function downloadPptx() {
        const pres = new PptxGenJS(); pres.layout = 'LAYOUT_16x9';
        const songT = document.getElementById('valTitle').value.trim() || "Song";
        const align = document.getElementById('slideAlign').value;
        const colT = document.getElementById('colTitle').value.replace('#',''), colL = document.getElementById('colLyrics').value.replace('#',''), 
              colC = document.getElementById('colChords').value.replace('#',''), colCp = document.getElementById('colCopy').value.replace('#','');

        lyricInput.value.split(/(?=\[)/).filter(s => s.trim()).forEach(section => {
            let slide = pres.addSlide(); slide.background = { path: selectedBgPath };
            slide.addNotes(section);
            
            slide.addText(songT, { 
                x:"5%", y:document.getElementById('yTitle').value+"%", w:"90%", 
                fontSize:SIZE_TITLE, fontFace:MAIN_FONT, bold:true, align, color:colT, margin:0
            });

            let textObjs = [];
            const lines = section.trim().split('\n');
            lines.forEach((line, i) => {
                if (line.trim().startsWith('[') && line.trim().endsWith(']')) {
                    textObjs.push({ text: line+"\n", options: { fontSize:SIZE_SECTION, bold:true, color:colL } });
                } else if (isChordLine(line)) {
                    // Complexity: In PPTX chords and lyrics are separate runs
                    // We will simplify by exporting chords as a distinct line
                    textObjs.push({ text: line+"\n", options: { fontSize:SIZE_CHORD, color:colC, fontFace:"Courier New" } });
                } else {
                    textObjs.push({ text: (line||" ")+"\n", options: { fontSize:SIZE_LYRIC, color:colL } });
                }
            });

            slide.addText(textObjs, { 
                x:"5%", y:document.getElementById('yLyrics').value+"%", w:"90%", h:"70%", 
                fontFace:MAIN_FONT, valign:'middle', align, margin:0
            });

            slide.addText(document.getElementById('valCopy').value, { 
                x:"5%", y:document.getElementById('yCopy').value+"%", w:"90%", 
                fontSize:SIZE_COPY, fontFace:MAIN_FONT, italic:true, align, color:colCp, margin:0
            });
        });
        await pres.writeFile({ fileName: `${songT}.pptx` });
    }

    // --- LISTENERS ---
    document.querySelectorAll('input, textarea, select').forEach(el => el.addEventListener('input', updatePreview));
    document.getElementById('importPptx').addEventListener('change', handleImport);
    
    document.getElementById('btnUp').onclick = () => { 
        lyricInput.value = lyricInput.value.split('\n').map(l => isChordLine(l) ? l.replace(/\S+/g, c => transposeChord(c, 1)) : l).join('\n'); 
        currentShift++; document.getElementById('keyShift').innerText = `Shift: ${currentShift}`; updatePreview(); 
    };
    
    document.getElementById('btnDown').onclick = () => { 
        lyricInput.value = lyricInput.value.split('\n').map(l => isChordLine(l) ? l.replace(/\S+/g, c => transposeChord(c, -1)) : l).join('\n'); 
        currentShift--; document.getElementById('keyShift').innerText = `Shift: ${currentShift}`; updatePreview(); 
    };

    document.getElementById('downloadBtn').onclick = downloadPptx;
    document.getElementById('nextSlide').onclick = () => { currentPreviewIndex++; updatePreview(); };
    document.getElementById('prevSlide').onclick = () => { if(currentPreviewIndex > 0) { currentPreviewIndex--; updatePreview(); }};
    
    window.addEventListener('resize', updatePreview);
    updatePreview();
});
