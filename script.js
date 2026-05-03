document.addEventListener('DOMContentLoaded', () => {
    const VERSION = "2.2.0-PRO";
    const MAIN_FONT = "Times New Roman";
    const SIZE_TITLE = 32, SIZE_LYRIC = 24, SIZE_CHORD = 14, SIZE_SECTION = 16, SIZE_COPY = 14;
    const PT_TO_PX = 96 / 72; 
    
    let currentPreviewIndex = 0, currentShift = 0, selectedBgPath = "assets/bg-default.png";
    const SCALE = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B'];
    const FLAT_MAP = { 'Db': 'C#', 'Eb': 'D#', 'Gb': 'F#', 'Ab': 'G#', 'Bb': 'A#' };

    const lyricInput = document.getElementById('valLyrics');
    const slideMock = document.getElementById('slideMock');
    const fullScreenBtn = document.getElementById('fullScreenBtn');
    const fullScreenTarget = document.getElementById('fullScreenTarget');

    if (document.getElementById('vBadge')) document.getElementById('vBadge').innerText = `v${VERSION}`;

    // --- 1. BG GALLERY ---
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

    const btnBgUp = document.getElementById('btnBgUp');
    const btnBgDown = document.getElementById('btnBgDown');
    btnBgDown.onclick = () => bgSelector.scrollBy({ top: 100, behavior: 'smooth' });
    btnBgUp.onclick = () => bgSelector.scrollBy({ top: -100, behavior: 'smooth' });

    // --- 2. RESIZER LOGIC ---
    const resizer = document.getElementById('dragMe');
    const leftSide = document.getElementById('editorPanel');
    let x = 0, leftWidth = 0;

    const mouseDownHandler = (e) => {
        x = e.clientX;
        leftWidth = leftSide.getBoundingClientRect().width;
        document.addEventListener('mousemove', mouseMoveHandler);
        document.addEventListener('mouseup', mouseUpHandler);
        document.body.style.cursor = 'col-resize';
        document.body.style.userSelect = 'none';
    };
    const mouseMoveHandler = (e) => {
        const dx = e.clientX - x;
        const newLeftWidth = ((leftWidth + dx) * 100) / resizer.parentNode.getBoundingClientRect().width;
        leftSide.style.width = `${newLeftWidth}%`;
        updatePreview();
    };
    const mouseUpHandler = () => {
        document.removeEventListener('mousemove', mouseMoveHandler);
        document.removeEventListener('mouseup', mouseUpHandler);
        document.body.style.removeProperty('cursor');
        document.body.style.removeProperty('user-select');
    };
    resizer.addEventListener('mousedown', mouseDownHandler);

    // --- 3. FULLSCREEN & KEYBOARD ---
    const toggleFullScreen = () => {
        if (!document.fullscreenElement) fullScreenTarget.requestFullscreen().catch(e => console.error(e));
        else document.exitFullscreen();
    };
    fullScreenBtn.onclick = toggleFullScreen;
    document.addEventListener('fullscreenchange', () => setTimeout(updatePreview, 100));

    document.addEventListener('keydown', (e) => {
        if (document.activeElement.tagName === 'TEXTAREA' || document.activeElement.tagName === 'INPUT') return;
        const sections = lyricInput.value.split(/(?=\[)/).filter(s => s.trim());
        if (e.key === "ArrowRight" || e.key === " ") {
            if (currentPreviewIndex < sections.length - 1) { currentPreviewIndex++; updatePreview(); }
        } else if (e.key === "ArrowLeft") {
            if (currentPreviewIndex > 0) { currentPreviewIndex--; updatePreview(); }
        } else if (e.key.toLowerCase() === "f") {
            toggleFullScreen();
        }
    });

    // --- 4. PREVIEW & TRANSPOSE ---
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
        return s.replace(/[A-G]|[m|maj|min|dim|aug|sus|2|4|5|7|9]|#|b|\s|\/|v|i|\[|\]/gi, "").length === 0;
    }

    function updatePreview() {
        if(!slideMock) return;
        const sections = lyricInput.value.split(/(?=\[)/).filter(s => s.trim());
        if (currentPreviewIndex >= sections.length) currentPreviewIndex = Math.max(0, sections.length - 1);
        document.getElementById('slideIndicator').innerText = `Slide ${sections.length > 0 ? currentPreviewIndex + 1 : 0} / ${sections.length}`;
        
        const scale = (slideMock.offsetWidth / 960) * PT_TO_PX;
        slideMock.style.backgroundImage = `url(${selectedBgPath})`;

        const colT = document.getElementById('colTitle').value, colL = document.getElementById('colLyrics').value,
              colC = document.getElementById('colChords').value, colCp = document.getElementById('colCopy').value;
        const align = document.getElementById('slideAlign').value;

        const pt = document.getElementById('prevTitle'), pc = document.getElementById('prevCopy'), pl = document.getElementById('prevLyrics');
        [pt, pc, pl].forEach(el => { el.style.textAlign = align; el.style.fontFamily = MAIN_FONT; });

        pt.innerText = document.getElementById('valTitle').value;
        pt.style.top = document.getElementById('yTitle').value + "%";
        pt.style.fontSize = (SIZE_TITLE * scale) + "px"; pt.style.fontWeight = "bold"; pt.style.color = colT;

        pc.innerText = document.getElementById('valCopy').value;
        pc.style.top = document.getElementById('yCopy').value + "%";
        pc.style.fontSize = (SIZE_COPY * scale) + "px"; pc.style.fontStyle = "italic"; pc.style.color = colCp;

        pl.style.top = document.getElementById('yLyrics').value + "%";
        pl.style.height = "70%"; pl.style.display = "flex"; pl.style.flexDirection = "column"; pl.style.justifyContent = "center";
        
        const active = (sections[currentPreviewIndex] || "").replace(/^[\n\r]+|[\n\r]+$/g, '');
        pl.innerHTML = ""; const inner = document.createElement('div'); inner.style.width = "100%";
        active.split('\n').forEach((line, i, arr) => {
            const div = document.createElement('div'); div.style.whiteSpace = "pre";
            if (line.trim().startsWith('[') && line.trim().endsWith(']')) {
                div.style.fontSize = (SIZE_SECTION * scale) + "px"; div.style.fontWeight = "bold"; div.innerText = line; div.style.color = colL;
            } else if (isChordLine(line)) {
                div.style.fontSize = (SIZE_CHORD * scale) + "px"; div.style.lineHeight = "0.7"; 
                div.style.marginBottom = (parseInt(document.getElementById('chordGap').value) * scale) + "px";
                div.innerHTML = createHtmlLine(line, arr[i+1] || "", scale, align, colC);
            } else {
                div.style.fontSize = (SIZE_LYRIC * scale) + "px"; div.style.lineHeight = "1"; div.innerText = line || " "; div.style.color = colL;
            }
            inner.appendChild(div);
        });
        pl.appendChild(inner);
    }

    function createHtmlLine(chords, lyrics, scale, align, cCol) {
        let h = ""; const len = align === 'center' ? Math.max(chords.length, lyrics.length) : chords.length;
        for (let i = 0; i < len; i++) {
            const c = chords[i] || " ", l = lyrics[i] || " ";
            h += `<span style="position:relative; display:inline-block; font-size:${SIZE_LYRIC*scale}px; color:transparent;">${l===" "?"\u00A0":l}`;
            if (c!==" ") h += `<span style="position:absolute; left:0; bottom:0; font-size:${SIZE_CHORD*scale}px; color:${cCol}; visibility:visible; font-family:monospace;">${c}</span>`;
            h += `</span>`;
        }
        return h;
    }

    // --- 5. PPTX DOWNLOAD ---
    async function downloadPptx() {
        const pres = new PptxGenJS(); pres.layout = 'LAYOUT_16x9';
        const songT = document.getElementById('valTitle').value.trim() || "Song";
        const align = document.getElementById('slideAlign').value;
        const colT = document.getElementById('colTitle').value.replace('#',''), colL = document.getElementById('colLyrics').value.replace('#',''), 
              colC = document.getElementById('colChords').value.replace('#',''), colCp = document.getElementById('colCopy').value.replace('#','');

        lyricInput.value.split(/(?=\[)/).filter(s => s.trim()).forEach(section => {
            let slide = pres.addSlide(); slide.background = { path: selectedBgPath };
            slide.addNotes(section);
            slide.addText(songT, { x:"5%", y:document.getElementById('yTitle').value+"%", w:"90%", fontSize:SIZE_TITLE, fontFace:MAIN_FONT, bold:true, align, color:colT, margin: 0, valign: 'top' });
            let textObjs = [];
            section.replace(/^[\n\r]+|[\n\r]+$/g, '').split('\n').forEach((line, i, arr) => {
                if (line.trim().startsWith('[') && line.trim().endsWith(']')) textObjs.push({ text: line+"\n", options: { fontSize:SIZE_SECTION, bold:true, color:colL } });
                else if (isChordLine(line)) textObjs.push(...createPptxLine(line, arr[i+1] || "", align, colC));
                else textObjs.push({ text: (line||" ")+"\n", options: { fontSize:SIZE_LYRIC, color:colL } });
            });
            slide.addText(textObjs, { x:"5%", y:document.getElementById('yLyrics').value+"%", w:"90%", h:"70%", fontFace:MAIN_FONT, valign:'middle', align, lineSpacing: SIZE_LYRIC * 0.9, margin: 0 });
            slide.addText(document.getElementById('valCopy').value, { x:"5%", y:document.getElementById('yCopy').value+"%", w:"90%", fontSize:SIZE_COPY, fontFace:MAIN_FONT, italic:true, align, color:colCp, margin: 0, valign: 'top' });
        });
        await pres.writeFile({ fileName: `${songT}.pptx` });
    }

    function createPptxLine(chords, lyrics, align, cCol) {
        let r = []; const len = align === 'center' ? Math.max(chords.length, lyrics.length) : chords.length;
        for (let i = 0; i < len; i++) {
            const c = chords[i] || " ", l = lyrics[i] || " ";
            if (c !== " ") r.push({ text: c, options: { color: cCol, fontSize: SIZE_CHORD } });
            else r.push({ text: l===""?" ":l, options: { transparency: 100, fontSize: SIZE_LYRIC } });
        }
        r.push({ text: "\n" }); return r;
    }

    // --- 6. LISTENERS ---
    document.querySelectorAll('input, textarea, select').forEach(el => el.addEventListener('input', updatePreview));
    document.getElementById('btnUp').onclick = () => { 
        lyricInput.value = lyricInput.value.split('\n').map(l => isChordLine(l) ? l.replace(/\S+/g, c => transposeChord(c, 1)) : l).join('\n'); 
        currentShift++; document.getElementById('keyShift').innerText = `Shift: ${currentShift}`; updatePreview(); 
    };
    document.getElementById('btnDown').onclick = () => { 
        lyricInput.value = lyricInput.value.split('\n').map(l => isChordLine(l) ? l.replace(/\S+/g, c => transposeChord(c, -1)) : l).join('\n'); 
        currentShift--; document.getElementById('keyShift').innerText = `Shift: ${currentShift}`; updatePreview(); 
    };
    document.getElementById('downloadBtn').onclick = downloadPptx;
    document.getElementById('nextSlide').onclick = () => { 
        const sections = lyricInput.value.split(/(?=\[)/).filter(s => s.trim());
        if(currentPreviewIndex < sections.length - 1) { currentPreviewIndex++; updatePreview(); }
    };
    document.getElementById('prevSlide').onclick = () => { if(currentPreviewIndex > 0) { currentPreviewIndex--; updatePreview(); }};
    window.addEventListener('resize', updatePreview);
    updatePreview();
});
