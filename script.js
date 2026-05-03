document.addEventListener('DOMContentLoaded', () => {
    const VERSION = "2.0-PRO";
    const MAIN_FONT = "Times New Roman";
    const SIZE_TITLE = 32, SIZE_LYRIC = 24, SIZE_CHORD = 14, SIZE_SECTION = 16, SIZE_COPY = 14;
    const PT_TO_PX = 96 / 72; 
    
    let currentPreviewIndex = 0, currentShift = 0, selectedBgPath = "assets/bg-default.png";
    const SCALE = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B'];
    const FLAT_MAP = { 'Db': 'C#', 'Eb': 'D#', 'Gb': 'F#', 'Ab': 'G#', 'Bb': 'A#' };

    const lyricInput = document.getElementById('valLyrics');
    const bgSelector = document.getElementById('bgSelector');
    if (document.getElementById('vBadge')) document.getElementById('vBadge').innerText = VERSION;

    const bgOptions = [
        { name: 'Default', path: 'assets/bg-default.png' },
        { name: 'Holy', path: 'assets/bg-HolySpirit.png' },
        { name: 'Linen', path: 'assets/bg-linen.png' },
        { name: 'Soft', path: 'assets/bg-soft.png' },
        { name: 'Dark', path: 'assets/bg-dark.png' },
        { name: 'Cloud', path: 'assets/bg-cloud.png' }
    ];

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
        const mock = document.getElementById('slideMock');
        const sections = lyricInput.value.split(/(?=\[)/).filter(s => s.trim());
        if (currentPreviewIndex >= sections.length) currentPreviewIndex = Math.max(0, sections.length - 1);
        
        document.getElementById('slideIndicator').innerText = `Slide ${sections.length > 0 ? currentPreviewIndex + 1 : 0} / ${sections.length}`;
        
        const scale = (mock.offsetWidth / 960) * PT_TO_PX;
        mock.style.backgroundImage = `url(${selectedBgPath})`;

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
        pl.innerHTML = "";
        
        const active = (sections[currentPreviewIndex] || "").replace(/^[\n\r]+|[\n\r]+$/g, '');
        const inner = document.createElement('div'); inner.style.width = "100%";
        active.split('\n').forEach((line, i, arr) => {
            const div = document.createElement('div'); div.style.whiteSpace = "pre";
            if (line.trim().startsWith('[') && line.trim().endsWith(']')) {
                div.style.fontSize = (SIZE_SECTION * scale) + "px"; div.style.fontWeight = "bold"; div.innerText = line; div.style.color = colL;
            } else if (isChordLine(line)) {
                div.style.fontSize = (SIZE_CHORD * scale) + "px"; div.style.lineHeight = "0.7"; 
                div.style.marginBottom = (parseInt(document.getElementById('chordGap').value) * scale) + "px";
                div.innerHTML = createHtmlLine(line, arr[i+1] || "", scale, align, colC);
            } else {
                div.style.fontSize = (SIZE_LYRIC * scale) + "px"; div.innerText = line || " "; div.style.color = colL;
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

    const fsBtn = document.getElementById('fullscreenBtn');
    const fsWrapper = document.getElementById('fullscreenWrapper');
    fsBtn.onclick = () => {
        if (!document.fullscreenElement) fsWrapper.requestFullscreen().catch(console.error);
        else document.exitFullscreen();
    };
    document.addEventListener('fullscreenchange', () => setTimeout(updatePreview, 100));

    async function downloadPptx() {
        const pres = new PptxGenJS(); pres.layout = 'LAYOUT_16x9';
        const songT = document.getElementById('valTitle').value.trim() || "Song";
        const align = document.getElementById('slideAlign').value;
        const colT = document.getElementById('colTitle').value.replace('#',''), colL = document.getElementById('colLyrics').value.replace('#',''), 
              colC = document.getElementById('colChords').value.replace('#',''), colCp = document.getElementById('colCopy').value.replace('#','');

        lyricInput.value.split(/(?=\[)/).filter(s => s.trim()).forEach(section => {
            let slide = pres.addSlide(); slide.background = { path: selectedBgPath };
            slide.addNotes(section);
            slide.addText(songT, { x:"5%", y:document.getElementById('yTitle').value+"%", w:"90%", fontSize:SIZE_TITLE, fontFace:MAIN_FONT, bold:true, align, color:colT });
            let textObjs = [];
            section.replace(/^[\n\r]+|[\n\r]+$/g, '').split('\n').forEach((line, i, arr) => {
                if (line.trim().startsWith('[') && line.trim().endsWith(']')) textObjs.push({ text: line+"\n", options: { fontSize:SIZE_SECTION, bold:true, color:colL } });
                else if (isChordLine(line)) textObjs.push(...createPptxLine(line, arr[i+1] || "", align, colC));
                else textObjs.push({ text: (line||" ")+"\n", options: { fontSize:SIZE_LYRIC, color:colL } });
            });
            slide.addText(textObjs, { x:"5%", y:document.getElementById('yLyrics').value+"%", w:"90%", h:"70%", fontFace:MAIN_FONT, valign:'middle', align });
            slide.addText(document.getElementById('valCopy').value, { x:"5%", y:document.getElementById('yCopy').value+"%", w:"90%", fontSize:SIZE_COPY, fontFace:MAIN_FONT, italic:true, align, color:colCp });
        });
        pres.writeFile({ fileName: `${songT}.pptx` });
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

    document.getElementById('importPptx').onchange = async (e) => {
        const file = e.target.files[0]; if (!file) return;
        const zip = await JSZip.loadAsync(file); const parser = new DOMParser();
        let fullLyrics = "";
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
        lyricInput.value = fullLyrics.trim(); updatePreview();
    };

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
    document.getElementById('nextSlide').onclick = () => { currentPreviewIndex++; updatePreview(); };
    document.getElementById('prevSlide').onclick = () => { if(currentPreviewIndex>0) { currentPreviewIndex--; updatePreview(); }};
    
    document.getElementById('btnBgDown').onclick = () => bgSelector.scrollBy({ top: 100, behavior: 'smooth' });
    document.getElementById('btnBgUp').onclick = () => bgSelector.scrollBy({ top: -100, behavior: 'smooth' });

    window.onresize = updatePreview;
    updatePreview();
});
