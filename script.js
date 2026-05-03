document.addEventListener('DOMContentLoaded', () => {
    const VERSION = "2.0-SETLIST";
    const MAIN_FONT = "Times New Roman";
    const SIZE_TITLE = 32, SIZE_LYRIC = 24, SIZE_CHORD = 14, SIZE_SECTION = 16, SIZE_COPY = 14;
    const PT_TO_PX = 96 / 72; 
    const SCALE = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B'];
    const FLAT_MAP = { 'Db': 'C#', 'Eb': 'D#', 'Gb': 'F#', 'Ab': 'G#', 'Bb': 'A#' };

    // --- STATE ---
    let setlist = [{
        title: "Song Title",
        lyrics: "[Verse 1]\nG           C\nLyrics start here...",
        copy: "© Copyright Info",
        shift: 0,
        align: 'center'
    }];
    let activeIndex = 0;
    let currentPreviewIndex = 0; 
    let selectedBgPath = "assets/bg-default.png";

    const lyricInput = document.getElementById('valLyrics');
    const titleInput = document.getElementById('valTitle');
    const copyInput = document.getElementById('valCopy');
    const alignInput = document.getElementById('slideAlign');
    if (document.getElementById('vBadge')) document.getElementById('vBadge').innerText = VERSION;

    // --- PPTX IMPORT LOGIC (The "Missing" Part) ---
    async function handleImport(event) {
        const file = event.target.files[0];
        if (!file) return;
        try {
            const zip = await JSZip.loadAsync(file);
            const parser = new DOMParser();
            let extractedTitle = "", extractedCopy = "", fullLyrics = "";

            // 1. Extract Meta (Title/Copy) from slides
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
                        if (rPr && rPr.getAttribute("sz")) sz = parseInt(rPr.getAttribute("sz")) / 100;
                        if (rPr && rPr.getAttribute("b") === "1") isB = true;
                    }
                    if (sz === 32 && isB && !extractedTitle) extractedTitle = txt.trim();
                    if (sz === 14 && !extractedCopy) extractedCopy = txt.trim();
                }
            }

            // 2. Extract Lyrics from Notes
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

            // 3. Add as NEW song to setlist
            setlist.push({
                title: extractedTitle || file.name.replace('.pptx',''),
                lyrics: fullLyrics.trim() || "No lyrics found in notes.",
                copy: extractedCopy || "©",
                shift: 0,
                align: 'center'
            });
            activeIndex = setlist.length - 1;
            loadActiveSong();
        } catch (e) { 
            console.error(e); 
            alert("Error parsing PPTX. Make sure chords/lyrics are in the Slide Notes.");
        }
    }

    // --- BG GALLERY ---
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

    // --- SETLIST CORE ---
    function renderSetlist() {
        const container = document.getElementById('setlistItems');
        container.innerHTML = '';
        setlist.forEach((song, idx) => {
            const div = document.createElement('div');
            div.className = `set-item ${idx === activeIndex ? 'active' : ''}`;
            div.innerHTML = `<span>${idx + 1}. ${song.title || 'Untitled'}</span>
                             <button class="del-btn" data-idx="${idx}">×</button>`;
            div.onclick = (e) => { if (!e.target.classList.contains('del-btn')) selectSong(idx); };
            container.appendChild(div);
        });

        document.querySelectorAll('.del-btn').forEach(btn => {
            btn.onclick = (e) => {
                e.stopPropagation();
                if (setlist.length > 1) {
                    setlist.splice(parseInt(e.target.dataset.idx), 1);
                    activeIndex = Math.max(0, activeIndex - 1);
                    loadActiveSong();
                }
            };
        });
    }

    function selectSong(idx) {
        saveActiveSong();
        activeIndex = idx;
        currentPreviewIndex = 0;
        loadActiveSong();
    }

    function saveActiveSong() {
        setlist[activeIndex] = {
            ...setlist[activeIndex],
            title: titleInput.value,
            lyrics: lyricInput.value,
            copy: copyInput.value,
            align: alignInput.value
        };
    }

    function loadActiveSong() {
        const s = setlist[activeIndex];
        titleInput.value = s.title;
        lyricInput.value = s.lyrics;
        copyInput.value = s.copy;
        alignInput.value = s.align;
        document.getElementById('keyShift').innerText = `Shift: ${s.shift}`;
        renderSetlist();
        updatePreview();
    }

    // --- TRANSPOSE HELPERS ---
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

    function applyTranspose(steps, global = false) {
        if (global) {
            setlist.forEach(s => {
                s.lyrics = s.lyrics.split('\n').map(l => isChordLine(l) ? l.replace(/\S+/g, c => transposeChord(c, steps)) : l).join('\n');
                s.shift += steps;
            });
        } else {
            const s = setlist[activeIndex];
            s.lyrics = s.lyrics.split('\n').map(l => isChordLine(l) ? l.replace(/\S+/g, c => transposeChord(c, steps)) : l).join('\n');
            s.shift += steps;
        }
        loadActiveSong();
    }

    // --- RENDERING ---
    function updatePreview() {
        const mock = document.getElementById('slideMock');
        if(!mock) return;
        const sections = lyricInput.value.split(/(?=\[)/).filter(s => s.trim());
        
        if (currentPreviewIndex >= sections.length) currentPreviewIndex = Math.max(0, sections.length - 1);
        document.getElementById('slideIndicator').innerText = `Slide ${sections.length > 0 ? currentPreviewIndex + 1 : 0} / ${sections.length}`;
        
        const scale = (mock.offsetWidth / 960) * PT_TO_PX;
        mock.style.backgroundImage = `url(${selectedBgPath})`;

        const align = alignInput.value;
        const colT = document.getElementById('colTitle').value, colL = document.getElementById('colLyrics').value,
              colC = document.getElementById('colChords').value, colCp = document.getElementById('colCopy').value;

        const pt = document.getElementById('prevTitle'), pc = document.getElementById('prevCopy'), pl = document.getElementById('prevLyrics');
        [pt, pc, pl].forEach(el => { el.style.textAlign = align; el.style.fontFamily = MAIN_FONT; });

        pt.innerText = titleInput.value;
        pt.style.top = document.getElementById('yTitle').value + "%";
        pt.style.fontSize = (SIZE_TITLE * scale) + "px"; pt.style.fontWeight = "bold"; pt.style.color = colT;

        pc.innerText = copyInput.value;
        pc.style.top = document.getElementById('yCopy').value + "%";
        pc.style.fontSize = (SIZE_COPY * scale) + "px"; pc.style.fontStyle = "italic"; pc.style.color = colCp;

        pl.style.top = document.getElementById('yLyrics').value + "%";
        pl.style.height = "70%"; pl.style.display = "flex"; pl.style.flexDirection = "column"; pl.style.justifyContent = "center";
        
        const activeText = (sections[currentPreviewIndex] || "").replace(/^[\n\r]+|[\n\r]+$/g, '');
        pl.innerHTML = ""; const inner = document.createElement('div'); inner.style.width = "100%";

        activeText.split('\n').forEach((line, i, arr) => {
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

    // --- PPTX EXPORT ---
    async function downloadPptx(all = false) {
        saveActiveSong();
        const pres = new PptxGenJS();
        pres.layout = 'LAYOUT_16x9';
        const colT = document.getElementById('colTitle').value.replace('#',''), 
              colL = document.getElementById('colLyrics').value.replace('#',''), 
              colC = document.getElementById('colChords').value.replace('#',''), 
              colCp = document.getElementById('colCopy').value.replace('#','');

        const songs = all ? setlist : [setlist[activeIndex]];
        songs.forEach(song => {
            song.lyrics.split(/(?=\[)/).filter(s => s.trim()).forEach(section => {
                let slide = pres.addSlide();
                slide.background = { path: selectedBgPath };
                slide.addNotes(section);
                slide.addText(song.title, { x:"5%", y:document.getElementById('yTitle').value+"%", w:"90%", fontSize:SIZE_TITLE, fontFace:MAIN_FONT, bold:true, align:song.align, color:colT, valign:'top' });
                
                let textObjs = [];
                section.replace(/^[\n\r]+|[\n\r]+$/g, '').split('\n').forEach((line, i, arr) => {
                    if (line.trim().startsWith('[') && line.trim().endsWith(']')) textObjs.push({ text: line+"\n", options: { fontSize:SIZE_SECTION, bold:true, color:colL } });
                    else if (isChordLine(line)) {
                        let lineR = []; const len = song.align === 'center' ? Math.max(line.length, (arr[i+1]||"").length) : line.length;
                        for (let k = 0; k < len; k++) {
                            const c = line[k] || " ", l = (arr[i+1]||"")[k] || " ";
                            if (c !== " ") lineR.push({ text: c, options: { color: colC, fontSize: SIZE_CHORD } });
                            else lineR.push({ text: l===""?" ":l, options: { transparency: 100, fontSize: SIZE_LYRIC } });
                        }
                        lineR.push({ text: "\n" }); textObjs.push(...lineR);
                    }
                    else textObjs.push({ text: (line||" ")+"\n", options: { fontSize:SIZE_LYRIC, color:colL } });
                });
                slide.addText(textObjs, { x:"5%", y:document.getElementById('yLyrics').value+"%", w:"90%", h:"70%", fontFace:MAIN_FONT, valign:'middle', align:song.align, lineSpacing: SIZE_LYRIC * 0.9 });
                slide.addText(song.copy, { x:"5%", y:document.getElementById('yCopy').value+"%", w:"90%", fontSize:SIZE_COPY, fontFace:MAIN_FONT, italic:true, align:song.align, color:colCp, valign:'top' });
            });
        });
        await pres.writeFile({ fileName: all ? "Setlist.pptx" : `${setlist[activeIndex].title}.pptx` });
    }

    // --- LISTENERS ---
    document.getElementById('importPptx').addEventListener('change', handleImport);
    document.getElementById('btnAddSong').onclick = () => {
        saveActiveSong();
        setlist.push({ title: "New Song", lyrics: "[Verse 1]\n", copy: "©", shift: 0, align: 'center' });
        activeIndex = setlist.length - 1;
        loadActiveSong();
    };

    document.querySelectorAll('input, textarea, select').forEach(el => el.addEventListener('input', () => {
        saveActiveSong();
        updatePreview();
        if(el.id === 'valTitle') renderSetlist();
    }));

    document.getElementById('btnUp').onclick = () => applyTranspose(1);
    document.getElementById('btnDown').onclick = () => applyTranspose(-1);
    document.getElementById('btnGlobalTranspose').onclick = () => {
        const val = prompt("Steps to shift entire setlist (e.g. 2 or -1):", "1");
        if (val) applyTranspose(parseInt(val), true);
    };

    document.getElementById('nextSlide').onclick = () => { currentPreviewIndex++; updatePreview(); };
    document.getElementById('prevSlide').onclick = () => { if(currentPreviewIndex > 0) { currentPreviewIndex--; updatePreview(); }};
    document.getElementById('downloadCurrent').onclick = (e) => { e.preventDefault(); downloadPptx(false); };
    document.getElementById('downloadSetlist').onclick = (e) => { e.preventDefault(); downloadPptx(true); };
    document.getElementById('btnFullScreen').onclick = () => {
        const m = document.getElementById('slideMock');
        if (!document.fullscreenElement) m.requestFullscreen(); else document.exitFullscreen();
    };

    window.addEventListener('resize', updatePreview);
    loadActiveSong();
});