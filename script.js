document.addEventListener('DOMContentLoaded', () => {
    // --- 1. CONFIGURATION & STATE ---
    const VERSION = "3.0-PRO-SETLIST";
    const MAIN_FONT = "Times New Roman";
    const SIZE_TITLE = 32, SIZE_LYRIC = 24, SIZE_CHORD = 14, SIZE_SECTION = 16, SIZE_COPY = 14;
    const PT_TO_PX = 96 / 72;
    
    let currentPreviewIndex = 0, currentShift = 0;
    let selectedBgPath = "assets/bg-default.png";
    let setlist = []; 
    let lastLoadedIndex = -1; // -1 indicates no song is currently active/safe to save

    const SCALE = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B'];
    const FLAT_MAP = { 'Db': 'C#', 'Eb': 'D#', 'Gb': 'F#', 'Ab': 'G#', 'Bb': 'A#' };

    const lyricInput = document.getElementById('valLyrics');
    const bgSelector = document.getElementById('bgSelector');
    const dropdown = document.getElementById('setlistDropdown');
    const setlistBtn = document.getElementById('downloadSetlistBtn');

    if (document.getElementById('vBadge')) document.getElementById('vBadge').innerText = VERSION;

    // --- 2. BACKGROUND GALLERY SETUP ---
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

    // --- 3. TRANSPOSE ENGINE ---
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

    // --- 4. PREVIEW RENDERING ---
    function updatePreview() {
        const mock = document.getElementById('slideMock'); if (!mock) return;
        const sections = lyricInput.value.split(/(?=\[)/).filter(s => s.trim());
        if (currentPreviewIndex >= sections.length) currentPreviewIndex = Math.max(0, sections.length - 1);
        document.getElementById('slideIndicator').innerText = `Slide ${sections.length > 0 ? currentPreviewIndex + 1 : 0} / ${sections.length}`;
        const rect = mock.getBoundingClientRect(); const scale = (rect.width / 960) * PT_TO_PX;
        mock.style.backgroundImage = `url(${selectedBgPath})`;
        const colT = document.getElementById('colTitle').value, colL = document.getElementById('colLyrics').value, colC = document.getElementById('colChords').value, colCp = document.getElementById('colCopy').value, align = document.getElementById('slideAlign').value;
        const pt = document.getElementById('prevTitle'), pc = document.getElementById('prevCopy'), pl = document.getElementById('prevLyrics');
        [pt, pc, pl].forEach(el => { el.style.textAlign = align; el.style.fontFamily = MAIN_FONT; el.style.margin = "0"; el.style.padding = "0"; });
        pt.innerText = document.getElementById('valTitle').value; pt.style.top = document.getElementById('yTitle').value + "%";
        pt.style.fontSize = (SIZE_TITLE * scale) + "px"; pt.style.fontWeight = "bold"; pt.style.color = colT;
        pc.innerText = document.getElementById('valCopy').value; pc.style.top = document.getElementById('yCopy').value + "%";
        pc.style.fontSize = (SIZE_COPY * scale) + "px"; pc.style.fontStyle = "italic"; pc.style.color = colCp;
        pl.style.top = document.getElementById('yLyrics').value + "%"; pl.style.height = "70%"; pl.style.display = "flex"; pl.style.flexDirection = "column"; pl.style.justifyContent = "center";
        const active = (sections[currentPreviewIndex] || "").replace(/^[\n\r]+|[\n\r]+$/g, '');
        pl.innerHTML = ""; const inner = document.createElement('div'); inner.style.width = "100%";
        active.split('\n').forEach((line, i, arr) => {
            const div = document.createElement('div'); div.style.whiteSpace = "pre";
            if (line.trim().startsWith('[') && line.trim().endsWith(']')) {
                div.style.fontSize = (SIZE_SECTION * scale) + "px"; div.style.fontWeight = "bold"; div.innerText = line; div.style.color = colL;
            } else if (isChordLine(line)) {
                div.style.fontSize = (SIZE_CHORD * scale) + "px"; div.style.lineHeight = "0.7"; div.style.marginBottom = (parseInt(document.getElementById('chordGap').value) * scale) + "px";
                div.innerHTML = createHtmlLine(line, arr[i+1] || "", scale, align, colC);
            } else {
                div.style.fontSize = (SIZE_LYRIC * scale) + "px"; div.innerText = line || " "; div.style.color = colL;
            }
            inner.appendChild(div);
        });
        pl.appendChild(inner);
    }

    function createHtmlLine(chords, lyrics, scale, align, cCol) {
        let h = ""; const maxLen = Math.max(chords.length, lyrics.length);
        const lSize = SIZE_LYRIC * scale, cSize = SIZE_CHORD * scale;
        for (let i = 0; i < maxLen; i++) {
            const charL = lyrics[i] || " ", charC = chords[i] || "";
            h += `<span style="position:relative; display:inline-block; font-size:${lSize}px; color:transparent;">`;
            h += charL === " " ? "\u00A0" : charL;
            if (charC.trim() !== "") {
                let fullChord = charC, j = i + 1;
                while (j < chords.length && chords[j] !== " ") { fullChord += chords[j]; chords = chords.substring(0, j) + " " + chords.substring(j + 1); j++; }
                h += `<span style="position:absolute; left:0; bottom:0; font-size:${cSize}px; color:${cCol}; visibility:visible; font-family:serif; white-space:nowrap; transform:translateY(-20%); font-weight:600;">${fullChord}</span>`;
            }
            h += `</span>`;
        }
        return h;
    }

    // --- 6. FULLSCREEN API ---
    const fsBtn = document.getElementById('fullscreenBtn'), fsWrapper = document.getElementById('fullscreenWrapper');
    if (fsBtn) { fsBtn.onclick = () => { if (!document.fullscreenElement) fsWrapper.requestFullscreen().catch(console.error); else document.exitFullscreen(); }; }
    document.addEventListener('fullscreenchange', () => setTimeout(updatePreview, 100));

    // --- 7. PPTX EXPORT ENGINE ---
    function createPptxLine(chords, lyrics, align, cCol) {
        let r = []; const len = align === 'center' ? Math.max(chords.length, lyrics.length) : chords.length;
        for (let i = 0; i < len; i++) {
            const c = chords[i] || " ", l = lyrics[i] || " ";
            if (c !== " ") r.push({ text: c, options: { color: cCol, fontSize: SIZE_CHORD } });
            else r.push({ text: l===""?" ":l, options: { transparency: 100, fontSize: SIZE_LYRIC } });
        }
        r.push({ text: "\n" }); return r;
    }

    async function downloadPptx() {
        const pres = new PptxGenJS(); pres.layout = 'LAYOUT_16x9';
        addSongToPres(pres, document.getElementById('valTitle').value, document.getElementById('valCopy').value, document.getElementById('valLyrics').value);
        pres.writeFile({ fileName: `${document.getElementById('valTitle').value}.pptx` });
    }

    async function downloadFullSetlist() {
        saveCurrentState(); // Capture last edits
        const pres = new PptxGenJS(); pres.layout = 'LAYOUT_16x9';
        setlist.forEach(s => addSongToPres(pres, s.title, s.copy, s.lyrics));
        pres.writeFile({ fileName: `Full_Setlist.pptx` });
    }

    function addSongToPres(pres, songTitle, songCopy, songLyrics) {
        const align = document.getElementById('slideAlign').value, colT = document.getElementById('colTitle').value.replace('#',''), colL = document.getElementById('colLyrics').value.replace('#',''), colC = document.getElementById('colChords').value.replace('#',''), colCp = document.getElementById('colCopy').value.replace('#',''), yTitle = document.getElementById('yTitle').value, yLyrics = document.getElementById('yLyrics').value, yCopy = document.getElementById('yCopy').value;
        songLyrics.split(/(?=\[)/).filter(s => s.trim()).forEach(section => {
            let slide = pres.addSlide(); slide.background = { path: selectedBgPath }; slide.addNotes(section);
            slide.addText(songTitle, { x: "5%", y: yTitle + "%", w: "90%", fontSize: SIZE_TITLE, fontFace: MAIN_FONT, bold: true, align, color: colT, margin: 0, valign: 'top' });
            let textObjs = [];
            section.replace(/^[\n\r]+|[\n\r]+$/g, '').split('\n').forEach((line, i, arr) => {
                if (line.trim().startsWith('[') && line.trim().endsWith(']')) textObjs.push({ text: line + "\n", options: { fontSize: SIZE_SECTION, bold: true, color: colL } });
                else if (isChordLine(line)) textObjs.push(...createPptxLine(line, arr[i+1] || "", align, colC));
                else textObjs.push({ text: (line || " ") + "\n", options: { fontSize: SIZE_LYRIC, color: colL } });
            });
            slide.addText(textObjs, { x: "5%", y: yLyrics + "%", w: "90%", h: "70%", fontFace: MAIN_FONT, valign: 'middle', align, margin: 0, lineSpacing: SIZE_LYRIC * 1.1 });
            slide.addText(songCopy, { x: "5%", y: yCopy + "%", w: "90%", fontSize: SIZE_COPY, fontFace: MAIN_FONT, italic: true, align, color: colCp, margin: 0, valign: 'top' });
        });
    }

    // --- 8. PPTX IMPORT (FIXED FIRST SONG OVERWRITE) ---
    document.getElementById('importPptx').onchange = async (e) => {
        const file = e.target.files[0]; if (!file) return;
        try {
            const zip = await JSZip.loadAsync(file); const parser = new DOMParser();
            setlist = []; 
            lastLoadedIndex = -1; // SAFETY LOCK: prevents saving the current UI state during import

            const slideFiles = Object.keys(zip.files).filter(n => n.toLowerCase().includes('ppt/slides/slide')).sort((a,b) => {
                const numA = (a.match(/\d+/) || [0])[0], numB = (b.match(/\d+/) || [0])[0];
                return parseInt(numA) - parseInt(numB);
            });

            let currentSong = null;
            for (let i = 0; i < slideFiles.length; i++) {
                const xml = await zip.file(slideFiles[i]).async("text"), doc = parser.parseFromString(xml, "application/xml");
                let sT = "", sC = ""; const paras = doc.getElementsByTagNameNS("*", "p");
                for (let p of paras) {
                    let pTxt = ""; let isT = false, isC = false;
                    for (let r of p.getElementsByTagNameNS("*", "r")) {
                        const rPr = r.getElementsByTagNameNS("*", "rPr")[0], t = r.getElementsByTagNameNS("*", "t")[0];
                        if (t) pTxt += t.textContent;
                        if (rPr) {
                            const sz = parseInt(rPr.getAttribute("sz") || "0"), b = rPr.getAttribute("b") === "1" || rPr.getElementsByTagNameNS("*", "b").length > 0, it = rPr.getAttribute("i") === "1" || rPr.getElementsByTagNameNS("*", "i").length > 0;
                            if (sz >= 3200 && b) isT = true;
                            if (sz <= 1600 && it) isC = true;
                        }
                    }
                    if (isT && !sT) sT = pTxt.trim(); if (isC && !sC) sC = pTxt.trim();
                }

                if (sT && (!currentSong || sT !== currentSong.title)) {
                    if (currentSong) setlist.push(currentSong);
                    currentSong = { title: sT, copy: sC, lyrics: "" };
                } else if (currentSong && sC && !currentSong.copy) { currentSong.copy = sC; }
                else if (i === 0 && !currentSong) { currentSong = { title: sT || "Imported Song", copy: sC || "", lyrics: "" }; }

                const numMatch = slideFiles[i].match(/\d+/);
                if (numMatch && currentSong) {
                    const nF = zip.file(`ppt/notesSlides/notesSlide${numMatch[0]}.xml`);
                    if (nF) {
                        const nXml = await nF.async("text"), nDoc = parser.parseFromString(nXml, "application/xml");
                        for (let np of nDoc.getElementsByTagNameNS("*", "p")) {
                            let line = ""; for (let nt of np.getElementsByTagNameNS("*", "t")) line += nt.textContent;
                            if (line.trim() && !/^\d+$/.test(line.trim())) currentSong.lyrics += line + "\n";
                        }
                        currentSong.lyrics += "\n";
                    }
                }
            }
            if (currentSong) setlist.push(currentSong);

            dropdown.innerHTML = "";
            if (setlist.length > 1) {
                const placeholder = document.createElement('option'); placeholder.value = ""; placeholder.innerText = `-- Select Song (${setlist.length}) --`; dropdown.appendChild(placeholder);
                setlist.forEach((s, idx) => { const opt = document.createElement('option'); opt.value = idx; opt.innerText = s.title; dropdown.appendChild(opt); });
                dropdown.style.display = "block"; setlistBtn.style.display = "block";
            } else { dropdown.style.display = "none"; setlistBtn.style.display = "none"; }
            
            // LOAD FIRST SONG IMMEDIATELY
            if (setlist.length > 0) loadSong(0);

        } catch (err) { alert("Import Failed: " + err.message); }
    };

    // --- 9. UI EVENT LISTENERS & HELPERS ---
    function saveCurrentState() {
        // Only save if a song from the setlist is actually active
        if (lastLoadedIndex !== -1 && setlist[lastLoadedIndex]) {
            setlist[lastLoadedIndex].title = document.getElementById('valTitle').value;
            setlist[lastLoadedIndex].copy = document.getElementById('valCopy').value;
            setlist[lastLoadedIndex].lyrics = document.getElementById('valLyrics').value;
        }
    }

    function loadSong(idx) {
        saveCurrentState(); 
        lastLoadedIndex = idx;
        const s = setlist[idx]; if (!s) return;
        document.getElementById('valTitle').value = s.title;
        document.getElementById('valCopy').value = s.copy;
        document.getElementById('valLyrics').value = s.lyrics.trim();
        currentPreviewIndex = 0; currentShift = 0;
        document.getElementById('keyShift').innerText = `Shift: 0`;
        updatePreview();
    }

    dropdown.onchange = (e) => { if(e.target.value !== "") loadSong(parseInt(e.target.value)); };
    document.querySelectorAll('input, textarea, select').forEach(el => el.addEventListener('input', updatePreview));
    document.getElementById('btnUp').onclick = () => { lyricInput.value = lyricInput.value.split('\n').map(l => isChordLine(l) ? l.replace(/\S+/g, c => transposeChord(c, 1)) : l).join('\n'); currentShift++; document.getElementById('keyShift').innerText = `Shift: ${currentShift}`; updatePreview(); };
    document.getElementById('btnDown').onclick = () => { lyricInput.value = lyricInput.value.split('\n').map(l => isChordLine(l) ? l.replace(/\S+/g, c => transposeChord(c, -1)) : l).join('\n'); currentShift--; document.getElementById('keyShift').innerText = `Shift: ${currentShift}`; updatePreview(); };
    document.getElementById('downloadBtn').onclick = downloadPptx;
    document.getElementById('downloadSetlistBtn').onclick = downloadFullSetlist;
    document.getElementById('nextSlide').onclick = () => { currentPreviewIndex++; updatePreview(); };
    document.getElementById('prevSlide').onclick = () => { if(currentPreviewIndex>0) { currentPreviewIndex--; updatePreview(); }};
    window.addEventListener('keydown', (e) => { if (document.activeElement.tagName === 'TEXTAREA' || document.activeElement.tagName === 'INPUT') return; if (e.key === "ArrowRight" || e.key === "PageDown" || e.key === " ") { e.preventDefault(); currentPreviewIndex++; updatePreview(); } else if (e.key === "ArrowLeft" || e.key === "PageUp") { if (currentPreviewIndex > 0) { currentPreviewIndex--; updatePreview(); } } });

    // --- 10. GALLERY ROW-BY-ROW LOGIC ---
    const bgSelectorDiv = document.getElementById('bgSelector');
    function lockGalleryToSingleRow() { const ft = bgSelectorDiv.querySelector('.bg-thumb'); if (ft && ft.offsetHeight > 0) bgSelectorDiv.style.height = ft.offsetHeight + "px"; }
    document.getElementById('btnBgDown').onclick = () => { const ft = bgSelectorDiv.querySelector('.bg-thumb'); if (ft) bgSelectorDiv.scrollBy({ top: ft.offsetHeight + 10, behavior: 'smooth' }); };
    document.getElementById('btnBgUp').onclick = () => { const ft = bgSelectorDiv.querySelector('.bg-thumb'); if (ft) bgSelectorDiv.scrollBy({ top: -(ft.offsetHeight + 10), behavior: 'smooth' }); };
    window.addEventListener('resize', () => { updatePreview(); lockGalleryToSingleRow(); });
    updatePreview(); setTimeout(lockGalleryToSingleRow, 300);
});
