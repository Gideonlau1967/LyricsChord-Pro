document.addEventListener('DOMContentLoaded', () => {
    // --- 1. STATE & CONFIG ---
    const MAIN_FONT = "Times New Roman";
    const SIZE_TITLE = 32, SIZE_LYRIC = 24, SIZE_CHORD = 14, SIZE_SECTION = 16, SIZE_COPY = 14;
    const PT_TO_PX = 96 / 72;
    const SCALE = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B'];
    const FLAT_MAP = { 'Db': 'C#', 'Eb': 'D#', 'Gb': 'F#', 'Ab': 'G#', 'Bb': 'A#' };

    let setlist = [{ title: "Song 1", copy: "© Copyright Info", lyrics: "[Verse 1]\nG           C\nLyrics here...", shift: 0 }];
    let currentSongIdx = 0;
    let currentPreviewIndex = 0;
    let selectedBgPath = "assets/bg-default.png";

    const lyricInput = document.getElementById('valLyrics');
    const bgSelector = document.getElementById('bgSelector');

    // --- 2. SETLIST MANAGEMENT ---
    function renderSetlist() {
        const container = document.getElementById('setlistItems');
        container.innerHTML = "";
        setlist.forEach((song, idx) => {
            const tab = document.createElement('div');
            tab.className = `song-tab ${idx === currentSongIdx ? 'active' : ''}`;
            tab.innerText = song.title || `Song ${idx + 1}`;
            tab.onclick = () => switchSong(idx);
            container.appendChild(tab);
        });
        document.getElementById('songCount').innerText = setlist.length;
    }

    function switchSong(idx) {
        // Save current editor state to memory
        setlist[currentSongIdx].title = document.getElementById('valTitle').value;
        setlist[currentSongIdx].copy = document.getElementById('valCopy').value;
        setlist[currentSongIdx].lyrics = lyricInput.value;

        currentSongIdx = idx;
        currentPreviewIndex = 0;

        // Load new song state to UI
        const song = setlist[idx];
        document.getElementById('valTitle').value = song.title;
        document.getElementById('valCopy').value = song.copy;
        lyricInput.value = song.lyrics;
        document.getElementById('keyShift').innerText = `Shift: ${song.shift || 0}`;

        renderSetlist();
        updatePreview();
    }

    document.getElementById('addSongBtn').onclick = () => {
        setlist.push({ title: "New Song", copy: "© Copyright", lyrics: "[Verse 1]\nLyrics here...", shift: 0 });
        switchSong(setlist.length - 1);
    };

    document.getElementById('delSongBtn').onclick = () => {
        if (setlist.length > 1) {
            setlist.splice(currentSongIdx, 1);
            currentSongIdx = Math.max(0, currentSongIdx - 1);
            switchSong(currentSongIdx);
        }
    };

    // --- 3. IMPORT LOGIC (MULTI-SONG DETECTION) ---
    document.getElementById('importPptx').onchange = async (e) => {
        const file = e.target.files[0]; if (!file) return;
        const zip = await JSZip.loadAsync(file);
        const parser = new DOMParser();
        let newSetlist = [];
        let tempSong = null;

        const slidePaths = Object.keys(zip.files).filter(n => n.startsWith('ppt/slides/slide')).sort((a,b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));

        for (let path of slidePaths) {
            const num = path.match(/\d+/)[0];
            const slideXml = await zip.file(path).async("text");
            const slideDoc = parser.parseFromString(slideXml, "application/xml");

            // Identify Title Slides (Large Bold Text)
            let slideTitle = "";
            let isNewSong = false;
            for (let sp of slideDoc.getElementsByTagNameNS("*", "sp")) {
                let txt = "", sz = 0, bold = false;
                for (let r of sp.getElementsByTagNameNS("*", "r")) {
                    const t = r.getElementsByTagNameNS("*", "t")[0];
                    if (t) txt += t.textContent;
                    const rPr = r.getElementsByTagNameNS("*", "rPr")[0];
                    if (rPr) {
                        sz = parseInt(rPr.getAttribute("sz")) / 100;
                        bold = rPr.getAttribute("b") === "1";
                    }
                }
                if (sz >= 30 && bold && txt.trim()) { slideTitle = txt; isNewSong = true; break; }
            }

            // Extract Notes
            const noteFile = `ppt/notesSlides/notesSlide${num}.xml`;
            let notesText = "";
            if (zip.file(noteFile)) {
                const noteXml = await zip.file(noteFile).async("text");
                const paras = parser.parseFromString(noteXml, "application/xml").getElementsByTagNameNS("*", "p");
                for (let p of paras) {
                    let line = "";
                    for (let t of p.getElementsByTagNameNS("*", "t")) line += t.textContent;
                    if (line.trim() && !/^\d+$/.test(line.trim())) notesText += line + "\n";
                }
            }

            if (isNewSong) {
                if (tempSong) newSetlist.push(tempSong);
                tempSong = { title: slideTitle, copy: "© Imported", lyrics: "", shift: 0 };
            } else if (tempSong && notesText) {
                tempSong.lyrics += notesText + "\n";
            }
        }
        if (tempSong) newSetlist.push(tempSong);
        if (newSetlist.length > 0) { setlist = newSetlist; switchSong(0); }
    };

    // --- 4. PREVIEW RENDER ---
    function updatePreview() {
        const mock = document.getElementById('slideMock');
        if (!mock) return;
        const currentSong = setlist[currentSongIdx];
        const sections = (lyricInput.value || "").split(/(?=\[)/).filter(s => s.trim());
        
        if (currentPreviewIndex >= sections.length) currentPreviewIndex = Math.max(0, sections.length - 1);
        document.getElementById('slideIndicator').innerText = `Slide ${sections.length > 0 ? currentPreviewIndex + 1 : 0} / ${sections.length}`;
        
        const rect = mock.getBoundingClientRect();
        const scale = (rect.width / 960) * PT_TO_PX;
        mock.style.backgroundImage = `url(${selectedBgPath})`;

        const colT = document.getElementById('colTitle').value, colL = document.getElementById('colLyrics').value,
              colC = document.getElementById('colChords').value, colCp = document.getElementById('colCopy').value,
              align = document.getElementById('slideAlign').value;

        const pt = document.getElementById('prevTitle'), pc = document.getElementById('prevCopy'), pl = document.getElementById('prevLyrics');
        [pt, pc, pl].forEach(el => { el.style.textAlign = align; el.style.fontFamily = MAIN_FONT; el.style.padding = "0"; });

        pt.innerText = document.getElementById('valTitle').value;
        pt.style.top = document.getElementById('yTitle').value + "%";
        pt.style.fontSize = (SIZE_TITLE * scale) + "px"; pt.style.fontWeight = "bold"; pt.style.color = colT;

        pc.innerText = document.getElementById('valCopy').value;
        pc.style.top = document.getElementById('yCopy').value + "%";
        pc.style.fontSize = (SIZE_COPY * scale) + "px"; pc.style.fontStyle = "italic"; pc.style.color = colCp;

        pl.style.top = document.getElementById('yLyrics').value + "%";
        pl.style.height = "70%"; pl.style.display = "flex"; pl.style.flexDirection = "column"; pl.style.justifyContent = "center";
        
        pl.innerHTML = ""; const inner = document.createElement('div'); inner.style.width = "100%";
        const activeContent = (sections[currentPreviewIndex] || "").trim();
        
        activeContent.split('\n').forEach((line, i, arr) => {
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
        let h = ""; const maxLen = Math.max(chords.length, lyrics.length);
        const lSize = SIZE_LYRIC * scale, cSize = SIZE_CHORD * scale;
        for (let i = 0; i < maxLen; i++) {
            const charL = lyrics[i] || " ", charC = chords[i] || "";
            h += `<span style="position:relative; display:inline-block; font-size:${lSize}px; color:transparent;">${charL===" "?"\u00A0":charL}`;
            if (charC.trim() !== "") {
                let full = charC, j = i + 1;
                while (j < chords.length && chords[j] !== " ") { full += chords[j]; chords = chords.substring(0,j)+" "+chords.substring(j+1); j++; }
                h += `<span style="position:absolute; left:0; bottom:0; font-size:${cSize}px; color:${cCol}; visibility:visible; font-family:serif; white-space:nowrap; transform:translateY(-20%); font-weight:600;">${full}</span>`;
            }
            h += `</span>`;
        }
        return h;
    }

    // --- 5. TRANSPOSE LOGIC (PER SONG) ---
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

    document.getElementById('btnUp').onclick = () => {
        lyricInput.value = lyricInput.value.split('\n').map(l => isChordLine(l) ? l.replace(/\S+/g, c => transposeChord(c, 1)) : l).join('\n');
        setlist[currentSongIdx].shift++;
        document.getElementById('keyShift').innerText = `Shift: ${setlist[currentSongIdx].shift}`;
        updatePreview();
    };

    document.getElementById('btnDown').onclick = () => {
        lyricInput.value = lyricInput.value.split('\n').map(l => isChordLine(l) ? l.replace(/\S+/g, c => transposeChord(c, -1)) : l).join('\n');
        setlist[currentSongIdx].shift--;
        document.getElementById('keyShift').innerText = `Shift: ${setlist[currentSongIdx].shift}`;
        updatePreview();
    };

    // --- 6. EXPORT LOGIC (GLOBAL) ---
    async function downloadPptx() {
        setlist[currentSongIdx].title = document.getElementById('valTitle').value;
        setlist[currentSongIdx].lyrics = lyricInput.value;
        
        const pres = new PptxGenJS(); pres.layout = 'LAYOUT_16x9';
        const colT = document.getElementById('colTitle').value.replace('#',''), colL = document.getElementById('colLyrics').value.replace('#',''), colC = document.getElementById('colChords').value.replace('#',''), colCp = document.getElementById('colCopy').value.replace('#','');
        const align = document.getElementById('slideAlign').value;

        setlist.forEach(song => {
            const sections = song.lyrics.split(/(?=\[)/).filter(s => s.trim());
            sections.forEach(section => {
                let slide = pres.addSlide(); slide.background = { path: selectedBgPath };
                slide.addNotes(section);
                slide.addText(song.title, { x:"5%", y:document.getElementById('yTitle').value+"%", w:"90%", fontSize:SIZE_TITLE, fontFace:MAIN_FONT, bold:true, align, color:colT, margin:0, valign:'top' });
                
                let textObjs = [];
                section.replace(/^[\n\r]+|[\n\r]+$/g, '').split('\n').forEach((line, i, arr) => {
                    if (line.trim().startsWith('[') && line.trim().endsWith(']')) textObjs.push({ text: line+"\n", options: { fontSize:SIZE_SECTION, bold:true, color:colL } });
                    else if (isChordLine(line)) {
                        let r = []; const len = align === 'center' ? Math.max(line.length, (arr[i+1]||"").length) : line.length;
                        for (let k = 0; k < len; k++) {
                            const c = line[k] || " ", l = (arr[i+1]||"")[k] || " ";
                            if (c !== " ") r.push({ text: c, options: { color: colC, fontSize: SIZE_CHORD } });
                            else r.push({ text: l===""?" ":l, options: { transparency: 100, fontSize: SIZE_LYRIC } });
                        }
                        r.push({ text: "\n" }); textObjs.push(...r);
                    } else textObjs.push({ text: (line||" ")+"\n", options: { fontSize:SIZE_LYRIC, color:colL } });
                });
                slide.addText(textObjs, { x:"5%", y:document.getElementById('yLyrics').value+"%", w:"90%", h:"70%", fontFace:MAIN_FONT, valign:'middle', align, margin:0, lineSpacing: SIZE_LYRIC * 1.1 });
                slide.addText(song.copy, { x:"5%", y:document.getElementById('yCopy').value+"%", w:"90%", fontSize:SIZE_COPY, fontFace:MAIN_FONT, italic:true, align, color:colCp, margin:0, valign:'top' });
            });
        });
        pres.writeFile({ fileName: `Setlist.pptx` });
    }

    // --- 7. UTILS & UI ---
    bgOptions.forEach(opt => {
        const thumb = document.createElement('div');
        thumb.className = `bg-thumb ${opt.path === selectedBgPath ? 'active' : ''}`;
        thumb.style.backgroundImage = `url(${opt.path})`;
        thumb.onclick = () => {
            document.querySelectorAll('.bg-thumb').forEach(t => t.classList.remove('active'));
            thumb.classList.add('active'); selectedBgPath = opt.path; updatePreview();
        };
        bgSelector.appendChild(thumb);
    });

    const fsBtn = document.getElementById('fullscreenBtn');
    const fsWrapper = document.getElementById('fullscreenWrapper');
    fsBtn.onclick = () => { if (!document.fullscreenElement) fsWrapper.requestFullscreen(); else document.exitFullscreen(); };
    document.addEventListener('fullscreenchange', () => setTimeout(updatePreview, 100));

    window.addEventListener('keydown', (e) => {
        if (document.activeElement.tagName === 'TEXTAREA' || document.activeElement.tagName === 'INPUT') return;
        if (e.key === "ArrowRight" || e.key === " ") { currentPreviewIndex++; updatePreview(); }
        else if (e.key === "ArrowLeft" && currentPreviewIndex > 0) { currentPreviewIndex--; updatePreview(); }
    });

    document.getElementById('nextSlide').onclick = () => { currentPreviewIndex++; updatePreview(); };
    document.getElementById('prevSlide').onclick = () => { if(currentPreviewIndex > 0) { currentPreviewIndex--; updatePreview(); }};
    document.getElementById('downloadBtn').onclick = downloadPptx;

    function lockGallery() {
        const first = bgSelector.querySelector('.bg-thumb');
        if (first) bgSelector.style.height = first.offsetHeight + "px";
    }

    window.onresize = () => { updatePreview(); lockGallery(); };
    setTimeout(lockGallery, 500);
    renderSetlist();
    updatePreview();
});
