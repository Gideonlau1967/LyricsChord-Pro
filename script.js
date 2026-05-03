/**
 * SONG STUDIO PRO V2.0 - CORE SCRIPT
 * 
 * MAIN FUNCTIONALITIES:
 * 1. UI SCALING: Dynamic calculation of font sizes based on responsive container widths.
 * 2. TRANSPOSE ENGINE: Mathematical shifting of musical keys (A-G) using 12-semitone scale.
 * 3. CHORD GROUPING: Intelligent preview rendering that keeps complex chords (e.g., F#m) 
 *    as single units above lyrics.
 * 4. PPTX IMPORT (NOTES): Extracts text from "ppt/notesSlides/" using JSZip and DOMParser.
 * 5. PPTX EXPORT: Generates high-quality PowerPoint decks with PptxGenJS.
 * 6. FULLSCREEN API: Presentation mode toggle with automatic font re-scaling.
 * 7. BG GALLERY: Thumbnail-based background selection with smooth-scroll navigation.
 */

document.addEventListener('DOMContentLoaded', () => {
    // --- 1. CONFIGURATION & STATE ---
    const VERSION = "2.0-PRO";
    const MAIN_FONT = "Times New Roman";
    const SIZE_TITLE = 32, SIZE_LYRIC = 24, SIZE_CHORD = 14, SIZE_SECTION = 16, SIZE_COPY = 14;
    const PT_TO_PX = 96 / 72; // Conversion factor for accurate point-to-pixel rendering
    
    let currentPreviewIndex = 0;
    let currentShift = 0;
    let selectedBgPath = "assets/bg-default.png";
    let importedLibrary = []; // NEW: Stores setlist songs

    const SCALE = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B'];
    const FLAT_MAP = { 'Db': 'C#', 'Eb': 'D#', 'Gb': 'F#', 'Ab': 'G#', 'Bb': 'A#' };

    const lyricInput = document.getElementById('valLyrics');
    const bgSelector = document.getElementById('bgSelector');
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
        // Regex filters lines containing ONLY chords, symbols, and whitespace
        return s.replace(/[A-G]|[m|maj|min|dim|aug|sus|2|4|5|7|9]|#|b|\s|\/|v|i|\[|\]/gi, "").length === 0;
    }

    // --- 4. PREVIEW RENDERING ---
    function updatePreview() {
        const mock = document.getElementById('slideMock');
        if (!mock) return;

        const sections = lyricInput.value.split(/(?=\[)/).filter(s => s.trim());
        if (currentPreviewIndex >= sections.length) currentPreviewIndex = Math.max(0, sections.length - 1);
        
        document.getElementById('slideIndicator').innerText = `Slide ${sections.length > 0 ? currentPreviewIndex + 1 : 0} / ${sections.length}`;
        
        // 1. GET ACCURATE SCALE
        const rect = mock.getBoundingClientRect();
        const scale = (rect.width / 960) * PT_TO_PX;
        mock.style.backgroundImage = `url(${selectedBgPath})`;

        const colT = document.getElementById('colTitle').value;
        const colL = document.getElementById('colLyrics').value;
        const colC = document.getElementById('colChords').value;
        const colCp = document.getElementById('colCopy').value;
        const align = document.getElementById('slideAlign').value;

        const pt = document.getElementById('prevTitle'), pc = document.getElementById('prevCopy'), pl = document.getElementById('prevLyrics');
        [pt, pc, pl].forEach(el => { 
            el.style.textAlign = align; 
            el.style.fontFamily = MAIN_FONT; 
            el.style.margin = "0"; // Sync with PPTX margin:0
            el.style.padding = "0";
        });

        // 2. POSITION TITLE & COPYRIGHT (Top Aligned)
        pt.innerText = document.getElementById('valTitle').value;
        pt.style.top = document.getElementById('yTitle').value + "%";
        pt.style.fontSize = (SIZE_TITLE * scale) + "px"; pt.style.fontWeight = "bold"; pt.style.color = colT;

        pc.innerText = document.getElementById('valCopy').value;
        pc.style.top = document.getElementById('yCopy').value + "%";
        pc.style.fontSize = (SIZE_COPY * scale) + "px"; pc.style.fontStyle = "italic"; pc.style.color = colCp;

        // 3. POSITION LYRIC BLOCK (Middle Aligned inside 70% height)
        pl.style.top = document.getElementById('yLyrics').value + "%";
        pl.style.height = "70%"; // Matches PPTX h: "70%"
        pl.style.display = "flex"; 
        pl.style.flexDirection = "column"; 
        pl.style.justifyContent = "center"; // Mimics PPTX valign: 'middle'
        
        const active = (sections[currentPreviewIndex] || "").replace(/^[\n\r]+|[\n\r]+$/g, '');
        pl.innerHTML = ""; 
        const inner = document.createElement('div'); 
        inner.style.width = "100%";
        
        active.split('\n').forEach((line, i, arr) => {
            const div = document.createElement('div'); 
            div.style.whiteSpace = "pre";
            div.style.lineHeight = "1.1"; // Sync with PPTX lineSpacing

            if (line.trim().startsWith('[') && line.trim().endsWith(']')) {
                div.style.fontSize = (SIZE_SECTION * scale) + "px"; 
                div.style.fontWeight = "bold"; 
                div.innerText = line; 
                div.style.color = colL;
            } else if (isChordLine(line)) {
                div.style.fontSize = (SIZE_CHORD * scale) + "px"; 
                div.style.lineHeight = "0.7"; 
                div.style.marginBottom = (parseInt(document.getElementById('chordGap').value) * scale) + "px";
                div.innerHTML = createHtmlLine(line, arr[i+1] || "", scale, align, colC);
            } else {
                div.style.fontSize = (SIZE_LYRIC * scale) + "px"; 
                div.innerText = line || " "; 
                div.style.color = colL;
            }
            inner.appendChild(div);
        });
        pl.appendChild(inner);
    }

    // --- 5. CHORD ALIGNMENT ENGINE (PREVIEW ONLY) ---
    function createHtmlLine(chords, lyrics, scale, align, cCol) {
        let h = ""; 
        const maxLen = Math.max(chords.length, lyrics.length);
        const lSize = SIZE_LYRIC * scale;
        const cSize = SIZE_CHORD * scale;

        for (let i = 0; i < maxLen; i++) {
            const charL = lyrics[i] || " ";
            const charC = chords[i] || "";
            
            h += `<span style="position:relative; display:inline-block; font-size:${lSize}px; color:transparent;">`;
            h += charL === " " ? "\u00A0" : charL;
            
            if (charC.trim() !== "") {
                // Grouping Logic: Detect if next characters belong to the same chord string
                let fullChord = charC;
                let j = i + 1;
                while (j < chords.length && chords[j] !== " ") {
                    fullChord += chords[j];
                    chords = chords.substring(0, j) + " " + chords.substring(j + 1); // Mark character as used
                    j++;
                }
                h += `<span style="position:absolute; left:0; bottom:0; font-size:${cSize}px; color:${cCol}; visibility:visible; font-family:serif; white-space:nowrap; transform:translateY(-20%); font-weight:600;">${fullChord}</span>`;
            }
            h += `</span>`;
        }
        return h;
    }

    // --- 6. FULLSCREEN API ---
    const fsBtn = document.getElementById('fullscreenBtn');
    const fsWrapper = document.getElementById('fullscreenWrapper');
    if (fsBtn) {
        fsBtn.onclick = () => {
            if (!document.fullscreenElement) fsWrapper.requestFullscreen().catch(console.error);
            else document.exitFullscreen();
        };
    }
    document.addEventListener('fullscreenchange', () => setTimeout(updatePreview, 100));

    // --- 7. PPTX EXPORT ENGINE ---
    async function downloadPptx() {
        const pres = new PptxGenJS(); 
        pres.layout = 'LAYOUT_16x9';
        
        const songT = document.getElementById('valTitle').value.trim() || "Song";
        const align = document.getElementById('slideAlign').value;
        const colT = document.getElementById('colTitle').value.replace('#','');
        const colL = document.getElementById('colLyrics').value.replace('#','');
        const colC = document.getElementById('colChords').value.replace('#','');
        const colCp = document.getElementById('colCopy').value.replace('#','');

        const yTitle = document.getElementById('yTitle').value;
        const yLyrics = document.getElementById('yLyrics').value;
        const yCopy = document.getElementById('yCopy').value;

        lyricInput.value.split(/(?=\[)/).filter(s => s.trim()).forEach(section => {
            let slide = pres.addSlide(); 
            slide.background = { path: selectedBgPath };
            slide.addNotes(section);
            
            // TITLE - Explicitly set margin 0 and valign top
            slide.addText(songT, { 
                x: "5%", y: yTitle + "%", w: "90%", 
                fontSize: SIZE_TITLE, fontFace: MAIN_FONT, bold: true, 
                align, color: colT, margin: 0, valign: 'top' 
            });

            // LYRICS & CHORDS
            let textObjs = [];
            section.replace(/^[\n\r]+|[\n\r]+$/g, '').split('\n').forEach((line, i, arr) => {
                if (line.trim().startsWith('[') && line.trim().endsWith(']')) {
                    textObjs.push({ text: line + "\n", options: { fontSize: SIZE_SECTION, bold: true, color: colL } });
                } else if (isChordLine(line)) {
                    textObjs.push(...createPptxLine(line, arr[i+1] || "", align, colC));
                } else {
                    textObjs.push({ text: (line || " ") + "\n", options: { fontSize: SIZE_LYRIC, color: colL } });
                }
            });

            // LYRIC BLOCK - Explicitly set margin 0 and valign middle
            slide.addText(textObjs, { 
                x: "5%", y: yLyrics + "%", w: "90%", h: "70%", 
                fontFace: MAIN_FONT, valign: 'middle', align, 
                margin: 0, 
                lineSpacing: SIZE_LYRIC * 1.1 // Matches CSS line-height 1.1
            });

            // COPYRIGHT - Explicitly set margin 0 and valign top
            slide.addText(document.getElementById('valCopy').value, { 
                x: "5%", y: yCopy + "%", w: "90%", 
                fontSize: SIZE_COPY, fontFace: MAIN_FONT, italic: true, 
                align, color: colCp, margin: 0, valign: 'top' 
            });
        });

        pres.writeFile({ fileName: `${songT}.pptx` });
    }

    function createPptxLine(chords, lyrics, align, cCol) {
        let r = []; 
        const len = align === 'center' ? Math.max(chords.length, lyrics.length) : chords.length;
        for (let i = 0; i < len; i++) {
            const c = chords[i] || " ", l = lyrics[i] || " ";
            if (c !== " ") r.push({ text: c, options: { color: cCol, fontSize: SIZE_CHORD } });
            else r.push({ text: l===""?" ":l, options: { transparency: 100, fontSize: SIZE_LYRIC } });
        }
        r.push({ text: "\n" }); return r;
    }

    // --- 8. PPTX IMPORT (READ NOTES AS SETLIST) ---
    document.getElementById('importPptx').onchange = async (e) => {
        const file = e.target.files[0]; if (!file) return;
        try {
            const zip = await JSZip.loadAsync(file); 
            const parser = new DOMParser();
            let rawSongs = [];

            const notes = Object.keys(zip.files)
                .filter(n => n.startsWith('ppt/notesSlides/notesSlide'))
                .sort((a,b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));

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
                
                if (noteTxt.trim()) {
                    // Start new song entry if notes look like a header or if it's the very first entry
                    if (noteTxt.includes("[Title]") || rawSongs.length === 0) {
                        rawSongs.push(noteTxt.trim());
                    } else {
                        rawSongs[rawSongs.length - 1] += "\n\n" + noteTxt.trim();
                    }
                }
            }

            importedLibrary = rawSongs.map((content, index) => {
                const firstLine = content.split('\n')[0].replace(/[\[\]]/g, '');
                return { title: firstLine || `Song ${index + 1}`, data: content };
            });

            renderSetlist();
            alert(`Imported ${importedLibrary.length} songs to setlist.`);
        } catch (err) {
            console.error("Failed to import PPTX:", err);
            alert("Error importing PPTX notes.");
        }
    };

    // --- NEW: RENDER SETLIST TRAY ---
    function renderSetlist() {
        const container = document.getElementById('setlistItems');
        if (importedLibrary.length === 0) {
            container.innerHTML = '<p class="empty-msg">No songs imported yet.</p>';
            return;
        }
        container.innerHTML = "";
        importedLibrary.forEach((song, index) => {
            const item = document.createElement('div');
            item.className = 'set-item';
            item.innerHTML = `<span>${song.title}</span><button class="st-btn-sm" style="padding:1px 5px">Load</button>`;
            item.onclick = () => {
                document.getElementById('valTitle').value = song.title;
                document.getElementById('valLyrics').value = song.data;
                updatePreview();
            };
            container.appendChild(item);
        });
    }

    document.getElementById('clearSetlist').onclick = () => {
        importedLibrary = [];
        renderSetlist();
    };

    // --- 9. UI EVENT LISTENERS ---
    document.querySelectorAll('input, textarea, select').forEach(el => el.addEventListener('input', updatePreview));

    document.getElementById('btnUp').onclick = () => { 
        lyricInput.value = lyricInput.value.split('\n').map(l => 
            isChordLine(l) ? l.replace(/\S+/g, c => transposeChord(c, 1)) : l
        ).join('\n'); 
        currentShift++; document.getElementById('keyShift').innerText = `Shift: ${currentShift}`; updatePreview(); 
    };

    document.getElementById('btnDown').onclick = () => { 
        lyricInput.value = lyricInput.value.split('\n').map(l => 
            isChordLine(l) ? l.replace(/\S+/g, c => transposeChord(c, -1)) : l
        ).join('\n'); 
        currentShift--; document.getElementById('keyShift').innerText = `Shift: ${currentShift}`; updatePreview(); 
    };

    document.getElementById('downloadBtn').onclick = downloadPptx;
    
    document.getElementById('nextSlide').onclick = () => { currentPreviewIndex++; updatePreview(); };
    document.getElementById('prevSlide').onclick = () => { if(currentPreviewIndex>0) { currentPreviewIndex--; updatePreview(); }};

    window.addEventListener('keydown', (e) => {
        // Prevent navigation if the user is currently typing in an input or textarea
        if (document.activeElement.tagName === 'TEXTAREA' || document.activeElement.tagName === 'INPUT') {
            return;
        }

        // NEXT SLIDE: Right Arrow, Page Down, or Space
        if (e.key === "ArrowRight" || e.key === "PageDown" || e.key === " ") {
            e.preventDefault(); // Stop page from jumping down on Spacebar
            currentPreviewIndex++;
            updatePreview();
        } 
        // PREVIOUS SLIDE: Left Arrow or Page Up
        else if (e.key === "ArrowLeft" || e.key === "PageUp") {
            if (currentPreviewIndex > 0) {
                currentPreviewIndex--;
                updatePreview();
            }
        }
    });
    
 // --- 10. GALLERY ROW-BY-ROW LOGIC ---
    const btnBgDown = document.getElementById('btnBgDown');
    const btnBgUp = document.getElementById('btnBgUp');

    // Function to make the gallery exactly 1 row high
    function lockGalleryToSingleRow() {
        const firstThumb = bgSelector.querySelector('.bg-thumb');
        if (firstThumb && firstThumb.offsetHeight > 0) {
            // Set the container height to exactly one thumbnail's height
            bgSelector.style.height = firstThumb.offsetHeight + "px";
        }
    }

    if (btnBgDown && btnBgUp) {
        btnBgDown.onclick = () => {
            const firstThumb = bgSelector.querySelector('.bg-thumb');
            if (firstThumb) {
                const gap = parseFloat(getComputedStyle(bgSelector).rowGap) || 0;
                const step = firstThumb.offsetHeight + gap;
                bgSelector.scrollBy({ top: step, behavior: 'smooth' });
            }
        };

        btnBgUp.onclick = () => {
            const firstThumb = bgSelector.querySelector('.bg-thumb');
            if (firstThumb) {
                const gap = parseFloat(getComputedStyle(bgSelector).rowGap) || 0;
                const step = firstThumb.offsetHeight + gap;
                bgSelector.scrollBy({ top: -step, behavior: 'smooth' });
            }
        };
    }

    // --- INITIALIZE & RESIZE ---
    // Single listener to handle all resize logic
    window.addEventListener('resize', () => {
        updatePreview();
        lockGalleryToSingleRow(); 
    });

    // Final Initialization
    updatePreview();
    
    // Use a timeout to ensure CSS and images are rendered before measuring height
    setTimeout(lockGalleryToSingleRow, 300);
});
