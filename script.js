document.addEventListener('DOMContentLoaded', () => {
    const inputs = document.querySelectorAll('input, textarea, select, [type="range"]');
    const downloadBtn = document.getElementById('downloadBtn');
    const bgSelector = document.getElementById('bgSelector');

    // CONFIG - Shared between Preview and PPTX
    const VERSION = "1.1.0"; 
    const MAIN_FONT = "Times New Roman";
    const SIZE_TITLE = 32;
    const SIZE_LYRIC = 24;
    const SIZE_CHORD = 14;
    const SIZE_COPY = 14;
    
    // Sync Factors
    const PT_TO_PX = 96 / 72; // Browser (96dpi) vs PPTX (72pt/inch)
    const LINE_HEIGHT_MULT = 1.15; // Prevents vertical drift

    const versionDisplay = document.getElementById('appVersion');
    if (versionDisplay) versionDisplay.innerText = `v${VERSION}`;

    const bgOptions = [
        { name: 'Plain', path: '' },
        { name: 'Modern', path: 'assets/bg-modern.png' },
        { name: 'Linen', path: 'assets/bg-linen.png' },
        { name: 'Soft', path: 'assets/bg-soft.png' }
    ];

    let selectedBgPath = "";

    // Generate Background Thumbnails
    bgOptions.forEach((opt, i) => {
        const thumb = document.createElement('div');
        thumb.className = `bg-thumb ${i === 0 ? 'active' : ''}`;
        if(opt.path) thumb.style.backgroundImage = `url(${opt.path})`;
        thumb.onclick = () => {
            document.querySelectorAll('.bg-thumb').forEach(t => t.classList.remove('active'));
            thumb.classList.add('active');
            selectedBgPath = opt.path;
            updatePreview();
        };
        bgSelector.appendChild(thumb);
    });

    function updatePreview() {
        const mock = document.getElementById('slideMock');
        
        // Calibration: 960px is the internal logical width of a 16:9 PPTX slide at 96 DPI
        const ratio = mock.offsetWidth / 960; 
        const scale = ratio * PT_TO_PX;

        const title = document.getElementById('valTitle').value;
        const lyrics = document.getElementById('valLyrics').value;
        const copy = document.getElementById('valCopy').value;
        const align = document.getElementById('slideAlign').value;

        mock.style.backgroundImage = selectedBgPath ? `url(${selectedBgPath})` : 'none';

        // --- TITLE PREVIEW ---
        const pt = document.getElementById('prevTitle');
        pt.innerText = title;
        pt.style.top = document.getElementById('yTitle').value + "%";
        pt.style.textAlign = align;
        pt.style.fontSize = (SIZE_TITLE * scale) + "px"; 
        pt.style.fontFamily = MAIN_FONT;
        pt.style.fontWeight = "bold";
        pt.style.lineHeight = LINE_HEIGHT_MULT;

        // --- COPYRIGHT PREVIEW ---
        const pc = document.getElementById('prevCopy');
        pc.innerText = copy;
        pc.style.top = document.getElementById('yCopy').value + "%";
        pc.style.textAlign = align;
        pc.style.fontSize = (SIZE_COPY * scale) + "px";
        pc.style.fontFamily = MAIN_FONT;
        pc.style.fontStyle = "italic";

        // --- LYRICS PREVIEW ---
        const pl = document.getElementById('prevLyrics');
        pl.style.top = document.getElementById('yLyrics').value + "%";
        pl.style.textAlign = align;
        pl.style.fontFamily = MAIN_FONT;
        
        // Split and preview only the first section
        const firstSection = lyrics.split(/\n?\s*(?=\[)/)[0] || "";
        pl.innerHTML = ""; 
        
        const lines = firstSection.split('\n');
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            const nextLine = lines[i+1] || "";
            const lineDiv = document.createElement('div');
            lineDiv.style.whiteSpace = "pre";
            lineDiv.style.lineHeight = LINE_HEIGHT_MULT;

            if (isChordLine(line)) {
                lineDiv.style.fontSize = (SIZE_CHORD * scale) + "px";
                lineDiv.style.color = "#808080";
                // Chords need to be spaced based on the Lyric font size
                lineDiv.innerHTML = createHtmlGhostLine(line, nextLine, scale);
            } else {
                lineDiv.style.fontSize = (SIZE_LYRIC * scale) + "px";
                lineDiv.style.color = "#000";
                lineDiv.innerText = line || " "; 
            }
            pl.appendChild(lineDiv);
        }
    }

    inputs.forEach(input => input.addEventListener('input', updatePreview));
    window.addEventListener('resize', updatePreview);

    // --- PPTX GENERATION ---
    downloadBtn.onclick = async () => {
        const pres = new PptxGenJS();
        pres.layout = 'LAYOUT_16x9';
        const align = document.getElementById('slideAlign').value;
        const sections = document.getElementById('valLyrics').value.split(/\n?\s*(?=\[)/).filter(s => s.trim());

        sections.forEach(section => {
            let slide = pres.addSlide();
            slide.background = selectedBgPath ? { path: selectedBgPath } : { fill: "FFFFFF" };

            // Title
            slide.addText(document.getElementById('valTitle').value, {
                x: "5%", y: document.getElementById('yTitle').value + "%", w: "90%",
                fontSize: SIZE_TITLE, color: "000000", fontFace: MAIN_FONT, bold: true, 
                align: align, valign: 'top', margin: 0
            });

            // Lyrics & Chords
            const lines = section.split('\n');
            let textObjects = [];
            for (let i = 0; i < lines.length; i++) {
                const line = lines[i];
                const nextLine = lines[i+1] || "";

                if (isChordLine(line)) {
                    textObjects.push(...createPptxGhostLine(line, nextLine));
                } else {
                    textObjects.push({ 
                        text: (line || " ") + "\n", 
                        options: { color: "000000", fontSize: SIZE_LYRIC, fontFace: MAIN_FONT } 
                    });
                }
            }

            slide.addText(textObjects, {
                x: "5%", y: document.getElementById('yLyrics').value + "%", w: "90%", h: "70%",
                fontFace: MAIN_FONT, valign: 'top', align: align, 
                lineSpacing: SIZE_LYRIC * LINE_HEIGHT_MULT, 
                margin: 0
            });

            // Copyright
            slide.addText(document.getElementById('valCopy').value, {
                x: "5%", y: document.getElementById('yCopy').value + "%", w: "90%",
                fontSize: SIZE_COPY, color: "000000", italic: true, fontFace: MAIN_FONT, 
                align: align, valign: 'top', margin: 0
            });
        });

        await pres.writeFile({ fileName: "Song_Slides.pptx" });
    };

    function isChordLine(str) {
        if (!str.trim()) return false;
        const chordChars = str.replace(/[A-G]|[m|maj|dim|aug|sus|2|4|5|7|9]|#|b|\s|\/|v|i|\[|\]/gi, "");
        return chordChars.length === 0;
    }

    function createHtmlGhostLine(chords, lyrics, currentScale) {
        let html = "", lastIdx = 0, re = /\S+/g, match;
        while ((match = re.exec(chords)) !== null) {
            let spacerText = lyrics.substring(lastIdx, match.index);
            if (spacerText) {
                // Invisible text pushes the chord to the right letter
                html += `<span style="visibility:hidden; font-size:${SIZE_LYRIC * currentScale}px">${spacerText}</span>`;
            }
            html += `<span>${match[0]}</span>`;
            lastIdx = match.index + match[0].length;
        }
        return html || " ";
    }

    function createPptxGhostLine(chords, lyrics) {
        let result = [], lastIdx = 0, re = /\S+/g, match;
        while ((match = re.exec(chords)) !== null) {
            let spacer = lyrics.substring(lastIdx, match.index);
            if (spacer) {
                // transparency: 100 makes the text invisible in PPTX
                result.push({ 
                    text: spacer, 
                    options: { transparency: 100, fontSize: SIZE_LYRIC, fontFace: MAIN_FONT } 
                });
            }
            result.push({ 
                text: match[0], 
                options: { color: "808080", fontSize: SIZE_CHORD, fontFace: MAIN_FONT } 
            });
            lastIdx = match.index + match[0].length;
        }
        result.push({ text: "\n" });
        return result;
    }

    updatePreview();
});
