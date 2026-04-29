document.addEventListener('DOMContentLoaded', () => {
    const inputs = document.querySelectorAll('input, textarea, select, [type="range"]');
    const downloadBtn = document.getElementById('downloadBtn');
    const bgSelector = document.getElementById('bgSelector');

    // CONFIG
    const VERSION = "1.2.2-DEBUG"; 
    const MAIN_FONT = "Times New Roman";
    const SIZE_TITLE = 32, SIZE_LYRIC = 24, SIZE_CHORD = 14, SIZE_COPY = 14;
    const PT_TO_PX = 96 / 72; 

    const versionDisplay = document.getElementById('appVersion');
    if (versionDisplay) versionDisplay.innerText = `v${VERSION}`;

    const bgOptions = [
        { name: 'Plain', path: '' },
        { name: 'Modern', path: 'assets/bg-modern.png' },
        { name: 'Linen', path: 'assets/bg-linen.png' },
        { name: 'Soft', path: 'assets/bg-soft.png' }
    ];

    let selectedBgPath = "";

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
        const ratio = mock.offsetWidth / 960; 
        const scale = ratio * PT_TO_PX;

        const title = document.getElementById('valTitle').value;
        const lyrics = document.getElementById('valLyrics').value;
        const copy = document.getElementById('valCopy').value;
        const align = document.getElementById('slideAlign').value;
        const chordGap = 1 + (document.getElementById('chordGap').value / 100);

        mock.style.backgroundImage = selectedBgPath ? `url(${selectedBgPath})` : 'none';

        // --- TITLE ---
        const pt = document.getElementById('prevTitle');
        pt.innerText = title;
        pt.style.top = document.getElementById('yTitle').value + "%";
        pt.style.textAlign = align;
        pt.style.fontSize = (SIZE_TITLE * scale) + "px"; 
        pt.style.fontFamily = MAIN_FONT;
        pt.style.fontWeight = "bold";

        // --- COPYRIGHT ---
        const pc = document.getElementById('prevCopy');
        pc.innerText = copy;
        pc.style.top = document.getElementById('yCopy').value + "%";
        pc.style.textAlign = align;
        pc.style.fontSize = (SIZE_COPY * scale) + "px";
        pc.style.fontFamily = MAIN_FONT;
        pc.style.fontStyle = "italic";

        // --- LYRICS ---
        const pl = document.getElementById('prevLyrics');
        pl.style.top = document.getElementById('yLyrics').value + "%";
        pl.style.textAlign = align;
        pl.style.fontFamily = MAIN_FONT;
        
        const firstSection = lyrics.split(/\n?\s*(?=\[)/)[0] || "";
        pl.innerHTML = ""; 
        
        const lines = firstSection.split('\n');
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            const nextLine = lines[i+1] || "";
            const lineDiv = document.createElement('div');
            lineDiv.style.whiteSpace = "pre";
            lineDiv.style.lineHeight = chordGap;

            if (isChordLine(line)) {
                lineDiv.style.fontSize = (SIZE_CHORD * scale) + "px";
                lineDiv.style.color = "#808080";
                lineDiv.innerHTML = createHtmlGhostLine(line, nextLine, scale, align);
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

    downloadBtn.onclick = async () => {
        const pres = new PptxGenJS();
        pres.layout = 'LAYOUT_16x9';
        const align = document.getElementById('slideAlign').value;
        const chordGap = 1 + (document.getElementById('chordGap').value / 100);
        const sections = document.getElementById('valLyrics').value.split(/\n?\s*(?=\[)/).filter(s => s.trim());

        sections.forEach(section => {
            let slide = pres.addSlide();
            slide.background = selectedBgPath ? { path: selectedBgPath } : { fill: "FFFFFF" };

            slide.addText(document.getElementById('valTitle').value, {
                x: "5%", y: document.getElementById('yTitle').value + "%", w: "90%",
                fontSize: SIZE_TITLE, color: "000000", fontFace: MAIN_FONT, bold: true, align: align,
                valign: 'top', margin: 0
            });

            const lines = section.split('\n');
            let textObjects = [];
            for (let i = 0; i < lines.length; i++) {
                const line = lines[i];
                const nextLine = lines[i+1] || "";
                if (isChordLine(line)) {
                    textObjects.push(...createPptxGhostLine(line, nextLine, align));
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
                lineSpacing: SIZE_LYRIC * chordGap, margin: 0
            });

            slide.addText(document.getElementById('valCopy').value, {
                x: "5%", y: document.getElementById('yCopy').value + "%", w: "90%",
                fontSize: SIZE_COPY, color: "000000", italic: true, fontFace: MAIN_FONT, align: align,
                valign: 'top', margin: 0
            });
        });

        await pres.writeFile({ fileName: "Song_Slides_Debug.pptx" });
    };

    function isChordLine(str) {
        if (!str.trim()) return false;
        const chordChars = str.replace(/[A-G]|[m|maj|dim|aug|sus|2|4|5|7|9]|#|b|\s|\/|v|i|\[|\]/gi, "");
        return chordChars.length === 0;
    }

    /**
     * GHOST TEXT METHOD - DEBUG VERSION
     * The lyric character is duplicated but rendered with low opacity.
     */
    function createHtmlGhostLine(chords, lyrics, scale, align) {
        let html = "";
        const targetLen = align === 'center' ? Math.max(chords.length, lyrics.length) : chords.length;

        for (let i = 0; i < targetLen; i++) {
            const c = chords[i] || " ";
            const l = lyrics[i] || " ";

            if (c !== " ") {
                html += `<span>${c}</span>`;
            } else {
                // TROUBLESHOOTING: Showing the lyric duplication with 15% opacity
                const char = l === " " ? "\u00A0" : l;
                html += `<span style="color: rgba(255, 0, 0, 0.15); font-size:${SIZE_LYRIC * scale}px">${char}</span>`;
            }
        }
        return html || " ";
    }

    function createPptxGhostLine(chords, lyrics, align) {
        let result = [];
        const targetLen = align === 'center' ? Math.max(chords.length, lyrics.length) : chords.length;

        for (let i = 0; i < targetLen; i++) {
            const c = chords[i] || " ";
            const l = lyrics[i] || " ";

            if (c !== " ") {
                result.push({ 
                    text: c, 
                    options: { color: "808080", fontSize: SIZE_CHORD, fontFace: MAIN_FONT } 
                });
            } else {
                // TROUBLESHOOTING: Set transparency to 90 (10% visible) and color to Red
                result.push({ 
                    text: l === "" ? " " : l, 
                    options: { 
                        transparency: 90, 
                        color: "FF0000", 
                        fontSize: SIZE_LYRIC, 
                        fontFace: MAIN_FONT 
                    } 
                });
            }
        }
        result.push({ text: "\n" });
        return result;
    }

    updatePreview();
});
