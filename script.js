document.addEventListener('DOMContentLoaded', () => {
    const inputs = document.querySelectorAll('input, textarea, select, [type="range"]');
    const downloadBtn = document.getElementById('downloadBtn');
    const bgSelector = document.getElementById('bgSelector');

    const VERSION = "1.0.7"; 
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

        mock.style.backgroundImage = selectedBgPath ? `url(${selectedBgPath})` : 'none';

        // --- TITLE ---
        const pt = document.getElementById('prevTitle');
        pt.innerText = title;
        pt.style.top = document.getElementById('yTitle').value + "%";
        pt.style.textAlign = align;
        pt.style.fontSize = (SIZE_TITLE * scale) + "px"; 
        pt.style.fontFamily = MAIN_FONT;

        // --- COPYRIGHT ---
        const pc = document.getElementById('prevCopy');
        pc.innerText = copy;
        pc.style.top = document.getElementById('yCopy').value + "%";
        pc.style.textAlign = align;
        pc.style.fontSize = (SIZE_COPY * scale) + "px";
        pc.style.fontFamily = MAIN_FONT;

        // --- LYRICS ---
        const pl = document.getElementById('prevLyrics');
        pl.style.top = document.getElementById('yLyrics').value + "%";
        pl.style.textAlign = align;
        pl.style.fontFamily = MAIN_FONT;
        
        // Only preview the first section (before next [Verse/Chorus])
        const firstSection = lyrics.split(/\n?\s*(?=\[)/)[0] || "";
        pl.innerHTML = ""; 
        
        const lines = firstSection.split('\n');
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            const nextLine = lines[i+1] || "";
            const lineDiv = document.createElement('div');
            lineDiv.style.whiteSpace = "pre";
            
            if (isChordLine(line)) {
                lineDiv.style.fontSize = (SIZE_CHORD * scale) + "px";
                lineDiv.style.color = "#808080";
                lineDiv.style.lineHeight = "1";
                // Important: Pass the lyric line to calculate ghost spacing
                lineDiv.innerHTML = createHtmlGhostLine(line, nextLine, scale);
                pl.appendChild(lineDiv);
            } else {
                lineDiv.style.fontSize = (SIZE_LYRIC * scale) + "px";
                lineDiv.style.color = "#000";
                // Adjust margin to simulate PPTX line spacing
                lineDiv.style.marginBottom = (SIZE_LYRIC * 0.2 * scale) + "px";
                lineDiv.innerText = line || " "; 
                pl.appendChild(lineDiv);
            }
        }
    }

    inputs.forEach(input => input.addEventListener('input', updatePreview));
    window.addEventListener('resize', updatePreview);

    downloadBtn.onclick = async () => {
        const pres = new PptxGenJS();
        pres.layout = 'LAYOUT_16x9';
        const align = document.getElementById('slideAlign').value;
        const rawText = document.getElementById('valLyrics').value;
        
        // Split by bracket sections e.g. [Verse 1]
        const sections = rawText.split(/\n?\s*(?=\[)/).filter(s => s.trim());

        sections.forEach(section => {
            let slide = pres.addSlide();
            slide.background = selectedBgPath ? { path: selectedBgPath } : { fill: "FFFFFF" };

            // Title
            slide.addText(document.getElementById('valTitle').value, {
                x: "5%", y: document.getElementById('yTitle').value + "%", w: "90%",
                fontSize: SIZE_TITLE, color: "000000", fontFace: MAIN_FONT, bold: true, align: align
            });

            // Lyrics & Chords
            const lines = section.split('\n');
            let textObjects = [];
            
            for (let i = 0; i < lines.length; i++) {
                const line = lines[i];
                const nextLine = lines[i+1] || "";

                if (isChordLine(line)) {
                    // Create Chord Line with Ghost Spacers
                    textObjects.push(...createPptxGhostLine(line, nextLine));
                } else {
                    // Normal Lyric Line
                    textObjects.push({ 
                        text: (line || " ") + "\n", 
                        options: { color: "000000", fontSize: SIZE_LYRIC, fontFace: MAIN_FONT } 
                    });
                }
            }

            slide.addText(textObjects, {
                x: "5%", y: document.getElementById('yLyrics').value + "%", w: "90%", h: "auto",
                fontFace: MAIN_FONT, valign: 'top', align: align,
                lineSpacing: SIZE_LYRIC * 1.2, // Match the 1.2 height used in preview
                margin: 0
            });

            // Copyright
            slide.addText(document.getElementById('valCopy').value, {
                x: "5%", y: document.getElementById('yCopy').value + "%", w: "90%",
                fontSize: SIZE_COPY, color: "000000", italic: true, fontFace: MAIN_FONT, align: align
            });
        });

        await pres.writeFile({ fileName: `Song_Slides.pptx` });
    };

    function isChordLine(str) {
        if (!str.trim()) return false;
        // Check if line contains mostly chords/spaces
        const chordChars = str.replace(/[A-G]|[m|maj|dim|aug|sus|2|4|5|7|9]|#|b|\s|\/|v|i|\[|\]/gi, "");
        return chordChars.length === 0;
    }

    function createHtmlGhostLine(chords, lyrics, currentScale) {
        let html = "", lastIdx = 0, re = /\S+/g, match;
        while ((match = re.exec(chords)) !== null) {
            let spacerText = lyrics.substring(lastIdx, match.index);
            if (spacerText) {
                // Ghost text must match Lyric Size and Font to push chords correctly
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
                // Transparency 100 makes it a ghost spacer
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
