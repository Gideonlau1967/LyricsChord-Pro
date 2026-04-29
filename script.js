document.addEventListener('DOMContentLoaded', () => {
    const inputs = document.querySelectorAll('input, textarea, select, [type="range"]');
    const downloadBtn = document.getElementById('downloadBtn');
    const bgSelector = document.getElementById('bgSelector');

    const VERSION = "1.1.2"; 
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
        
        // Link the "Chord Gap" slider to line-height
        const chordGapVal = document.getElementById('chordGap').value;
        const lineSpacingMult = 1 + (chordGapVal / 100);

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
            lineDiv.style.lineHeight = lineSpacingMult;

            if (isChordLine(line)) {
                lineDiv.style.fontSize = (SIZE_CHORD * scale) + "px";
                lineDiv.style.color = "#808080";
                // Pass the next line (lyrics) to calculate ghost spacing
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

    downloadBtn.onclick = async () => {
        const pres = new PptxGenJS();
        pres.layout = 'LAYOUT_16x9';
        const align = document.getElementById('slideAlign').value;
        const lineSpacingMult = 1 + (document.getElementById('chordGap').value / 100);
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
                const nextLine = lines[i+1] || "";
                if (isChordLine(lines[i])) {
                    textObjects.push(...createPptxGhostLine(lines[i], nextLine));
                } else {
                    textObjects.push({ 
                        text: (lines[i] || " ") + "\n", 
                        options: { color: "000000", fontSize: SIZE_LYRIC, fontFace: MAIN_FONT } 
                    });
                }
            }

            slide.addText(textObjects, {
                x: "5%", y: document.getElementById('yLyrics').value + "%", w: "90%", h: "70%",
                fontFace: MAIN_FONT, valign: 'top', align: align, 
                lineSpacing: SIZE_LYRIC * lineSpacingMult, margin: 0
            });

            slide.addText(document.getElementById('valCopy').value, {
                x: "5%", y: document.getElementById('yCopy').value + "%", w: "90%",
                fontSize: SIZE_COPY, color: "000000", italic: true, fontFace: MAIN_FONT, align: align,
                valign: 'top', margin: 0
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
        const lyricLen = lyrics.length;

        while ((match = re.exec(chords)) !== null) {
            // 1. Add ghost text for the gap BEFORE the chord
            let spacerText = lyrics.substring(lastIdx, match.index);
            
            // If the chord is placed further right than the lyrics exist, use real spaces
            if (match.index > lyricLen) {
                const spacesNeeded = match.index - Math.max(lastIdx, lyricLen);
                spacerText += " ".repeat(spacesNeeded);
            }

            if (spacerText) {
                html += `<span style="visibility:hidden; font-size:${SIZE_LYRIC * currentScale}px">${spacerText}</span>`;
            }

            // 2. Add the actual Chord
            html += `<span>${match[0]}</span>`;
            lastIdx = match.index + match[0].length;
        }
        
        // 3. IMPORTANT: Add ghost text for the rest of the lyric line
        // This ensures the chord line is the same width as the lyric line for Center Alignment
        if (lastIdx < lyricLen) {
            const trailingPadding = lyrics.substring(lastIdx);
            html += `<span style="visibility:hidden; font-size:${SIZE_LYRIC * currentScale}px">${trailingPadding}</span>`;
        }

        return html || " ";
    }

    function createPptxGhostLine(chords, lyrics) {
        let result = [], lastIdx = 0, re = /\S+/g, match;
        const lyricLen = lyrics.length;

        while ((match = re.exec(chords)) !== null) {
            let spacer = lyrics.substring(lastIdx, match.index);
            if (match.index > lyricLen) {
                spacer += " ".repeat(match.index - Math.max(lastIdx, lyricLen));
            }
            if (spacer) {
                result.push({ text: spacer, options: { transparency: 100, fontSize: SIZE_LYRIC, fontFace: MAIN_FONT } });
            }
            result.push({ text: match[0], options: { color: "808080", fontSize: SIZE_CHORD, fontFace: MAIN_FONT } });
            lastIdx = match.index + match[0].length;
        }

        // Add trailing ghost padding for PPTX alignment
        if (lastIdx < lyricLen) {
            result.push({ text: lyrics.substring(lastIdx), options: { transparency: 100, fontSize: SIZE_LYRIC, fontFace: MAIN_FONT } });
        }
        
        result.push({ text: "\n" });
        return result;
    }

    updatePreview();
});
