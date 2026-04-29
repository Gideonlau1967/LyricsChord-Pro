--- START OF FILE script (2).js ---

document.addEventListener('DOMContentLoaded', () => {
    const inputs = document.querySelectorAll('input, textarea, select, [type="range"]');
    const downloadBtn = document.getElementById('downloadBtn');
    const bgSelector = document.getElementById('bgSelector');

    // CONFIG
    const VERSION = "1.4.4 Final Stable"; 
    const MAIN_FONT = "Times New Roman";
    const SIZE_TITLE = 32, SIZE_LYRIC = 24, SIZE_CHORD = 14, SIZE_SECTION = 16, SIZE_COPY = 14;
    const PT_TO_PX = 96 / 72; 

    // UI: Set the input to monospace for chord alignment
    const lyricInput = document.getElementById('valLyrics');
    if (lyricInput) {
        lyricInput.style.fontFamily = "'Consolas', 'Monaco', 'Courier New', monospace";
    }

    const versionDisplay = document.querySelector('.version-badge');
    if (versionDisplay) versionDisplay.innerText = `v${VERSION}`;

    const bgOptions = [
        { name: 'Plain', path: '' },
        { name: 'Modern', path: 'assets/bg-default.png' },
        { name: 'Linen', path: 'assets/bg-linen.png' },
        { name: 'Soft', path: 'assets/bg-soft.png' },
        { name: 'Dark', path: 'assets/bg-modern.png' },
        { name: 'Paper', path: 'assets/bg-paper.png' }
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
        if (!mock) return;
        const ratio = mock.offsetWidth / 960; 
        const scale = ratio * PT_TO_PX;

        const title = document.getElementById('valTitle').value;
        const lyrics = document.getElementById('valLyrics').value;
        const copy = document.getElementById('valCopy').value;
        const align = document.getElementById('slideAlign').value;
        const chordGapSlider = parseInt(document.getElementById('chordGap').value);

        mock.style.backgroundImage = selectedBgPath ? `url(${selectedBgPath})` : 'none';

        const pt = document.getElementById('prevTitle');
        const pc = document.getElementById('prevCopy');
        const pl = document.getElementById('prevLyrics');

        [pt, pc, pl].forEach(el => {
            el.style.textAlign = align;
            el.style.fontFamily = MAIN_FONT;
        });

        pt.innerText = title;
        pt.style.top = document.getElementById('yTitle').value + "%";
        pt.style.fontSize = (SIZE_TITLE * scale) + "px"; 
        pt.style.fontWeight = "bold";

        pc.innerText = copy;
        pc.style.top = document.getElementById('yCopy').value + "%";
        pc.style.fontSize = (SIZE_COPY * scale) + "px";
        pc.style.fontStyle = "italic";

        // PREVIEW: Align middle using Flexbox and height
        pl.style.top = document.getElementById('yLyrics').value + "%";
        pl.style.height = "70%"; // Match the PPTX height
        pl.style.display = "flex";
        pl.style.flexDirection = "column";
        pl.style.justifyContent = "center"; // This centers content vertically in preview
        
        const firstSectionRaw = lyrics.split(/(?=\[)/)[0] || "";
        const firstSection = firstSectionRaw.replace(/^[\n\r]+|[\n\r]+$/g, '');
        pl.innerHTML = ""; 
        
        // Inner container for the lines to maintain alignment
        const innerContent = document.createElement('div');
        innerContent.style.width = "100%";

        const lines = firstSection.split('\n');
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            const nextLine = lines[i+1] || "";
            const lineDiv = document.createElement('div');
            lineDiv.style.whiteSpace = "pre";

            if (line.trim().startsWith('[') && line.trim().endsWith(']')) {
                lineDiv.style.fontSize = (SIZE_SECTION * scale) + "px";
                lineDiv.style.lineHeight = "1.2";
                lineDiv.style.marginBottom = (2 * scale) + "px";
                lineDiv.innerText = line;
            } else if (isChordLine(line)) {
                lineDiv.style.fontSize = (SIZE_CHORD * scale) + "px";
                lineDiv.style.lineHeight = "0.7"; 
                lineDiv.style.marginBottom = (chordGapSlider * scale) + "px";
                lineDiv.innerHTML = createHtmlGhostLine(line, nextLine, scale, align);
            } else {
                lineDiv.style.fontSize = (SIZE_LYRIC * scale) + "px";
                lineDiv.style.lineHeight = "1";
                lineDiv.style.marginBottom = (5 * scale) + "px"; 
                lineDiv.innerText = line || " "; 
            }
            innerContent.appendChild(lineDiv);
        }
        pl.appendChild(innerContent);
    }

    inputs.forEach(input => input.addEventListener('input', updatePreview));
    window.addEventListener('resize', updatePreview);

    downloadBtn.onclick = async () => {
        const pres = new PptxGenJS();
        pres.layout = 'LAYOUT_16x9';
        
        const songTitleInput = document.getElementById('valTitle').value.trim();
        const songTitle = songTitleInput || "Song_Slides";
        
        const align = document.getElementById('slideAlign').value;
        const gapVal = parseInt(document.getElementById('chordGap').value);
        const rawText = document.getElementById('valLyrics').value;

        const sections = rawText.split(/(?=\[)/).filter(s => s.trim());

        sections.forEach(section => {
            let slide = pres.addSlide();
            slide.background = selectedBgPath ? { path: selectedBgPath } : { fill: "FFFFFF" };

            slide.addNotes(section);

            slide.addText(songTitleInput || "Untitled", {
                x: "5%", y: document.getElementById('yTitle').value + "%", w: "90%",
                fontSize: SIZE_TITLE, color: "000000", fontFace: MAIN_FONT, bold: true, align: align, valign: 'top'
            });

            const cleanSectionForSlide = section.replace(/^[\n\r]+|[\n\r]+$/g, '');
            const lines = cleanSectionForSlide.split('\n');
            let textObjects = [];
            for (let i = 0; i < lines.length; i++) {
                const line = lines[i];
                if (line.trim().startsWith('[') && line.trim().endsWith(']')) {
                    textObjects.push({ text: line.trim() + "\n", options: { fontSize: SIZE_SECTION } });
                } else if (isChordLine(line)) {
                    textObjects.push(...createPptxGhostLine(line, lines[i+1] || "", align));
                } else {
                    textObjects.push({ text: (line || " ") + "\n", options: { fontSize: SIZE_LYRIC } });
                }
            }
            const spacingMult = 0.85 + (gapVal / 100);
            
            // SET VALIGN TO MIDDLE HERE
            slide.addText(textObjects, {
                x: "5%", y: document.getElementById('yLyrics').value + "%", w: "90%", h: "70%",
                fontFace: MAIN_FONT, valign: 'middle', align: align, lineSpacing: SIZE_LYRIC * spacingMult
            });

            slide.addText(document.getElementById('valCopy').value, {
                x: "5%", y: document.getElementById('yCopy').value + "%", w: "90%",
                fontSize: SIZE_COPY, fontFace: MAIN_FONT, italic: true, align: align, valign: 'top'
            });
        });

        const safeFileName = songTitle.replace(/[/\\?%*:|"<>]/g, '-') + ".pptx";
        await pres.writeFile({ fileName: safeFileName });
    };

    function isChordLine(str) {
        if (!str.trim() || (str.trim().startsWith('[') && str.trim().endsWith(']'))) return false;
        return str.replace(/[A-G]|[m|maj|dim|aug|sus|2|4|5|7|9]|#|b|\s|\/|v|i|\[|\]/gi, "").length === 0;
    }

    function createHtmlGhostLine(chords, lyrics, scale, align) {
        let html = "";
        const targetLen = align === 'center' ? Math.max(chords.length, lyrics.length) : chords.length;
        for (let i = 0; i < targetLen; i++) {
            const c = chords[i] || " ", l = lyrics[i] || " ", char = l === " " ? "\u00A0" : l;
            html += `<span style="position:relative; display:inline-block; font-size:${SIZE_LYRIC * scale}px; color: transparent;">`;
            html += char; 
            if (c !== " ") html += `<span style="position:absolute; left:0; bottom:0; font-size:${SIZE_CHORD * scale}px; color:#808080; visibility:visible;">${c}</span>`;
            html += `</span>`;
        }
        return html || " ";
    }

    function createPptxGhostLine(chords, lyrics, align) {
        let result = [];
        const targetLen = align === 'center' ? Math.max(chords.length, lyrics.length) : chords.length;
        for (let i = 0; i < targetLen; i++) {
            const c = chords[i] || " ", l = lyrics[i] || " ";
            if (c !== " ") {
                result.push({ text: c, options: { color: "808080", fontSize: SIZE_CHORD } });
            } else {
                result.push({ text: l === "" ? " " : l, options: { transparency: 100, color: "FFFFFF", fontSize: SIZE_LYRIC } });
            }
        }
        result.push({ text: "\n" });
        return result;
    }

    updatePreview();
});
