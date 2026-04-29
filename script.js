document.addEventListener('DOMContentLoaded', () => {
    const inputs = document.querySelectorAll('input, textarea, select, [type="range"]');
    const downloadBtn = document.getElementById('downloadBtn');
    const bgSelector = document.getElementById('bgSelector');

    // CONFIG
    const VERSION = "1.0.3"; 
    const MAIN_FONT = "Times New Roman";
    const SIZE_TITLE = 38, SIZE_LYRIC = 24, SIZE_CHORD = 16, SIZE_COPY = 18;

    // Display version
    const versionDisplay = document.getElementById('appVersion');
    if (versionDisplay) versionDisplay.innerText = `v${VERSION}`;

    const bgOptions = [
        { name: 'Plain', path: '' },
        { name: 'Modern', path: 'assets/bg-modern.png' },
        { name: 'Linen', path: 'assets/bg-linen.png' },
        { name: 'Soft', path: 'assets/bg-soft.png' }
    ];

    let selectedBgPath = "";

    // 1. Build Bg Selector
    bgOptions.forEach((opt, i) => {
        const thumb = document.createElement('div');
        thumb.className = `bg-thumb ${i === 0 ? 'active' : ''}`;
        thumb.style.backgroundColor = opt.path ? 'transparent' : '#fff';
        if(opt.path) thumb.style.backgroundImage = `url(${opt.path})`;
        thumb.onclick = () => {
            document.querySelectorAll('.bg-thumb').forEach(t => t.classList.remove('active'));
            thumb.classList.add('active');
            selectedBgPath = opt.path;
            updatePreview();
        };
        bgSelector.appendChild(thumb);
    });

    // 2. Calibrated Dynamic Preview
    function updatePreview() {
        const mock = document.getElementById('slideMock');
        
        // CALIBRATION: Calculate ratio based on a 960px wide virtual PowerPoint slide
        const mockWidth = mock.offsetWidth;
        const ratio = mockWidth / 960; 

        const title = document.getElementById('valTitle').value;
        const lyrics = document.getElementById('valLyrics').value;
        const copy = document.getElementById('valCopy').value;
        const align = document.getElementById('slideAlign').value;

        mock.style.backgroundImage = selectedBgPath ? `url(${selectedBgPath})` : 'none';

        // Title
        const pt = document.getElementById('prevTitle');
        pt.innerText = title;
        pt.style.top = document.getElementById('yTitle').value + "%";
        pt.style.textAlign = align;
        pt.style.fontSize = (SIZE_TITLE * ratio) + "px"; // Calibrated
        pt.style.fontFamily = MAIN_FONT;
        pt.style.fontWeight = "bold";

        // Copyright
        const pc = document.getElementById('prevCopy');
        pc.innerText = copy;
        pc.style.top = document.getElementById('yCopy').value + "%";
        pc.style.textAlign = align;
        pc.style.fontSize = (SIZE_COPY * ratio) + "px"; // Calibrated
        pc.style.fontFamily = MAIN_FONT;
        pc.style.fontStyle = "italic";

        // Lyrics Area
        const pl = document.getElementById('prevLyrics');
        pl.style.top = document.getElementById('yLyrics').value + "%";
        pl.style.textAlign = align;
        pl.style.fontFamily = MAIN_FONT;
        
        const firstSection = lyrics.split(/\n?\s*(?=\[)/)[0] || "";
        pl.innerHTML = ""; 
        
        const lines = firstSection.split('\n');
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            const nextLine = lines[i+1];
            const lineDiv = document.createElement('div');
            lineDiv.style.whiteSpace = "pre";
            lineDiv.style.lineHeight = "1.1";

            if (isChordLine(line) && nextLine) {
                lineDiv.style.fontSize = (SIZE_CHORD * ratio) + "px"; // Calibrated
                lineDiv.style.color = "#808080";
                lineDiv.innerHTML = createHtmlGhostLine(line, nextLine, ratio);
            } else {
                lineDiv.style.fontSize = (SIZE_LYRIC * ratio) + "px"; // Calibrated
                lineDiv.style.color = "#000";
                lineDiv.innerText = line || " "; 
            }
            pl.appendChild(lineDiv);
        }
    }

    // Update preview when inputs change or window resizes
    inputs.forEach(input => input.addEventListener('input', updatePreview));
    window.addEventListener('resize', updatePreview);

    // 3. PPTX Download
    downloadBtn.onclick = async () => {
        const pres = new PptxGenJS();
        pres.layout = 'LAYOUT_16x9';
        const align = document.getElementById('slideAlign').value;
        const sections = document.getElementById('valLyrics').value.split(/\n?\s*(?=\[)/).filter(s => s.trim());

        sections.forEach(section => {
            let slide = pres.addSlide();
            slide.background = selectedBgPath ? { path: selectedBgPath } : { fill: "FFFFFF" };

            slide.addText(document.getElementById('valTitle').value, {
                x: "5%", y: document.getElementById('yTitle').value + "%", w: "90%",
                fontSize: SIZE_TITLE, color: "000000", fontFace: MAIN_FONT, bold: true, align: align
            });

            const lines = section.split('\n');
            let textObjects = [];
            for (let i = 0; i < lines.length; i++) {
                if (isChordLine(lines[i]) && lines[i+1]) {
                    textObjects.push(...createGhostLine(lines[i], lines[i+1], SIZE_CHORD, SIZE_LYRIC));
                } else if (lines[i].trim() || lines[i] === "") {
                    textObjects.push({ text: lines[i] + "\n", options: { color: "000000", fontSize: SIZE_LYRIC, fontFace: MAIN_FONT } });
                }
            }

            slide.addText(textObjects, {
                x: "5%", y: document.getElementById('yLyrics').value + "%", w: "90%", h: "60%",
                fontFace: MAIN_FONT, valign: 'top', align: align, lineSpacing: 24
            });

            slide.addText(document.getElementById('valCopy').value, {
                x: "5%", y: document.getElementById('yCopy').value + "%", w: "90%",
                fontSize: SIZE_COPY, color: "000000", italic: true, fontFace: MAIN_FONT, align: align
            });
        });

        await pres.writeFile({ fileName: "Song_Slides.pptx" });
    };

    function isChordLine(str) {
        const chordChars = str.replace(/[A-G]|[m|maj|dim|aug|sus|2|4|5|7|9]|#|b|\s|\/|v|i|\[|\]/gi, "");
        return str.trim() && chordChars.length === 0;
    }

    function createHtmlGhostLine(chords, lyrics, ratio) {
        let html = "", lastIdx = 0, re = /\S+/g, match;
        while ((match = re.exec(chords)) !== null) {
            let spacerText = lyrics.substring(lastIdx, match.index);
            if (match.index > lyrics.length) {
                spacerText = lyrics.substring(lastIdx) + " ".repeat(match.index - lyrics.length);
            }
            if (spacerText) {
                // Spacer font size must match Lyric font size for perfect horizontal alignment
                html += `<span style="opacity:0; font-size:${SIZE_LYRIC * ratio}px">${spacerText}</span>`;
            }
            html += `<span>${match[0]}</span>`;
            lastIdx = match.index + match[0].length;
        }
        return html || " ";
    }

    function createGhostLine(chords, lyrics, cSize, lSize) {
        let result = [], lastIdx = 0, re = /\S+/g, match;
        while ((match = re.exec(chords)) !== null) {
            let spacer = lyrics.substring(lastIdx, match.index);
            if (match.index > lyrics.length) spacer = lyrics.substring(lastIdx) + " ".repeat(match.index - lyrics.length);
            if (spacer) result.push({ text: spacer, options: { transparency: 100, fontSize: lSize, fontFace: MAIN_FONT } });
            result.push({ text: match[0], options: { color: "#808080", fontSize: cSize, fontFace: MAIN_FONT } });
            lastIdx = match.index + match[0].length;
        }
        result.push({ text: "\n" });
        return result;
    }

    updatePreview();
});
