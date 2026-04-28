document.addEventListener('DOMContentLoaded', () => {
    const downloadBtn = document.getElementById('downloadBtn');

    // PPT Configuration
    const MAIN_FONT = "Times New Roman"; 
    const COLOR_BLACK = "000000";
    const COLOR_GREY  = "808080";
    const SIZE_TITLE  = 38;
    const SIZE_LYRIC  = 28;
    const SIZE_CHORD  = 18;
    const SIZE_COPY   = 18;

    downloadBtn.onclick = async () => {
        const status = document.getElementById('status');
        status.innerText = "Generating PowerPoint...";

        try {
            const pres = new PptxGenJS();
            pres.layout = 'LAYOUT_16x9';

            const rawText = document.getElementById('valLyrics').value;
            const alignSetting = document.getElementById('slideAlign').value;
            
            // Split sections by [Verse], [Chorus] etc.
            const sections = rawText.split(/\n?\s*(?=\[)/).filter(s => s.trim() !== "");

            sections.forEach((section) => {
                let slide = pres.addSlide();
                
                // Background
                const bgUrl = document.getElementById('bgUrl').value;
                if (bgUrl) slide.background = { path: bgUrl };
                else slide.background = { fill: "FFFFFF" };

                // 1. Add Title (38pt, Bold, Black)
                slide.addText(document.getElementById('valTitle').value, {
                    x: "5%", y: "8%", w: "90%",
                    fontSize: SIZE_TITLE, color: COLOR_BLACK, 
                    fontFace: MAIN_FONT, bold: true, align: alignSetting
                });

                // 2. Process Lyrics & Chords (Ghost Text Method)
                const lines = section.split('\n');
                let textObjects = [];
                const gapSize = SIZE_LYRIC * 0.25;

                for (let i = 0; i < lines.length; i++) {
                    let line = lines[i];
                    if (isChordLine(line) && lines[i+1]) {
                        textObjects.push(...createGhostLine(line, lines[i+1], SIZE_CHORD, SIZE_LYRIC));
                        textObjects.push({ text: "\n", options: { fontSize: gapSize } });
                    } else if (!isChordLine(line) && line.trim() !== "") {
                        textObjects.push({ 
                            text: line + "\n", 
                            options: { color: COLOR_BLACK, fontSize: SIZE_LYRIC, fontFace: MAIN_FONT } 
                        });
                        textObjects.push({ text: "\n", options: { fontSize: SIZE_LYRIC * 0.4 } });
                    }
                }

                slide.addText(textObjects, {
                    x: "5%", y: "30%", w: "90%", h: "55%", 
                    fontFace: MAIN_FONT, valign: 'top', align: alignSetting
                });
                
                // 3. Add Copyright (18pt, Black, Italic)
                slide.addText(document.getElementById('valCopy').value, {
                    x: "5%", y: "88%", w: "90%",
                    fontSize: SIZE_COPY, color: COLOR_BLACK, 
                    italic: true, fontFace: MAIN_FONT, align: alignSetting
                });
            });

            await pres.writeFile({ fileName: `${document.getElementById('valTitle').value || 'Song'}.pptx` });
            status.innerText = "PowerPoint Downloaded Successfully!";
        } catch (e) { 
            console.error(e); 
            status.innerText = "Error generating file."; 
        }
    };

    function isChordLine(str) {
        if (!str.trim()) return false;
        const chordChars = str.replace(/[A-G]|[m|maj|dim|aug|sus|2|4|5|7|9]|#|b|\s|\/|v|i|\[|\]/gi, "");
        return chordChars.length === 0;
    }

    function createGhostLine(chords, lyrics, cSize, lSize) {
        let result = [];
        let lastIdx = 0;
        const re = /\S+/g;
        let match;
        while ((match = re.exec(chords)) !== null) {
            let chord = match[0];
            let pos = match.index;
            let spacerText = lyrics.substring(lastIdx, pos);
            if (pos > lyrics.length) spacerText = lyrics.substring(lastIdx) + " ".repeat(pos - lyrics.length);
            
            if (spacerText) {
                result.push({ text: spacerText, options: { transparency: 100, fontSize: lSize, fontFace: MAIN_FONT } });
            }
            result.push({ text: chord, options: { color: COLOR_GREY, fontSize: cSize, fontFace: MAIN_FONT } });
            lastIdx = pos + chord.length;
        }
        result.push({ text: "\n" });
        return result;
    }
});