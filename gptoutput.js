// functions.js (or wherever your exported functions live)
import nlp from 'compromise'; // Make sure you installed this via npm

export async function run() {
    await Word.run(async (context) => {
        const body = context.document.body;
        body.load("text"); // Load all text from the document
        await context.sync();

        const text = body.text;

        // ----------------------------
        // 1️⃣ Split text into sentences
        // ----------------------------
        const sentences = text.match(/[^.!?]+[.!?]+/g) || [];
        const sentenceLengths = sentences.map(s => s.trim().split(/\s+/).length);

        // ----------------------------
        // 2️⃣ Calculate average sentence length
        // ----------------------------
        const totalWords = sentenceLengths.reduce((a, b) => a + b, 0);
        const averageLength = sentenceLengths.length ? (totalWords / sentenceLengths.length) : 0;

        // ----------------------------
        // 3️⃣ Identify long sentences (>30 words)
        // ----------------------------
        const longSentences = sentences.filter((s, i) => sentenceLengths[i] > 30);

        // Optional: Highlight long sentences in Word
        longSentences.forEach(sentence => {
            const range = body.search(sentence, { matchCase: false, matchWholeWord: true });
            range.load("font");
            range.items.forEach(r => r.font.highlightColor = "#FFFF00"); // Yellow highlight
        });
        await context.sync();

        // ----------------------------
        // 4️⃣ Word type breakdown using compromise
        // ----------------------------
        const doc = nlp(text);
        const nounCount = doc.nouns().length;
        const verbCount = doc.verbs().length;
        const adverbCount = doc.adverbs().length;
        const articleCount = doc.match('#Determiner').length; // includes a, an, the

        // ----------------------------
        // 5️⃣ Display results in taskpane
        // ----------------------------
        const resultsDiv = document.getElementById("results");
        if (resultsDiv) {
            resultsDiv.innerHTML = `
                <p><strong>Average sentence length:</strong> ${averageLength.toFixed(2)} words</p>
                <p><strong>Number of long sentences (>30 words):</strong> ${longSentences.length}</p>
                <p><strong>Word types:</strong> Nouns: ${nounCount}, Verbs: ${verbCount}, Adverbs: ${adverbCount}, Articles: ${articleCount}</p>
            `;
        }
    });
}