await Word.run(async (context) => {

  // --- First sync: read the document ---
  const body = context.document.body;
  context.load(body, 'text');
  await context.sync();

  // Now we can use body.text to find wordy sentences
  const wordySentences = nlp(body.text)
    .sentences()
    .filter(s => s.wordCount() > 30)
    .out('array');

  // --- Queue up highlights based on what we found ---
  for (const sentence of wordySentences) {
    const results = body.search(sentence);
    context.load(results, 'items');
    await context.sync();  // Second sync: needed to get search results back

    if (results.items.length > 0) {
      results.items[0].font.highlightColor = '#FFFF00';
    }
  }

  // --- Third sync: push highlights to Word ---
  await context.sync();
});