async function logParagraphStyles() {
    await Word.run(async (context) => {
        const body = context.document.body;
        context.load(body, 'paragraphs');
        await context.sync();

        body.paragraphs.items.forEach(para => {
            context.load(para, 'text, style, font');
        });
        await context.sync();

        body.paragraphs.items.forEach(para => {
            console.log('${para.style}: ${para.text}')
        })
    });
};
async function InsertToc() {
    await Word.run(async (context) => {
        const body = context.document.body;
        // 1. Insert a blank paragraph at the start to hold the TOC
        const tocParagraph = body.insertParagraph('', Word.InsertLocation.start);
        // 2. Insert the native TOC field into that paragraph
        tocParagraph.insertField(Word.FieldType.toc, true, Word.InsertLocation.replace);
        await context.sync();
    });
}