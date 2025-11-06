Office.onReady(() => {
  if (Office.context.host === Office.HostType.Word) {
    replaceHelloWord();
  } else if (Office.context.host === Office.HostType.Excel) {
    replaceHelloExcel();
  }
});

function replaceHelloWord() {
  Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    if (body.text.includes("@hello")) {
      body.insertText(body.text.replace(/@hello/g, "hellohellohellohellohello"), Word.InsertLocation.replace);
    }
    await context.sync();
  });
}

function replaceHelloExcel() {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    range.load("values");
    await context.sync();

    const newValues = range.values.map(row =>
      row.map(cell => typeof cell === "string" ? cell.replace(/@hello/g, "hellohellohellohellohello") : cell)
    );

    range.values = newValues;
    await context.sync();
  });
}
