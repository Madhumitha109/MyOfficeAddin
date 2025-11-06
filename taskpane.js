Office.onReady(() => {
  document.getElementById("runButton").onclick = runReplacement;
});

async function runReplacement() {
  const host = Office.context.host;

  if (host === Office.HostType.Word) {
    await replaceInWord();
  } else if (host === Office.HostType.Excel) {
    await replaceInExcel();
  } else if (host === Office.HostType.PowerPoint) {
    await replaceInPowerPoint();
  } else {
    console.log("Host not supported");
  }
}

// ------------------ Word ------------------
async function replaceInWord() {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const newText = body.text.replace(/@helloworld/g, "hello hello hello hello hello");
    body.insertText(newText, Word.InsertLocation.replace);
    await context.sync();
  });
}

// ------------------ Excel ------------------
async function replaceInExcel() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    usedRange.load("values");
    await context.sync();

    const newValues = usedRange.values.map(row => 
      row.map(cell => typeof cell === "string" ? cell.replace(/@helloworld/g, "hello hello hello hello hello") : cell)
    );

    usedRange.values = newValues;
    await context.sync();
  });
}

// ------------------ PowerPoint ------------------
async function replaceInPowerPoint() {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    slides.items.forEach(slide => {
      slide.shapes.load("items");
    });
    await context.sync();

    slides.items.forEach(slide => {
      slide.shapes.items.forEach(shape => {
        if (shape.textFrame && shape.textFrame.textRange) {
          shape.textFrame.textRange.text = shape.textFrame.textRange.text.replace(/@helloworld/g, "hello hello hello hello hello");
        }
      });
    });
    await context.sync();
  });
}
