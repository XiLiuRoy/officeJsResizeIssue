// 1. Copy the following to script lab Script Tab
$("#setBoldAndResize").click(() => tryCatch(setBoldAndResize));
let columnSeed = 0;

async function setBoldAndResize() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    const testString =
      "This is a bold text and the display language needs to set to Chinese or some language other than English";



    const cell = sheet.getCell(1, columnSeed++);
    cell.values = [[testString]];
    cell.format.font.bold = true;

    await context.sync();
    cell.format.autofitColumns();
    await context.sync();
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

// 2. Copy the following to scripit lab HTML tab
<section class="ms-font-m">
	<p class="ms-font-m">This sample demonstrates basic Excel API calls.</p>
</section>

<section class="samples ms-font-m">
	<h3>Try it out</h3>
	<p class="ms-font-m">Select some cells in the worksheet, then press <b>Highlight selected range</b>.</p>
	<button id="setBoldAndResize" class="ms-Button">
			        <span class="ms-Button-label">setBoldAndResize</span>
		</button>
</section>
