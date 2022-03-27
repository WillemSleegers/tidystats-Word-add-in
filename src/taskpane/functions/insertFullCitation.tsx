/* global Word, OfficeExtension */

const insertFullCitation = async () => {
  Word.run(async (context) => {
    const range = context.document.getSelection();
    range.insertText(
      "Sleegers, W. W. A. (2021). tidystats: Save output of statistical tests (Version 0.51) [Computer software]. https://doi.org/10.5281/zenodo.4041859",
      "End"
    );

    return context.sync;
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
};

export { insertFullCitation };
