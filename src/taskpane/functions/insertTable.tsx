import { Group } from "../classes/Group";
import { formatValue } from "../functions/formatValue";

/* global Word, OfficeExtension */

const insertTable = async (name: string, groups?: Group[]) => {
  Word.run(async (context) => {
    // Make sure there are groups and that there are statistics
    if (groups && groups[0].statistics) {
      const rows = groups.length + 1;
      const columns = groups[0].statistics?.length + 1;

      const range = context.document.getSelection();
      const table = range.insertTable(rows, columns, Word.InsertLocation.after);

      // Some styling
      table.getBorder("All").type = "None";
      table.getBorder("Top").type = "Double";
      table.getBorder("Bottom").type = "Single";

      // Set the first cell's content to the group name
      const cellName = table.getCell(0, 0);
      cellName.getBorder("Bottom").type = "Single";
      cellName.body.getRange().insertText(name, Word.InsertLocation.replace);

      // Set the content of the remaining cells in the first row to the names of the statistics
      groups[0].statistics.forEach((statistic, i) => {
        const cellStatisticsName = table.getCell(0, i + 1);
        cellStatisticsName.getBorder("Bottom").type = "Single";
        cellStatisticsName.body.getRange().font.italic = true;
        cellStatisticsName.body
          .getRange()
          .insertText(statistic.symbol ? statistic.symbol : statistic.name, Word.InsertLocation.replace);
      });

      // Loop over each group and set the name and values
      groups.forEach((group, i) => {
        table
          .getCell(i + 1, 0)
          .body.getRange()
          .insertText(group.name, Word.InsertLocation.replace);

        group.statistics?.forEach((statistic, j) => {
          const value = table
            .getCell(i + 1, j + 1)
            .body.getRange()
            .insertContentControl();
          value.tag = statistic.identifier;
          value.insertText(formatValue(statistic, 2), Word.InsertLocation.replace);
        });
      });
    }
    return context.sync;
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
};

export { insertTable };
