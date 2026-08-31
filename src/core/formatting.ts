import { DEMO_TABLE_NAME } from "./types";

/**
 * Adds a conditional format that highlights every row of the demo table whose
 * quantity is below the given threshold.
 *
 * @param threshold Quantity below which a row is highlighted.
 */
export async function addRowConditionalFormat(threshold = 50): Promise<void> {
  await Excel.run(async (context) => {
    const table = context.workbook.tables.getItem(DEMO_TABLE_NAME);
    const bodyRange = table.getDataBodyRange();
    const format = bodyRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);

    format.custom.rule.formula = `=$B2<${threshold}`;
    format.custom.format.fill.color = "#FFF2CC";
    format.custom.format.font.color = "#7F6000";

    await context.sync();
  });
}

/**
 * Adds a colour scale conditional format to a single column of the demo table.
 *
 * @param columnName Name of the table column, for example "Price".
 */
export async function addColumnConditionalFormat(columnName = "Price"): Promise<void> {
  await Excel.run(async (context) => {
    const table = context.workbook.tables.getItem(DEMO_TABLE_NAME);
    const columnRange = table.columns.getItem(columnName).getDataBodyRange();
    const format = columnRange.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);

    format.colorScale.criteria = {
      minimum: { type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "#F8696B" },
      midpoint: { formula: "=50%", type: Excel.ConditionalFormatColorCriterionType.percentile, color: "#FFEB84" },
      maximum: { type: Excel.ConditionalFormatColorCriterionType.highestValue, color: "#63BE7B" },
    };

    await context.sync();
  });
}

/**
 * Removes every conditional format from the demo table.
 */
export async function clearConditionalFormats(): Promise<void> {
  await Excel.run(async (context) => {
    const table = context.workbook.tables.getItem(DEMO_TABLE_NAME);
    table.getRange().conditionalFormats.clearAll();

    await context.sync();
  });
}
