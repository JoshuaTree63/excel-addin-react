import * as React from "react";
import { Input, Label, useId, makeStyles, Button, Textarea, Spinner } from "@fluentui/react-components";
import { getCellAddress } from "../taskpane";
import { useState } from "react";

const useStyles = makeStyles({
  root: {
    backgroundColor: "#f4f4f4",
    minHeight: "100vh",
    display: "flex",
    justifyContent: "center",
    flexDirection: "column",
  },
  form: {
    display: "flex",
    flexDirection: "column",
    gap: "5px",
    padding: "10px",
  },
  inputContainer: {
    display: "flex",
    flexDirection: "row",
    margin: "auto",
    gap: "10px",
  },
  inputWrapper: {
    display: "flex",
    flexDirection: "column",
  },
  formContent: {
    display: "flex",
    flexDirection: "column",
    gap: "10px",
  },
});

const App = () => {
  const styles = useStyles();
  // Read section
  const inputId = useId("input");
  const inputId2 = useId("input");
  const targetUrlId = useId("input");
  const [processedRows, setProcessedRows] = useState(0);
  const [currentWorksheet, setCurrentWorksheet] = useState(0);
  const [worksheetRows, setWorksheetRows] = useState(0);
  const [isReadLoading, setIsReadLoading] = useState(false);
  const [readErrorMessage, setReadErrorMessage] = useState<string>();
  const sheetPercent = worksheetRows > 0 ? ((processedRows / worksheetRows) * 100).toFixed(2) : 0;

  const incrementProcessedRows = () => {
    setProcessedRows((state) => state + 1);
  };

  const incrementWorksheet = () => {
    setCurrentWorksheet((state) => state + 1);
  };

  const resetProgress = () => {
    setCurrentWorksheet(0);
    setWorksheetRows(0);
    setProcessedRows(0);
  };

  const resetProcessedRows = () => {
    setProcessedRows(0);
  };

  const readWorkbook = async () => {
    resetProgress();

    try {
      const workbook = await Excel.run(async (context) => {
        var sheets = context.workbook.worksheets;
        sheets.load("items");
        await context.sync();

        const worksheets = [];

        for (var worksheet of sheets.items) {
          resetProcessedRows();
          incrementWorksheet();

          const usedRange = worksheet.getUsedRange();
          usedRange.load();
          const usedRow = usedRange.getLastRow();
          const usedCol = usedRange.getLastColumn();

          usedRow.load("rowIndex");
          usedCol.load("columnIndex");
          await context.sync();

          const range = worksheet.getRangeByIndexes(0, 0, usedRow.rowIndex + 1, usedCol.columnIndex);
          range.load([
            "values",
            "formulas",
            "formulasR1C1",
            "address",
            "numberFormat",
            "format/font",
            "rowCount",
            "columnCount",
          ]);

          await context.sync();

          const addresses = range.address.split(":")[0].split("!")[1]; // Get starting address
          const baseColumn = addresses.replace(/[0-9]/g, "");
          const baseRow = parseInt(addresses.replace(/[^0-9]/g, ""));
          const worksheetData = {
            name: worksheet.name,
            cells: {},
          };
          setWorksheetRows(range.rowCount);
          for (let row = 0; row < range.rowCount; row++) {
            incrementProcessedRows();
            for (let col = 0; col < range.columnCount; col++) {
              const cellAddress = getCellAddress(baseColumn, baseRow, row, col);
              const cellFont = range.getCell(row, col).format.font;
              const cellFill = range.getCell(row, col).format.fill;
              cellFont.load(["bold", "color", "italic", "name", "size", "underline", "backgroundColor"]);
              cellFill.load(["color"]);
              await context.sync();

              const cellData = {
                value: range.values[row][col],
                formula: range.formulas[row][col],
                formulaR1C1: range.formulasR1C1[row][col],
                address: cellAddress,
                rowIndex: row,
                columnIndex: col,                
                format: {
                  font: {
                    name: cellFont.name,
                    size: cellFont.size,
                    bold: cellFont.bold,
                    italic: cellFont.italic,
                    underline: cellFont.underline,
                    color: cellFont.color,
                  },
                  numberFormat: range.numberFormat[row][col],
                  backgroundColor: cellFill.color,
                },
              };
              worksheetData.cells[cellAddress] = cellData;
            }
          }
          worksheets.push(worksheetData);
        }
        console.log(worksheets);
        return worksheets;
      });
      return workbook;
    } catch (error) {
      setReadErrorMessage(error.message);
      throw error;
    }
  };

  const handleSubmitRead = async (e: React.FormEvent<HTMLFormElement>) => {
    setIsReadLoading(true);
    setReadErrorMessage("");
    e.preventDefault();
    const formData = new FormData(e.currentTarget);
    const param1 = formData.get("param_1");
    const param2 = formData.get("param_2");
    const targetUrl = formData.get("targetUrl");

    const workbook = await readWorkbook();

    await fetch(targetUrl as string, {
      method: "POST",
      body: JSON.stringify({
        worksheets: workbook,
        params: { param1, param2 },
      }),
    });
    setIsReadLoading(false);
  };

  // Write section
  const textAreaId = useId("textarea");
  const [isWriteLoading, setIsWriteLoading] = useState(false);
  const [writeErrorMessage, setWriteErrorMessage] = useState<string>();

  const writeWorkbook = async (jsonData: any) => {
    await Excel.run(async (context) => {
      try {
        const workbook = context.workbook;
        const worksheets = workbook.worksheets;

        // Loop through each worksheet in the JSON data
        for (const sheetData of jsonData.worksheets) {
          let worksheet = worksheets.getItemOrNullObject(sheetData.name);
          await context.sync();

          // If the worksheet does not exist, create it
          if (worksheet.isNullObject) {
            worksheet = worksheets.add(sheetData.name);
          }

          // Loop through each cell in the worksheet
          for (const cellAddress in sheetData.cells) {
            const cellData = sheetData.cells[cellAddress];
            const range = worksheet.getRange(cellAddress);

            // Set cell value
            if (cellData.value !== undefined) {
              range.values = [[cellData.value]];
            }

            // Set cell formula
            if (cellData.formula !== undefined) {
              range.formulas = [[cellData.formula]];
            }

            // Set cell formulaR1C1
            if (cellData.formulaR1C1 !== undefined) {
              range.formulasR1C1 = [[cellData.formulaR1C1]];
            }

            // Set cell format
            if (cellData.format) {
              const format = cellData.format;

              if (format.font) {
                range.format.font.name = format.font.name;
                range.format.font.size = format.font.size;
                range.format.font.bold = format.font.bold;
                range.format.font.italic = format.font.italic;
                range.format.font.underline = format.font.underline;
                range.format.font.strikethrough = format.font.strikethrough;
                range.format.font.color = format.font.color;
              }

              if (format.backgroundColor) {
                range.format.fill.color = format.backgroundColor;
              }

              if (format.numberFormat) {
                range.numberFormat = [[format.numberFormat]];
              }
            }
          }
        }
        await context.sync();
      } catch (error) {
        setWriteErrorMessage(error.message);
        throw error;
      }
    });
  };

  const handleSubmitWrite = async (e: React.FormEvent<HTMLFormElement>) => {
    setIsWriteLoading(true);
    e.preventDefault();
    const formData = new FormData(e.currentTarget);
    const textArea = formData.get("textArea");
    const jsonData = JSON.parse(textArea as string);
    await writeWorkbook(jsonData);
    setIsWriteLoading(false);
  };

  return (
    <div className={styles.root}>
      <form className={styles.form} onSubmit={handleSubmitRead}>
        <div className={styles.formContent}>
          <div className={styles.inputWrapper}>
            <Label htmlFor={targetUrlId}>Target URL</Label>
            <Input required id={targetUrlId} name="targetUrl" />
          </div>
          <div className={styles.inputContainer}>
            <div className={styles.inputWrapper}>
              <Label htmlFor={inputId}>Param 1</Label>
              <Input required id={inputId} name="param_1" />
            </div>
            <div className={styles.inputWrapper}>
              <Label htmlFor={inputId2}>Param 2</Label>
              <Input required id={inputId2} name="param_2" />
            </div>
          </div>
          <Button type="submit" appearance="primary" disabled={isReadLoading}>
            {isReadLoading ? (
              <span>
                &nbsp;{sheetPercent}%&nbsp;of&nbsp;sheet&nbsp;{currentWorksheet}
              </span>
            ) : (
              "Run"
            )}
          </Button>
          <span className="">{readErrorMessage}</span>
        </div>
      </form>
      <form className={styles.form} onSubmit={handleSubmitWrite}>
        <div className={styles.inputWrapper}>
          <Label htmlFor={targetUrlId}>JSON</Label>
          <Textarea resize="both" required id={textAreaId} name="textArea" />
        </div>
        <Button type="submit" appearance="primary" disabled={isWriteLoading}>
          {isWriteLoading ? <Spinner /> : "Run"}
        </Button>
        <span className="">{writeErrorMessage}</span>
      </form>
    </div>
  );
};

export default App;
