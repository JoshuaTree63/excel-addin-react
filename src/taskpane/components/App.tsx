import * as React from "react";
import { 
  Input, 
  Label, 
  useId, 
  makeStyles, 
  Button, 
  Textarea, 
  Spinner, 
  TabList, 
  Tab, 
  tokens, 
  Card, 
  Title3, 
  Text,
  ProgressBar,
  shorthands
} from "@fluentui/react-components";
import { getCellAddress } from "../taskpane";
import { useState } from "react";

// ProJets brand color
const brandColor = "#4B0DFF";
const accentColor = "#FF6B00"; // Orange accent color from the image

const useStyles = makeStyles({
  root: {
    backgroundColor: "#000000",
    color: "#ffffff",
    minHeight: "100vh",
    display: "flex",
    flexDirection: "column",
    padding: "16px",
  },
  header: {
    marginBottom: "16px",
    color: brandColor,
    borderBottom: `2px solid ${brandColor}`,
    paddingBottom: "8px"
  },
  tabContainer: {
    marginBottom: "16px",
  },
  brandTab: {
    color: "#ffffff",
    "&:hover": {
      color: accentColor,
    },
    "&[data-selected]": {
      color: accentColor,
      borderBottomColor: accentColor,
    },
  },
  tabLabel: {
    color: accentColor,
    fontWeight: "600",
  },
  form: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    width: "100%",
  },
  card: {
    padding: "16px",
    marginBottom: "16px",
    borderTop: `3px solid ${brandColor}`,
    backgroundColor: "#121212",
    color: "#ffffff",
  },
  inputContainer: {
    display: "flex",
    flexDirection: "row",
    gap: "12px",
    width: "100%",
  },
  inputWrapper: {
    display: "flex",
    flexDirection: "column",
    width: "100%",
    gap: "4px",
  },
  label: {
    color: "#ffffff",
  },
  input: {
    backgroundColor: "#333333",
    color: "#ffffff",
    ...shorthands.border("1px", "solid", "#555555"),
    ":focus-within": {
      ...shorthands.border("1px", "solid", brandColor),
    },
  },
  formContent: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    width: "100%",
  },
  buttonContainer: {
    marginTop: "16px",
  },
  primaryButton: {
    backgroundColor: brandColor,
    color: "white",
    "&:hover": {
      backgroundColor: `${brandColor}dd`,
    },
    "&:active": {
      backgroundColor: `${brandColor}bb`,
    },
  },
  errorMessage: {
    color: "#ff6b6b",
    fontSize: tokens.fontSizeBase200,
    marginTop: "8px",
  },
  progressContainer: {
    marginTop: "16px",
  },
  progressText: {
    fontSize: tokens.fontSizeBase200,
    marginBottom: "8px",
    color: "#ffffff",
  },
  brandTitle: {
    fontSize: "28px",
    fontWeight: "bold",
    color: brandColor,
  }
});

const App = () => {
  const styles = useStyles();
  const [selectedTab, setSelectedTab] = useState("read");
  
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

  // Write section
  const textAreaId = useId("textarea");
  const [isWriteLoading, setIsWriteLoading] = useState(false);
  const [writeErrorMessage, setWriteErrorMessage] = useState<string>();

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

    try {
      const workbook = await readWorkbook();

      await fetch(targetUrl as string, {
        method: "POST",
        body: JSON.stringify({
          worksheets: workbook,
          params: { param1, param2 },
        }),
      });
    } catch (error) {
      setReadErrorMessage(error.message || "An error occurred during the read operation");
    } finally {
      setIsReadLoading(false);
    }
  };

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
    setWriteErrorMessage("");
    e.preventDefault();
    
    try {
      const formData = new FormData(e.currentTarget);
      const textArea = formData.get("textArea");
      const jsonData = JSON.parse(textArea as string);
      await writeWorkbook(jsonData);
    } catch (error) {
      setWriteErrorMessage(error.message || "An error occurred during the write operation");
    } finally {
      setIsWriteLoading(false);
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Title3 className={styles.brandTitle}>ProJets</Title3>
      </div>
      
      <div className={styles.tabContainer}>
        <TabList selectedValue={selectedTab} onTabSelect={(_, data) => setSelectedTab(data.value as string)}>
          <Tab className={styles.brandTab} value="read">
            <span className={styles.tabLabel}>Read Data</span>
          </Tab>
          <Tab className={styles.brandTab} value="write">
            <span className={styles.tabLabel}>Write Data</span>
          </Tab>
        </TabList>
      </div>
      
      {selectedTab === "read" && (
        <Card className={styles.card}>
          <form className={styles.form} onSubmit={handleSubmitRead}>
            <div className={styles.formContent}>
              <div className={styles.inputWrapper}>
                <Label className={styles.label} htmlFor={targetUrlId}>Target URL</Label>
                <Input 
                  required 
                  id={targetUrlId} 
                  name="targetUrl" 
                  placeholder="https://your-api-endpoint.com/data"
                  className={styles.input}
                />
              </div>
              
              <div className={styles.inputContainer}>
                <div className={styles.inputWrapper}>
                  <Label className={styles.label} htmlFor={inputId}>Parameter 1</Label>
                  <Input 
                    required 
                    id={inputId} 
                    name="param_1" 
                    placeholder="Enter first parameter"
                    className={styles.input}
                  />
                </div>
                <div className={styles.inputWrapper}>
                  <Label className={styles.label} htmlFor={inputId2}>Parameter 2</Label>
                  <Input 
                    required 
                    id={inputId2} 
                    name="param_2" 
                    placeholder="Enter second parameter"
                    className={styles.input}
                  />
                </div>
              </div>
              
              <div className={styles.buttonContainer}>
                <Button 
                  type="submit" 
                  appearance="primary" 
                  disabled={isReadLoading}
                  className={styles.primaryButton}
                >
                  {isReadLoading ? "Processing..." : "Read Workbook Data"}
                </Button>
              </div>
              
              {isReadLoading && (
                <div className={styles.progressContainer}>
                  <Text className={styles.progressText}>
                    Processing Sheet {currentWorksheet}: {processedRows} of {worksheetRows} rows ({sheetPercent}%)
                  </Text>
                  <ProgressBar value={Number(sheetPercent) / 100} />
                </div>
              )}
              
              {readErrorMessage && (
                <Text className={styles.errorMessage}>{readErrorMessage}</Text>
              )}
            </div>
          </form>
        </Card>
      )}
      
      {selectedTab === "write" && (
        <Card className={styles.card}>
          <form className={styles.form} onSubmit={handleSubmitWrite}>
            <div className={styles.formContent}>
              <div className={styles.inputWrapper}>
                <Label className={styles.label} htmlFor={textAreaId}>JSON Data</Label>
                <Textarea 
                  resize="vertical" 
                  required 
                  id={textAreaId} 
                  name="textArea" 
                  placeholder="Paste your JSON data here..."
                  style={{ minHeight: "200px" }}
                  className={styles.input}
                />
              </div>
              
              <div className={styles.buttonContainer}>
                <Button 
                  type="submit" 
                  appearance="primary" 
                  disabled={isWriteLoading}
                  className={styles.primaryButton}
                >
                  {isWriteLoading ? (
                    <>
                      <Spinner size="tiny" style={{ marginRight: "8px" }} />
                      Writing Data...
                    </>
                  ) : (
                    "Write to Workbook"
                  )}
                </Button>
              </div>
              
              {writeErrorMessage && (
                <Text className={styles.errorMessage}>{writeErrorMessage}</Text>
              )}
            </div>
          </form>
        </Card>
      )}
    </div>
  );
};

export default App;
