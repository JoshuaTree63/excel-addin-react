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
const backgroundColor = "#2D2D2D"; // Dark gray background
const cardColor = "#3D3D3D"; // Slightly lighter gray for cards
const inputColor = "#4D4D4D"; // Even lighter gray for inputs

const useStyles = makeStyles({
  root: {
    backgroundColor: backgroundColor,
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
    backgroundColor: cardColor,
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
    backgroundColor: inputColor,
    color: "#ffffff",
    ...shorthands.border("1px", "solid", "#666666"),
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
  const [fullPluginMode, setFullPluginMode] = useState("main"); // Tracks the mode within Full Plugin tab

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
  
  const readRange = async (worksheetData, range) => {
    range.load([
            "columnCount",
            "formulasR1C1",
            "numberFormat",
            "rowCount",
          ]);
          var properties = range.getCellProperties({
            address: true,
            format: {
              fill: {
                color: true
              },
              font: {
                bold: true,
                color: true,
                italic: true,
                name: true,
                size: true,
                strikethrough: true,
                underline: true
              }
            }
          });
          await range.context.sync();

          for (let row = 0; row < range.rowCount; row++) {
            for (let col = 0; col < range.columnCount; col++) {
              const cellData = {
                formulaR1C1: range.formulasR1C1[row][col],
                address: properties.value[row][col].address,
                rowIndex: row,
                columnIndex: col,
                format: {
                  font: {
                    bold: properties.value[row][col].format.font.bold,
                    color: properties.value[row][col].format.font.color,
                    italic: properties.value[row][col].format.font.italic,
                    name: properties.value[row][col].format.font.name,
                    size: properties.value[row][col].format.font.size,
                    strikethrough: properties.value[row][col].format.font.strikethrough,
                    underline: properties.value[row][col].format.font.underline,
                  },
                  numberFormat: range.numberFormat[row][col],
                  backgroundColor: properties.value[row][col].format.fill.color,
                },
              };
              worksheetData.cells[properties.value[row][col].address] = cellData;

            }
          }

          return worksheetData;
  }

    try {
      const workbook = await Excel.run(async (context) => {
        var sheets = context.workbook.worksheets;

        const worksheets = [];
        var worksheet = sheets.getFirst();

        do {
          await context.sync();
          if (worksheet.isNullObject) {
            break;
          }
          resetProcessedRows();
          incrementWorksheet();
          worksheet.load("name")

          var usedRange = worksheet.getUsedRange();
          var lastColumn = usedRange.getLastColumn();
          var lastRow = usedRange.getLastRow();
          lastColumn.load("columnIndex");
          lastRow.load("RowIndex");
          await context.sync();

          var worksheetData = {
            name: worksheet.name,
            cells: {},
          };

          var rowsCount = lastRow.rowIndex;
          var columnsCount = lastColumn.columnIndex;
          usedRange.untrack();
          lastColumn.untrack();
          lastRow.untrack();

          var currentRow = 0;
          var currentColumn = 0;
          setWorksheetRows(rowsCount);

          const COLUMN_CHUNK_SIZE = 100;
          const ROW_CHUNK_SIZE = 2;

          while (currentRow < rowsCount) {
            let lastRowIndex = Math.min(currentRow + ROW_CHUNK_SIZE, rowsCount);
            currentColumn = 0;
            while (currentColumn < columnsCount) {
              let lastColumnIndex = Math.min(currentColumn + COLUMN_CHUNK_SIZE, columnsCount);
              const range = worksheet.getRangeByIndexes(currentRow, currentColumn, lastRowIndex - currentRow, lastColumnIndex - currentColumn);
              worksheetData = await readRange(worksheetData, range);
              range.untrack();
              await context.sync();

              currentColumn = lastColumnIndex;
            }
            currentRow = lastRowIndex;
            incrementProcessedRows();
            incrementProcessedRows();
          }

          if (lastRow.rowIndex == 0)
          {
            continue;
          }

          worksheets.push(worksheetData);
        } while (worksheet = worksheet.getNextOrNullObject());
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

  // Handler for model selection
  const handleSelectModel = () => {
    setFullPluginMode("modelSelection");
  };

  // Handler to go back to main full plugin view
  const handleBackToFullPlugin = () => {
    setFullPluginMode("main");
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
          <Tab className={styles.brandTab} value="full">
            <span className={styles.tabLabel}>Full Plugin</span>
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

      {selectedTab === "full" && fullPluginMode === "main" && (
        <div style={{ padding: "16px" }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr", gap: "12px" }}>
            <Button 
              appearance="outline"
              style={{ 
                backgroundColor: "white", 
                color: "#000000", 
                textAlign: "left",
                justifyContent: "flex-start"
              }}
              icon={<span style={{ marginRight: "8px" }}>🔍</span>}
              onClick={handleSelectModel}
            >
              Select Model
            </Button>
            
            <Button 
              appearance="outline"
              style={{ 
                backgroundColor: "white", 
                color: "#000000", 
                textAlign: "left",
                justifyContent: "flex-start"
              }}
              icon={<span style={{ marginRight: "8px" }}>➕</span>}
            >
              Insert Module
            </Button>
            
            <Button 
              appearance="outline"
              style={{ 
                backgroundColor: "white", 
                color: "#000000", 
                textAlign: "left",
                justifyContent: "flex-start"
              }}
              icon={<span style={{ marginRight: "8px" }}>➖</span>}
            >
              Remove Module
            </Button>
            
            <Button 
              appearance="outline"
              style={{ 
                backgroundColor: "white", 
                color: "#000000", 
                textAlign: "left",
                justifyContent: "flex-start"
              }}
              icon={<span style={{ marginRight: "8px" }}>🐍</span>}
            >
              Python AI Assistant
            </Button>
            
            <Button 
              appearance="outline"
              style={{ 
                backgroundColor: "white", 
                color: "#000000", 
                textAlign: "left",
                justifyContent: "flex-start"
              }}
              icon={<span style={{ marginRight: "8px" }}>📊</span>}
            >
              Analyze Government Tenders
            </Button>
          </div>
        </div>
      )}

      {selectedTab === "full" && fullPluginMode === "modelSelection" && (
        <div style={{ padding: "16px" }}>
          <div style={{ marginBottom: "16px", display: "flex", alignItems: "center" }}>
            <Button 
              appearance="subtle"
              onClick={handleBackToFullPlugin}
              style={{ marginRight: "12px" }}
            >
              ← Back
            </Button>
            <Text weight="semibold">Select a Model</Text>
          </div>
          
          <div style={{ display: "grid", gridTemplateColumns: "1fr", gap: "12px" }}>
            <Button 
              appearance="outline"
              style={{ 
                backgroundColor: "white", 
                color: "#000000", 
                textAlign: "left",
                justifyContent: "flex-start"
              }}
            >
              Photovoltaic (PV) Solar Farms
            </Button>
            
            <Button 
              appearance="outline"
              style={{ 
                backgroundColor: "white", 
                color: "#000000", 
                textAlign: "left",
                justifyContent: "flex-start"
              }}
            >
              Toll Road Concessions
            </Button>
            
            <Button 
              appearance="outline"
              style={{ 
                backgroundColor: "white", 
                color: "#000000", 
                textAlign: "left",
                justifyContent: "flex-start"
              }}
            >
              Data Center
            </Button>
            
            <Button 
              appearance="outline"
              style={{ 
                backgroundColor: "white", 
                color: "#000000", 
                textAlign: "left",
                justifyContent: "flex-start"
              }}
            >
              High-Speed Rail Infrastructure
            </Button>
            
            <Button 
              appearance="outline"
              style={{ 
                backgroundColor: "white", 
                color: "#000000", 
                textAlign: "left",
                justifyContent: "flex-start"
              }}
            >
              Autonomous Freight Corridors
            </Button>
            
            <Button 
              appearance="outline"
              style={{ 
                backgroundColor: "white", 
                color: "#000000", 
                textAlign: "left",
                justifyContent: "flex-start"
              }}
            >
              LNG Regasification Terminals
            </Button>
            
            <Button 
              appearance="outline"
              style={{ 
                backgroundColor: "white", 
                color: "#000000", 
                textAlign: "left",
                justifyContent: "flex-start"
              }}
            >
              Smart Highway Infrastructure
            </Button>
            
            <Button 
              appearance="outline"
              style={{ 
                backgroundColor: "white", 
                color: "#000000", 
                textAlign: "left",
                justifyContent: "flex-start"
              }}
            >
              Desalination Plants
            </Button>
            
            <Button 
              appearance="outline"
              style={{ 
                backgroundColor: "white", 
                color: "#000000", 
                textAlign: "left",
                justifyContent: "flex-start"
              }}
            >
              Floating Solar Installations
            </Button>
            
            <Button 
              appearance="outline"
              style={{ 
                backgroundColor: "white", 
                color: "#000000", 
                textAlign: "left",
                justifyContent: "flex-start"
              }}
            >
              Seaport Logistics Hubs
            </Button>
          </div>
        </div>
      )}
    </div>
  );
};

export default App;
