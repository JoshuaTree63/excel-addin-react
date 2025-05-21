using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Net.Http;

namespace ProJetsAddin
{
    public partial class MainForm : Form
    {
        private Application excelApp;
        private TabControl tabControl;
        private TabPage readTab;
        private TabPage writeTab;
        private TabPage fullPluginTab;
        private ProgressBar progressBar;
        private Label progressLabel;
        private TextBox targetUrlTextBox;
        private TextBox param1TextBox;
        private TextBox param2TextBox;
        private TextBox jsonTextBox;
        private Button readButton;
        private Button writeButton;
        private Label errorLabel;

        // Brand colors
        private readonly Color brandColor = Color.FromArgb(75, 13, 255); // #4B0DFF
        private readonly Color accentColor = Color.FromArgb(255, 107, 0); // #FF6B00
        private readonly Color backgroundColor = Color.FromArgb(45, 45, 45); // #2D2D2D
        private readonly Color cardColor = Color.FromArgb(61, 61, 61); // #3D3D3D

        public MainForm()
        {
            InitializeComponent();
            InitializeUI();
            excelApp = Globals.ThisAddIn.Application;
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // Form settings
            this.Text = "ProJets";
            this.Size = new Size(400, 600);
            this.BackColor = backgroundColor;
            this.ForeColor = Color.White;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ResumeLayout(false);
        }

        private void InitializeUI()
        {
            // Create tab control
            tabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Appearance = TabAppearance.FlatButtons,
                ItemSize = new Size(0, 1),
                SizeMode = TabSizeMode.Fixed
            };

            // Create tabs
            readTab = CreateReadTab();
            writeTab = CreateWriteTab();
            fullPluginTab = CreateFullPluginTab();

            tabControl.TabPages.Add(readTab);
            tabControl.TabPages.Add(writeTab);
            tabControl.TabPages.Add(fullPluginTab);

            // Add tab control to form
            this.Controls.Add(tabControl);

            // Create tab buttons
            CreateTabButtons();
        }

        private TabPage CreateReadTab()
        {
            var tab = new TabPage();
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(16)
            };

            // Create form controls
            targetUrlTextBox = new TextBox
            {
                Location = new Point(16, 16),
                Size = new Size(350, 20),
                BackColor = cardColor,
                ForeColor = Color.White
            };

            param1TextBox = new TextBox
            {
                Location = new Point(16, 56),
                Size = new Size(170, 20),
                BackColor = cardColor,
                ForeColor = Color.White
            };

            param2TextBox = new TextBox
            {
                Location = new Point(196, 56),
                Size = new Size(170, 20),
                BackColor = cardColor,
                ForeColor = Color.White
            };

            readButton = new Button
            {
                Location = new Point(16, 96),
                Size = new Size(350, 30),
                Text = "Read Workbook Data",
                BackColor = brandColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            readButton.Click += ReadButton_Click;

            progressBar = new ProgressBar
            {
                Location = new Point(16, 136),
                Size = new Size(350, 20),
                Style = ProgressBarStyle.Continuous
            };

            progressLabel = new Label
            {
                Location = new Point(16, 166),
                Size = new Size(350, 20),
                ForeColor = Color.White
            };

            errorLabel = new Label
            {
                Location = new Point(16, 196),
                Size = new Size(350, 40),
                ForeColor = Color.Red,
                AutoSize = true
            };

            // Add controls to panel
            panel.Controls.AddRange(new Control[] {
                new Label { Text = "Target URL", Location = new Point(16, 0), ForeColor = Color.White },
                targetUrlTextBox,
                new Label { Text = "Parameter 1", Location = new Point(16, 40), ForeColor = Color.White },
                param1TextBox,
                new Label { Text = "Parameter 2", Location = new Point(196, 40), ForeColor = Color.White },
                param2TextBox,
                readButton,
                progressBar,
                progressLabel,
                errorLabel
            });

            tab.Controls.Add(panel);
            return tab;
        }

        private TabPage CreateWriteTab()
        {
            var tab = new TabPage();
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(16)
            };

            jsonTextBox = new TextBox
            {
                Location = new Point(16, 16),
                Size = new Size(350, 200),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                BackColor = cardColor,
                ForeColor = Color.White
            };

            writeButton = new Button
            {
                Location = new Point(16, 226),
                Size = new Size(350, 30),
                Text = "Write to Workbook",
                BackColor = brandColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            writeButton.Click += WriteButton_Click;

            panel.Controls.AddRange(new Control[] {
                new Label { Text = "JSON Data", Location = new Point(16, 0), ForeColor = Color.White },
                jsonTextBox,
                writeButton
            });

            tab.Controls.Add(panel);
            return tab;
        }

        private TabPage CreateFullPluginTab()
        {
            var tab = new TabPage();
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(16)
            };

            // Add model selection buttons
            string[] models = {
                "Photovoltaic (PV) Solar Farms",
                "Toll Road Concessions",
                "Data Center",
                "High-Speed Rail Infrastructure",
                "Autonomous Freight Corridors",
                "LNG Regasification Terminals",
                "Smart Highway Infrastructure",
                "Desalination Plants",
                "Floating Solar Installations",
                "Seaport Logistics Hubs"
            };

            int buttonHeight = 40;
            int spacing = 10;
            int currentY = 16;

            foreach (string model in models)
            {
                var button = new Button
                {
                    Location = new Point(16, currentY),
                    Size = new Size(350, buttonHeight),
                    Text = model,
                    BackColor = Color.White,
                    ForeColor = Color.Black,
                    FlatStyle = FlatStyle.Flat,
                    TextAlign = ContentAlignment.MiddleLeft
                };
                button.Click += ModelButton_Click;
                panel.Controls.Add(button);
                currentY += buttonHeight + spacing;
            }

            tab.Controls.Add(panel);
            return tab;
        }

        private void CreateTabButtons()
        {
            var buttonPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 40,
                BackColor = backgroundColor
            };

            string[] tabNames = { "Read Data", "Write Data", "Full Plugin" };
            int buttonWidth = 100;
            int spacing = 10;

            for (int i = 0; i < tabNames.Length; i++)
            {
                var button = new Button
                {
                    Text = tabNames[i],
                    Location = new Point(i * (buttonWidth + spacing), 5),
                    Size = new Size(buttonWidth, 30),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = backgroundColor,
                    ForeColor = Color.White
                };

                int tabIndex = i;
                button.Click += (s, e) => tabControl.SelectedIndex = tabIndex;
                buttonPanel.Controls.Add(button);
            }

            this.Controls.Add(buttonPanel);
        }

        private async void ReadButton_Click(object sender, EventArgs e)
        {
            try
            {
                readButton.Enabled = false;
                errorLabel.Text = "";
                progressBar.Value = 0;
                progressLabel.Text = "Reading workbook...";

                var workbook = await ReadWorkbookAsync();
                var jsonData = JsonConvert.SerializeObject(workbook);

                using (var client = new HttpClient())
                {
                    var content = new StringContent(jsonData);
                    await client.PostAsync(targetUrlTextBox.Text, content);
                }

                progressLabel.Text = "Completed!";
            }
            catch (Exception ex)
            {
                errorLabel.Text = $"Error: {ex.Message}";
            }
            finally
            {
                readButton.Enabled = true;
            }
        }

        private async void WriteButton_Click(object sender, EventArgs e)
        {
            try
            {
                writeButton.Enabled = false;
                errorLabel.Text = "";

                var jsonData = JsonConvert.DeserializeObject(jsonTextBox.Text);
                await WriteWorkbookAsync(jsonData);

                MessageBox.Show("Data written successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                errorLabel.Text = $"Error: {ex.Message}";
            }
            finally
            {
                writeButton.Enabled = true;
            }
        }

        private void ModelButton_Click(object sender, EventArgs e)
        {
            var button = sender as Button;
            MessageBox.Show($"Selected model: {button.Text}", "Model Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private async Task<object> ReadWorkbookAsync()
        {
            return await Task.Run(() =>
            {
                var workbook = excelApp.ActiveWorkbook;
                var worksheets = new System.Collections.Generic.List<object>();

                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    var usedRange = worksheet.UsedRange;
                    var data = new
                    {
                        name = worksheet.Name,
                        cells = new System.Collections.Generic.Dictionary<string, object>()
                    };

                    for (int row = 1; row <= usedRange.Rows.Count; row++)
                    {
                        for (int col = 1; col <= usedRange.Columns.Count; col++)
                        {
                            var cell = usedRange.Cells[row, col];
                            var address = cell.Address;
                            var cellData = new
                            {
                                value = cell.Value2,
                                formula = cell.Formula,
                                formulaR1C1 = cell.FormulaR1C1,
                                format = new
                                {
                                    font = new
                                    {
                                        name = cell.Font.Name,
                                        size = cell.Font.Size,
                                        bold = cell.Font.Bold,
                                        italic = cell.Font.Italic,
                                        underline = cell.Font.Underline,
                                        color = cell.Font.Color
                                    },
                                    backgroundColor = cell.Interior.Color,
                                    numberFormat = cell.NumberFormat
                                }
                            };
                            data.cells[address] = cellData;
                        }
                    }
                    worksheets.Add(data);
                }

                return new { worksheets };
            });
        }

        private async Task WriteWorkbookAsync(object jsonData)
        {
            await Task.Run(() =>
            {
                var workbook = excelApp.ActiveWorkbook;
                dynamic data = jsonData;

                foreach (var worksheetData in data.worksheets)
                {
                    Worksheet worksheet = null;
                    try
                    {
                        worksheet = workbook.Worksheets[worksheetData.name.ToString()];
                    }
                    catch
                    {
                        worksheet = workbook.Worksheets.Add();
                        worksheet.Name = worksheetData.name.ToString();
                    }

                    foreach (var cellData in worksheetData.cells)
                    {
                        var address = cellData.Name;
                        var cell = worksheet.Range[address];
                        var cellInfo = cellData.Value;

                        if (cellInfo.value != null)
                            cell.Value2 = cellInfo.value;
                        if (cellInfo.formula != null)
                            cell.Formula = cellInfo.formula;
                        if (cellInfo.formulaR1C1 != null)
                            cell.FormulaR1C1 = cellInfo.formulaR1C1;

                        if (cellInfo.format != null)
                        {
                            if (cellInfo.format.font != null)
                            {
                                cell.Font.Name = cellInfo.format.font.name;
                                cell.Font.Size = cellInfo.format.font.size;
                                cell.Font.Bold = cellInfo.format.font.bold;
                                cell.Font.Italic = cellInfo.format.font.italic;
                                cell.Font.Underline = cellInfo.format.font.underline;
                                cell.Font.Color = cellInfo.format.font.color;
                            }

                            if (cellInfo.format.backgroundColor != null)
                                cell.Interior.Color = cellInfo.format.backgroundColor;

                            if (cellInfo.format.numberFormat != null)
                                cell.NumberFormat = cellInfo.format.numberFormat;
                        }
                    }
                }
            });
        }
    }
} 