using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Devices;
using NAudio.Wave;
using System.Text.RegularExpressions;
using Whisper.net;

namespace WhisperSingle
{
    public partial class Form1 : Form
    {
        private XLWorkbook mainWorkbook;
        private IXLWorksheet mainWorksheet;
        public Form1()
        {
            InitializeComponent();
        }

        private async void btnConvert_Click(object sender, EventArgs e)
        {
            btnConvert.Enabled= false;
            var mp3Files = Directory.GetFiles(txtTargetFolder.Text, "*.mp3");
            lblTotalCount.Invoke((MethodInvoker)(() => lblTotalCount.Text = $"Total files: {mp3Files.Count()}"));
            if (mp3Files.Length == 0)
            {
                MessageBox.Show("No MP3 files found in the specified folder.");
                return;
            }
            int mp3Index = 0;//要處理第幾個資料
            var modelFilePath = GetModelFilePath("large");
            if (modelFilePath == null)
            {
                MessageBox.Show("The specified model file path does not exist.");
                return;
            }

            var whisperFactory = WhisperFactory.FromPath(modelFilePath);
            var processor = whisperFactory.CreateBuilder().WithLanguage("Mandarin").Build();
            bool isfinished = mp3Files.Count()< mp3Index+1;
            int processedCount = 0;
            while (!isfinished) 
            {
                //開始處理
                var mp3FilePath= mp3Files[mp3Index];
                using (var mp3Stream = File.OpenRead(mp3FilePath))
                {
                    using (var mp3Reader = new Mp3FileReader(mp3FilePath))
                    {
                        // Resample the MP3 stream to 16 kHz
                        var resampledStream = new WaveFormatConversionStream(new WaveFormat(16000, 16, mp3Reader.WaveFormat.Channels), mp3Reader);

                        // Use a temporary MemoryStream to write the WAV data
                        using (var tempStream = new MemoryStream())
                        {
                            using (var waveWriter = new WaveFileWriter(tempStream, resampledStream.WaveFormat))
                            {
                                resampledStream.CopyTo(waveWriter);
                            }

                            // Create a new MemoryStream for the output
                            var outputStream = new MemoryStream(tempStream.ToArray());

                            var allMessages = new System.Text.StringBuilder();
                            await foreach (var result in processor.ProcessAsync(outputStream))
                            {
                                allMessages.AppendLine(result.Text);
                            }
                            bool matchesKeyword = allMessages.ToString().Contains(txtKeyWord.Text);
                            var csvlines = File.ReadAllLines(txtSourceExcelPath.Text);
                            string MatchFilaPath = Path.Combine(txtMatchFilePath.Text, "MatchFile.xlsx");
                            string NotMatchFilePAth = Path.Combine(txtNotmatchFilePath.Text, "NotMatchFile.xlsx");
                            string workbookPath = matchesKeyword ? MatchFilaPath : NotMatchFilePAth;
                            var workbook = File.Exists(workbookPath) ? new XLWorkbook(workbookPath) : new XLWorkbook();
                            var worksheet = workbook.Worksheets.Count > 0 ? workbook.Worksheet(1) : workbook.Worksheets.Add("Sheet1");
                            
                            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
                            var newRow = worksheet.Row(lastRow + 1);

                            string mp3fileName = Path.GetFileName(mp3FilePath);
                            string identifier = ExtractIdentifierFromFileName(mp3fileName);

                            bool IsMatchNumber = false;
                            int matchIndex = 0;
                            string ApplyDatetime = string.Empty;//預約時間
                            string TotalTime = string.Empty;//撥出時間
                            string Tel = string.Empty;//電話
                            string TalkLength = string.Empty;//通話長度
                            string status = string.Empty;//狀態
                            string CustomLevel = string.Empty;//客戶等級
                            string FileNumber = string.Empty;//檔案編號
                            string SubPhone = string.Empty;//應答分機
                            string audioLink = mp3FilePath;
                            foreach (var csvline in csvlines)
                            {
                                var spiltLine = csvline.Split(",");
                                if (spiltLine[6].Replace("'","")== identifier)
                                {
                                    IsMatchNumber = true;
                                    ApplyDatetime = spiltLine[0];
                                    TotalTime = spiltLine[1];
                                    Tel = spiltLine[2];
                                    TalkLength= spiltLine[3];
                                    status= spiltLine[4];
                                    CustomLevel= spiltLine[5];
                                    FileNumber= spiltLine[6];
                                    SubPhone= spiltLine[7];
                                    break;
                                }
                                matchIndex++;
                            }

                            if (IsMatchNumber)
                            {
                                newRow.Cell(1).Value = ApplyDatetime;
                                newRow.Cell(2).Value = TotalTime;
                                newRow.Cell(3).Value = Tel;
                                newRow.Cell(4).Value = TalkLength;
                                newRow.Cell(5).Value = CustomLevel;
                                newRow.Cell(6).Value = FileNumber;
                                newRow.Cell(7).Value = SubPhone;
                                newRow.Cell(8).Value = "音訊檔";
                                newRow.Cell(8).SetHyperlink(new XLHyperlink(audioLink));
                                
                                newRow.Cell(9).Value = allMessages.ToString();
                                worksheet.Column(9).Width = 40;
                                workbook.SaveAs(workbookPath);
                            }
                            workbook.SaveAs(workbookPath);
                            mainWorkbook?.Dispose();
                            workbook.Dispose();
                        }
                    }
                }
                mp3Index++;
                isfinished = mp3Files.Count() < mp3Index + 1;
                processedCount++;
                lblProcessedCount.Invoke((MethodInvoker)(() => lblProcessedCount.Text = $"Processed: {processedCount}"));
            }
            MessageBox.Show("完成");
            btnConvert.Enabled = true;
        }
        private string RemoveTrailingZeros(string input)
        {
            if (decimal.TryParse(input, out decimal num))
            {
                return num.ToString("G29");  // "G29" format specifier to avoid scientific notation
            }
            return input;
        }
        private string FormatIdentifier(string identifier)
        {
            // Check if the identifier contains a decimal point
            if (identifier.Contains("."))
            {
                var parts = identifier.Split('.');
                string integerPart = parts[0];
                string decimalPart = parts[1];

                // Pad the decimal part with zeros to ensure it is exactly 6 digits
                decimalPart = decimalPart.PadRight(6, '0').Substring(0, 6);

                // Combine the integer and decimal parts
                return integerPart + "." + decimalPart;
            }

            // If no decimal point is found, append ".000000"
            return identifier + ".000000";
        }
        //最後一個字改為0
        private string ChangeLastDigitToZero(string numberString)
        {
            // Check if the string contains a decimal point
            int decimalIndex = numberString.LastIndexOf('.');
            if (decimalIndex != -1 && decimalIndex < numberString.Length - 1)
            {
                // There is a decimal point, and it's not at the end of the string
                char[] chars = numberString.ToCharArray();
                if (numberString.Split(".")[1].Length == 6)
                {
                    chars[chars.Length - 1] = '0';
                }
                // Change the last character to '0'
                return new string(chars);
            }
            return numberString; // Return the original string if no decimal point is found
        }
        private string ExtractIdentifierFromFileName(string fileName)
        {
            // Extracts the numeric identifier from the file name, assuming it follows the last dash '-'
            var match = Regex.Match(fileName, @"-(\d+\.\d+)");
            return match.Success ? match.Groups[1].Value : null;
        }
        private string GetModelFilePath(string modelName)
        {
            var modelDictionary = new Dictionary<string, string>
            {
                { "large", "ggml-large-v2.bin" },
                { "medium", "ggml-medium.bin" },
                { "small", "ggml-small.bin" },
                { "tiny", "ggml-tiny.bin" }
            };

            if (modelDictionary.TryGetValue(modelName, out var modelPath))
            {
                return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, modelPath);
            }

            return null; // Handle the case where the model is not found appropriately
        }

        private void btnAudioFilePath_Click(object sender, EventArgs e)
        {
            using (var folderBrowserDialog = new FolderBrowserDialog())
            {
                // Set the initial directory and description for the dialog
                folderBrowserDialog.Description = "選擇目標資料夾";
                folderBrowserDialog.RootFolder = Environment.SpecialFolder.MyComputer;

                // Show the dialog and check if the user selected a folder
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    // Set the selected folder path to the TextBox
                    txtTargetFolder.Text = folderBrowserDialog.SelectedPath;
                }
            }
        }

        private void btnSourceExcelPath_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Set the filter to Excel files only
            //openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;.csv";

            // Set the title of the dialog window
            openFileDialog.Title = "Select an Excel File";

            // Show the dialog and check if the result is OK
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Get the selected file name
                string fileName = openFileDialog.FileName;
                txtSourceExcelPath.Text = fileName;
            }
        }

        private void btnMatchFilePath_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "選擇符合關鍵字之檔案存檔位置";
                DialogResult result = folderBrowserDialog.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(folderBrowserDialog.SelectedPath))
                {
                    txtMatchFilePath.Text = folderBrowserDialog.SelectedPath;
                }
            }
        }

        private void btnNotMatchFilePath_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "選擇不符合關鍵字之檔案存檔位置";
                DialogResult result = folderBrowserDialog.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(folderBrowserDialog.SelectedPath))
                {
                    txtNotmatchFilePath.Text = folderBrowserDialog.SelectedPath;
                }
            }
        }
    }
}