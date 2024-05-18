using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Devices;
using Microsoft.VisualBasic.Logging;
using NAudio.Wave;
using System.Text.RegularExpressions;
using Whisper.net;

namespace WhisperSingle
{
    public partial class Form1 : Form
    {
        private System.Windows.Forms.Timer timer;
        private int dotCount = 0;
        private XLWorkbook mainWorkbook;
        private IXLWorksheet mainWorksheet;
        public Form1()
        {
            InitializeComponent();
            InitializeProcessingTimer();
        }
        private void InitializeProcessingTimer()
        {
            timer = new System.Windows.Forms.Timer();
            timer.Interval = 1000; // Timer will tick every 1000 milliseconds (1 second)
            timer.Tick += new EventHandler(OnTimerTick);
            
        }
        private void OnTimerTick(object sender, EventArgs e)
        {
            UpdateFormTitle();
        }
        private async void btnConvert_Click(object sender, EventArgs e)
        {
            timer.Start();
            btnConvert.Enabled= false;
            var mp3Files = Directory.GetFiles(txtTargetFolder.Text, "*.mp3");
            lblTotalCount.Invoke((MethodInvoker)(() => lblTotalCount.Text = $"�`�ɮ׼�: {mp3Files.Count()}"));
            if (mp3Files.Length == 0)
            {
                MessageBox.Show("�Ы��wmp3���T�ӷ�.");
                return;
            }
            int mp3Index = 0;//�n�B�z�ĴX�Ӹ��
            var modelFilePath = GetModelFilePath("large");
            if (modelFilePath == null)
            {
                MessageBox.Show("�ҫ����|���s�b�A�бNggml-large-v2.bin�ҫ���J�D�{����Ƨ�.");
                return;
            }

            var whisperFactory = WhisperFactory.FromPath(modelFilePath);
            var processor = whisperFactory.CreateBuilder().WithLanguage("Mandarin").Build();
            bool isfinished = mp3Files.Count()< mp3Index+1;
            int processedCount = 0;
            while (!isfinished) 
            {

                if (ckOnlymatchExcel.Checked)
                {
                    await Task.Run(() => ProcessFilesMathcFileNumber(mp3Files, mp3Index));
                }
                else 
                {
                    //�}�l�B�z
                    var mp3FilePath = mp3Files[mp3Index];
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

                                //�P�_�O�_��r��������r
                                bool matchesKeyword = allMessages.ToString().Contains(txtKeyWord.Text);
                                var csvlines = File.ReadAllLines(txtSourceExcelPath.Text);
                                string MatchFilaPath = Path.Combine(txtMatchFilePath.Text, "MatchFile.xlsx");
                                string NotMatchFilePAth = Path.Combine(txtNotmatchFilePath.Text, "NotMatchFile.xlsx");
                                
                                //�P�_�O�_���s�bExcel�ɮ�
                                string workbookPath = matchesKeyword ? MatchFilaPath : NotMatchFilePAth;
                                var workbook = File.Exists(workbookPath) ? new XLWorkbook(workbookPath) : new XLWorkbook();
                                var worksheet = workbook.Worksheets.Count > 0 ? workbook.Worksheet(1) : workbook.Worksheets.Add("Sheet1");
                                int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
                                var newRow = worksheet.Row(lastRow + 1);

                                //Ū�����T�ɦW
                                string mp3fileName = Path.GetFileName(mp3FilePath);
                                string identifier = ExtractIdentifierFromFileName(mp3fileName);

                                bool IsMatchNumber = false;
                                int matchIndex = 0;
                                string ApplyDatetime = string.Empty;//�w���ɶ�
                                string TotalTime = string.Empty;//���X�ɶ�
                                string Tel = string.Empty;//�q��
                                string TalkLength = string.Empty;//�q�ܪ���
                                string status = string.Empty;//���A
                                string CustomLevel = string.Empty;//�Ȥᵥ��
                                string FileNumber = string.Empty;//�ɮ׽s��
                                string SubPhone = string.Empty;//��������
                                string audioLink = mp3FilePath;
                                int spiltLineIndex = 1;
                                try
                                {
                                    foreach (var csvline in csvlines)
                                    {
                                        var spiltLine = csvline.Split(",");
                                        if (spiltLine[6].Replace("'", "") == identifier)
                                        {
                                            IsMatchNumber = true;
                                            ApplyDatetime = spiltLine[0];
                                            TotalTime = spiltLine[1];
                                            Tel = spiltLine[2];
                                            TalkLength = spiltLine[3];
                                            status = spiltLine[4];
                                            CustomLevel = spiltLine[5];
                                            FileNumber = spiltLine[6];
                                            SubPhone = spiltLine[7];
                                            break;
                                        }
                                        matchIndex++;
                                        spiltLineIndex++;
                                    }
                                }
                                catch (Exception ex)
                                {

                                    Log("���Τ�r��o�Ϳ��~�A���ˬd��" + spiltLineIndex.ToString() + "���ƬO�_����");
                                    Log(ex.Message);
                                    if (ex.InnerException != null)
                                    {
                                        Log(ex.InnerException.Message);
                                    }
                                    IsMatchNumber = false;
                                    spiltLineIndex++;
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
                                    newRow.Cell(8).Value = "���T��";
                                    newRow.Cell(8).SetHyperlink(new XLHyperlink(audioLink));

                                    newRow.Cell(9).Value = allMessages.ToString();
                                    worksheet.Column(9).Width = 40;
                                }
                                workbook.SaveAs(workbookPath);
                                mainWorkbook?.Dispose();
                                workbook.Dispose();
                            }
                        }
                    }
                   
                }
                mp3Index++;
                isfinished = mp3Files.Count() < mp3Index + 1;
                processedCount++;
                lblProcessedCount.Invoke((MethodInvoker)(() => lblProcessedCount.Text = $"�w���ɮ�: {processedCount}"));
            }
            MessageBox.Show("����");
            btnConvert.Enabled = true;
            timer.Stop();
            this.Text = "���T���ɻ��U�{��";


        }

        /// <summary>
        /// �����r�u�������ɦW
        /// </summary>
        private void ProcessFilesMathcFileNumber(string[] mp3Files,int mp3Index) 
        {
            //Ū��csv�ӷ��ɮ�
            var csvlines = File.ReadAllLines(txtSourceExcelPath.Text);
            //�ǳƷs��Excel�ɮ�
            string workbookPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OnlyMatchNumberFile.xlsx");
            var workbook = File.Exists(workbookPath) ? new XLWorkbook(workbookPath) : new XLWorkbook();
            var worksheet = workbook.Worksheets.Count > 0 ? workbook.Worksheet(1) : workbook.Worksheets.Add("�M��");
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
            var newRow = worksheet.Row(lastRow + 1);

            //Ū�����T�ɦW
            var mp3FilePath = mp3Files[mp3Index];
            string mp3fileName = Path.GetFileName(mp3FilePath);
            string identifier = ExtractIdentifierFromFileName(mp3fileName);

            //�ɦW�����üg�J
            bool IsMatchNumber = false;
            int matchIndex = 0;
            string ApplyDatetime = string.Empty;//�w���ɶ�
            string TotalTime = string.Empty;//���X�ɶ�
            string Tel = string.Empty;//�q��
            string TalkLength = string.Empty;//�q�ܪ���
            string status = string.Empty;//���A
            string CustomLevel = string.Empty;//�Ȥᵥ��
            string FileNumber = string.Empty;//�ɮ׽s��
            string SubPhone = string.Empty;//��������
            string audioLink = mp3FilePath;
            int spiltLineIndex = 1;
            try
            {

                foreach (var csvline in csvlines)
                {
                    var spiltLine = csvline.Split(",");
                    if (spiltLine[6].Replace("'", "") == identifier)
                    {
                        IsMatchNumber = true;
                        ApplyDatetime = spiltLine[0];
                        TotalTime = spiltLine[1];
                        Tel = spiltLine[2];
                        TalkLength = spiltLine[3];
                        status = spiltLine[4];
                        CustomLevel = spiltLine[5];
                        FileNumber = spiltLine[6];
                        SubPhone = spiltLine[7];
                        break;
                    }
                    matchIndex++;
                    spiltLineIndex++;//�O����e���
                }
            }
            catch (Exception ex)
            {
                Log("���Τ�r��o�Ϳ��~�A���ˬd��" + spiltLineIndex.ToString() + "���ƬO�_����");
                Log(ex.Message);
                if (ex.InnerException != null)
                {
                    Log(ex.InnerException.Message);
                }
                IsMatchNumber = false;
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
                newRow.Cell(8).Value = "���T��";
                newRow.Cell(8).SetHyperlink(new XLHyperlink(audioLink));
                worksheet.Column(9).Width = 40;
            }
            workbook.SaveAs(workbookPath);
            mainWorkbook?.Dispose();
            workbook.Dispose();

        }

        //�̫�@�Ӧr�אּ0

        private string ExtractIdentifierFromFileName(string fileName)
        {
            // Extracts the numeric identifier from the file name, assuming it follows the last dash '-'
            var match = Regex.Match(fileName, @"-(\d+\.\d+)");
            return match.Success ? match.Groups[1].Value : null;
        }

        /// <summary>
        /// ���ܤ�r
        /// </summary>
        private void UpdateFormTitle()
        {
            string baseText = "�B�z��";
            string dots = new string('.', dotCount);
            this.Text = baseText + dots;

            dotCount = (dotCount + 1) % 5; // Reset after four dots
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
                folderBrowserDialog.Description = "��ܥؼи�Ƨ�";
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
                folderBrowserDialog.Description = "��ܲŦX����r���ɮצs�ɦ�m";
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
                folderBrowserDialog.Description = "��ܤ��ŦX����r���ɮצs�ɦ�m";
                DialogResult result = folderBrowserDialog.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(folderBrowserDialog.SelectedPath))
                {
                    txtNotmatchFilePath.Text = folderBrowserDialog.SelectedPath;
                }
            }
        }

        private void  Log(string message) 
        {
            string logDirectory = Path.Combine(Application.StartupPath, "Logs");
            if (!Directory.Exists(logDirectory))
            {
                Directory.CreateDirectory(logDirectory);
            }
            string logFileName = $"log_{DateTime.Now:yyyy-MM-dd}.txt";
            string logFilePath = Path.Combine(logDirectory, logFileName);
            string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}{Environment.NewLine}";
            File.AppendAllText(logFilePath, logMessage);
        }

        private void ckOnlymatchExcel_Click(object sender, EventArgs e)
        {
            if (ckOnlymatchExcel.Checked) 
            {
                txtKeyWord.Enabled = false;
                txtMatchFilePath.Enabled = false;
                txtNotmatchFilePath.Enabled = false;
                btnMatchFilePath.Enabled = false;
                btnNotMatchFilePath.Enabled = false;
            }
            else 
            {
                txtKeyWord.Enabled = true;
                txtMatchFilePath.Enabled = true;
                txtNotmatchFilePath.Enabled = true;
                btnMatchFilePath.Enabled = true;
                btnNotMatchFilePath.Enabled = true;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}