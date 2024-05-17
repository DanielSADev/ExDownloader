using ClosedXML.Excel;

namespace ExcelImageDownloader
{
    public partial class Form1 : Form
    {
        private string ExFilePath { get; set; }
        private string LoadDirectory { get; set; }
        

        public Form1()
        {
            InitializeComponent();
        }


        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ExFilePath = openFileDialog.FileName;
                    txtFilePath.Text = ExFilePath;
                    lblStatus.Text = "Arquivo Excel selecionado.";
                }
            }
        }

        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    LoadDirectory = folderBrowserDialog.SelectedPath;
                    txtFolderPath.Text = LoadDirectory;
                    lblStatus.Text = "Pasta de destino selecionada.";
                }
            }
        }

        private async void btnDownload_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(ExFilePath) || string.IsNullOrEmpty(LoadDirectory))
            {
                MessageBox.Show("Por favor, selecione um arquivo Excel e um diretório para salvar as imagens.");
                return;
            }

            try
            {
                lblStatus.Text = "Iniciando o download das imagens...";
                progressBar.Value = 0;

                
                int totalImages = CountHyperlinks(ExFilePath);
                progressBar.Maximum = totalImages;

                await DownloadImagesAsync(ExFilePath, LoadDirectory);
                lblStatus.Text = "Download completo.";
                
            }
            catch (Exception ex)
            {
                
                MessageBox.Show($"Erro ao realizar download das imagens: {ex.Message}");
            }
        }

        private int CountHyperlinks(string hiperPath)
        {
            int count = 0;
            using (var workbook = new XLWorkbook(hiperPath))
            {
                var worksheet = workbook.Worksheet(1);

                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        if (cell.HasHyperlink)
                        {
                            count++;
                        }
                    }
                }
            }
            return count;
        }

        private async Task DownloadImagesAsync(string filePath, string downloadDirectory)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var httpClient = new HttpClient();
                int downloadedImages = 0;

                foreach (var row in worksheet.RowsUsed().Skip(1))
                {
                    var tucValue = row.Cell("B").GetString(); 
                    string folderPath = Path.Combine(downloadDirectory, tucValue);

                    
                    Directory.CreateDirectory(folderPath);

                    foreach (var cell in row.CellsUsed())
                    {
                        if (cell.HasHyperlink)
                        {
                            string imageUrl = cell.GetHyperlink().ExternalAddress.ToString();
                            string imageName = cell.GetString();
                            string safeImageName = GetSafeFileName(folderPath, imageName);
                            string imagePath = Path.Combine(folderPath, safeImageName);
                            await DownloadImageAsync(httpClient, imageUrl, imagePath);
                            downloadedImages++;
                            progressBar.Value = downloadedImages;
                        }
                    }
                }
            }
        }

        private string GetSafeFileName(string directory, string fileName)
        {
            string safeFileName = fileName;
            int count = 1;

            while (File.Exists(Path.Combine(directory, safeFileName)))
            {
                string nameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
                string extension = Path.GetExtension(fileName);
                safeFileName = $"{nameWithoutExtension}_{count++}{extension}";
            }

            return safeFileName;
        }

        private async Task DownloadImageAsync(HttpClient httpClient, string url, string path)
        {
            try
            {
                var response = await httpClient.GetAsync(url);
                response.EnsureSuccessStatusCode();

                using (var stream = await response.Content.ReadAsStreamAsync())
                using (var fileStream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    await stream.CopyToAsync(fileStream);
                }
                lblStatus.Text = $"Imagem salva em {path}.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao baixar a imagem: {ex.Message}");
            }
        }
    }
}
