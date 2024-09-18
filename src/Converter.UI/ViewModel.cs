using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Timers;
using System.Windows;
using System.Windows.Input;
using Converter.Types;
using Converter.View;
using Microsoft.Win32;
using Timer = System.Timers.Timer;

namespace Converter
{
    public class ViewModel : ViewModelBase
    {
        private readonly Timer progressTimer;

        private bool openWhenGenerated = true;
        private string? templateFilePath = string.Empty;
        private string? filePath;
        private string? outputMessage;
        private int? progress;
        private ICommand? selectExcelCommand;
        private ICommand? selectTemplateCommand;

        public ViewModel()
        {
            this.progressTimer = new Timer(1000); // 1 second interval
            this.progressTimer.Elapsed += this.OnProgressTimerElapsed;
        }

        public bool OpenWhenGenerated
        {
            get
            {
                return this.openWhenGenerated;
            }

            set
            {
                this.openWhenGenerated = value;
                this.OnPropertyChanged(nameof(this.OpenWhenGenerated));
            }
        }

        public string? FilePath
        {
            get => this.filePath;
            set
            {
                this.filePath = value;
                this.OnPropertyChanged(nameof(this.FilePath));
            }
        }

        public string? TemplateFilePath
        {
            get => this.templateFilePath;
            set
            {
                this.templateFilePath = value;
                this.OnPropertyChanged(nameof(this.templateFilePath));
            }
        }

        public int? Progress
        {
            get => this.progress;
            set
            {
                this.progress = value;
                this.OnPropertyChanged(nameof(this.Progress));
            }
        }

        public string? OutputMessage
        {
            get => this.outputMessage;
            set
            {
                this.outputMessage = value;
                this.OnPropertyChanged(nameof(this.OutputMessage));
            }
        }

        public ICommand SelectExcelCommand
        {
            get
            {
                this.selectExcelCommand ??=
                    new GenerateReportCommand(async () => await this.SelectExcelAndGenerateReportAsync());
                return this.selectExcelCommand;
            }
        }

        public ICommand SelectTemplateCommand
        {
            get
            {
                this.selectTemplateCommand ??=
                    new GenerateReportCommand(async () => await this.SelectTemplateAsync());
                return this.selectTemplateCommand;
            }
        }

        private static async Task<HttpResponseMessage> CallAzureFunctionAsync(string? url, string jsonContent)
        {
            StringContent content = new (jsonContent, Encoding.UTF8, "application/json");
            HttpResponseMessage response = await Client.PostAsync(url, content);
            return response;
        }

        private static async Task WriteAllBytesAsync(string path, byte[] bytes)
        {
            using FileStream? fs = new (path, FileMode.Create, FileAccess.Write, FileShare.None, 4096, true);
            await fs.WriteAsync(bytes);
        }

        private static string GetUrl() =>
#if !DEBUG
            "https://endepreciation20240816152948.azurewebsites.net/api/WordInserter?code=" +
            "HmsQBKdfNgyIauysjVGX-ahtlKDz1uAozQK9YRUU1HIZAzFu-UxD_Q%3D%3D";
#else
            "http://localhost:7044/api/WordInserter";
#endif

        private async Task SelectExcelAndGenerateReportAsync()
        {
            OpenFileDialog openFileDialog = new ()
            {
                Title = "Choose a Spreadsheet file:",
                Filter = "*.xlsm",
                Multiselect = false,
                CheckFileExists = true,
            };
            if (openFileDialog.ShowDialog() != true)
            {
                return;
            }

            this.FilePath = openFileDialog.FileName;

            this.Progress = 0;
            this.progressTimer.Start();

            try
            {
                if (this.TemplateFilePath == string.Empty || this.TemplateFilePath is null)
                {
                    MessageBox.Show(
                    "Missing Template",
                    "Failed",
                    MessageBoxButton.OK);
                    throw new Exception("missing output file");
                }

                Request request = Reader.Program.ReadExcel(
                    this.FilePath,
                    this.TemplateFilePath);

                byte[] fileContent = Writer.Program.InsertDataToWord(request, this.TemplateFilePath);
                string outputFilePath = Path.ChangeExtension(this.FilePath, ".docx");
                await WriteAllBytesAsync(outputFilePath, fileContent);
                this.progressTimer.Stop();
                this.Progress = 100;

                MessageBox.Show(
                    "Report generated successfully!",
                    "Success",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                if (this.OpenWhenGenerated)
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = outputFilePath,
                        UseShellExecute = true,
                    });
                }
            }
            catch (Exception ex)
            {
                this.progressTimer.Stop();
                MessageBox.Show(
                    $"An error occurred while processing the file: {ex.Message}",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private Task SelectTemplateAsync()
        {
            OpenFileDialog openFileDialog = new()
            {
                Title = "Choose your template:",
                Filter = "*.docx",
                Multiselect = false,
                CheckFileExists = true,
            };
            if (openFileDialog.ShowDialog() != true)
            {
                return Task.CompletedTask;
            }

            TemplateFilePath = openFileDialog.FileName;
            return Task.CompletedTask;
        }

        private void OnProgressTimerElapsed(object? sender, ElapsedEventArgs e)
        {
            if (this.Progress < 95)
            {
                this.Progress += 5;
            }
            else
            {
                this.progressTimer.Stop();
            }
        }
    }
}
