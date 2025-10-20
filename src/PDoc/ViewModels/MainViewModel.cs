using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows.Input;
using Microsoft.Win32;
using PDoc.Services;
using System.IO;

namespace PDoc.ViewModels
{
    public enum FileStatus
    {
        Pending,
        Converting,
        Completed,
        Failed
    }

    public class MainViewModel : INotifyPropertyChanged
    {
        private readonly PythonHostingService _pythonService;
        private bool _isConverting;
        private string? _errorMessage;

        public ObservableCollection<DocFile> Files { get; } = new();

        public ICommand SelectFilesCommand { get; }
        public ICommand ConvertCommand { get; }

        public bool IsConverting
        {
            get => _isConverting;
            set
            {
                if (_isConverting != value)
                {
                    _isConverting = value;
                    OnPropertyChanged();
                }
            }
        }

        public MainViewModel(PythonHostingService pythonService)
        {
            _pythonService = pythonService;
            SelectFilesCommand = new RelayCommand(_ => SelectFiles());
            ConvertCommand = new RelayCommand(async _ => await Convert(), _ => CanConvert());
        }

        private void SelectFiles()
        {
            var dialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Word Documents|*.doc;*.docx"
            };

            if (dialog.ShowDialog() == true)
            {
                foreach (var file in dialog.FileNames)
                {
                    Files.Add(new DocFile { FilePath = file });
                }
            }
        }

        private bool CanConvert()
        {
            return Files.Count > 0 && !IsConverting;
        }

        private async Task Convert()
        {
            if (Files.Count == 0) return;

            IsConverting = true;
            ErrorMessage = null;

            try
            {
                foreach (var file in Files.ToList())
                {
                    try
                    {
                        var pdfPath = Path.ChangeExtension(file.FilePath, ".pdf");
                        await _pythonService.ConvertToPdf(file.FilePath, pdfPath);
                        file.PdfPath = pdfPath;
                        file.Status = FileStatus.Completed;
                    }
                    catch (Exception ex)
                    {
                        file.Status = FileStatus.Failed;
                        file.ErrorMessage = ex.Message;
                        ErrorMessage = $"Error converting {file.FileName}: {ex.Message}";
                    }
                }
            }
            finally
            {
                IsConverting = false;
            }
        }

        public string? ErrorMessage
        {
            get => _errorMessage;
            set
            {
                if (_errorMessage != value)
                {
                    _errorMessage = value;
                    OnPropertyChanged();
                }
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class DocFile : INotifyPropertyChanged
    {
        private FileStatus _status = FileStatus.Pending;
        private string? _errorMessage;
        private string? _pdfPath;
        private ICommand? _openPdfCommand;
        private ICommand? _openFolderCommand;

        public required string FilePath { get; set; }
        public string FileName => Path.GetFileName(FilePath);

        public string? PdfPath
        {
            get => _pdfPath;
            set
            {
                if (_pdfPath != value)
                {
                    _pdfPath = value;
                    OnPropertyChanged();
                }
            }
        }
        
        public FileStatus Status
        {
            get => _status;
            set
            {
                if (_status != value)
                {
                    _status = value;
                    OnPropertyChanged();
                }
            }
        }

        public string? ErrorMessage
        {
            get => _errorMessage;
            set
            {
                if (_errorMessage != value)
                {
                    _errorMessage = value;
                    OnPropertyChanged();
                }
            }
        }

        public ICommand OpenPdfCommand => _openPdfCommand ??= new RelayCommand(_ =>
        {
            if (File.Exists(PdfPath))
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = PdfPath,
                    UseShellExecute = true
                });
        });

        public ICommand OpenFolderCommand => _openFolderCommand ??= new RelayCommand(_ =>
        {
            if (File.Exists(PdfPath))
                System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{PdfPath}\"");
        });

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
