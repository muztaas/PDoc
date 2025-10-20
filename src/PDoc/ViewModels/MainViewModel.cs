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
        private bool _areAllFilesSelected = true;
        private ICommand? _selectSaveLocationCommand;

        public ObservableCollection<DocFile> Files { get; } = new();

        public ICommand SelectFilesCommand { get; }
        public ICommand ConvertCommand { get; }

        public ICommand SelectSaveLocationCommand => _selectSaveLocationCommand ??= new RelayCommand(file =>
        {
            var docFile = file as DocFile;
            if (docFile == null) return;

            var dialog = new SaveFileDialog
            {
                Filter = "PDF Files|*.pdf",
                FileName = Path.GetFileNameWithoutExtension(docFile.FilePath) + ".pdf",
                InitialDirectory = Path.GetDirectoryName(docFile.FilePath)
            };

            if (dialog.ShowDialog() == true)
            {
                docFile.CustomPdfPath = dialog.FileName;
            }
        });

        public bool AreAllFilesSelected
        {
            get => _areAllFilesSelected;
            set
            {
                if (_areAllFilesSelected != value)
                {
                    _areAllFilesSelected = value;
                    foreach (var file in Files)
                    {
                        file.IsSelected = value;
                    }
                    OnPropertyChanged();
                }
            }
        }

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
            var selectedFiles = Files.Where(f => f.IsSelected).ToList();
            if (selectedFiles.Count == 0)
            {
                ErrorMessage = "Please select at least one file to convert.";
                return;
            }

            IsConverting = true;
            ErrorMessage = null;

            try
            {
                foreach (var file in selectedFiles)
                {
                    try
                    {
                        if (!File.Exists(file.FilePath))
                        {
                            file.Status = FileStatus.Failed;
                            file.ErrorMessage = "File not found.";
                            continue;
                        }

                        var fileInfo = new FileInfo(file.FilePath);
                        if (fileInfo.Length == 0)
                        {
                            file.Status = FileStatus.Failed;
                            file.ErrorMessage = "File is empty.";
                            continue;
                        }

                        file.Status = FileStatus.Converting;
                        var pdfPath = file.CustomPdfPath ?? Path.ChangeExtension(file.FilePath, ".pdf");

                        var pdfDirectory = Path.GetDirectoryName(pdfPath);
                        if (!Directory.Exists(pdfDirectory))
                        {
                            try
                            {
                                Directory.CreateDirectory(pdfDirectory!);
                            }
                            catch (Exception ex)
                            {
                                file.Status = FileStatus.Failed;
                                file.ErrorMessage = $"Failed to create output directory: {ex.Message}";
                                continue;
                            }
                        }

                        if (File.Exists(pdfPath))
                        {
                            try
                            {
                                File.Delete(pdfPath);
                            }
                            catch (Exception ex)
                            {
                                file.Status = FileStatus.Failed;
                                file.ErrorMessage = $"Failed to overwrite existing PDF: {ex.Message}";
                                continue;
                            }
                        }

                        await _pythonService.ConvertToPdf(file.FilePath, pdfPath);

                        if (!File.Exists(pdfPath))
                        {
                            file.Status = FileStatus.Failed;
                            file.ErrorMessage = "Conversion completed but PDF file was not created.";
                            continue;
                        }

                        file.PdfPath = pdfPath;
                        file.Status = FileStatus.Completed;
                    }
                    catch (Exception ex)
                    {
                        file.Status = FileStatus.Failed;
                        file.ErrorMessage = $"Conversion failed: {ex.Message}";
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
        private string? _customPdfPath;
        private bool _isSelected = true;
        private ICommand? _openPdfCommand;
        private ICommand? _openFolderCommand;

        public string FilePath { get; set; }

        public string FileName
        {
            get
            {
                var fileName = Path.GetFileName(FilePath);
                if (fileName.Length > 50)
                {
                    return $"{fileName.Substring(0, 40)}...{fileName.Substring(fileName.Length - 10)}";
                }
                return fileName;
            }
        }

        public string? CustomPdfPath
        {
            get => _customPdfPath;
            set
            {
                if (_customPdfPath != value)
                {
                    _customPdfPath = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged();
                }
            }
        }

        public string? PdfPath
        {
            get => _pdfPath;
            set
            {
                if (_pdfPath != value)
                {
                    _pdfPath = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(TruncatedPdfPath));
                }
            }
        }

        public string? TruncatedPdfPath
        {
            get
            {
                if (string.IsNullOrEmpty(PdfPath))
                    return null;

                const int maxLength = 60;
                const int tailLength = 10;

                if (PdfPath.Length <= maxLength)
                    return PdfPath;

                string start = PdfPath.Substring(0, maxLength - tailLength - 3);
                string end = PdfPath.Substring(PdfPath.Length - tailLength);
                return $"{start}...{end}";
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
