using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;

namespace PDoc.Services
{
    public class PythonHostingService
    {
        private readonly string _pythonPath;
        private readonly string _scriptPath;
        private readonly string _pythonDir;

        public PythonHostingService()
        {
            var baseDir = AppContext.BaseDirectory;
            _pythonDir = Path.Combine(baseDir, "Python");
            _pythonPath = Path.Combine(_pythonDir, "python.exe");
            _scriptPath = Path.Combine(_pythonDir, "convert.py");
        }

        public async Task ConvertToPdf(string inputFile, string outputFile)
        {
            if (!File.Exists(_pythonPath))
                throw new FileNotFoundException("Python runtime not found", _pythonPath);

            if (!File.Exists(_scriptPath))
                throw new FileNotFoundException("Conversion script not found", _scriptPath);

            var startInfo = new ProcessStartInfo
            {
                FileName = _pythonPath,
                Arguments = $"\"{_scriptPath}\" \"{inputFile}\" \"{outputFile}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = _pythonDir
            };

            // Set up Python environment variables
            startInfo.Environment["PYTHONHOME"] = _pythonDir;
            startInfo.Environment["PYTHONPATH"] = Path.Combine(_pythonDir, "Lib", "site-packages");
            
            try 
            {
                using var process = Process.Start(startInfo);
                if (process == null)
                    throw new Exception("Failed to start Python process");

                var output = await process.StandardOutput.ReadToEndAsync();
                var error = await process.StandardError.ReadToEndAsync();
                await process.WaitForExitAsync();

                if (process.ExitCode != 0)
                    throw new Exception($"Python conversion failed: {error}\nOutput: {output}");
            }
            catch (Exception ex)
            {
                throw new Exception($"Error during conversion: {ex.Message}", ex);
            }
        }
    }
}