using System.Windows;
using PDoc.Services;
using PDoc.ViewModels;

namespace PDoc
{
    public partial class App : Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            var pythonService = new PythonHostingService();
            var viewModel = new MainViewModel(pythonService);
            var mainWindow = new MainWindow { DataContext = viewModel };
            mainWindow.Show();
        }
    }
}