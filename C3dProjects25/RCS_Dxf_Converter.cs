#nullable enable
using System;
using System.Collections.ObjectModel;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using C3dProjects25.ViewModels;

public interface IDwgToDxfExporter
{
    Task<BatchExportResult> ExportAsync(BatchExportRequest request, IProgress<string> progress, CancellationToken cancellationToken);
}

namespace C3dProjects25.ViewModels
{
    public abstract class ViewModelBase
    {
        // Add common ViewModel functionality here if needed.
    }

    public sealed class BatchExportViewModel : ViewModelBase
    {
        private readonly IDwgToDxfExporter _exporter;
        private CancellationTokenSource? _cts;

        public BatchExportViewModel(IDwgToDxfExporter exporter)
        {
            _exporter = exporter;
            RunCommand = new AsyncRelayCommand(RunAsync, CanRun);
        }

        public string InputFolder { get; set; } = "";
        public string OutputFolder { get; set; } = "";
        public bool RecurseSubfolders { get; set; } = true;

        public ObservableCollection<string> Messages { get; } = new();

        public ICommand RunCommand { get; }

        public async Task RunAsync()
        {
            _cts = new CancellationTokenSource();
            Messages.Clear();

            var request = new BatchExportRequest
            {
                InputFolder = InputFolder,
                OutputFolder = OutputFolder,
                RecurseSubfolders = RecurseSubfolders
            };

            var progress = new Progress<string>(msg => Messages.Add(msg));

            try
            {
                BatchExportResult result = await _exporter.ExportAsync(request, progress, _cts.Token);
                Messages.Add($"Log: {result.LogPath}");
            }
            catch (Exception ex)
            {
                Messages.Add($"Failed: {ex.Message}");
            }
        }

        private bool CanRun()
        {
            return !string.IsNullOrWhiteSpace(InputFolder)
                   && !string.IsNullOrWhiteSpace(OutputFolder);
        }
    }

    public sealed class BatchExportRequest
    {
        public string InputFolder { get; set; } = string.Empty;
        public string OutputFolder { get; set; } = string.Empty;
        public bool RecurseSubfolders { get; set; } = true;
    }

    public sealed class BatchExportResult
    {
        public string LogPath { get; set; } = string.Empty;
    }

    // Add this class to provide an AsyncRelayCommand implementation compatible with ICommand.
    // This is a minimal implementation for async command support in WPF/MVVM.

    public sealed class AsyncRelayCommand : ICommand
    {
        private readonly Func<Task> _execute;
        private readonly Func<bool>? _canExecute;
        private bool _isExecuting;

        public AsyncRelayCommand(Func<Task> execute, Func<bool>? canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public bool CanExecute(object? parameter)
        {
            return !_isExecuting && (_canExecute?.Invoke() ?? true);
        }

        public async void Execute(object? parameter)
        {
            if (!CanExecute(parameter))
                return;

            _isExecuting = true;
            RaiseCanExecuteChanged();

            try
            {
                await _execute();
            }
            finally
            {
                _isExecuting = false;
                RaiseCanExecuteChanged();
            }
        }

        public event EventHandler? CanExecuteChanged;

        public void RaiseCanExecuteChanged()
        {
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}