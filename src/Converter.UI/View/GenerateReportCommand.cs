using System.Windows.Input;

namespace Converter.View
{
    public class GenerateReportCommand : ICommand
    {
        private readonly Action execute;
        private readonly Func<bool>? canExecute;

        public GenerateReportCommand(Action execute, Func<bool>? canExecute = null)
        {
            this.execute = execute ?? throw new ArgumentNullException(nameof(execute));
            this.canExecute = canExecute;
        }

        public event EventHandler? CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }

        public bool CanExecute(object? parameter) => this.canExecute == null || this.canExecute();

        public void Execute(object? parameter) => this.execute();
    }
}
