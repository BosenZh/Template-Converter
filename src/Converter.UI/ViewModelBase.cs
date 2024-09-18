using System.ComponentModel;
using System.Net.Http;
using System.Runtime.CompilerServices;

namespace Converter
{
    public class ViewModelBase : INotifyPropertyChanged
    {
        protected static readonly HttpClient Client = new ();

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}