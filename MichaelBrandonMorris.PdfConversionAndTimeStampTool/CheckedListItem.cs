using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace MichaelBrandonMorris.PdfConversionAndTimeStampTool
{
    public class CheckedListItem<T> : INotifyPropertyChanged
    {
        private bool _isChecked;
        private T _item;

        public CheckedListItem(T item, bool isChecked = false)
        {
            _item = item;
            _isChecked = isChecked;
        }

        public bool IsChecked
        {
            get
            {
                return _isChecked;
            }
            set
            {
                if (_isChecked == value)
                {
                    return;
                }

                _isChecked = value;
                NotifyPropertyChanged();
            }
        }

        public T Item
        {
            get
            {
                return _item;
            }
            set
            {
                _item = value;
                NotifyPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged(
            [CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(
                this,
                new PropertyChangedEventArgs(propertyName));
        }
    }
}