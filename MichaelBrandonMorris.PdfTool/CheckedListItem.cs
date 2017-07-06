using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace MichaelBrandonMorris.PdfTool
{
    /// <summary>
    ///     Class CheckedListItem.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <seealso cref="INotifyPropertyChanged" />
    /// TODO Edit XML Comment Template for CheckedListItem`1
    public class CheckedListItem<T> : INotifyPropertyChanged
    {
        /// <summary>
        ///     The is checked
        /// </summary>
        /// TODO Edit XML Comment Template for _isChecked
        private bool _isChecked;

        /// <summary>
        ///     The item
        /// </summary>
        /// TODO Edit XML Comment Template for _item
        private T _item;

        /// <summary>
        ///     Initializes a new instance of the
        ///     <see cref="CheckedListItem{T}" /> class.
        /// </summary>
        /// <param name="item">The item.</param>
        /// <param name="isChecked">if set to <c>true</c> [is checked].</param>
        /// TODO Edit XML Comment Template for #ctor
        public CheckedListItem(T item, bool isChecked = false)
        {
            _item = item;
            _isChecked = isChecked;
        }

        /// <summary>
        ///     Gets or sets a value indicating whether this instance
        ///     is checked.
        /// </summary>
        /// <value>
        ///     <c>true</c> if this instance is checked; otherwise,
        ///     <c>false</c>.
        /// </value>
        /// TODO Edit XML Comment Template for IsChecked
        public bool IsChecked
        {
            get => _isChecked;
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

        /// <summary>
        ///     Gets or sets the item.
        /// </summary>
        /// <value>The item.</value>
        /// TODO Edit XML Comment Template for Item
        public T Item
        {
            get => _item;
            set
            {
                _item = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        ///     Occurs when a property value changes.
        /// </summary>
        /// TODO Edit XML Comment Template for PropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        ///     Notifies the property changed.
        /// </summary>
        /// <param name="propertyName">Name of the property.</param>
        /// TODO Edit XML Comment Template for NotifyPropertyChanged
        private void NotifyPropertyChanged(
            [CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(
                this,
                new PropertyChangedEventArgs(propertyName));
        }
    }
}