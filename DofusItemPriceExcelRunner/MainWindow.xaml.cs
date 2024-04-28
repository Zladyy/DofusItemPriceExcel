using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;

namespace DofusItemPriceExcelRunner
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static readonly Regex _numbersOnlyRegex = new Regex("[^0-9.-]+"); //regex that allows only numbers
        private MainWindowController ViewModel => DataContext as MainWindowController;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = new MainWindowController(delegate (bool value)
            {
                RunBtn.IsEnabled = value;
            });
            ViewModel.OnWorkDone += delegate
            {
                Close();
            };
        }

        private void SelectButton_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.OnSelectBtnClicked();
        }

        private void RunButton_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.OnRunButtonClicked();
        }

        private void TextBox_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = !IsTextAllowed(e.Text);
            if(!e.Handled
                && sender is TextBox tb
                && int.TryParse(string.Concat(!string.IsNullOrEmpty(tb.SelectedText) ? tb.Text.Replace(tb.SelectedText, "") : tb.Text, e.Text), out var parsed)
                && (parsed < 0 || parsed > 100))
            {
                e.Handled = true;
            }
        }

        private static bool IsTextAllowed(string text)
        {
            return !_numbersOnlyRegex.IsMatch(text);
        }

        private void TextBox_Pasting(object sender, DataObjectPastingEventArgs e)
        {
            if(e.DataObject.GetDataPresent(typeof(string)))
            {
                string text = (string)e.DataObject.GetData(typeof(string));
                if(!IsTextAllowed(text))
                {
                    e.CancelCommand();
                }
            }
            else
            {
                e.CancelCommand();
            }
        }
    }
}
