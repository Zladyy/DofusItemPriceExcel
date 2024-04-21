using System.Text.RegularExpressions;
using System.Windows;

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
        }

        private static bool IsTextAllowed(string text)
        {
            return !_numbersOnlyRegex.IsMatch(text);
        }
    }
}
