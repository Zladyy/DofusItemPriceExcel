using System.Windows;

namespace DofusItemPriceExcelRunner
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
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
    }
}
