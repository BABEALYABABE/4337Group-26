using System.Windows;

namespace Group4337
{
    public partial class MainWindow : Window
    {
        public MainWindow()
            => InitializeComponent();

        private void Akhmedova_4337_Click(object sender, RoutedEventArgs e)
        {
            var Akhmedova_4337 = new _4337_Akhmedova();
            Akhmedova_4337.ShowDialog();
        }
    }
}