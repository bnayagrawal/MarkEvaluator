using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shell;
using System.Collections.Generic;

namespace markevaluator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Windows.setWindowChrome(this);
            Windows.loginWindow = this;

            MouseDown += login_window_MouseDown;
            try
            {
                Medatabase.OpenConnection();
            }
            catch(databaseException de)
            {
                MessageBox.Show(de.Message,"Connection Failed",MessageBoxButton.OK,MessageBoxImage.Exclamation);
                Environment.Exit(0);
            }

            //SET APP THEME COLOR BASED ON WINDOWS ACCENT COLOR
            Brush LightBrush = (Brush)(new BrushConverter()).ConvertFromString(WindowsTheme.getHexColorFromRegistry());
            Brush DarkBrush = (Brush)(new BrushConverter()).ConvertFromString(WindowsTheme.getHexDarkColorFromRegistry());
            Application.Current.Resources["ApplicationThemeLight"] = LightBrush;
            Application.Current.Resources["ApplicationThemeDark"] = DarkBrush;

            //SET SOIS LABLE BACKGROUND
            System.Drawing.Color c = WindowsTheme.getColorFromRegistry();
            LinearGradientBrush gradient = new LinearGradientBrush();
            gradient.StartPoint = new Point(0.5, 0);
            gradient.EndPoint = new Point(0.5, 1);
            gradient.RelativeTransform = lblBanner.Background.RelativeTransform;
            gradient.GradientStops.Add(new GradientStop(Colors.White, 0));
            gradient.GradientStops.Add(new GradientStop(Color.FromRgb(c.R,c.G,c.B), 1));
            lblBanner.Background = gradient;
        }

        private void login_window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                DragMove();
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            Environment.Exit(0);
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (txtUsername.GetLineLength(0) == 0 || txtPassword.Password.Length == 0)
                {
                    MessageBox.Show("Please input all the fields!", "WAIT", MessageBoxButton.OK, MessageBoxImage.Stop);
                    return;
                }

                string tmp = null;
                List<Row> rows = Medatabase.fetchRecords("SELECT password FROM admin_login WHERE username='" + txtUsername.Text + "'");
                foreach(Row row in rows)
                    tmp = (string)row.column["password"];

                if (string.Compare(tmp, txtPassword.Password) != 0 || tmp == null)
                {
                    MessageBox.Show("Invalid username or password!", "Error", MessageBoxButton.OK, MessageBoxImage.Stop);
                    return;
                }
                admin_window obj = new admin_window();
                obj.Show();
                this.Hide();
            }
            catch(databaseException de)
            {
                MessageBox.Show(de.Message, "Error Occured", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch(Exception ex)
            {
                MessageBox.Show("Something went wrong :(", "Error Occured", MessageBoxButton.OK, MessageBoxImage.Error);
                LogWriter.WriteError("Login button clicked", ex.Message);
            }
        }

        private void txtUsername_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.OemQuotes) 
                e.Handled = true; 
        }
    }
}
