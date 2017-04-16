using System;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;

namespace markevaluator
{
    /// <summary>
    /// Interaction logic for parser_window.xaml
    /// </summary>
    public partial class parser_window : Window
    {
        private IParse cat;
        private string fileName;
        public parser_window(Object cat,String title)
        {
            InitializeComponent();
            Windows.setWindowChrome(this);
            Windows.parserWindow = this;

            this.cat = (IParse)cat;
            this.label.Content += " (" + title + ")";
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnBrowseFile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.FileName = "";
                ofd.Title = "Please select a xlsx file";
                ofd.Filter = "Excel Worksheet files|*.xlsx";
                ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                if (ofd.ShowDialog() == true)
                {
                    if (System.IO.File.Exists(ofd.FileName))
                    {
                        lblTitle.Content = "File is ready to be validated...";
                        lblSelectedFile.Content = ofd.FileName;
                        lblFileSize.Content = "File Size: " + (((new System.IO.FileInfo(ofd.FileName)).Length) / 1024) + " Kb";
                        if (MessageBox.Show("Procced with the selected file?", "Your file is going to be validated", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                        {
                            fileName = ofd.FileName;
                            StartValidation();
                        }
                        else
                        {
                            lblSelectedFile.Content = "";
                            lblFileSize.Content = "File Size: 0";
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid File!");
                        return;
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Something went wrong :(" + Environment.NewLine + "Please retry.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                LogWriter.WriteError("While selecting a file (parser window)", ex.Message);
            }
        }

        private void label_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        void StartValidation()
        {
            btnBrowseFile.IsEnabled = false;
            btnCancel.IsEnabled = false;
            txtOut.Cursor = Cursors.Wait;
            this.Cursor = Cursors.Wait;
            prgsbar.Value = 0;
            txtOut.Clear();

            lblTitle.Content = "Please wait, your file is being validated";
            cat.ValidateWorksheet(fileName);
        }

        /// <summary>
        /// This method will be invoked by the class for which
        /// validation is performed
        /// </summary>
        /// <param name="message">Some message about validation</param>
        /// <param name="isValid">Is the worksheet is valid and error free</param>
        public void ValidationCompleted(string message,bool isValid)
        {
            btnCancel.IsEnabled = true;
            btnBrowseFile.IsEnabled = true;
            txtOut.Cursor = Cursors.Arrow;
            this.Cursor = Cursors.Arrow;
            lblTitle.Content = "Validation completed";

            if (isValid)
            {
                if(MessageBox.Show(message, "Validation Result", MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                {
                    btnBrowseFile.IsEnabled = false;
                    btnCancel.IsEnabled = false;
                    txtOut.Cursor = Cursors.Wait;
                    this.Cursor = Cursors.Wait;
                    lblTitle.Content = "Pushing information to database.";

                    //perferom database update
                    prgsbar.Value = 0;
                    cat.PushToDatabase(fileName);
                }
            }
            else
            {
                MessageBox.Show(message, "Validation Result", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
        }

        /// <summary>
        /// After pushing all records to database from excel sheet
        /// </summary>
        /// <param name="message">Message containing description of database push result</param>
        /// <param name="isAllOk">Whether all records where successfully pushed or not</param>
        public void PushToDatabaseCompleted(string message,bool isAllOk)
        {
            btnCancel.IsEnabled = true;
            btnBrowseFile.IsEnabled = true;
            txtOut.Cursor = Cursors.Arrow;
            this.Cursor = Cursors.Arrow;
            lblTitle.Content = "Records pushed to database";
            MessageBox.Show(message, "Pushed to database", MessageBoxButton.OK, isAllOk ? MessageBoxImage.Information : MessageBoxImage.Exclamation);
            Windows.adminWindow.populateData();
        }
    }
}
