using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace markevaluator
{
    /// <summary>
    /// Interaction logic for generator_window.xaml
    /// </summary>
    public partial class generator_window : Window
    {
        public generator_window()
        {
            InitializeComponent();
            Windows.setWindowChrome(this);
            Windows.generatorWindow = this;
            
            //Set default directory to desktop
            lblOutputFolder.Content = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            cbxCourse.ItemsSource = Courses.getCourseCodes(StudentType.REGULAR);
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void label_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void btnBrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            fbd.Description = "Select a folder where output files will be saved";
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                lblOutputFolder.Content = fbd.SelectedPath;
        }

        private void cbxCourse_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cbxCourse.SelectedIndex != -1)
            {
                cbxSemester.ItemsSource = Results.getSemesterList((String)cbxCourse.SelectedItem);
                cbxSemester.Text = "-- SEMESTER --";
                cbxYear.ItemsSource = null;
                cbxYear.Text = "-- YEAR --";
                cbxMonth.ItemsSource = null;
                cbxMonth.Text = "-- MONTH --";
            }
        }

        private void cbxSemester_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cbxSemester.SelectedIndex != -1)
            {
                cbxYear.ItemsSource = Results.getYearList((String)cbxCourse.SelectedItem, (int)cbxSemester.SelectedItem);
                cbxYear.Text = "-- YEAR --";
                cbxMonth.ItemsSource = null;
                cbxMonth.Text = "-- MONTH --";
            }
        }

        private void cbxYear_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(cbxYear.SelectedIndex != -1)
            {
                cbxMonth.ItemsSource = Results.getMonthList((String)cbxCourse.SelectedItem, (int)cbxSemester.SelectedItem, (int)cbxYear.SelectedItem);
            }
        }

        private void btnGenerateSheet_Click(object sender, RoutedEventArgs e)
        {
            String errors = null;

            if (cbxCourse.SelectedIndex == -1)
                errors += "Please select course" + Environment.NewLine;
            if (cbxSemester.SelectedIndex == -1)
                errors += "Please select semester" + Environment.NewLine;
            if (cbxYear.SelectedIndex == -1)
                errors += "Please select year" + Environment.NewLine;
            if (cbxDegree.SelectedIndex == -1)
                errors += "Please select degree type" + Environment.NewLine;
            if (cbxMonth.SelectedIndex == -1)
                errors += "Please select month" + Environment.NewLine;
            if (rdoCustomStudents.IsChecked == true)
                if (System.IO.File.Exists((String)lblSheetPath.Content) == false)
                    errors += "Please select student list.xlsx file";

            if (errors != null)
            {
                MessageBox.Show("Please correct the following error(s)!" + Environment.NewLine + Environment.NewLine + errors, "Invalid inputs", MessageBoxButton.OK, MessageBoxImage.Stop);
                return;
            }

            toggleControlState();
            this.txtOutput.Cursor = Cursors.Wait;
            this.Cursor = Cursors.Wait;
            progressOne.Value = 0;
            progressTwo.Value = 0;
            txtOutput.Clear();

            //get values from combobox
            String output_folder = (String)lblOutputFolder.Content;
            String course = (String)cbxCourse.SelectedItem;
            String degree = (cbxDegree.SelectedItem as ComboBoxItem).Content.ToString();
            String month = (String)cbxMonth.SelectedItem;
            String fileName = (String)lblSheetPath.Content;
            int semester = (int)cbxSemester.SelectedItem;
            int year = (int)cbxYear.SelectedItem;

            if (rdoGradeSheet.IsChecked == true)
                if (rdoCustomStudents.IsChecked == false)
                    (new Results()).generateResultSheet(course, semester, year, month, output_folder);
                else
                    (new Results()).getRegIdFromSheet(course, semester, month, year, fileName);
            else
                if (rdoCustomStudents.IsChecked == false)
                    (new Results()).generateGradeSheets(course, semester, year, month, output_folder, degree);
                else
                    (new Results()).getRegIdFromSheet(course, semester, month, year, fileName);
        }

        public void exportCompleted(String message,bool has_error)
        {
            toggleControlState(true);
            this.txtOutput.Cursor = Cursors.Arrow;
            this.Cursor = Cursors.Arrow;

            if (has_error)
                MessageBox.Show(message, "Task completed", MessageBoxButton.OK, MessageBoxImage.Warning);
            else
                MessageBox.Show(message, "Task completed", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        /// <summary>
        /// For temporarily enabling or disabling child controls of this window
        /// while performing a time consuming operation
        /// </summary>
        /// <param name="toggle">True to enable else disable. default is false</param>
        void toggleControlState(bool toggle = false)
        {
            this.rdoGradeSheet.IsEnabled = toggle;
            this.rdoMarkSheet.IsEnabled = toggle;
            this.rdoAllStudents.IsEnabled = toggle;
            this.rdoCustomStudents.IsEnabled = toggle;
            this.btnBrowseSheet.IsEnabled = toggle;
            this.btnBrowseFolder.IsEnabled = toggle;
            this.cbxCourse.IsEnabled = toggle;
            this.cbxSemester.IsEnabled = toggle;
            this.cbxYear.IsEnabled = toggle;
            this.cbxMonth.IsEnabled = toggle;
            this.cbxDegree.IsEnabled = toggle;
            this.btnGenerateSheet.IsEnabled = toggle;
            this.btnClose.IsEnabled = toggle;
        }

        private void rdoCustomStudents_Click(object sender, RoutedEventArgs e)
        {
            bool toggle = false;

            if (rdoCustomStudents.IsChecked == true)
                toggle = true;

            this.btnBrowseSheet.IsEnabled = toggle;
            this.lblSheetPath.IsEnabled = toggle;
        }

        private void rdoAllStudents_Click(object sender, RoutedEventArgs e)
        {
            bool toggle = true;

            if (rdoAllStudents.IsChecked == true)
                toggle = false;

            this.btnBrowseSheet.IsEnabled = toggle;
            this.lblSheetPath.IsEnabled = toggle;
        }

        private void btnBrowseSheet_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog();
            ofd.FileName = "";
            ofd.Title = "Please select custom student list .xlsx file";
            ofd.Filter = "Excel Worksheet files|*.xlsx";

            if (ofd.ShowDialog() == true)
            {
                lblSheetPath.Content = ofd.FileName;
            }
            else
            {
                if (System.IO.File.Exists((String)lblSheetPath.Content) == false)
                    lblSheetPath.Content = "Internal Marksheet.xlsx";
            }
        }

        public void regIdCollected(string message,bool isvalid,bool error_free,List<long> reg_id_list)
        {            
            //get values from combobox
            String output_folder = (String)lblOutputFolder.Content;
            String course = (String)cbxCourse.SelectedItem;
            String degree = (cbxDegree.SelectedItem as ComboBoxItem).Content.ToString();
            String month = (String)cbxMonth.SelectedItem;
            int semester = (int)cbxSemester.SelectedItem;
            int year = (int)cbxYear.SelectedItem;

            if (!isvalid)
            {
                MessageBox.Show(message, "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
                toggleControlState(true);
                this.txtOutput.Cursor = Cursors.Arrow;
                this.Cursor = Cursors.Arrow;
            }
            else if (!error_free)
            {
                MessageBox.Show(message, "WARNING", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                if (rdoGradeSheet.IsChecked == true)
                    (new Results()).generateResultSheet(course, semester, year, month, output_folder, reg_id_list);
                else
                    (new Results()).generateGradeSheets(course, semester, year, month, output_folder, degree, reg_id_list);
            }
            else
                if (rdoGradeSheet.IsChecked == true)
                    (new Results()).generateResultSheet(course, semester, year, month, output_folder, reg_id_list);
                else
                    (new Results()).generateGradeSheets(course, semester, year, month, output_folder, degree, reg_id_list);
        }
    }
}
