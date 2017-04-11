using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Shell;
using Microsoft.Win32;

namespace markevaluator
{
    /// <summary>
    /// Interaction logic for admin_window.xaml
    /// </summary>
    public partial class admin_window : Window
    {
        OpenFileDialog ofd;
        public admin_window()
        {
            InitializeComponent();
            Windows.setWindowChrome(this);
            Windows.adminWindow = this;
            populateData();
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to exit?", "Mark Evaluator", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                Environment.Exit(0);
        }

        private void btnMinimize_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private void btnUploadCourse_Click(object sender, RoutedEventArgs e)
        {
            parser_window pw = new parser_window(new Courses(),"Upload course");
            pw.ShowDialog();     
        }

        private void lblWindowTitle_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                DragMove();
        }

        private void lstvCourse_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                lstvSubjectList.ItemsSource = Subjects.getSubjectList(((CourseCols)lstvCourse.SelectedItems[0]).c_code);

                CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(lstvSubjectList.ItemsSource);
                PropertyGroupDescription groupDescription = new PropertyGroupDescription("semNo");
                view.GroupDescriptions.Add(groupDescription);
            }
            catch (Exception ex) { LogWriter.WriteError("During course listview selection change", ex.Message);  }
        }

        private void btnUploadSubjects_Click(object sender, RoutedEventArgs e)
        {
            parser_window pw = new parser_window(new Subjects(), "Upload subjects");
            pw.ShowDialog();
        }

        private void lblWindowTitle_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (this.WindowState == WindowState.Maximized)
                this.WindowState = WindowState.Normal;
            else
                this.WindowState = WindowState.Maximized;
        }

        private void btnUploadIEMark_Click(object sender, RoutedEventArgs e)
        {
            ofd = new OpenFileDialog();
            ofd.FileName = "";
            ofd.Title = "Please select internal marks xlsx file";
            ofd.Filter = "Excel Worksheet files|*.xlsx";
            //ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (ofd.ShowDialog() == true)
            {
                txtIMFile.Text = ofd.FileName;
            }
            else
            {
                if (System.IO.File.Exists(txtIMFile.Text) == false)
                    txtIMFile.Text = "Internal Marksheet.xlsx";
            }
        }

        private void btnUploadEEMark_Click(object sender, RoutedEventArgs e)
        {
            ofd = new OpenFileDialog();
            ofd.FileName = "";
            ofd.Title = "Please select external marks xlsx file";
            ofd.Filter = "Excel Worksheet files|*.xlsx";
            //ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (ofd.ShowDialog() == true)
            {
                txtEMFile.Text = ofd.FileName;
            }
            else
            {
                if(System.IO.File.Exists(txtEMFile.Text) == false)
                    txtEMFile.Text = "External Marksheet.xlsx";
            }
        }

        private void btnUploadCC_Click(object sender, RoutedEventArgs e)
        {
            ofd = new OpenFileDialog();
            ofd.FileName = "";
            ofd.Title = "Please select cut-off xlsx file";
            ofd.Filter = "Excel Worksheet files|*.xlsx";
            //ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (ofd.ShowDialog() == true)
            {
                txtCutoffFile.Text = ofd.FileName;
            }
            else
            {
                if (System.IO.File.Exists(txtCutoffFile.Text) == false)
                    txtCutoffFile.Text = "Cutt off.xlsx";
            }
        }

        private void btnEvaluate_Click(object sender, RoutedEventArgs e)
        {
            bool isAllOk = true;
            
            //validation
            String error_message = null;
            if (System.IO.File.Exists(txtIMFile.Text) == false)
                error_message = "Please select internal marksheet file" + Environment.NewLine;
            if (System.IO.File.Exists(txtEMFile.Text) == false)
                error_message += "Please select external marksheet file" + Environment.NewLine;
            if (System.IO.File.Exists(txtCutoffFile.Text) == false)
                error_message += "Please select cut off file" + Environment.NewLine;
            if (cbxECourse.SelectedIndex == -1)
                error_message += "Please select course" + Environment.NewLine;
            if (cbxESem.SelectedIndex == -1)
                error_message += "Please select semester" + Environment.NewLine;
            if (cbxEMonth.SelectedIndex == -1)
                error_message += "Please select month" + Environment.NewLine;
            if (cbxEType.SelectedIndex == -1)
                error_message += "Please select student type" + Environment.NewLine;
            if (cbxEYear.SelectedIndex == -1)
                error_message += "Please select year" + Environment.NewLine;

            if (error_message != null)
            {
                isAllOk = false;
                MessageBox.Show("Please correct the following error(s)!" + Environment.NewLine + Environment.NewLine + error_message, "Invalid inputs", MessageBoxButton.OK, MessageBoxImage.Stop);
            }

            if (isAllOk) //Call for validation and evaluation
            {
                //get values from controls
                int semester = (int)cbxESem.SelectedItem;
                int year = Convert.ToInt16(cbxEYear.SelectedItem);
                string course = (String)cbxECourse.SelectedItem;

                string month = (cbxEMonth.SelectedItem as ComboBoxItem).Content.ToString(); //As the item is comboboxitem
                    
                try
                {
                    List<Row> rows = Medatabase.fetchRecords("SELECT * FROM exam_master WHERE course_code='" + course + "' AND semester=" + semester + " AND year=" + year + " AND month='" + month + "'");
                    
                    if (cbxEType.SelectedIndex == 0) //check if evaluation has already finished for given course, semester and date
                    {
                        if (rows.Count > 0)
                        {
                            MessageBox.Show("Evaluation has been finished already for semester " + semester + " of " + course + " on " + month + " " + year, "Stop", MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }
                    }
                    else //check if evaluation has been done or not for given course, semester and date
                    {
                        if (rows.Count == 0)
                        {
                            MessageBox.Show("No record of evaluation for semester " + semester + " of " + course + " on " + month + " " + year, "No records of exam for given values", MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Something went wrong!");
                    return;
                }

                MessageBox.Show("Your files are going to be validated!", "Evaluator", MessageBoxButton.OK, MessageBoxImage.Information);
                tabControl.IsEnabled = false;
                this.Cursor = Cursors.Wait;
                txtEvalOutput.Clear();

                StudentType stype = StudentType.REGULAR;
                if (String.Compare("Regular", (cbxEType.SelectedItem as ComboBoxItem).Content.ToString(), true) == 0)
                    stype = StudentType.REGULAR;
                else if (String.Compare("Failed", (cbxEType.SelectedItem as ComboBoxItem).Content.ToString(), true) == 0)
                    stype = StudentType.FAILURE;
                else
                    stype = StudentType.ABSENT;

                (new Evaluator(txtIMFile.Text, txtEMFile.Text, txtCutoffFile.Text)).validateWorksheets(course, semester, year, month, stype);
            }
        }

        /// <summary>
        /// After marksheet and cutoff validation is completed 
        /// </summary>
        /// <param name="message">Some message about validation</param>
        /// <param name="isValid">Is the worksheet is valid and error free</param>
        public void ValidationCompleted(string message, bool isValid, StudentType stype)
        {
            tabControl.IsEnabled = true;
            this.Cursor = Cursors.Arrow;

            if (isValid)
            {
                if (MessageBox.Show(message, "Validation Result", MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                {
                    tabControl.IsEnabled = false;
                    this.Cursor = Cursors.Wait;

                    //get values from controls
                    int semester = (int)cbxESem.SelectedItem;
                    int year = Convert.ToInt16(cbxEYear.SelectedItem);
                    string course = (String)cbxECourse.SelectedItem;
                    string month = (cbxEMonth.SelectedItem as ComboBoxItem).Content.ToString(); //As the item is comboboxitem

                    //perform evaluate and push to database
                    (new Evaluator(txtIMFile.Text, txtEMFile.Text, txtCutoffFile.Text)).EvaluateAndPush(course,semester,year,month,stype);
                }
            }
            else
            {
                MessageBox.Show(message, "Validation Result", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
        }

        public void EvaluationCompleted(string message, bool isAllOk)
        {
            tabControl.IsEnabled = true;
            this.Cursor = Cursors.Arrow;

            MessageBox.Show(message, "Evaluation completed", MessageBoxButton.OK, isAllOk ? MessageBoxImage.Information : MessageBoxImage.Exclamation);
        }

        private void cbxECourse_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cbxECourse.SelectedIndex != -1)
                cbxESem.ItemsSource = Courses.getInSemList(cbxECourse.SelectedItem.ToString());
        }

        private void cbxEMonth_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //if (cbxEMonth.SelectedIndex != -1)
            //    MessageBox.Show((cbxEMonth.SelectedItem as ComboBoxItem).Content.ToString());
        }

        private void btnGenerateSheets_Click(object sender, RoutedEventArgs e)
        {
            (new generator_window()).ShowDialog();
        }

        private void cbxSCourse_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cbxSCourse.SelectedIndex != -1)
            {
                cbxSYear.ItemsSource = Students.getBatchList(cbxSCourse.SelectedItem.ToString());
                cbxSYear.Text = "-- SELECT YEAR --";
            }
        }

        private void btnShowStudents_Click(object sender, RoutedEventArgs e)
        {
            if (cbxSCourse.SelectedIndex == -1 || cbxSYear.SelectedIndex == -1)
            {
                MessageBox.Show("Select course and year first!", "Wait", MessageBoxButton.OK, MessageBoxImage.Stop);
                return;
            }

            lstvStudentList.ItemsSource = Students.getStudentList(cbxSCourse.SelectedItem.ToString(), (int)cbxSYear.SelectedItem);
        }

        private void cbxMCourse_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cbxMCourse.SelectedIndex != -1)
            {
                cbxMSem.ItemsSource = Results.getSemesterList(cbxMCourse.SelectedItem.ToString());
                cbxMSem.Text = "-- SEMESTER --";
                cbxMRegId.ItemsSource = null;
                cbxMYear.ItemsSource = null;
                cbxMMonth.ItemsSource = null;
                cbxMMonth.Text = "-- MONTH --";
                cbxMYear.Text = "-- YEAR --";
                cbxMRegId.Text = "-- REG ID --";
            }
        }

        private void cbxMSem_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(cbxMSem.SelectedIndex != -1 && cbxMCourse.SelectedIndex != -1)
            {
                cbxMYear.ItemsSource = Results.getYearList(cbxMCourse.SelectedItem.ToString(),(int)cbxMSem.SelectedItem);
                cbxMMonth.ItemsSource = null;
                cbxMMonth.Text = "-- MONTH --";
                cbxMRegId.ItemsSource = null;
                cbxMRegId.Text = "-- REG ID --";
            }
        }

        private void cbxMYear_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(cbxMYear.SelectedIndex != -1 && cbxMSem.SelectedIndex != -1)
            {
                cbxMMonth.ItemsSource = Results.getMonthList(cbxMCourse.SelectedItem.ToString(), (int)cbxMSem.SelectedItem,(int)cbxMYear.SelectedItem);
                cbxMRegId.ItemsSource = null;
                cbxMRegId.Text = "-- REG ID --";
            }
        }

        private void cbxMMonth_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cbxMMonth.SelectedIndex != -1)
            {
                cbxMRegId.ItemsSource = Results.getRegidList(cbxMCourse.SelectedItem.ToString(), (int)cbxMSem.SelectedItem, (int)cbxMYear.SelectedItem, (String)cbxMMonth.SelectedItem);
                cbxMRegId.Text = "-- REG ID --";
            }
        }

        private void btnShowMarks_Click(object sender, RoutedEventArgs e)
        {
            String errors = null;

            if (cbxMCourse.SelectedIndex == -1)
                errors += "Please select course" + Environment.NewLine;
            if (cbxMSem.SelectedIndex == -1)
                errors += "Please select semester" + Environment.NewLine;
            if (cbxMYear.SelectedIndex == -1)
                errors += "Please select year" + Environment.NewLine;
            if (cbxMRegId.SelectedIndex == -1)
                errors += "Please select registration id" + Environment.NewLine;

            if(errors!= null)
            {
                MessageBox.Show("Please correct the following error(s)!" + Environment.NewLine + Environment.NewLine + errors, "Invalid inputs", MessageBoxButton.OK, MessageBoxImage.Stop);
                return;
            }

            //fill listview
            lstvMStudentList.ItemsSource = Results.getExamResult(cbxMCourse.SelectedItem.ToString(), (int)cbxMSem.SelectedItem, (int)cbxMYear.SelectedItem,(long)cbxMRegId.SelectedItem,(String)cbxMMonth.SelectedItem);
            lblStudentGpa.Content = Results.getStudentGpa(cbxMCourse.SelectedItem.ToString(), (int)cbxMSem.SelectedItem, (int)cbxMYear.SelectedItem, (long)cbxMRegId.SelectedItem);
        }

        /// <summary>
        /// Fills listviews and comboboxes with relative data
        /// </summary>
        public void populateData()
        {
            lstvCourse.ItemsSource = Courses.getCourseList();
            lstvSubjectList.ItemsSource = null;
            cbxSCourse.ItemsSource = Courses.getCourseCodes(StudentType.REGULAR);
            cbxMCourse.ItemsSource = cbxSCourse.ItemsSource;
        }

        private void cbxEType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(cbxEType.SelectedIndex != -1)
            {
                if (cbxEType.SelectedIndex == 0)
                    cbxECourse.ItemsSource = Courses.getCourseCodes(StudentType.REGULAR);
                else if (cbxEType.SelectedIndex == 1)
                    cbxECourse.ItemsSource = Courses.getCourseCodes(StudentType.ABSENT);
                else
                    cbxECourse.ItemsSource = Courses.getCourseCodes(StudentType.FAILURE);

                //clear other combobox selection
                cbxECourse.Text = "-- SELECT COURSE --";
                cbxESem.ItemsSource = null;
                cbxESem.Text = "-- SEMESTER --";
                cbxEYear.ItemsSource = null;
                cbxEYear.Text = "-- YEAR --";
                cbxEMonth.ItemsSource = null;
                cbxEMonth.Text = "-- MONTH --";
            }
        }

        private void cbxESem_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(cbxESem.SelectedIndex != -1)
                cbxEYear.ItemsSource = Results.getExamYearsList();
        }

        private void cbxAType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cbxAType.SelectedIndex != -1)
            {
                if (cbxAType.SelectedIndex == 0)
                    cbxACourse.ItemsSource = Courses.getCourseCodes(StudentType.ABSENT);
                else
                    cbxACourse.ItemsSource = Courses.getCourseCodes(StudentType.FAILURE);

                //clear other combobox selection
                cbxACourse.Text = "-- SELECT COURSE --";
                cbxASem.ItemsSource = null;
                cbxASem.Text = "-- SEMESTER --";
                cbxAYear.ItemsSource = null;
                cbxAYear.Text = "-- YEAR --";
                cbxAMonth.ItemsSource = null;
                cbxAMonth.Text = "-- MONTH --";
            }
        }

        private void cbxACourse_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(cbxACourse.SelectedIndex != -1)
            {
                List<int> sems = new List<int>();
                foreach (Row row in Medatabase.fetchRecords("SELECT DISTINCT semester FROM exam_master WHERE course_code='" + (String)cbxACourse.SelectedItem + "'"))
                    sems.Add(Convert.ToInt16(row.column["semester"]));
                cbxASem.ItemsSource = sems;
                cbxASem.Text = "-- SEMESTER --";
                cbxAYear.ItemsSource = null;
                cbxAYear.Text = "-- YEAR --";
                cbxAMonth.ItemsSource = null;
                cbxAMonth.Text = "-- MONTH --";
            }
        }

        private void cbxASem_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(cbxASem.SelectedIndex != -1)
            {
                List<int> years = new List<int>();
                foreach (Row row in Medatabase.fetchRecords("SELECT DISTINCT year FROM exam_master WHERE course_code='" + (String)cbxACourse.SelectedItem + "' AND semester=" + Convert.ToInt16(cbxASem.SelectedItem)))
                    years.Add(Convert.ToInt16(row.column["year"]));
                cbxAYear.ItemsSource = years;
                cbxAYear.Text = "-- YEAR --";
                cbxAMonth.ItemsSource = null;
                cbxAMonth.Text = "-- MONTH --";
            }
        }

        private void cbxAYear_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cbxAYear.SelectedIndex != -1)
            {
                List<String> months = new List<String>();
                foreach (Row row in Medatabase.fetchRecords("SELECT DISTINCT month FROM exam_master WHERE course_code='" + (String)cbxACourse.SelectedItem + "' AND semester=" + Convert.ToInt16(cbxASem.SelectedItem)))
                    months.Add((String)row.column["month"]);
                cbxAMonth.ItemsSource = months;
            }
        }

        private void btnAShow_Click(object sender, RoutedEventArgs e)
        {
            bool isAllOk = true;

            //validation
            String error_message = null;
            if (cbxACourse.SelectedIndex == -1)
                error_message += "Please select course" + Environment.NewLine;
            if (cbxASem.SelectedIndex == -1)
                error_message += "Please select semester" + Environment.NewLine;
            if (cbxAMonth.SelectedIndex == -1)
                error_message += "Please select month" + Environment.NewLine;
            if (cbxAType.SelectedIndex == -1)
                error_message += "Please select student type" + Environment.NewLine;
            if (cbxAYear.SelectedIndex == -1)
                error_message += "Please select year" + Environment.NewLine;

            if (error_message != null)
            {
                isAllOk = false;
                MessageBox.Show("Please correct the following error(s)!" + Environment.NewLine + Environment.NewLine + error_message, "Invalid inputs", MessageBoxButton.OK, MessageBoxImage.Stop);
            }

            if (isAllOk)
            {
                //get values from controls
                int semester = (int)cbxASem.SelectedItem;
                int year = Convert.ToInt16(cbxAYear.SelectedItem);
                string course = (String)cbxACourse.SelectedItem;
                string month = (String)cbxAMonth.SelectedItem;

                StudentType stype = StudentType.FAILURE;
                if (String.Compare("Failed", (cbxAType.SelectedItem as ComboBoxItem).Content.ToString(), true) == 0)
                    stype = StudentType.FAILURE;
                else
                    stype = StudentType.ABSENT;

                List<NonregularCols> items = Results.getNonregularStudentList(course, semester, year, month, stype);
                if(items == null)
                    MessageBox.Show("Something went wrong while fetching data :(", "ERROR", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                else
                    lstvAnalysisList.ItemsSource = items;
            }
        }

        /// <summary>
        /// To be called after analysis report has been generated
        /// </summary>
        /// <param name="message">Information</param>
        /// <param name="success">Whether failed to generate excel sheet</param>
        public void analysisReportResult(String message,bool success)
        {
            tabControl.IsEnabled = true;
            this.Cursor = Cursors.Arrow;
            if (success)
                MessageBox.Show(message, "Analysis result", MessageBoxButton.OK, MessageBoxImage.Information);
            else
                MessageBox.Show(message, "Analysis result", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private void btnGenerateAnalysisSheet_Click(object sender, RoutedEventArgs e)
        {
            bool isAllOk = true;

            //validation
            String error_message = null;
            if (cbxACourse.SelectedIndex == -1)
                error_message += "Please select course" + Environment.NewLine;
            if (cbxASem.SelectedIndex == -1)
                error_message += "Please select semester" + Environment.NewLine;
            if (cbxAMonth.SelectedIndex == -1)
                error_message += "Please select month" + Environment.NewLine;
            if (cbxAType.SelectedIndex == -1)
                error_message += "Please select student type" + Environment.NewLine;
            if (cbxAYear.SelectedIndex == -1)
                error_message += "Please select year" + Environment.NewLine;

            if (error_message != null)
            {
                isAllOk = false;
                MessageBox.Show("Please correct the following error(s)!" + Environment.NewLine + Environment.NewLine + error_message, "Invalid inputs", MessageBoxButton.OK, MessageBoxImage.Stop);
            }

            if (isAllOk)
            {
                //get values from controls
                int semester = (int)cbxASem.SelectedItem;
                int year = Convert.ToInt16(cbxAYear.SelectedItem);
                string course = (String)cbxACourse.SelectedItem;
                string month = (String)cbxAMonth.SelectedItem;

                StudentType stype = StudentType.FAILURE;
                if (String.Compare("Failed", (cbxAType.SelectedItem as ComboBoxItem).Content.ToString(), true) == 0)
                    stype = StudentType.FAILURE;
                else
                    stype = StudentType.ABSENT;

                String output_folder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
                fbd.Description = "Select a folder where output files will be saved";
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    output_folder = fbd.SelectedPath;
                else
                    return;

                tabControl.IsEnabled = false;
                this.Cursor = Cursors.Wait;
                (new Results()).generateAnalysisResultSheet(course, semester, month, year, output_folder, stype);
            }
        }
    }

}
