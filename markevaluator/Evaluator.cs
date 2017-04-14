using System;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace markevaluator
{
    class Evaluator
    {
        //For updating the validation output 
        //and progressbar of admin_window.
        System.Windows.Controls.TextBox txtOutput = Windows.adminWindow.txtEvalOutput;
        System.Windows.Controls.ProgressBar progress  = Windows.adminWindow.prgsbarEval;

        //Background workers for performing validation and evaluation
        BackgroundWorker bgwValidate,bgwValidateNonRegular;
        BackgroundWorker bgwEvaluateAndPush,bgwEvaluateAndPushNonRegular;
        
        //Worksheet files
        String in_marksheet_file;
        String ex_marksheet_file;
        String cutoff_file;

        //Variables required for pushing to database
        String Ecourse;
        int Esemester;
        int Eyear;
        String Emonth;

         //Temporary Variables
        Application excelApp;
        private string tempProcessLog = "";
        private bool isValid = true;
        int invalidRows = 0;
        int cutoff_records_pushed = 0;
        int student_records_pushed = 0;
        int result_records_pushed = 0;
        string ln = Environment.NewLine;
        StudentType stype;

        public Evaluator(String in_marksheet_file, String ex_marksheet_file, String cutoff_file)
        {
            this.in_marksheet_file = in_marksheet_file;
            this.ex_marksheet_file = ex_marksheet_file;
            this.cutoff_file = cutoff_file;
        }

        /// <summary>
        /// For validating marksheets and cutoff file
        /// </summary>
        public void validateWorksheets(String Ecourse, int Esemester, int Eyear, String Emonth, StudentType stype)
        {
            this.stype = stype;
            this.Ecourse = Ecourse;
            this.Esemester = Esemester;
            this.Eyear = Eyear;
            this.Emonth = Emonth;

            if (stype == StudentType.REGULAR)
            {
                //Intantiate background worker
                bgwValidate = new BackgroundWorker();
                bgwValidate.DoWork += BgwValidate_DoWork;
                bgwValidate.ProgressChanged += BgwValidate_ProgressChanged;
                bgwValidate.RunWorkerCompleted += BgwValidate_RunWorkerCompleted;

                //Call background worker to perform validation
                bgwValidate.RunWorkerAsync();
                bgwValidate.WorkerReportsProgress = true;
                bgwValidate.WorkerSupportsCancellation = true;
            }
            else // For absentees and failures
            {
                //Intantiate background worker
                bgwValidateNonRegular = new BackgroundWorker();
                bgwValidateNonRegular.DoWork += BgwValidateNonRegular_DoWork;
                bgwValidateNonRegular.ProgressChanged += BgwValidateNonRegular_ProgressChanged;
                bgwValidateNonRegular.RunWorkerCompleted += BgwValidateNonRegular_RunWorkerCompleted;

                //Call background worker to perform validation
                bgwValidateNonRegular.RunWorkerAsync();
                bgwValidateNonRegular.WorkerReportsProgress = true;
                bgwValidateNonRegular.WorkerSupportsCancellation = true;
            }
        }

        private void BgwValidate_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //on validation complete
            //All ended
            invalidRows = 0;
            if (isValid)
                Windows.adminWindow.ValidationCompleted("Your files are validated. There are no error(s)!" + Environment.NewLine + "Evaluate marks and push data to database?", true,stype);
            else
                Windows.adminWindow.ValidationCompleted("Your worksheet file contains some errors, please see the logs." + Environment.NewLine + "Please correct the errors.", false,stype);
        }

        private void BgwValidate_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //Update progressbar value
            progress.Value = e.ProgressPercentage;   
        }

        private void BgwValidate_DoWork(object sender, DoWorkEventArgs e)
        {
            DateTime begin = DateTime.Now;
            try
            {
                //Temp variable required till validation
                int sheetCount = 0, invalid_sub_codes = 0;
                int rowCount = 0, columnCount = 0,electiveSubCount = 0;
                bool has_electives = false;

                List<String> sub_codes;
                List<long> reg_id;

                isValid = true;
                excelApp = new Application();


                /******************* MARKSHEETS VALIDATION *********************/

                //open excelsheet
                Workbook MarksheetWorkbook = excelApp.Workbooks.Open(in_marksheet_file, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet MarkWorksheet = (Worksheet)MarksheetWorkbook.Sheets.get_Item(1); //select the first sheet from excel file
                Range MarkWorksheetRange; //Sets used rows range

                //temp variable required till validation process
                sub_codes = new List<string>();
                reg_id = new List<long>();
                invalid_sub_codes = 0;

                //Internal Marksheet will be validated in the 1st loop and then External marksheet will be validated after 1st loop
                //For validating internal & external marksheet files as the validation rules are same.

                for (int k = 1; k <= 2; k++) //Exactly twice (1st loop for Internal, 2nd loop for external)
                {
                    MarkWorksheetRange = MarkWorksheet.UsedRange;
                    tempProcessLog = ln + ln + "============= VALIDATING " + ((k == 1) ? "INTERNAL" : "EXTERNAL") + " MARKSHEET ==============" + ln;
                    sheetCount = MarksheetWorkbook.Sheets.Count;
                    tempProcessLog += ln + "Reading worksheet..." + ln + "Total sheets : " + sheetCount;

                    tempProcessLog += ln + "Using Sheet " + MarkWorksheet.Name + "...";

                    rowCount = MarkWorksheetRange.Rows.Count;
                    columnCount = MarkWorksheetRange.Columns.Count;

                    //if there are no rows
                    if (rowCount <= 1)
                    {
                        tempProcessLog += ln + "There are no records to validate...";

                        App.Current.Dispatcher.Invoke(new System.Action(() =>
                        {
                            txtOutput.AppendText(tempProcessLog);
                            txtOutput.ScrollToEnd();
                            progress.Value = 100;
                        }));

                        isValid = false;
                        return;
                    }

                    //if columns are less than 4
                    if (columnCount < 4)
                    {
                        tempProcessLog += ln + "Expecting columns to be more than 3, " + columnCount + " is present." + ln + "This excelsheet is not meeting the desired format.";

                        App.Current.Dispatcher.Invoke(new System.Action(() =>
                        {
                            txtOutput.AppendText(tempProcessLog);
                            txtOutput.ScrollToEnd();
                            progress.Value = 100;
                        }));

                        isValid = false;
                        return;
                    }

                    tempProcessLog += ln + "Rows " + rowCount + " | Columns " + columnCount + ln;
                    tempProcessLog += ln + "Expecting first row to be column names. " + ln + "reading from 2nd row...";
                    tempProcessLog += ln + "Checking course_codes...";

                    //validate subject codes present in the first row
                    for (int i = 4; i <= columnCount; i++)
                    {
                        if ((String)MarkWorksheet.Cells[1, i].Value != null)
                        {
                            //check if the given subject code is present in database or not
                            if (!Medatabase.isPresent("subject_master", "sub_code", (String)MarkWorksheet.Cells[1, i].Value))
                            {
                                invalid_sub_codes++;
                                tempProcessLog += ln + "subject code '" + (string)MarkWorksheet.Cells[1, i].Value + "' is not present in database in row " + i;
                                isValid = false;
                                invalidRows++;
                            }

                            //Check for duplicate subject code in the same excel sheet
                            if (sub_codes.IndexOf((string)MarkWorksheet.Cells[1, i].Value) != -1)
                            {
                                tempProcessLog += ln + "duplicate subject code '" + (string)MarkWorksheet.Cells[1, i].Value + "' already present in the same excel sheet in row " + i;
                                isValid = false;
                                invalidRows++;
                            }
                            sub_codes.Add((string)MarkWorksheet.Cells[1, i].Value);
                        }
                        else //Means no value given for subject code
                        {
                            invalid_sub_codes++;
                            tempProcessLog += ln + "no value given for subject code in column " + i;
                            isValid = false;
                            invalidRows++;
                        }
                    }

                    // if there's something wrong with subject code....
                    if(invalid_sub_codes > 0)
                    {
                        App.Current.Dispatcher.Invoke(new System.Action(() =>
                        {
                            txtOutput.AppendText(tempProcessLog);
                            txtOutput.ScrollToEnd();
                            progress.Value = 100;
                        }));

                        isValid = false;
                        return;
                    }
                    
                    //cycle through all the rows starting from 2nd row as first row is heading.
                    for(int i = 2; i <= rowCount; i++)
                    {
                        try
                        {
                            //validate registration id check for duplicate values in the same sheet
                            if (MarkWorksheet.Cells[i,2].Value != null)
                            {
                                if (reg_id.IndexOf((long)MarkWorksheet.Cells[i, 2].Value) != -1)
                                {
                                    tempProcessLog += ln + "duplicate registration id '" + (long)MarkWorksheet.Cells[i, 2].Value + "' already present in the same excel sheet in row " + i;
                                    isValid = false;
                                    invalidRows++;
                                }
                                else //if not duplicate
                                    reg_id.Add((long)MarkWorksheet.Cells[i, 2].Value);
                            }
                            else
                            {
                                tempProcessLog += ln + "no value given for registration id in row " + i;
                                isValid = false;
                                invalidRows++;
                            }

                            //Check if student name is empty
                            if((String)MarkWorksheet.Cells[i,3].Value == null)
                            {
                                tempProcessLog += ln + "no value given for student name in row " + i;
                                isValid = false;
                                invalidRows++;
                            }

                            electiveSubCount = 0; //for each student
                            //cycle through all the cells containing marks in the current row for validation
                            for(int j = 4; j <= columnCount; j++)
                            {
                                try
                                {
                                    //check if subject is elective or not
                                    if (((String)MarkWorksheet.Cells[1, j].Value).IndexOf(".") == -1)
                                    { 
                                        //check if null or empty value is given in place of mark
                                        if (MarkWorksheet.Cells[i, j].Value != null)
                                        {
                                            if (k == 1) //Means internal marksheet is being validated
                                            {
                                                //Check if mark is more than 50 or less than 0
                                                if ((int)MarkWorksheet.Cells[i, j].Value > 50 || (int)MarkWorksheet.Cells[i, j].Value < 0)
                                                {
                                                    tempProcessLog += ln + "max mark for internal is 50, '" + (int)MarkWorksheet.Cells[i, j].Value + "' given in row " + i + " column " + j;
                                                    isValid = false;
                                                    invalidRows++;
                                                }
                                            }
                                            else //Means External marksheet is being validated
                                            {
                                                //Check if mark is more than 100 or less than 0
                                                if ((int)MarkWorksheet.Cells[i, j].Value > 100 || (int)MarkWorksheet.Cells[i, j].Value < 0)
                                                {
                                                    tempProcessLog += ln + "max mark for external is 100, '" + (int)MarkWorksheet.Cells[i, j].Value + "' given in row " + i + " column " + j;
                                                    isValid = false;
                                                    invalidRows++;
                                                }
                                            }
                                        }
                                        else //Means null or empty value is given
                                        {
                                            //means not elective subject but mark value not given
                                            tempProcessLog += ln + "no value given for mark in row " + i + " column " + j;
                                            isValid = false;
                                            invalidRows++;
                                        }
                                    }
                                    else //means subject is elective
                                    {
                                        has_electives = true;
                                        if (MarkWorksheet.Cells[i, j].Value != null) //if value is not null
                                        {
                                            if (k == 1) //Means internal marksheet is being validated
                                            {
                                                //Check if mark is more than 50 or less than 0
                                                if ((int)MarkWorksheet.Cells[i, j].Value > 50 || (int)MarkWorksheet.Cells[i, j].Value < 0)
                                                {
                                                    tempProcessLog += ln + "[elective] max mark for internal is 50, '" + (int)MarkWorksheet.Cells[i, j].Value + "' given in row " + i + " column " + j;
                                                    isValid = false;
                                                    invalidRows++;
                                                }
                                            }
                                            else //Means External marksheet is being validated
                                            {
                                                //Check if mark is more than 100 or less than 0
                                                if ((int)MarkWorksheet.Cells[i, j].Value > 100 || (int)MarkWorksheet.Cells[i, j].Value < 0)
                                                {
                                                    tempProcessLog += ln + "[elective] max mark for external is 100, '" + (int)MarkWorksheet.Cells[i, j].Value + "' given in row " + i + " column " + j;
                                                    isValid = false;
                                                    invalidRows++;
                                                }
                                            }
                                            electiveSubCount++;
                                        } //means null value given for elective
                                    }
                                }
                                catch(Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                                {
                                    //means value is string or character (student is [incompleted exam | absent])
                                    if (String.Compare((String)MarkWorksheet.Cells[i, j].Value, "I", true) != 0)
                                        tempProcessLog += ln + "in row [" + i + "] neither 'I' nor mark is given. [TREATED AS INCOMPLETE]";
                                }
                            } //finished validating all marks

                            if (has_electives) //if electives present
                            {
                                //if more than one elective mark is present or no elective mark is given
                                if (electiveSubCount == 0 || electiveSubCount > 1)
                                {
                                    tempProcessLog += ln + "expected marks for 1 elective subject, " + electiveSubCount + " elective subject(s) mark is present in excel sheet";
                                    invalidRows++;
                                    isValid = false;
                                }
                            }
                        }
                        catch(Microsoft.CSharp.RuntimeBinder.RuntimeBinderException be)
                        {
                            invalidRows++;
                            isValid = false;
                            tempProcessLog += ln + "In row [" + i + "] expecting numeric value, string given!";
                            LogWriter.WriteError("During course excel sheet validation", be.Message);
                        }

                        //Update output textbox
                        App.Current.Dispatcher.Invoke(new System.Action(() =>
                        {
                            txtOutput.AppendText(tempProcessLog);
                            txtOutput.ScrollToEnd();
                        }));

                        //report progress
                        bgwValidate.ReportProgress(Convert.ToInt16(((i * 100) / rowCount)));
                        tempProcessLog = ""; //clear process log for each row as it is printed in output textbox

                    } //finished validating all records or rows

                    //Change wrokbook and worksheet from Internal Marksheet to External Marksheet
                    MarksheetWorkbook = excelApp.Workbooks.Open(ex_marksheet_file, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    MarkWorksheet = (Worksheet)MarksheetWorkbook.Sheets.get_Item(1);

                    //reset
                    sub_codes = new List<string>();
                    reg_id = new List<long>();

                } //finished validating Internal and External Marksheets



                /******************* CUTOFF WORKSHEET VALIDATION *********************/


                //open excelsheet
                Workbook CutoffWorkbook = excelApp.Workbooks.Open(cutoff_file, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet CutoffWorksheet = (Worksheet)CutoffWorkbook.Sheets.get_Item(1); //select the first sheet from excel file
                Range CutoffWorksheetRange = CutoffWorksheet.UsedRange; //Sets used rows range

                //temp variable reuired till validation process
                sub_codes = new List<string>();
                reg_id = new List<long>();
                invalid_sub_codes = 0;

                tempProcessLog = ln + ln + "============= VALIDATING CUT OFF FILE ==============" + ln;
                sheetCount = CutoffWorkbook.Sheets.Count;
                tempProcessLog += ln + "Reading worksheet..." + ln + "Total sheets : " + sheetCount;

                tempProcessLog += ln + "Using Sheet " + CutoffWorksheet.Name + "...";

                rowCount = CutoffWorksheetRange.Rows.Count;
                columnCount = CutoffWorksheetRange.Columns.Count;

                //if there are no rows
                if (rowCount <= 1)
                {
                    tempProcessLog += ln + "There are no records to validate...";

                    App.Current.Dispatcher.Invoke(new System.Action(() =>
                    {
                        txtOutput.AppendText(tempProcessLog);
                        txtOutput.ScrollToEnd();
                        progress.Value = 100;
                    }));

                    isValid = false;
                    return;
                }

                //if columns are less than 8
                if (columnCount < 8)
                {
                    tempProcessLog += ln + "Expecting columns to be 8, " + columnCount + " is present." + ln + "This excelsheet is not meeting the desired format.";

                    App.Current.Dispatcher.Invoke(new System.Action(() =>
                    {
                        txtOutput.AppendText(tempProcessLog);
                        txtOutput.ScrollToEnd();
                        progress.Value = 100;
                    }));

                    isValid = false;
                    return;
                }

                tempProcessLog += ln + "Rows " + rowCount + " | Columns " + columnCount + ln;
                tempProcessLog += ln + "Expecting first row to be column names. " + ln + "reading from 2nd row...";
                tempProcessLog += ln + "Checking course_codes...";

                //cycle through each row for valiation
                for(int i = 2; i <= rowCount; i++)
                {
                    //SUBJECT CODE VALIDATION
                    if ((String)CutoffWorksheet.Cells[i, 1].Value != null)
                    {
                        //check if the given subject code is present in database or not
                        if (!Medatabase.isPresent("subject_master", "sub_code", (String)CutoffWorksheet.Cells[i, 1].Value))
                        {
                            invalid_sub_codes++;
                            tempProcessLog += ln + "subject code '" + (string)CutoffWorksheet.Cells[i, 1].Value + "' is not present in database in row " + i;
                            isValid = false;
                            invalidRows++;
                        }

                        //Check for duplicate subject code in the same excel sheet
                        if (sub_codes.IndexOf((string)CutoffWorksheet.Cells[i, 1].Value) != -1)
                        {
                            tempProcessLog += ln + "duplicate subject code '" + (string)CutoffWorksheet.Cells[i, 1].Value + "' already present in the same excel sheet in row " + i;
                            isValid = false;
                            invalidRows++;
                        }
                        sub_codes.Add((string)CutoffWorksheet.Cells[i, 1].Value);
                    }
                    else //Means no value given for subject code
                    {
                        invalid_sub_codes++;
                        tempProcessLog += ln + "no value given for subject code in row " + i;
                        isValid = false;
                        invalidRows++;
                    }

                    // if there's something wrong with subject code....
                    if (invalid_sub_codes > 0)
                    {
                        App.Current.Dispatcher.Invoke(new System.Action(() =>
                        {
                            txtOutput.AppendText(tempProcessLog);
                            txtOutput.ScrollToEnd();
                            progress.Value = 100;
                        }));

                        isValid = false;
                        return;
                    }

                    //CUTOFF RANGE VALIDATION
                    float PREV_RANGE = 101f; //initial value for A+

                    //cycle through each columns in the rows starting from 2nd column, 2nd row
                    for(int j = 2; j <= columnCount; j++)
                    {
                        try
                        {
                            //check if empty or null value is present
                            if (CutoffWorksheet.Cells[i, j].Value != null)
                            {
                                if((float)CutoffWorksheet.Cells[i, j].Value >= PREV_RANGE)
                                {
                                    tempProcessLog += ln + (String)CutoffWorksheet.Cells[1, j].Value + "'s value cannot be greater than or equal to " + (String)CutoffWorksheet.Cells[1, j - 1].Value + " in row " +  i + " column " + j;
                                    isValid = false;
                                    invalidRows++;
                                    break; //because the whole range will become invalid
                                }
                                else if((float)CutoffWorksheet.Cells[i, j].Value < 0)
                                {
                                    tempProcessLog += ln + (String)CutoffWorksheet.Cells[1, j].Value + "'s value cannot be less than 0 in row " + i + " column " + j;
                                    isValid = false;
                                    invalidRows++;
                                    break; //because the whole range will become invalid
                                }
                                PREV_RANGE = (float)CutoffWorksheet.Cells[i, j].Value; //to compare with next range
                            }
                            else //Means no value given 
                            {
                                tempProcessLog += ln + "no value given for '" + (String)CutoffWorksheet.Cells[1, j].Value + "' in column " + i;
                                isValid = false;
                                invalidRows++;
                            }
                        }
                        catch(Microsoft.CSharp.RuntimeBinder.RuntimeBinderException be)
                        {
                            invalidRows++;
                            isValid = false;
                            tempProcessLog += ln + "In row " + i + " column " + j + " expecting numeric value, string given!";
                            LogWriter.WriteError("During cutoff excel sheet validation", be.Message);
                        }
                    } //cutoff range validation completed

                    //Update output textbox
                    App.Current.Dispatcher.Invoke(new System.Action(() =>
                    {
                        txtOutput.AppendText(tempProcessLog);
                        txtOutput.ScrollToEnd();
                    }));

                    //report progress
                    bgwValidate.ReportProgress(Convert.ToInt16(((i * 100) / rowCount)));
                    tempProcessLog = ""; //clear process log for each row as it is printed int output textbox

                } //finished validating all rows
            }
            catch (System.IO.IOException ioe)
            {
                tempProcessLog += ln + "ERROR HANDLING FILE :(" + ln + "Please retry";
                LogWriter.WriteError("During marksheet and cutoff file validation", ioe.Message);
                isValid = false;
                return;
            }
            catch (databaseException)
            {
                tempProcessLog += ln + "ERROR OCCURED IN DATABASE OPERATION :(" + ln + "Please retry.";
                isValid = false;
                return;
            }
            catch (Exception ex)
            {
                tempProcessLog += ln + "UNEXPECTED ERROR OCCURED :(" + ln + "Please retry.";
                LogWriter.WriteError("During marksheet and cutoff file validation", ex.Message);
                isValid = false;
                return;
            }
            finally
            {
                tempProcessLog += ln + "Took " + String.Format("{0:0.00}", (DateTime.Now - begin).TotalSeconds) + " seconds";
                App.Current.Dispatcher.Invoke(new System.Action(() =>
                {
                    tempProcessLog += ln + "There are total " + invalidRows + " errors in all 3 files";
                    txtOutput.AppendText(tempProcessLog);
                    txtOutput.ScrollToEnd();
                    progress.Value = 100;
                }));

                excelApp.Workbooks.Close();
            }
        }

        /// <summary>
        /// For evaluating and pushing to database
        /// </summary>
        /// <param name="Ecourse">course code</param>
        /// <param name="Esemester">semester number</param>
        /// <param name="Eyear">year</param>
        /// <param name="Emonth">month</param>
        public void EvaluateAndPush(String Ecourse, int Esemester, int Eyear, String Emonth, StudentType stype)
        {
            this.stype = stype;
            this.Ecourse = Ecourse;
            this.Esemester = Esemester;
            this.Eyear = Eyear;
            this.Emonth = Emonth;

            if (stype == StudentType.REGULAR)
            {
                //Intantiate background worker
                bgwEvaluateAndPush = new BackgroundWorker();
                bgwEvaluateAndPush.DoWork += BgwEvaluateAndPush_DoWork;
                bgwEvaluateAndPush.ProgressChanged += BgwEvaluateAndPush_ProgressChanged;
                bgwEvaluateAndPush.RunWorkerCompleted += BgwEvaluateAndPush_RunWorkerCompleted;

                //Call background worker to perform validation
                bgwEvaluateAndPush.RunWorkerAsync();
                bgwEvaluateAndPush.WorkerReportsProgress = true;
                bgwEvaluateAndPush.WorkerSupportsCancellation = true;
            }
            else
            {
                bgwEvaluateAndPushNonRegular = new BackgroundWorker();
                bgwEvaluateAndPushNonRegular.DoWork += BgwEvaluateAndPushNonRegular_DoWork;
                bgwEvaluateAndPushNonRegular.ProgressChanged += BgwEvaluateAndPushNonRegular_ProgressChanged;
                bgwEvaluateAndPushNonRegular.RunWorkerCompleted += BgwEvaluateAndPushNonRegular_RunWorkerCompleted;

                bgwEvaluateAndPushNonRegular.RunWorkerAsync();
                bgwEvaluateAndPushNonRegular.WorkerReportsProgress = true;
                bgwEvaluateAndPushNonRegular.WorkerSupportsCancellation = true;
            }
        }


        /**************************************
            EVALUATION FOR NON-REGULAR STUDENTS 
          
         **************************************/

        private void BgwEvaluateAndPushNonRegular_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (!isValid)
                Windows.adminWindow.EvaluationCompleted("There were some errors during evaluation." + ln + "Please see the logs", false);
            else
                Windows.adminWindow.EvaluationCompleted("Students mark has been evaluated and updated successfully." + ln, true);
        }

        private void BgwEvaluateAndPushNonRegular_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //Update progressbar value
            progress.Value = e.ProgressPercentage;
        }

        private void BgwEvaluateAndPushNonRegular_DoWork(object sender, DoWorkEventArgs e)
        {
            DateTime begin = DateTime.Now;
            try
            {
                //Temp variable required till validation
                int sheetCount = 0;
                int rowCount = 0, columnCount = 0;
                tempProcessLog = "";
                float AP, A, B, C, D, E, F;
                string subject_code;

                isValid = true;
                excelApp = new Application();

                //open excelsheet
                Workbook CutoffWorkbook = excelApp.Workbooks.Open(cutoff_file, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Workbook INMarkWorkbook = excelApp.Workbooks.Open(in_marksheet_file, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Workbook EXMarkWorkbook = excelApp.Workbooks.Open(ex_marksheet_file, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                //select the first sheet from excel file
                Worksheet CutoffWorksheet = (Worksheet)CutoffWorkbook.Sheets.get_Item(1);
                Worksheet INMarkWorksheet = (Worksheet)INMarkWorkbook.Sheets.get_Item(1);
                Worksheet EXMarkWorksheet = (Worksheet)EXMarkWorkbook.Sheets.get_Item(1);

                //Sets used rows range
                Range CutoffWorksheetRange = CutoffWorksheet.UsedRange;
                Range INMarkWorksheetRange = INMarkWorksheet.UsedRange;
                Range EXMarkWorksheetRange = EXMarkWorksheet.UsedRange;


                /*************** PUSH CUTOFF TO DATABASE *******************/


                tempProcessLog = ln + ln + "============ PUSHING CUTOFF DETAILS TO DATABASE ==============" + ln;
                sheetCount = CutoffWorkbook.Sheets.Count;
                tempProcessLog += ln + "Reading worksheet..." + ln + "Total sheets : " + sheetCount;

                tempProcessLog += ln + "Using Sheet " + CutoffWorksheet.Name + "...";

                rowCount = CutoffWorksheetRange.Rows.Count;
                columnCount = CutoffWorksheetRange.Columns.Count;

                tempProcessLog += ln + "Rows " + rowCount + " | Columns " + columnCount + ln;

                //Cycle through is each row to push to database
                for (int i = 2; i <= rowCount; i++)
                {
                    subject_code = (String)CutoffWorksheet.Cells[i, 1].Value;
                    AP = (float)CutoffWorksheet.Cells[i, 2].Value;
                    A = (float)CutoffWorksheet.Cells[i, 3].Value;
                    B = (float)CutoffWorksheet.Cells[i, 4].Value;
                    C = (float)CutoffWorksheet.Cells[i, 5].Value;
                    D = (float)CutoffWorksheet.Cells[i, 6].Value;
                    E = (float)CutoffWorksheet.Cells[i, 7].Value;
                    F = (float)CutoffWorksheet.Cells[i, 8].Value;

                    tempProcessLog += ln + "> Processing Row: " + i;
                    cutoff_records_pushed += Medatabase.ExecuteQuery("UPDATE cutoff_master SET A+=" + AP + ", A=" + A + ", B=" + B + ", C=" + C + ", D=" + D + ", E=" + E + ", F=" + F + " WHERE course_code='" + Ecourse + "' AND sub_code='" + subject_code + "'");

                    //Update output textbox
                    App.Current.Dispatcher.Invoke(new System.Action(() =>
                    {
                        txtOutput.AppendText(tempProcessLog);
                        txtOutput.ScrollToEnd();
                    }));

                    bgwEvaluateAndPushNonRegular.ReportProgress(Convert.ToInt16(((i * 100) / rowCount)));
                    tempProcessLog = "";
                }

                /********************** EVALUATION *************************/

                String subject_type = null;
                float internal_mark = 0, external_mark = 0, final_mark, gpa = 0, cgpa;
                List<Row> rows = null,rows2;
                Random rnd = new Random();
                long reg_id,exam_id;
                String grade, pass_fail;
                int subject_credits, points, total_credits;

                columnCount = INMarkWorksheetRange.Columns.Count;
                rowCount = INMarkWorksheetRange.Rows.Count;

                tempProcessLog = ln + ln + "============= EVALUATION IN PROGRESS ==============" + ln;

                //cycle through each rows starting from 2nd row
                for (int i = 2; i <= rowCount; i++)
                {
                    gpa = 0;
                    reg_id = (long)INMarkWorksheet.Cells[i, 2].Value;

                    //fetch exam_id of current student(registration_id)
                    rows2 = Medatabase.fetchRecords("SELECT exam_id FROM exam_master WHERE registration_id=" + reg_id + " AND course_code='" + Ecourse + "' AND semester=" + Esemester + " AND month='" + Emonth + "'");
                    exam_id = (long)rows2[0].column["exam_id"];
                    tempProcessLog += ln + "> Processing Row: " + i;

                    //cycle through each column to retrive different subject marks starting from 4th column
                    for (int j = 4; j <= columnCount; j++)
                    {
                        if (INMarkWorksheet.Cells[1, j].Value != null)
                        {
                            subject_code = (String)INMarkWorksheet.Cells[1, j].Value;
                            //get cuttoff values for current (row and column) subject
                            rows = Medatabase.fetchRecords("SELECT * FROM cutoff_master WHERE course_code='" + Ecourse + "' AND sub_code='" + subject_code + "'");
                            AP = float.Parse(rows[0].column["A+"].ToString());
                            A = float.Parse(rows[0].column["A"].ToString());
                            B = float.Parse(rows[0].column["B"].ToString());
                            C = float.Parse(rows[0].column["C"].ToString());
                            D = float.Parse(rows[0].column["D"].ToString());
                            E = float.Parse(rows[0].column["E"].ToString());
                            F = float.Parse(rows[0].column["F"].ToString());

                            rows = Medatabase.fetchRecords("SELECT type FROM subject_master WHERE sub_code='" + subject_code + "'");
                            subject_type = (String)rows[0].column["type"];

                            //if subject is elective mark may contain null hence continue
                            if(subject_code.IndexOf(".") != -1)
                                if (INMarkWorksheet.Cells[i, j].Value == null && EXMarkWorksheet.Cells[i, j].Value == null)
                                    continue;

                            try
                            {
                                //Calculate final marks [if subject is not lab then (external mark / 2) else  (external marks / 1)]
                                internal_mark = (float)INMarkWorksheet.Cells[i, j].Value;

                                //external marks may be null for seminar and project, if so value will be given 0
                                if (EXMarkWorksheet.Cells[i, j].Value != null)
                                    external_mark = (float)EXMarkWorksheet.Cells[i, j].Value;
                                else
                                    external_mark = 0;

                                //valuation will be different based on subject type like theory,lab,seminar or project
                                if (String.Compare("Theory", subject_type, true) == 0)
                                    final_mark = internal_mark + (external_mark / 2);
                                else if (String.Compare("Lab", subject_type, true) == 0)
                                    final_mark = internal_mark + external_mark;
                                else //for project or seminar
                                    final_mark = internal_mark;
                            }
                            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) //means "I" is present [EXPECTED] in cell
                            {
                                final_mark = -1;
                                try
                                {
                                    if (String.Compare((String)EXMarkWorksheet.Cells[i, j].Value, "I", true) == 0)
                                        external_mark = -1;
                                }
                                catch (Exception)
                                {
                                    external_mark = (float)EXMarkWorksheet.Cells[i, j].Value;
                                }
                                try
                                {
                                    if (String.Compare((String)INMarkWorksheet.Cells[i, j].Value, "I", true) == 0)
                                        internal_mark = -1;
                                }
                                catch (Exception)
                                {
                                    internal_mark = (float)INMarkWorksheet.Cells[i, j].Value;
                                }
                            }

                            grade = calculateGrade(external_mark, final_mark, AP, A, B, C, D, E, F, subject_type);

                            //if student has failed then grade will not be given above C even if he/she scores more than that
                            if(stype == StudentType.FAILURE)
                            {
                                switch(grade)
                                {
                                    case "A+":
                                        grade = "C";
                                        break;
                                    case "A":
                                        grade = "C";
                                        break;
                                    case "B":
                                        grade = "C";
                                        break;
                                    default:
                                        break;
                                }
                            }

                            if (string.Compare(grade, "F") == 0)
                                pass_fail = "Fail";
                            else if (string.Compare(grade, "I") == 0)
                                pass_fail = "Absent";
                            else
                                pass_fail = "Pass";

                            rows = Medatabase.fetchRecords("SELECT credits FROM subject_master WHERE sub_code='" + subject_code + "'");
                            subject_credits = (int)rows[0].column["credits"];
                            points = generatePoints(grade);

                            gpa = gpa + (points * subject_credits);

                            result_records_pushed += Medatabase.ExecuteQuery("UPDATE student_marks SET internal_marks=" + internal_mark + ", external_marks=" + external_mark + ", final_marks=" + final_mark + ", grade='" + grade + "', status='" + pass_fail + "' WHERE sub_code='" + subject_code + "' AND exam_id=" + exam_id);
                        }
                    }

                    //cgpa calculation
                    rows = Medatabase.fetchRecords("SELECT total_credits FROM course_master WHERE course_code='" + Ecourse + "'");
                    total_credits = Decimal.ToInt32((Decimal)rows[0].column["total_credits"]);
                    cgpa = (gpa / total_credits);

                    Medatabase.ExecuteQuery("UPDATE student_cgpa SET cgpa=" + cgpa + " WHERE exam_id=" + exam_id);

                    //Update output textbox
                    App.Current.Dispatcher.Invoke(new System.Action(() =>
                    {
                        txtOutput.AppendText(tempProcessLog);
                        txtOutput.ScrollToEnd();
                    }));

                    bgwEvaluateAndPushNonRegular.ReportProgress(Convert.ToInt16(((i * 100) / rowCount)));
                    tempProcessLog = "";
                }
            }
            catch (databaseException)
            {
                tempProcessLog += ln + "ERROR OCCURED IN DATABASE OPERATION :(" + ln + "Please retry.";
                isValid = false;
                return;
            }
            catch (System.IO.IOException ioe)
            {
                tempProcessLog += ln + "ERROR HANDLING FILE :(" + ln + "Please retry";
                LogWriter.WriteError("During upload of marks [non-regular]", ioe.Message);
                isValid = false;
                return;
            }
            catch (Exception ex)
            {
                tempProcessLog += ln + "UNEXPECTED ERROR OCCURED :(" + ln + "Please retry.";
                LogWriter.WriteError("During upload of marks [non-regular]", ex.Message);
                isValid = false;
                return;
            }
            finally
            {
                tempProcessLog += ln + ln + result_records_pushed + " record(s) pushed to database (exam result)";
                tempProcessLog += ln + "Took " + String.Format("{0:0.00}", (DateTime.Now - begin).TotalSeconds) + " seconds";
                App.Current.Dispatcher.Invoke(new System.Action(() =>
                {
                    txtOutput.AppendText(tempProcessLog);
                    txtOutput.ScrollToEnd();
                    progress.Value = 100;
                }));

                excelApp.Workbooks.Close();
            }
        }


        /**********************************
            EVALUATION FOR REGULAR STUDENTS 
          
         **********************************/

        private void BgwEvaluateAndPush_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (!isValid)
                Windows.adminWindow.EvaluationCompleted("There were some errors during evaluation." + ln + "Please see the logs", false);
            else
                Windows.adminWindow.EvaluationCompleted("Marks has been evaluated and pushed to database." + ln, true);
        }

        private void BgwEvaluateAndPush_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //Update progressbar value
            progress.Value = e.ProgressPercentage;
        }

        private void BgwEvaluateAndPush_DoWork(object sender, DoWorkEventArgs e)
        {
            DateTime begin = DateTime.Now;
            try
            {
                //Temp variable required till validation
                int sheetCount = 0;
                int rowCount = 0, columnCount = 0;
                tempProcessLog = "";
                float AP, A, B, C, D, E, F;
                string subject_code;

                isValid = true;
                excelApp = new Application();

                //open excelsheet
                Workbook CutoffWorkbook = excelApp.Workbooks.Open(cutoff_file, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Workbook INMarkWorkbook = excelApp.Workbooks.Open(in_marksheet_file, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Workbook EXMarkWorkbook = excelApp.Workbooks.Open(ex_marksheet_file, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                //select the first sheet from excel file
                Worksheet CutoffWorksheet = (Worksheet)CutoffWorkbook.Sheets.get_Item(1); 
                Worksheet INMarkWorksheet = (Worksheet)INMarkWorkbook.Sheets.get_Item(1);
                Worksheet EXMarkWorksheet = (Worksheet)EXMarkWorkbook.Sheets.get_Item(1);

                //Sets used rows range
                Range CutoffWorksheetRange = CutoffWorksheet.UsedRange;
                Range INMarkWorksheetRange = INMarkWorksheet.UsedRange;
                Range EXMarkWorksheetRange = EXMarkWorksheet.UsedRange;

                
                /*************** PUSH CUTOFF TO DATABASE *******************/


                tempProcessLog = ln + ln + "============ PUSHING CUTOFF DETAILS TO DATABASE ==============" + ln;
                sheetCount = CutoffWorkbook.Sheets.Count;
                tempProcessLog += ln + "Reading worksheet..." + ln + "Total sheets : " + sheetCount;

                tempProcessLog += ln + "Using Sheet " + CutoffWorksheet.Name + "...";

                rowCount = CutoffWorksheetRange.Rows.Count;
                columnCount = CutoffWorksheetRange.Columns.Count;

                tempProcessLog += ln + "Rows " + rowCount + " | Columns " + columnCount + ln;

                //Cycle through is each row to push to database
                for(int i = 2; i <= rowCount; i++)
                {
                    subject_code = (String)CutoffWorksheet.Cells[i, 1].Value;
                    AP = (float)CutoffWorksheet.Cells[i, 2].Value;
                    A = (float)CutoffWorksheet.Cells[i, 3].Value;
                    B = (float)CutoffWorksheet.Cells[i, 4].Value;
                    C = (float)CutoffWorksheet.Cells[i, 5].Value;
                    D = (float)CutoffWorksheet.Cells[i, 6].Value;
                    E = (float)CutoffWorksheet.Cells[i, 7].Value;
                    F = (float)CutoffWorksheet.Cells[i, 8].Value;

                    tempProcessLog += ln + "> Processing Row: " + i;
                    cutoff_records_pushed += Medatabase.ExecuteQuery("INSERT INTO cutoff_master VALUES('" + Ecourse + "','" + subject_code + "'," + AP + "," + A + "," + B + "," + C + "," + D + "," + E + "," + F + ")");

                    //Update output textbox
                    App.Current.Dispatcher.Invoke(new System.Action(() =>
                    {
                        txtOutput.AppendText(tempProcessLog);
                        txtOutput.ScrollToEnd();
                    }));

                    bgwEvaluateAndPush.ReportProgress(Convert.ToInt16(((i * 100) / rowCount)));
                    tempProcessLog = "";
                }


                /********************** PUSH STUDENTS' DETAILS ***********************/

                //temp variables
                long reg_id;
                String name;

                columnCount = INMarkWorksheetRange.Columns.Count;
                rowCount = INMarkWorksheetRange.Rows.Count;

                tempProcessLog = ln + ln + "========= PUSHING STUDENTS' DETAILS TO DATABASE ==========" + ln;

                //cycle through each rows in both marskheet starting from 2nd row
                for (int i = 2; i <= rowCount; i++)
                {
                    reg_id = (long)INMarkWorksheet.Cells[i, 2].Value;
                    name = (String)INMarkWorksheet.Cells[i, 3].Value;

                    tempProcessLog += ln + "> Processing Row: " + i;

                    //check if a student detail already present in database
                    if (!Medatabase.isPresent("student_details", "registration_id", reg_id))
                        student_records_pushed += Medatabase.ExecuteQuery("INSERT INTO student_details VALUES(" + reg_id + ",'" + name + "','" + Ecourse + "'," + Eyear + ",'" + Emonth + "')");
                    else
                        tempProcessLog += ln + "[" + name + "][" + reg_id + "] Already present, skipped" + ln;

                    //Update output textbox
                    App.Current.Dispatcher.Invoke(new System.Action(() =>
                    {
                        txtOutput.AppendText(tempProcessLog);
                        txtOutput.ScrollToEnd();
                    }));

                    bgwEvaluateAndPush.ReportProgress(Convert.ToInt16(((i * 100) / rowCount)));
                    tempProcessLog = "";
                }

                /********************** EVALUATION *************************/

                String subject_type = null;
                float internal_mark = 0, external_mark = 0, final_mark,gpa=0,cgpa;
                List<Row> rows = null;
                Random rnd = new Random();
                long rnd_exam_id = 0;
                String grade, pass_fail;
                int subject_credits,points,total_credits;

                columnCount = INMarkWorksheetRange.Columns.Count;
                rowCount = INMarkWorksheetRange.Rows.Count;

                tempProcessLog = ln + ln + "============= EVALUATION IN PROGRESS ==============" + ln;
                
                //cycle through each rows starting from 2nd row
                for(int i = 2; i <= rowCount; i++)
                {
                    gpa = 0;
                    reg_id = (long)INMarkWorksheet.Cells[i, 2].Value;
                    tempProcessLog += ln + "> Processing Row: " + i;

                    //generate a random 10 digit long value for exam id and check if not already taken
                    while (true) //DANGER
                    {
                        rnd_exam_id = rnd.Next(1000000000, 1999999999);
                        if (!Medatabase.isPresent("exam_master", "exam_id", rnd_exam_id))
                            break;
                    }

                    //cycle through each column to retrive different subject marks starting from 4th column
                    for (int j = 4; j <= columnCount; j++)
                    {
                        subject_code = (String)INMarkWorksheet.Cells[1, j].Value;
                        //get cuttoff values for current (row and column) subject
                        rows = Medatabase.fetchRecords("SELECT * FROM cutoff_master WHERE course_code='" + Ecourse + "' AND sub_code='" + subject_code + "'");
                        AP = float.Parse(rows[0].column["A+"].ToString());
                        A = float.Parse(rows[0].column["A"].ToString());
                        B = float.Parse(rows[0].column["B"].ToString());
                        C = float.Parse(rows[0].column["C"].ToString());
                        D = float.Parse(rows[0].column["D"].ToString());
                        E = float.Parse(rows[0].column["E"].ToString());
                        F = float.Parse(rows[0].column["F"].ToString());

                        rows = Medatabase.fetchRecords("SELECT type FROM subject_master WHERE sub_code='" + subject_code + "'");
                        subject_type = (String)rows[0].column["type"];

                        //if mark is null then continue as this will be a elective subject
                        if (INMarkWorksheet.Cells[i, j].Value == null && EXMarkWorksheet.Cells[i, j].Value == null)
                            continue;

                        try
                        {
                            //Calculate final marks [if subject is not lab then (external mark / 2) else  (external marks / 1)]
                            internal_mark = (float)INMarkWorksheet.Cells[i, j].Value;

                            //external marks may be null for seminar and project, if so value will be given 0
                            if (EXMarkWorksheet.Cells[i, j].Value != null)
                                external_mark = (float)EXMarkWorksheet.Cells[i, j].Value;
                            else
                                external_mark = 0;

                            //valuation will be different based on subject type like theory,lab,seminar or project
                            if (String.Compare("Theory", subject_type, true) == 0)
                                final_mark = internal_mark + (external_mark / 2);
                            else if (String.Compare("Lab", subject_type, true) == 0)
                                final_mark = internal_mark + external_mark;
                            else //for project or seminar
                                final_mark = internal_mark;
                        }
                        catch(Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) //means "I" is present [EXPECTED] in cell
                        {
                            final_mark = -1;
                            try
                            {
                                 if(String.Compare((String)EXMarkWorksheet.Cells[i, j].Value, "I", true) == 0)
                                    external_mark = -1;
                            }
                            catch(Exception)
                            {
                                external_mark = (float)EXMarkWorksheet.Cells[i, j].Value;
                            }
                            try
                            {
                                if (String.Compare((String)INMarkWorksheet.Cells[i, j].Value, "I", true) == 0)
                                    internal_mark = -1;
                            }
                            catch(Exception)
                            {
                                internal_mark = (float)INMarkWorksheet.Cells[i, j].Value;
                            }
                        }

                        grade = calculateGrade(external_mark,final_mark,AP,A,B,C,D,E,F,subject_type);

                        if (string.Compare(grade, "F") == 0)
                            pass_fail = "Fail";
                        else if (string.Compare(grade, "I") == 0)
                            pass_fail = "Absent";
                        else
                            pass_fail = "Pass";

                        rows = Medatabase.fetchRecords("SELECT credits FROM subject_master WHERE sub_code='" + subject_code + "'");
                        subject_credits = (int)rows[0].column["credits"];
                        points = generatePoints(grade);

                        gpa = gpa + (points * subject_credits);

                        result_records_pushed += Medatabase.ExecuteQuery("INSERT INTO student_marks VALUES(" + rnd_exam_id + ",'" + subject_code + "','" + grade + "'," + internal_mark + "," + external_mark + "," + final_mark + ",'" + pass_fail + "')");
                    }

                    //cgpa calculation
                    rows = Medatabase.fetchRecords("SELECT SUM(credits) AS total_credits FROM subject_master WHERE course_code='" + Ecourse + "'");
                    total_credits = Decimal.ToInt32((Decimal)rows[0].column["total_credits"]);
                    cgpa = (gpa / total_credits);

                    Medatabase.ExecuteQuery("INSERT INTO exam_master VALUES(" + rnd_exam_id + "," + reg_id + ",'" + Ecourse + "'," + Esemester + "," + Eyear + ",'" + Emonth + "')");
                    Medatabase.ExecuteQuery("INSERT INTO student_cgpa VALUES(" + rnd_exam_id + "," + cgpa + ")");

                    //Update output textbox
                    App.Current.Dispatcher.Invoke(new System.Action(() =>
                    {
                        txtOutput.AppendText(tempProcessLog);
                        txtOutput.ScrollToEnd();
                    }));

                    bgwEvaluateAndPush.ReportProgress(Convert.ToInt16(((i * 100) / rowCount)));
                    tempProcessLog = "";
                }
            }
            catch (databaseException)
            {
                tempProcessLog += ln + "ERROR OCCURED IN DATABASE OPERATION :(" + ln + "Please retry.";
                isValid = false;
                return;
            }
            catch (System.IO.IOException ioe)
            {
                tempProcessLog += ln + "ERROR HANDLING FILE :(" + ln + "Please retry";
                LogWriter.WriteError("During upload of marks", ioe.Message);
                isValid = false;
                return;
            }
            catch (Exception ex)
            {
                tempProcessLog += ln + "UNEXPECTED ERROR OCCURED :(" + ln + "Please retry.";
                LogWriter.WriteError("During upload of marks", ex.Message);
                isValid = false;
                return;
            }
            finally
            {
                tempProcessLog += ln + ln + result_records_pushed + " record(s) pushed to database (exam result)";
                tempProcessLog += ln + "Took " + String.Format("{0:0.00}", (DateTime.Now - begin).TotalSeconds) + " seconds";
                App.Current.Dispatcher.Invoke(new System.Action(() =>
                {
                    txtOutput.AppendText(tempProcessLog);
                    txtOutput.ScrollToEnd();
                    progress.Value = 100;
                }));

                excelApp.Workbooks.Close();
            }
        }

        /***************************************
            VALIDATIONS FOR NON-REGULAR STUDENTS

        *****************************************/

        private void BgwValidateNonRegular_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //on validation complete
            //All ended
            invalidRows = 0;
            if (isValid)
                Windows.adminWindow.ValidationCompleted("Your files are validated. There are no error(s)!" + Environment.NewLine + "Evaluate marks and push data to database?", true, stype);
            else
                Windows.adminWindow.ValidationCompleted("Your worksheet file contains some errors, please see the logs." + Environment.NewLine + "Please correct the errors.", false, stype);
        }

        private void BgwValidateNonRegular_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //Update progressbar value
            progress.Value = e.ProgressPercentage;
        }

        private void BgwValidateNonRegular_DoWork(object sender, DoWorkEventArgs e)
        {
            DateTime begin = DateTime.Now;
            try
            {
                //Temp variable required till validation
                int sheetCount = 0, invalid_sub_codes = 0, missing_sub_codes = 0, missing_reg_id = 0;
                int rowCount = 0, columnCount = 0, failed_sub_mark_count = 0;

                List<String> sub_codes;
                List<String> failed_sub_codes;
                List<String> std_failed_subs;
                List<long> reg_id;
                List<long> failed_reg_id;
                List<Row> rows,rows2;

                isValid = true;
                excelApp = new Application();


                /******************* MARKSHEETS VALIDATION *********************/

                //open excelsheet
                Workbook MarksheetWorkbook = excelApp.Workbooks.Open(in_marksheet_file, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet MarkWorksheet = (Worksheet)MarksheetWorkbook.Sheets.get_Item(1); //select the first sheet from excel file
                Range MarkWorksheetRange; //Sets used rows range

                //temp variable required till validation process
                sub_codes = new List<string>();
                reg_id = new List<long>();
                invalid_sub_codes = 0;

                //Internal Marksheet will be validated in the 1st loop and then External marksheet will be validated after 1st loop
                //For validating internal & external marksheet files as the validation rules are same.

                for (int k = 1; k <= 2; k++) //Exactly twice (1st loop for Internal, 2nd loop for external)
                {
                    MarkWorksheetRange = MarkWorksheet.UsedRange;
                    tempProcessLog = ln + ln + "============= VALIDATING " + ((k == 1) ? "INTERNAL" : "EXTERNAL") + " MARKSHEET ==============" + ln;
                    sheetCount = MarksheetWorkbook.Sheets.Count;
                    tempProcessLog += ln + "Reading worksheet..." + ln + "Total sheets : " + sheetCount;

                    tempProcessLog += ln + "Using Sheet " + MarkWorksheet.Name + "...";

                    rowCount = MarkWorksheetRange.Rows.Count;
                    columnCount = MarkWorksheetRange.Columns.Count;

                    //if there are no rows
                    if (rowCount <= 1)
                    {
                        tempProcessLog += ln + "There are no records to validate...";

                        App.Current.Dispatcher.Invoke(new System.Action(() =>
                        {
                            txtOutput.AppendText(tempProcessLog);
                            txtOutput.ScrollToEnd();
                            progress.Value = 100;
                        }));

                        isValid = false;
                        return;
                    }

                    //if columns are less than 4
                    if (columnCount < 4)
                    {
                        tempProcessLog += ln + "Expecting columns to be more than 3, " + columnCount + " is present." + ln + "This excelsheet is not meeting the desired format.";

                        App.Current.Dispatcher.Invoke(new System.Action(() =>
                        {
                            txtOutput.AppendText(tempProcessLog);
                            txtOutput.ScrollToEnd();
                            progress.Value = 100;
                        }));

                        isValid = false;
                        return;
                    }

                    tempProcessLog += ln + "Rows " + rowCount + " | Columns " + columnCount + ln;
                    tempProcessLog += ln + "Expecting first row to be column names. " + ln + "reading from 2nd row...";
                    tempProcessLog += ln + "Checking course_codes...";

                    //validate subject codes present in the first row
                    for (int i = 4; i <= columnCount; i++)
                    {
                        if ((String)MarkWorksheet.Cells[1, i].Value != null)
                        {
                            //check if the given subject code is present in database or not
                            if (!Medatabase.isPresent("subject_master", "sub_code", (String)MarkWorksheet.Cells[1, i].Value))
                            {
                                invalid_sub_codes++;
                                tempProcessLog += ln + "subject code '" + (string)MarkWorksheet.Cells[1, i].Value + "' is not present in database in row " + i;
                                isValid = false;
                                invalidRows++;
                            }

                            //Check for duplicate subject code in the same excel sheet
                            if (sub_codes.IndexOf((string)MarkWorksheet.Cells[1, i].Value) != -1)
                            {
                                tempProcessLog += ln + "duplicate subject code '" + (string)MarkWorksheet.Cells[1, i].Value + "' already present in the same excel sheet in row " + i;
                                isValid = false;
                                invalidRows++;
                            }
                            sub_codes.Add((string)MarkWorksheet.Cells[1, i].Value);
                        }
                        else //Means no value given for subject code
                        {
                            invalid_sub_codes++;
                            tempProcessLog += ln + "no value given for subject code in column " + i;
                            isValid = false;
                            invalidRows++;
                        }
                    }

                    // if there's something wrong with subject code....
                    if (invalid_sub_codes > 0)
                    {
                        App.Current.Dispatcher.Invoke(new System.Action(() =>
                        {
                            txtOutput.AppendText(tempProcessLog);
                            txtOutput.ScrollToEnd();
                            progress.Value = 100;
                        }));

                        isValid = false;
                        return;
                    }

                    //list all subject codes in which student have failed
                    rows = Medatabase.fetchRecords("SELECT DISTINCT sub_code FROM student_marks AS s,exam_master AS e WHERE s.status!='Pass' AND s.exam_id=e.exam_id AND e.course_code='" + Ecourse + "' AND e.semester=" + Esemester + " AND e.month='" + Emonth + "'");
                    failed_sub_codes = new List<string>();
                    foreach (Row row in rows)
                        failed_sub_codes.Add((String)row.column["sub_code"]);

                    //check if subject codes present in the excel file in which students have failed
                    foreach (String failedsc in failed_sub_codes)
                    {
                        if (sub_codes.IndexOf(failedsc) == -1)
                        {
                            tempProcessLog += ln + "'" + failedsc + "' is required but not found in excel sheet!";
                            missing_sub_codes++;
                        }
                    }

                    // if required subcode is not present in excelsheet
                    if (missing_sub_codes > 0)
                    {
                        App.Current.Dispatcher.Invoke(new System.Action(() =>
                        {
                            txtOutput.AppendText(tempProcessLog);
                            txtOutput.ScrollToEnd();
                            progress.Value = 100;
                        }));

                        isValid = false;
                        break;
                    }

                    //list all failed student list
                    rows = Medatabase.fetchRecords("SELECT DISTINCT registration_id FROM exam_master AS e, student_marks AS s WHERE e.course_code='" + Ecourse + "' AND e.semester=" + Esemester + " AND e.exam_id=s.exam_id AND s.status!='Pass'");
                    failed_reg_id = new List<long>();
                    foreach (Row row in rows)
                        failed_reg_id.Add((long)row.column["registration_id"]);

                    if(failed_reg_id.Count == 0)
                    {
                        tempProcessLog += ln + "There are no students who have either failed or Absent in any sbuject.";
                        App.Current.Dispatcher.Invoke(new System.Action(() =>
                        {
                            txtOutput.AppendText(tempProcessLog);
                            txtOutput.ScrollToEnd();
                            progress.Value = 100;
                        }));

                        isValid = false;
                        return;
                    }

                    //cycle through all the rows starting from 2nd row as first row is heading.
                    for (int i = 2; i <= rowCount; i++)
                    {
                        try
                        {
                            //validate registration id check for duplicate values in the same sheet
                            if (MarkWorksheet.Cells[i, 2].Value != null)
                            {
                                //check if registration_id is present in the failed student list or not
                                if (failed_reg_id.IndexOf((long)MarkWorksheet.Cells[i, 2].Value) != -1)
                                {
                                    //means student has passed in exam
                                    tempProcessLog += ln + (long)MarkWorksheet.Cells[i, 2].Value + " is a regular and has passed the exam but still given";
                                    isValid = false;
                                    invalidRows++;
                                }
                                else //means student has not passed in exam
                                {
                                    if (reg_id.IndexOf((long)MarkWorksheet.Cells[i, 2].Value) != -1)
                                    {
                                        tempProcessLog += ln + "duplicate registration id '" + (long)MarkWorksheet.Cells[i, 2].Value + "' already present in the same excel sheet in row " + i;
                                        isValid = false;
                                        invalidRows++;
                                        continue; //can't procceed to validate marks for duplicate reg id's
                                    }
                                    else //if not duplicate
                                        reg_id.Add((long)MarkWorksheet.Cells[i, 2].Value);
                                }
                            }
                            else
                            {
                                tempProcessLog += ln + "no value given for registration id in row " + i;
                                isValid = false;
                                invalidRows++;
                                continue; //Can't proceed to validate marks as registration_id itself is missing
                            }

                            //failed sub_code list for current regristration_id
                            failed_sub_mark_count = 0;
                            std_failed_subs = new List<string>();
                            rows2 = Medatabase.fetchRecords("SELECT sub_code FROM student_marks AS s, exam_master AS e WHERE s.status!='Pass' AND s.exam_id=e.exam_id AND e.course_code='" + Ecourse + "' AND e.semester=" + Esemester + " AND e.month='" + Emonth + "' AND e.registration_id=" + (long)MarkWorksheet.Cells[i, 2].Value);

                            foreach (Row rtt in rows2)
                                std_failed_subs.Add((String)rtt.column["sub_code"]);

                            //if sub_count equals 0 then continue as student is not failed in any subject
                            if (std_failed_subs.Count == 0)
                                continue;

                            //cycle through all the cells containing marks in the current row for validation
                            for (int j = 4; j <= columnCount; j++)
                            {
                                if (MarkWorksheet.Cells[i, j].Value != null)
                                {
                                    //check if mark is given for failed subject or not
                                    if (std_failed_subs.IndexOf((String)MarkWorksheet.Cells[1, j].Value) != -1)
                                    {
                                        if (k == 1) //Means internal marksheet is being validated
                                        {
                                            //Check if mark is more than 50 or less than 0
                                            if ((int)MarkWorksheet.Cells[i, j].Value > 50 || (int)MarkWorksheet.Cells[i, j].Value < 0)
                                            {
                                                tempProcessLog += ln + "max mark for internal is 50, '" + (int)MarkWorksheet.Cells[i, j].Value + "' given in row " + i + " column " + j;
                                                isValid = false;
                                                invalidRows++;
                                            }
                                        }
                                        else //Means External marksheet is being validated
                                        {
                                            //Check if mark is more than 100 or less than 0
                                            if ((int)MarkWorksheet.Cells[i, j].Value > 100 || (int)MarkWorksheet.Cells[i, j].Value < 0)
                                            {
                                                tempProcessLog += ln + "max mark for external is 100, '" + (int)MarkWorksheet.Cells[i, j].Value + "' given in row " + i + " column " + j;
                                                isValid = false;
                                                invalidRows++;
                                            }
                                        }
                                        failed_sub_mark_count++;
                                    }
                                    else //means mark is given but not for failed subject
                                    {
                                        tempProcessLog += ln + (long)MarkWorksheet.Cells[i, 2].Value + " has not failed or incompleted  in " + (String)MarkWorksheet.Cells[1, j].Value + " but mark is still given.";
                                        invalidRows++;
                                        isValid = false;
                                    }
                                }
                                else //if null value given in place of mark
                                {
                                    //check if null value is given in place of mark for failed subject
                                    if (std_failed_subs.IndexOf((String)MarkWorksheet.Cells[1, j].Value) != -1)
                                    {
                                        //its ok to have null value as the mark is not of a failed subject
                                    }
                                    else //means null is given in place of mark for failed subject
                                    {
                                        tempProcessLog += ln + "no mark is given for the failed or incomplete subject " + (String)MarkWorksheet.Cells[1, j].Value + " of " + (long)MarkWorksheet.Cells[i, 2].Value;
                                        invalidRows++;
                                        isValid = false;
                                    }
                                }
                            } //finished validating all marks
                        }
                        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException be)
                        {
                            invalidRows++;
                            isValid = false;
                            tempProcessLog += ln + "In row [" + i + "] expecting numeric value, unknown type given!";
                            LogWriter.WriteError("During course excel sheet validation [non-regular]", be.Message);
                        }

                        //Update output textbox
                        App.Current.Dispatcher.Invoke(new System.Action(() =>
                        {
                            txtOutput.AppendText(tempProcessLog);
                            txtOutput.ScrollToEnd();
                        }));

                        //report progress
                        bgwValidateNonRegular.ReportProgress(Convert.ToInt16(((i * 100) / rowCount)));
                        tempProcessLog = ""; //clear process log for each row as it is printed in output textbox

                    } //finished validating all records or rows

                    //check if failed students' registration id is present or not
                    foreach (long failedrid in failed_reg_id)
                    {
                        if (failed_reg_id.IndexOf(failedrid) == -1)
                        {
                            tempProcessLog += ln + "'" + failedrid + "' is required but not found in excel sheet!";
                            missing_reg_id++;
                        }
                    }

                    //if required reg_id is not present in excelsheet or failed students count mismatch
                    if (missing_reg_id > 0 || failed_reg_id.Count != reg_id.Count)
                    {
                        if (failed_reg_id.Count != reg_id.Count)
                            tempProcessLog += ln + "there are " + failed_reg_id.Count + " non-regular student(s) but " + reg_id.Count + " is present in excel sheet!";
                        App.Current.Dispatcher.Invoke(new System.Action(() =>
                        {
                            txtOutput.AppendText(tempProcessLog);
                            txtOutput.ScrollToEnd();
                            progress.Value = 100;
                        }));

                        isValid = false;
                        //return;
                    }
                    //Change wrokbook and worksheet from Internal Marksheet to External Marksheet
                    MarksheetWorkbook = excelApp.Workbooks.Open(ex_marksheet_file, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    MarkWorksheet = (Worksheet)MarksheetWorkbook.Sheets.get_Item(1);

                    //reset
                    sub_codes = new List<string>();
                    reg_id = new List<long>();

                } //finished validating Internal and External Marksheets



                /******************* CUTOFF WORKSHEET VALIDATION *********************/


                //open excelsheet
                Workbook CutoffWorkbook = excelApp.Workbooks.Open(cutoff_file, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet CutoffWorksheet = (Worksheet)CutoffWorkbook.Sheets.get_Item(1); //select the first sheet from excel file
                Range CutoffWorksheetRange = CutoffWorksheet.UsedRange; //Sets used rows range

                //temp variable reuired till validation process
                sub_codes = new List<string>();
                reg_id = new List<long>();
                invalid_sub_codes = 0;

                tempProcessLog = ln + ln + "============= VALIDATING CUT OFF FILE ==============" + ln;
                sheetCount = CutoffWorkbook.Sheets.Count;
                tempProcessLog += ln + "Reading worksheet..." + ln + "Total sheets : " + sheetCount;

                tempProcessLog += ln + "Using Sheet " + CutoffWorksheet.Name + "...";

                rowCount = CutoffWorksheetRange.Rows.Count;
                columnCount = CutoffWorksheetRange.Columns.Count;

                //if there are no rows
                if (rowCount <= 1)
                {
                    tempProcessLog += ln + "There are no records to validate...";

                    App.Current.Dispatcher.Invoke(new System.Action(() =>
                    {
                        txtOutput.AppendText(tempProcessLog);
                        txtOutput.ScrollToEnd();
                        progress.Value = 100;
                    }));

                    isValid = false;
                    return;
                }

                //if columns are less than 8
                if (columnCount < 8)
                {
                    tempProcessLog += ln + "Expecting columns to be 8, " + columnCount + " is present." + ln + "This excelsheet is not meeting the desired format.";

                    App.Current.Dispatcher.Invoke(new System.Action(() =>
                    {
                        txtOutput.AppendText(tempProcessLog);
                        txtOutput.ScrollToEnd();
                        progress.Value = 100;
                    }));

                    isValid = false;
                    return;
                }

                tempProcessLog += ln + "Rows " + rowCount + " | Columns " + columnCount + ln;
                tempProcessLog += ln + "Expecting first row to be column names. " + ln + "reading from 2nd row...";
                tempProcessLog += ln + "Checking course_codes...";

                //cycle through each row for valiation
                for (int i = 2; i <= rowCount; i++)
                {
                    //SUBJECT CODE VALIDATION
                    if ((String)CutoffWorksheet.Cells[i, 1].Value != null)
                    {
                        //check if the given subject code is present in database or not
                        if (!Medatabase.isPresent("subject_master", "sub_code", (String)CutoffWorksheet.Cells[i, 1].Value))
                        {
                            invalid_sub_codes++;
                            tempProcessLog += ln + "subject code '" + (string)CutoffWorksheet.Cells[i, 1].Value + "' is not present in database in row " + i;
                            isValid = false;
                            invalidRows++;
                        }

                        //Check for duplicate subject code in the same excel sheet
                        if (sub_codes.IndexOf((string)CutoffWorksheet.Cells[i, 1].Value) != -1)
                        {
                            tempProcessLog += ln + "duplicate subject code '" + (string)CutoffWorksheet.Cells[i, 1].Value + "' already present in the same excel sheet in row " + i;
                            isValid = false;
                            invalidRows++;
                        }
                        sub_codes.Add((string)CutoffWorksheet.Cells[i, 1].Value);
                    }
                    else //Means no value given for subject code
                    {
                        invalid_sub_codes++;
                        tempProcessLog += ln + "no value given for subject code in row " + i;
                        isValid = false;
                        invalidRows++;
                    }

                    // if there's something wrong with subject code....
                    if (invalid_sub_codes > 0)
                    {
                        App.Current.Dispatcher.Invoke(new System.Action(() =>
                        {
                            txtOutput.AppendText(tempProcessLog);
                            txtOutput.ScrollToEnd();
                            progress.Value = 100;
                        }));

                        isValid = false;
                        return;
                    }

                    //CUTOFF RANGE VALIDATION
                    float PREV_RANGE = 101f; //initial value for A+

                    //cycle through each columns in the rows starting from 2nd column, 2nd row
                    for (int j = 2; j <= columnCount; j++)
                    {
                        try
                        {
                            //check if empty or null value is present
                            if (CutoffWorksheet.Cells[i, j].Value != null)
                            {
                                if ((float)CutoffWorksheet.Cells[i, j].Value >= PREV_RANGE)
                                {
                                    tempProcessLog += ln + (String)CutoffWorksheet.Cells[1, j].Value + "'s value cannot be greater than or equal to " + (String)CutoffWorksheet.Cells[1, j - 1].Value + " in row " + i + " column " + j;
                                    isValid = false;
                                    invalidRows++;
                                    break; //because the whole range will become invalid
                                }
                                else if ((float)CutoffWorksheet.Cells[i, j].Value < 0)
                                {
                                    tempProcessLog += ln + (String)CutoffWorksheet.Cells[1, j].Value + "'s value cannot be less than 0 in row " + i + " column " + j;
                                    isValid = false;
                                    invalidRows++;
                                    break; //because the whole range will become invalid
                                }
                                PREV_RANGE = (float)CutoffWorksheet.Cells[i, j].Value; //to compare with next range
                            }
                            else //Means no value given 
                            {
                                tempProcessLog += ln + "no value given for '" + (String)CutoffWorksheet.Cells[1, j].Value + "' in column " + i;
                                isValid = false;
                                invalidRows++;
                            }
                        }
                        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException be)
                        {
                            invalidRows++;
                            isValid = false;
                            tempProcessLog += ln + "In row " + i + " column " + j + " expecting numeric value, string given!";
                            LogWriter.WriteError("During cutoff excel sheet validation", be.Message);
                        }
                    } //cutoff range validation completed

                    //Update output textbox
                    App.Current.Dispatcher.Invoke(new System.Action(() =>
                    {
                        txtOutput.AppendText(tempProcessLog);
                        txtOutput.ScrollToEnd();
                    }));

                    //report progress
                    bgwValidateNonRegular.ReportProgress(Convert.ToInt16(((i * 100) / rowCount)));
                    tempProcessLog = ""; //clear process log for each row as it is printed int output textbox

                } //finished validating all rows
            }
            catch (System.IO.IOException ioe)
            {
                tempProcessLog += ln + "ERROR HANDLING FILE :(" + ln + "Please retry";
                LogWriter.WriteError("During marksheet and cutoff file validation [non-regular]", ioe.Message);
                isValid = false;
                return;
            }
            catch (databaseException)
            {
                tempProcessLog += ln + "ERROR OCCURED IN DATABASE OPERATION :(" + ln + "Please retry.";
                isValid = false;
                return;
            }
            catch (Exception ex)
            {
                tempProcessLog += ln + "UNEXPECTED ERROR OCCURED :(" + ln + "Please retry.";
                LogWriter.WriteError("During marksheet and cutoff file validation [non-regular]", ex.Message);
                isValid = false;
                return;
            }
            finally
            {
                tempProcessLog += ln + "Took " + String.Format("{0:0.00}", (DateTime.Now - begin).TotalSeconds) + " seconds";
                App.Current.Dispatcher.Invoke(new System.Action(() =>
                {
                    tempProcessLog += ln + "There are total " + invalidRows + " errors in all 3 files";
                    txtOutput.AppendText(tempProcessLog);
                    txtOutput.ScrollToEnd();
                    progress.Value = 100;
                }));

                excelApp.Workbooks.Close();
            }
        }

        /// <summary>
        /// Returns grade for given marks and grade ranges
        /// </summary>
        /// <param name="external_mark"> External Mark </param>
        /// <param name="final_mark"> Internal Mark </param>
        /// <param name="AP"></param>
        /// <param name="A"></param>
        /// <param name="B"></param>
        /// <param name="C"></param>
        /// <param name="D"></param>
        /// <param name="E"></param>
        /// <param name="F"></param>
        /// <param name="subject_type"> Theory or Lab </param>
        /// <returns>Grade as string</returns>
        String calculateGrade(float external_mark,float final_mark, float AP, float A, float B, float C, float D, float E, float F,String subject_type)
        {
            String grade = "I";

            if (external_mark < 0) //INCOMPLETE EXAM [I]
                return "I";

            if (String.Compare(subject_type, "Theory", true) == 0)
                if (external_mark < E)
                    return "F";
            else
                if (external_mark < 18)
                    return "F";

            if (final_mark >= AP) grade = "A+";
            if (final_mark >= A && final_mark < AP) grade = "A";
            if (final_mark >= B && final_mark < A) grade = "B";
            if (final_mark >= C && final_mark < B) grade = "C";
            if (final_mark >= D && final_mark < C) grade = "D";
            if (final_mark >= E && final_mark < D) grade = "E";
            if (final_mark >= F && final_mark < E) grade = "F";

            return grade;
        }

        /// <summary>
        /// generates point for a given grade value
        /// </summary>
        /// <param name="grade">Grade</param>
        /// <returns>point as int</returns>
        int generatePoints(String grade)
        {
            int points = 0;

            if (String.Compare("A+", grade, true) == 0) points = 10;
            if (String.Compare("A", grade, true) == 0) points = 9;
            if (String.Compare("B", grade, true) == 0) points = 8;
            if (String.Compare("C", grade, true) == 0) points = 7;
            if (String.Compare("D", grade, true) == 0) points = 6;
            if (String.Compare("E", grade, true) == 0) points = 5;
            if (String.Compare("F", grade, true) == 0) points = 0;
            if (String.Compare("I", grade, true) == 0) points = 0;

            //for rest grades points is 0
            return points;
        }

    }
}
