using System;
using System.Collections.Generic;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;

namespace markevaluator
{
    class Courses : IParse
    {
        private string fileName;

        //For updating the validation output 
        //and progressbar of parser_window.
        System.Windows.Controls.TextBox txtOutput;
        System.Windows.Controls.ProgressBar progress;

        private BackgroundWorker bgwValidate;
        private BackgroundWorker bgwPush;

        Application excelApp;
        private string tempProcessLog = "";
        private bool isValid = true;
        int invalidRows = 0;
        int recordsPushed = 0;
        string ln = Environment.NewLine;

        public Courses()
        {
            //Do something here
        }

        private void bgwValidate_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //update progressbar
            progress.Value = e.ProgressPercentage;
        }

        private void bgwValidate_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //All ended
            invalidRows = 0;
            if (isValid)
                Windows.parserWindow.ValidationCompleted("Your file is validated. There are no error(s)!" + Environment.NewLine + "Push data to database?", true);
            else
                Windows.parserWindow.ValidationCompleted("Your worksheet file contains some errors, please see the logs." + Environment.NewLine + "Please correct the errors.", false);
        }

        private void bgwValidate_DoWork(object sender, DoWorkEventArgs e)
        {
            //Do the validation here
            DateTime begin = DateTime.Now;
            try
            {
                const int MAX_SEMS = 8;
                int sheetCount = 0;
                int rowCount = 0, columnCount = 0, i;
                int isem = 0, tsem = 0;
                string[] courseCodes;
                isValid = true;

                excelApp = new Application();
                Workbook courseWorkbook = excelApp.Workbooks.Open(fileName, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet courseWorksheet = (Worksheet)courseWorkbook.Sheets.get_Item(1);
                Range courseWorksheetRange = courseWorksheet.UsedRange;

                sheetCount = courseWorkbook.Sheets.Count;
                tempProcessLog = ln + "Reading worksheet..." + ln + "Total sheets : " + sheetCount;

                tempProcessLog += ln + "Using Sheet1... ";

                rowCount = courseWorksheetRange.Rows.Count;
                columnCount = courseWorksheetRange.Columns.Count;
                courseCodes = new string[rowCount];

                tempProcessLog += ln + "Rows " + rowCount + " | Columns " + columnCount + ln;

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

                if (columnCount > 4 || columnCount < 4)
                {
                    tempProcessLog += ln + "Expecting columns to be 4, " + columnCount + " is present." + ln + "This excelsheet is not meeting the desired format.";

                    App.Current.Dispatcher.Invoke(new System.Action(() =>
                    {
                        txtOutput.AppendText(tempProcessLog);
                        txtOutput.ScrollToEnd();
                        progress.Value = 100;
                    }));

                    isValid = false;
                    return;
                }

                tempProcessLog += ln + "Expecting first row to be column names. " + ln + "reading from 2nd row...";

                List<String> tmpCourseCodes = new List<string>();

                for (i = 2; i <= rowCount; i++)
                {
                    try
                    {
                        tempProcessLog += ln + "> Processing Row: " + i + ln;

                        if (courseWorksheet.Cells[i, 1].Value != null)
                        {
                            //Check for duplicate course_code in database
                            if (Medatabase.isPresentLike("course_master", "course_code", (string)courseWorksheet.Cells[i, 1].Value))
                            {
                                tempProcessLog += ln + "duplicate course_code '" + (string)courseWorksheet.Cells[i, 1].Value + "' already present in database...";
                                isValid = false;
                                invalidRows++;
                            }

                            //Check for duplicate course_code in the same excel sheet
                            if (tmpCourseCodes.IndexOf((string)courseWorksheet.Cells[i, 1].Value) != -1)
                            {
                                tempProcessLog += ln + "duplicate course_code '" + (string)courseWorksheet.Cells[i, 1].Value + "' already present in the same excel sheet...";
                                isValid = false;
                                invalidRows++;
                            }
                            else //if not duplicate
                                tmpCourseCodes.Add((string)courseWorksheet.Cells[i, 1].Value);
                        }
                        else
                        {
                            tempProcessLog += ln + "empty course_code given...'" + ln;
                            isValid = false;
                            invalidRows++;
                        }

                        if (courseWorksheet.Cells[i, 2].Value != null)
                        {
                            tsem = Convert.ToInt16(courseWorksheet.Cells[i, 2].Value);
                            if (tsem > MAX_SEMS || tsem < 0)
                            {
                                tempProcessLog += ln + (string)courseWorksheet.Cells[i, 1].Value + " > total semesters must be between 0 and " + MAX_SEMS + ", " + tsem + " given";
                                isValid = false;
                                invalidRows++;
                            }
                        }
                        else
                        {
                            tempProcessLog += ln + "empty value given for total semesters...'" + ln;
                            isValid = false;
                            invalidRows++;
                        }

                        if (courseWorksheet.Cells[i, 3].Value != null)
                        {
                            isem = Convert.ToInt16(courseWorksheet.Cells[i, 3].Value);
                            if (isem < 1)
                            {
                                tempProcessLog += ln + (string)courseWorksheet.Cells[i, 1].Value + " > in semesters must be between 0 and " + MAX_SEMS + ", " + isem + " given";
                                isValid = false;
                                invalidRows++;
                            }

                            if (isem > tsem)
                            {
                                tempProcessLog += ln + (string)courseWorksheet.Cells[i, 1].Value + " > in semesters can't be > than total sems, " + isem + " given";
                                isValid = false;
                                invalidRows++;
                            }

                            if (isem > MAX_SEMS)
                            {
                                tempProcessLog += ln + (string)courseWorksheet.Cells[i, 1].Value + " > max in semesters possible is " + MAX_SEMS + ", " + isem + " given";
                                isValid = false;
                                invalidRows++;
                            }
                        }
                        else
                        {
                            tempProcessLog += ln + "empty value given for in semesters...'" + ln;
                            isValid = false;
                            invalidRows++;
                        }
                        if (courseWorksheet.Cells[i, 4].Value != null)
                        {
                            if((int)courseWorksheet.Cells[i, 4].Value < 1)
                            {
                                tempProcessLog += ln + "credits must be greater than 0";
                                isValid = false;
                                invalidRows++;
                            }
                        }
                        else
                        {
                            tempProcessLog += ln + "empty value given for total credits...'" + ln;
                            isValid = false;
                            invalidRows++;
                        }
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException be)
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

                    bgwValidate.ReportProgress(Convert.ToInt16(((i * 100) / rowCount)));
                    tempProcessLog = "";
                }
            }
            catch (System.IO.IOException ioe)
            {
                tempProcessLog += ln + "ERROR HANDLING FILE :(" + ln + "Please retry";
                LogWriter.WriteError("During course excel sheet validation", ioe.Message);
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
                LogWriter.WriteError("During course excel sheet validation", ex.Message);
                isValid = false;
                return;
            }
            finally
            {
                if (invalidRows > 0)
                {
                    tempProcessLog += ln + "There are total " + invalidRows + " errors.";
                    tempProcessLog += ln + "Took " + String.Format("{0:0.00}", (DateTime.Now - begin).TotalSeconds) + " seconds";
                    App.Current.Dispatcher.Invoke(new System.Action(() =>
                    {
                        txtOutput.AppendText(tempProcessLog);
                        txtOutput.ScrollToEnd();
                        progress.Value = 100;
                    }));
                }
                excelApp.Workbooks.Close();
            }
        }

        public void ValidateWorksheet(string fileName)
        {
            this.fileName = fileName;
            txtOutput = Windows.parserWindow.txtOut;
            progress = Windows.parserWindow.prgsbar;

            //Initialize background worker
            bgwValidate = new BackgroundWorker();
            bgwValidate.DoWork += bgwValidate_DoWork;
            bgwValidate.ProgressChanged += bgwValidate_ProgressChanged;
            bgwValidate.RunWorkerCompleted += bgwValidate_RunWorkerCompleted;

            //call bgWorker to perform validation
            bgwValidate.RunWorkerAsync();
            bgwValidate.WorkerReportsProgress = true;
            bgwValidate.WorkerSupportsCancellation = true;
        }

        /// <summary>
        /// This method will return the total no. of sems in a particular course
        /// </summary>
        /// <param name="course_code">course for which sem count is requested</param>
        /// <returns>Total sems count</returns>
        public static int getSemCount(string course_code)
        {
            int scount = 0;
            List<Row> rows = Medatabase.fetchRecords("SELECT total_semesters FROM course_master WHERE course_code='" + course_code + "'");
            foreach (Row row in rows)
                scount = (int)row.column["total_semesters"];
            return scount;
        }

        /// <summary>
        /// This method will return the no. of sems(in) in a particular course
        /// </summary>
        /// <param name="course_code">course for which sem count is requested</param>
        /// <returns>Total sems count</returns>
        public static int getInSemCount(string course_code)
        {
            int scount = 0;
            List<Row> rows = Medatabase.fetchRecords("SELECT in_semesters FROM course_master WHERE course_code='" + course_code + "'");
            foreach (Row row in rows)
                scount = (int)row.column["in_semesters"];
            return scount;
        }

        /// <summary>
        /// Generates a list of semseter numbers for given course_code
        /// </summary>
        /// <param name="course_code">course code</param>
        /// <returns>integer collection</returns>
        public static List<int> getInSemList(String course_code)
        {
            List<int> items = new List<int>();
            int in_sems = getInSemCount(course_code);

            if (in_sems > 0)
                for (int i = 1; i <= in_sems; i++)
                    items.Add(i);
            return items;
        }

        /// <summary>
        /// Returns the number of subject presents in a given course
        /// </summary>
        /// <param name="course_code">course code</param>
        /// <returns>integer</returns>
        public static int getSubjectCount(String course_code)
        {
            int count = 0;
            List<Row> rows = Medatabase.fetchRecords("SELECT COUNT(sub_code) AS scount FROM subject_master WHERE course_code='" + course_code + "'");
            if (rows.Count > 0)
                count = Convert.ToInt16(rows[0].column["scount"]);
            return count;
        }

        /// <summary>
        /// Call this method after validating excel file to push info. to database
        /// </summary>
        /// <param name="fileName">Path of the excel file</param>
        /// <param name="pwindow">the parser window for updating progress info</param>
        public void PushToDatabase(string fileName)
        {
            this.fileName = fileName;

            //Initialize background worker
            bgwPush = new BackgroundWorker();
            bgwPush.DoWork += bgwPush_DoWork;
            bgwPush.ProgressChanged += bgwPush_ProgressChanged;
            bgwPush.RunWorkerCompleted += bgwPush_RunWorkerCompleted;

            //call bgWorker to perform database push operation
            bgwPush.RunWorkerAsync();
            bgwPush.WorkerReportsProgress = true;
            bgwPush.WorkerSupportsCancellation = true;
        }

        private void bgwPush_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (!isValid)
                Windows.parserWindow.PushToDatabaseCompleted("There were some errors while pushing to database." + ln + "Please see the logs" + ln + "Records pushed : " + recordsPushed, false);
            else
                Windows.parserWindow.PushToDatabaseCompleted("All records successfully pushed to database." + ln + "Records pushed : " + recordsPushed, true);
        }

        private void bgwPush_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //update progressbar
            progress.Value = e.ProgressPercentage;
        }

        private void bgwPush_DoWork(object sender, DoWorkEventArgs e)
        {
            DateTime begin = DateTime.Now;
            try
            {
                int sheetCount = 0;
                int rowCount = 0, columnCount = 0, i;
                int iSem = 0, tSem = 0, tCred = 0;
                string cCode;
                isValid = true;

                excelApp = new Application();
                Workbook courseWorkbook = excelApp.Workbooks.Open(fileName, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet courseWorksheet = (Worksheet)courseWorkbook.Sheets.get_Item(1);
                Range courseWorksheetRange = courseWorksheet.UsedRange;

                sheetCount = courseWorkbook.Sheets.Count;
                tempProcessLog = "Database push operation...";
                tempProcessLog += ln + "Reading worksheet..." + ln + "Total sheets : " + sheetCount;

                tempProcessLog += ln + "Using Sheet1... ";

                rowCount = courseWorksheetRange.Rows.Count;
                columnCount = courseWorksheetRange.Columns.Count;

                for (i = 2; i <= rowCount; i++)
                {
                    cCode = (string)courseWorksheet.Cells[i, 1].Value;
                    tSem = (int)courseWorksheet.Cells[i, 2].Value;
                    iSem = (int)courseWorksheet.Cells[i, 3].Value;
                    tCred = (int)courseWorksheet.Cells[i, 4].Value;

                    tempProcessLog += ln + "> Processing Row: " + i + ln;
                    tempProcessLog += ln + "PUSHING > " + cCode + " | " + " | ts " + tSem + " | is " + iSem;
                    recordsPushed += Medatabase.ExecuteQuery("INSERT INTO course_master VALUES('" + cCode + "'," + tSem + "," + iSem + "," + tCred + ")");

                    //Update output textbox
                    App.Current.Dispatcher.Invoke(new System.Action(() =>
                    {
                        txtOutput.AppendText(tempProcessLog);
                        txtOutput.ScrollToEnd();
                    }));
                    
                    bgwPush.ReportProgress(Convert.ToInt16(((i * 100) / rowCount)));
                    tempProcessLog = "";
                }
            }
            catch (databaseException)
            {
                tempProcessLog += ln + "ERROR OCCURED IN DATABASE OPERATION :(" + ln + "Please retry.";
                isValid = false;
            }
            catch (System.IO.IOException ioe)
            {
                tempProcessLog += ln + "ERROR HANDLING FILE :(" + ln + "Please retry";
                LogWriter.WriteError("During upload of course excel sheet file", ioe.Message);
                isValid = false;
            }
            catch (Exception ex)
            {
                tempProcessLog += ln + "UNEXPECTED ERROR OCCURED :(" + ln + ex.Message + ln + "Please retry.";
                LogWriter.WriteError("During upload of course excel sheet file", ex.Message);
                isValid = false;
            }
            finally
            {
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

        /// <summary>
        /// For filling course data in listview
        /// </summary>
        /// <returns>CourseCols collection</returns>
        public static List<CourseCols> getCourseList()
        {
            List<CourseCols> items = null;
            try
            {
                List<Row> rows = Medatabase.fetchRecords("SELECT * FROM course_master");
                items = new List<CourseCols>();

                foreach (Row row in rows)
                    items.Add(new CourseCols
                    {
                        c_code = (string)row.column["course_code"],
                        t_semesters = (int)row.column["total_semesters"],
                        in_semesters = (int)row.column["in_semesters"]
                    });
            }
            catch (databaseException)
            {
                //Do Something
            }
            catch (Exception ex)
            {
                //Do Something
                LogWriter.WriteError("Fetching course list from database", ex.Message);
            }
            return items;
        }

        /// <summary>
        /// fetches all course_codes from database if regular student
        /// else it fetches all course_codes only for which exam has held
        /// </summary>
        /// <param name="stype">Regular or non-regular</param>
        /// <returns>course_code string collection</returns>
        public static List<string> getCourseCodes(StudentType stype)
        {
            List<String> ccodes = new List<String>();
            try
            {
                List<Row> rows;
                if (stype == StudentType.REGULAR)
                    rows = Medatabase.fetchRecords("SELECT course_code FROM course_master");
                else
                    rows = Medatabase.fetchRecords("SELECT DISTINCT course_code FROM exam_master");
                foreach (Row row in rows)
                    ccodes.Add((string)row.column["course_code"]);
            }
            catch (databaseException)
            {
                //Do something
            }
            catch (Exception ex)
            {
                //Do something
                LogWriter.WriteError("Fetching course code list from database", ex.Message);
            }
            return ccodes;
        }

    }

    class CourseCols
    {
        public string c_code { get; set; }
        public int t_semesters { get; set; }
        public int in_semesters { get; set; }
    }

}
