using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.ComponentModel;

namespace markevaluator
{
    class Subjects : IParse
    {
        private string fileName;

        //For updating the validation output 
        //and progressbar of parser_window.
        System.Windows.Controls.TextBox txtOutput;
        System.Windows.Controls.ProgressBar progress;

        BackgroundWorker bgwValidate;
        BackgroundWorker bgwPush;

        Application excelApp;
        private string tempProcessLog = "";
        private bool isValid = true;
        int invalidRows = 0;
        int recordsPushed = 0;
        string ln = Environment.NewLine;

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
                Windows.parserWindow.ValidationCompleted("Your file is validated." + Environment.NewLine + "Push data to database?", true);
            else
                Windows.parserWindow.ValidationCompleted("Your worksheet file contains some errors, please see the logs." + Environment.NewLine + "Please correct the errors.", false);
        }

        private void bgwValidate_DoWork(object sender, DoWorkEventArgs e)
        {
            //Do the validation here
            DateTime begin = DateTime.Now;
            try
            {
                excelApp = new Application();
                List<String> scodes = new List<string>();
                int rowCount = 0, columnCount = 0, i,j = 1;

                Workbook courseWorkbook = excelApp.Workbooks.Open(fileName, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet courseWorksheet;
                Range courseWorksheetRange;
                isValid = true;

                tempProcessLog = ln + "Reading worksheet..." + ln + "Total sheets : " + courseWorkbook.Sheets.Count;

                foreach (Worksheet cws in courseWorkbook.Sheets)
                {
                    if(!Medatabase.isPresentLike("course_master","course_code",cws.Name))
                    {
                        tempProcessLog += ln + "sheet with name \"" + cws.Name + "\" course code is not present in database";
                        isValid = false;
                        invalidRows++;
                    }
                    else
                    {
                        tempProcessLog += ln + "======== reading sheet " + cws.Name + "========" + ln;
                        courseWorksheet = cws;
                        courseWorksheetRange = courseWorksheet.UsedRange;

                        rowCount = courseWorksheetRange.Rows.Count;
                        columnCount = courseWorksheetRange.Columns.Count;

                        tempProcessLog += ln + "Rows " + rowCount + " | Columns " + columnCount + ln;

                        if (rowCount <= 1)
                        {
                            tempProcessLog += ln + "in sheet \"" + cws.Name + "\" There are no records to validate...";
                            isValid = false;
                            continue; //Move to next sheet
                        }

                        if (columnCount > 5 || columnCount < 5)
                        {
                            tempProcessLog += ln + "in sheet \"" + cws.Name + "\" Expecting columns to be 5, " + columnCount + " is present." + ln + "This excelsheet is not meeting the desired format.";
                            isValid = false;
                            continue; //Move to next sheet
                        }

                        tempProcessLog += ln + "Expecting first row to be column names. " + ln + "reading from 2nd row...";

                        for (i = 2; i <= rowCount; i++)
                        {
                            try
                            {
                                tempProcessLog += ln + ">> Sheet [" + cws.Name + "] > Processing Row: " + i + ln;

                                /** CANT CHECK FOR DUPLICATE AS ELECTIVE LAB CAN HAVE DUPLICATE SUBJECT CODE :( **/

                                //Check for duplicate subject_code in database
                                if (Medatabase.isPresent("subject_master", "sub_code", (string)courseWorksheet.Cells[i, 1].Value))
                                {
                                    tempProcessLog += ln + "duplicate subject_code '" + (string)courseWorksheet.Cells[i, 1].Value + "' already present in database in row " + i;
                                    isValid = false;
                                    invalidRows++;
                                }
                                
                                //if not a lab subject
                                if (String.Compare((string)courseWorksheet.Cells[i, 5].Value, "Lab", true) != 0)
                                {
                                    //Check for duplicate subject_code in the same excel sheet
                                    if (scodes.IndexOf((string)courseWorksheet.Cells[i, 1].Value) != -1)
                                    {
                                        tempProcessLog += ln + "duplicate subject_code '" + (string)courseWorksheet.Cells[i, 1].Value + "' already present in the same excel sheet in row " + i;
                                        isValid = false;
                                        invalidRows++;
                                    }
                                    scodes.Add((string)courseWorksheet.Cells[i, 1].Value);
                                }

                                if (courseWorksheet.Cells[i, 1].Value == null)
                                {
                                    tempProcessLog += ln + "Empty value given for subject_code in row " + i;
                                    isValid = false;
                                    invalidRows++;
                                }

                                if (courseWorksheet.Cells[i, 2].Value == null)
                                {
                                    tempProcessLog += ln + "Empty value given for subject_name in row " + i;
                                    isValid = false;
                                    invalidRows++;
                                }

                                if ((int)courseWorksheet.Cells[i, 3].Value <= 0 || (int)courseWorksheet.Cells[i, 3].Value > Courses.getSemCount(cws.Name))
                                {
                                    tempProcessLog += ln + "Invalid semester number \"" + (int)courseWorksheet.Cells[i, 3].Value + "\" in row " + i;
                                    isValid = false;
                                    invalidRows++;
                                }

                                if ((int)courseWorksheet.Cells[i, 4].Value <= 0)
                                {
                                    tempProcessLog += ln + "Subject credits cannot be less than or equal to 0 in row " + i;
                                    isValid = false;
                                    invalidRows++;
                                }

                                if (courseWorksheet.Cells[i, 5].Value == null)
                                {
                                    tempProcessLog += ln + "Empty value given for subject type in row " + i;
                                    isValid = false;
                                    invalidRows++;
                                }
                            }
                            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException be)
                            {
                                invalidRows++;
                                isValid = false;
                                tempProcessLog += ln + "In row [" + i + "] expecting int value, string given!";
                                LogWriter.WriteError("While validating subject details worksheet", be.Message);
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
                    bgwValidate.ReportProgress(Convert.ToInt16(((j * 100) / courseWorkbook.Sheets.Count)));
                    j++;
                }
            }
            catch (System.IO.IOException ioe)
            {
                tempProcessLog += ln + "ERROR HANDLING FILE :(" + ln + "Please retry";
                LogWriter.WriteError("While validating subject details worksheet", ioe.Message);
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
                LogWriter.WriteError("While validating subject details worksheet", ex.Message);
                isValid = false;
                return;
            }
            finally
            {
                if (invalidRows > 0)
                {
                    tempProcessLog += ln + ln + "There are total " + invalidRows + " errors.";
                }

                //Update output textbox
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
                int credits = 0, sem = 0;
                string sCode, sName,sType;
                isValid = true;

                excelApp = new Application();
                Workbook courseWorkbook = excelApp.Workbooks.Open(fileName, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Range courseWorksheetRange;

                sheetCount = courseWorkbook.Sheets.Count;
                tempProcessLog = "Database push operation...";
                tempProcessLog += ln + "Reading worksheet..." + ln + "Total sheets : " + sheetCount;

                foreach (Worksheet cws in courseWorkbook.Sheets)
                {
                    tempProcessLog += ln + ">>> Using Sheet \"" + cws.Name + "\"";

                    courseWorksheetRange = cws.UsedRange;
                    rowCount = courseWorksheetRange.Rows.Count;
                    columnCount = courseWorksheetRange.Columns.Count;

                    for (i = 2; i <= rowCount; i++)
                    {
                        sCode = (string)cws.Cells[i, 1].Value;
                        sName = (string)cws.Cells[i, 2].Value;
                        sem = (int)cws.Cells[i, 3].Value;
                        credits = (int)cws.Cells[i, 4].Value;
                        sType = (string)cws.Cells[i, 5].Value;

                        tempProcessLog += ln + "> Processing Row: " + i + ln;
                        tempProcessLog += ln + "PUSHING > " + sCode + " | " + sName + " | ts " + sem + " | is " + credits;
                        recordsPushed += Medatabase.ExecuteQuery("INSERT INTO subject_master VALUES('" + cws.Name + "'," + sem + ",'" + sCode + "','" + sName + "'," + credits + ",'" + sType + "')");

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
                LogWriter.WriteError("While pushing subject details to database", ioe.Message);
                isValid = false;
                return;
            }
            catch (Exception ex)
            {
                tempProcessLog += ln + "UNEXPECTED ERROR OCCURED :(" + ln + ex.Message + ln + "Please retry.";
                LogWriter.WriteError("While pushing subject details to database", ex.Message);
                isValid = false;
                return;
            }
            finally
            {
                //Update output textbox
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
        /// Call this function to get subject list
        /// </summary>
        /// <param name="course_code">course code</param>
        /// <returns>SubjectCols collection</returns>
        public static List<SubjectCols> getSubjectList(string course_code)
        {
            List<SubjectCols> items = null;
            try
            {
                List<Row> rows = null;
                items = new List<SubjectCols>(); ;

                for (int t = 1; t <= Courses.getSemCount(course_code); t++)
                {
                    rows = Medatabase.fetchRecords("SELECT * FROM subject_master WHERE course_code='" + course_code + "' AND semester=" + t);
                    foreach(Row row in rows)
                        items.Add(new SubjectCols {
                            subject_code = (string)row.column["sub_code"],
                            subject_name = (string)row.column["name"],
                            semNo = (semNo)row.column["semester"],
                            credits = (int)row.column["credits"]
                        });
                }
            }
            catch (databaseException)
            {
                //do something
            }
            catch (Exception ex)
            {
                //do something
                LogWriter.WriteError("While fetching subject list from database", ex.Message);
            }
            return items;
        }
    }

    public enum semNo
    {
        Semester1 = 1,
        Semester2,
        Semester3,
        Semester4,
        Semester5,
        Semester6,
        Semester7,
        Semester8
    }

    class SubjectCols
    {
        public string subject_code { get; set; }
        public string subject_name { get; set; }
        public semNo semNo { get; set; }
        public int credits { get; set; } 
    }
}
