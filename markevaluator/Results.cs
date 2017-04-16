using System;
using System.Collections.Generic;
using System.ComponentModel;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace markevaluator
{
    class Results
    {
        BackgroundWorker bgwGenResultSheet;
        BackgroundWorker bgwGenAnalysisSheet;
        BackgroundWorker bgwGenGradeSheet;
        BackgroundWorker bgwFetchRegid;

        //Variables required till sheet generation
        String course_code;
        String degree_type;
        int semester, year;
        String month;
        bool error_occured;
        StudentType stype;
        String tempProcessLog;
        String ln = Environment.NewLine;
        String output_folder;
        Excel.Application excelApp;
        List<long> regIdList;
        String fileName;
        bool isValid;
        int invalidRows;
        int incorrect_reg_ids;

        //for updating output
        System.Windows.Controls.TextBox txtOutput = (Windows.generatorWindow != null) ? Windows.generatorWindow.txtOutput : null;
        System.Windows.Controls.ProgressBar progress_one = (Windows.generatorWindow != null) ? Windows.generatorWindow.progressOne : null;
        System.Windows.Controls.ProgressBar progress_two = (Windows.generatorWindow != null) ? Windows.generatorWindow.progressTwo : null;

        /// <summary>
        /// Returns a list semesters of a particular course from results table
        /// </summary>
        /// <param name="course_code">Course code</param>
        /// <returns>int collection</returns>
        public static List<int> getSemesterList(String course_code)
        {
            List<int> items = new List<int>();
            try
            {
                List<Row> rows = Medatabase.fetchRecords("SELECT DISTINCT semester FROM exam_master WHERE course_code='" + course_code + "'");
                foreach (Row row in rows)
                    items.Add((int)row.column["semester"]);
            }
            catch(Exception ex)
            {
                LogWriter.WriteError("Fetching result sem list of a course", ex.Message);
            }
            return items;
        }

        /// <summary>
        /// Returns a list of years in which exams held of a given course and semester
        /// </summary>
        /// <param name="course_code">Course code</param>
        /// <param name="semester">Semester</param>
        /// <returns>int collection</returns>
        public static List<int> getYearList(String course_code,int semester)
        {
            List<int> items = new List<int>();
            try
            {
                List<Row> rows = Medatabase.fetchRecords("SELECT DISTINCT year FROM exam_master WHERE course_code='" + course_code + "' AND semester=" + semester);
                foreach (Row row in rows)
                    items.Add((int)row.column["year"]);
            }
            catch (Exception ex)
            {
                LogWriter.WriteError("Fetching result year list of a course and sem", ex.Message);
            }
            return items;
        }

        /// <summary>
        /// Returns a list of months in which semester exam held
        /// </summary>
        /// <param name="course_code">Course code</param>
        /// <param name="semester">Semester</param>
        /// <param name="year">year</param>
        /// <returns>String collection</returns>
        public static List<String> getMonthList(String course_code,int semester,int year)
        {
            List<String> items = new List<String>();
            try
            {
                List<Row> rows = Medatabase.fetchRecords("SELECT DISTINCT month FROM exam_master WHERE course_code='" + course_code + "' AND semester=" + semester + " AND year=" + year);
                foreach (Row row in rows)
                    items.Add((String)row.column["month"]);
            }
            catch (Exception ex)
            {
                LogWriter.WriteError("Fetching result month list of a course and sem", ex.Message);
            }
            return items;
        }

        /// <summary>
        /// Returns a list registration id(students) for a given exam details
        /// </summary>
        /// <param name="course_code">course code</param>
        /// <param name="semester">semester number</param>
        /// <param name="year">year in which exam held</param>
        /// <returns></returns>
        public static List<long> getRegidList(String course_code,int semester,int year, String month)
        {
            List<long> items = new List<long>();
            try
            {
                List<Row> rows = Medatabase.fetchRecords("SELECT registration_id FROM exam_master WHERE course_code='" + course_code + "' AND semester=" + semester + " AND year=" + year + " AND month='" + month + "'");
                foreach (Row row in rows)
                    items.Add((long)row.column["registration_id"]);
            }
            catch (Exception ex)
            {
                LogWriter.WriteError("Fetching result year list of a course and sem", ex.Message);
            }
            return items;
        }

        /// <summary>
        /// Returns students exam result
        /// </summary>
        /// <param name="course_code">course code</param>
        /// <param name="semester">semester number</param>
        /// <param name="year">exam year</param>
        /// <param name="registration_id">students registration id</param>
        /// <returns>collection</returns>
        public static List<ResultCols> getExamResult(String course_code,int semester, int year, long registration_id,String month)
        {
            List<ResultCols> items = new List<ResultCols>();
            try
            {
                List<Row> rows = Medatabase.fetchRecords("SELECT * FROM student_marks WHERE exam_id IN (" + "SELECT exam_id FROM exam_master WHERE course_code='" + course_code + "' AND semester=" + semester + " AND year=" + year + " AND registration_id=" + registration_id + " AND month='"+ month +"')");
                foreach (Row row in rows)
                    items.Add(new ResultCols()
                    {
                        mrk_s_code = (String)row.column["sub_code"],
                        mrk_s_grade = (String)row.column["grade"],
                        mrk_s_iamark = (float)row.column["internal_marks"],
                        mrk_s_eamark = (float)row.column["external_marks"],
                        mrk_s_fmark = (float)row.column["final_marks"],
                    });
            }
            catch(Exception ex)
            {
                LogWriter.WriteError("Fetching exam result of student", ex.Message);
            }
            return items;
        }

        /// <summary>
        /// returns student's gpa for a given exam details
        /// </summary>
        /// <param name="course_code"> course_code </param>
        /// <param name="semester">semester number</param>
        /// <param name="year">exam year</param>
        /// <param name="registration_id">students registration id</param>
        /// <returns></returns>
        public static float getStudentGpa(String course_code, int semester, int year, long registration_id)
        {
            float gpa = 0.0f;
            try
            {
                List<Row> rows = Medatabase.fetchRecords("SELECT gpa FROM student_cgpa WHERE exam_id IN (SELECT exam_id FROM exam_master WHERE course_code='"+ course_code + "' AND semester=" + semester + " AND year=" + year + " AND registration_id=" + registration_id + ")");
                gpa = (float)rows[0].column["gpa"];
            }
            catch(Exception ex)
            {
                LogWriter.WriteError("While fetching gpa of a student", ex.Message);
            }
            return gpa;
        }

        /// <summary>
        /// Creates an excel sheet containing grades and gpa of students (Result sheet)
        /// </summary>
        /// <param name="course_code">Course code</param>
        /// <param name="semester">Semester</param>
        /// <param name="year">exam year</param>
        /// <param name="output_folder">output folder to save excel file</param>
        public void generateResultSheet(String course_code,int semester,int year, String month, String output_folder, List<long> regIdList = null)
        {
            this.course_code = course_code;
            this.semester = semester;
            this.year = year;
            this.month = month;
            this.output_folder = output_folder;
            this.regIdList = regIdList; //for custom list of students
            
            //initialize background worker
            bgwGenResultSheet = new BackgroundWorker();
            bgwGenResultSheet.DoWork += bgwGenResultSheet_DoWork;
            bgwGenResultSheet.RunWorkerCompleted += bgwGenResultSheet_RunWorkerCompleted;
            bgwGenResultSheet.ProgressChanged += bgwGenResultSheet_ProgressChanged;

            bgwGenResultSheet.RunWorkerAsync();
            bgwGenResultSheet.WorkerReportsProgress = true;
            bgwGenResultSheet.WorkerSupportsCancellation = true;
        }

        /// <summary>
        /// Creates individual grade sheets for each student of a semester
        /// </summary>
        /// <param name="course_code">Course code</param>
        /// <param name="semester">Semester</param>
        /// <param name="year">exam year</param>
        /// <param name="output_folder">output folder to save individual marksheet files</param>
        public void generateGradeSheets(String course_code, int semester, int year,String month, String output_folder, String degree_type, List<long> regIdList = null)
        {
            this.course_code = course_code;
            this.degree_type = degree_type;
            this.semester = semester;
            this.year = year;
            this.month = month;
            this.output_folder = output_folder;
            this.regIdList = regIdList; //for custom list of students

            //initialize background worker
            bgwGenGradeSheet = new BackgroundWorker();
            bgwGenGradeSheet.DoWork += bgwGenGradeSheet_DoWork;
            bgwGenGradeSheet.RunWorkerCompleted += bgwGenGradeSheet_RunWorkerCompleted;
            bgwGenGradeSheet.ProgressChanged += bgwGenGradeSheet_ProgressChanged;

            bgwGenGradeSheet.RunWorkerAsync();
            bgwGenGradeSheet.WorkerReportsProgress = true;
            bgwGenGradeSheet.WorkerSupportsCancellation = true;
        }

        private void bgwGenGradeSheet_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //for one progress bar
            progress_one.Value = e.ProgressPercentage;
        }

        private void bgwGenGradeSheet_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //task completed
            if (error_occured)
                Windows.generatorWindow.exportCompleted("An error occured while generating individual grade sheets!" + ln + "Please see the logs", true);
            else
                Windows.generatorWindow.exportCompleted("Individual grade sheets has been generated!" + ln + ln + "Please find the file in" + ln + output_folder, false);
        }

        private void bgwGenGradeSheet_DoWork(object sender, DoWorkEventArgs e)
        {
            DateTime begin = DateTime.Now;
            try
            {
                //TEMP WORD VARIABLES
                object oMissing = System.Reflection.Missing.Value;
                object oEndOfDoc = "\\endofdoc";
                object output_filename;

                Word._Application objWord;
                Word._Document objDoc;
                Word.Paragraph objPara0, objPara1, objPara2, objPara3, objPara4, objPara5, objPara6;
                Word.Range pRng0, pRng1, pRng2, pRng3, pRng4, pRng5, pRng6;
                Word.Table objTable;
                Word.Range wrdRng;

                DateTime dt; String str;
                int xrow = 0, trows, prog_one_value = 1, prog_two_value = 1;

                
                List<Row> rows3, rows2, rows = new List<Row>();

                if (regIdList == null)
                    rows = Medatabase.fetchRecords("SELECT * FROM exam_master WHERE course_code='" + course_code + "' AND semester=" + semester + " AND year=" + year + " AND month='" + month + "'");
                else // for custom student list
                    foreach(long rid in regIdList)
                        foreach (Row row in Medatabase.fetchRecords("SELECT * FROM exam_master WHERE course_code='" + course_code + "' AND semester=" + semester + " AND year=" + year + " AND month='" + month + "'" + " AND registration_id=" + rid))
                            rows.Add(row);

                List<String> sub_codes;

                //strip end slash if present
                output_folder = (output_folder[output_folder.Length - 1] == '\\') ? output_folder.Substring(0, output_folder.Length - 1) : output_folder;

                //create if "Results" directory exists
                if (!Directory.Exists(output_folder + "\\Results"))
                    Directory.CreateDirectory(output_folder + "\\Results");

                output_folder = output_folder + "\\Results";

                //create directory with course_code inside Results folder
                if (!Directory.Exists(output_folder + "\\" + course_code))
                    Directory.CreateDirectory(output_folder + "\\" + course_code);

                tempProcessLog = "===== GENERATING INDIVIDUAL GRADE SHEETS =====" + ln;

                //cycle through each registration id
                foreach(Row row in rows)
                {
                    prog_two_value = 1;
                    objWord = new Word.Application();
                    objWord.Visible = false; //word program wont be visible
                    objDoc = objWord.Documents.Add(ref oMissing, ref oMissing,ref oMissing, ref oMissing);

                    objDoc.Content.Text += ln + ln;

                    dt = DateTime.Now;
                    str = dt.ToString("dd MMMMMMMMM yyyy");

                    tempProcessLog += ln + "Processing > " + (long)row.column["registration_id"];
                    //Paragraph oject one
                    objPara0 = objDoc.Content.Paragraphs.Add(ref oMissing);
                    pRng0 = objPara0.Range;
                    pRng0.Font.Size = 10;
                    objPara0.Range.Font.Bold = 1;
                    objPara0.Range.Text = "Date: " + str;
                    pRng0.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    pRng0.Font.Name = "Arial";
                    objPara0.Range.InsertParagraphAfter();

                    //Paragraph oject two
                    objPara1 = objDoc.Content.Paragraphs.Add(ref oMissing);
                    pRng1 = objPara1.Range;
                    pRng1.Font.Size = 14;
                    pRng1.Font.Name = "Arial";
                    pRng1.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                    objPara1.Range.Font.Bold = 1;
                    objPara1.Range.Text = "GRADE REPORT";
                    pRng1.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    objPara1.Range.InsertParagraphAfter();

                    //Paragraph object three
                    objPara2 = objDoc.Content.Paragraphs.Add(ref oMissing);
                    pRng2 = objPara2.Range;
                    pRng2.Font.Size = 13;
                    pRng2.Font.Name = "Arial";
                    objPara2.Range.Font.Bold = 1;
                    objPara2.Range.Text = NumberInWord(semester) + " Semester "+ degree_type + " - " + course_code;
                    pRng2.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    pRng2.ParagraphFormat.SpaceAfter = 0.0F;
                    objPara2.Range.InsertParagraphAfter();

                    //Paragraph object four
                    objPara3 = objDoc.Content.Paragraphs.Add(ref oMissing);
                    pRng3 = objPara3.Range;
                    pRng3.Font.Size = 13;
                    pRng3.Font.Name = "Arial";
                    objPara3.Range.Font.Bold = 1;
                    objPara3.Range.Text = "Degree Examination - " + (String)row.column["month"] + " " + year;
                    pRng3.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    pRng3.ParagraphFormat.SpaceAfter = 0.0F;
                    objPara3.Range.Text += Environment.NewLine;
                    objPara3.Range.InsertParagraphAfter();

                    //Paragraph object five
                    objPara4 = objDoc.Content.Paragraphs.Add(ref oMissing);
                    pRng4 = objPara4.Range;
                    pRng4.Font.Size = 10;
                    pRng4.Font.Name = "Arial";
                    objPara4.Range.Font.Bold = 1;
                    objPara4.Range.Text = "Registration No: " + (long)row.column["registration_id"];
                    pRng4.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    pRng4.ParagraphFormat.SpaceAfter = 4.0F;
                    pRng4.ParagraphFormat.SpaceBefore = 4.0F;
                    objPara4.Range.InsertParagraphAfter();

                    //fetch student names
                    rows2 = Medatabase.fetchRecords("SELECT name FROM student_details WHERE registration_id=" + (long)row.column["registration_id"] + "");

                    objPara5 = objDoc.Content.Paragraphs.Add(ref oMissing);
                    pRng5 = objPara5.Range;
                    pRng5.Font.Size = 10;
                    pRng5.Font.Name = "Arial";
                    objPara5.Range.Font.Bold = 1;
                    objPara5.Range.Text = "Name                 : " + (String)rows2[0].column["name"];
                    pRng5.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    pRng5.ParagraphFormat.SpaceAfter = 4.0F;
                    pRng5.ParagraphFormat.SpaceBefore = 4.0F;
                    objPara5.Range.InsertParagraphAfter();

                    //fetch the number of subjects of each student
                    rows3 = Medatabase.fetchRecords("SELECT COUNT(sub_code) AS count FROM student_marks AS s,exam_master AS e WHERE s.exam_id=e.exam_id AND e.registration_id=" + (long)row.column["registration_id"]);
                    trows = Convert.ToInt16(rows3[0].column["count"]) + 2;
                    
                    //create table
                    wrdRng = objDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    objTable = objDoc.Tables.Add(wrdRng, trows, 4, ref oMissing, ref oMissing);

                    //set column width
                    for (int i = 1; i <= trows; i++)
                    {
                        objTable.Cell(i, 1).Width = 90;
                        objTable.Cell(i, 2).Width = 260;
                        objTable.Cell(i, 3).Width = 60;
                        objTable.Cell(i, 4).Width = 60;
                    }

                    objTable.Cell(1, 1).Range.Text = "Subject Code";
                    objTable.Cell(1, 1).Range.Font.Name = "Arial";
                    objTable.Cell(1, 1).Range.Font.Size = 10;
                    objTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    objTable.Cell(1, 2).Range.Text = "Subject Title";
                    objTable.Cell(1, 2).Range.Font.Name = "Arial";
                    objTable.Cell(1, 2).Range.Font.Size = 10;
                    objTable.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    objTable.Cell(1, 3).Range.Text = "Credit";
                    objTable.Cell(1, 3).Range.Font.Name = "Arial";
                    objTable.Cell(1, 3).Range.Font.Size = 10;
                    objTable.Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    objTable.Cell(1, 4).Range.Text = "Grade";
                    objTable.Cell(1, 4).Range.Font.Name = "Arial";
                    objTable.Cell(1, 4).Range.Font.Size = 9;
                    objTable.Cell(1, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    objTable.Rows[1].Range.Font.Bold = 1;
                    objTable.Borders.Enable = 1;
                    objTable.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth150pt;
                    objTable.Borders.InsideLineWidth = Word.WdLineWidth.wdLineWidth025pt;

                    objTable.Cell(1, 1).Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth150pt;
                    objTable.Cell(1, 2).Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth150pt;
                    objTable.Cell(1, 3).Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth150pt;
                    objTable.Cell(1, 4).Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth150pt;

                    objTable.Rows[trows-1].Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth150pt;
                    objTable.Rows[trows].Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth150pt;
                    objTable.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth150pt;

                    objTable.Rows[trows].Cells[1].Merge(objTable.Rows[trows].Cells[2]);
                    objTable.Rows[trows].Cells[2].Borders[Word.WdBorderType.wdBorderLeft].Color = Word.WdColor.wdColorWhite;
                    objTable.Rows[trows].Cells[3].Borders[Word.WdBorderType.wdBorderLeft].Color = Word.WdColor.wdColorWhite;

                    xrow = 2;
                    sub_codes = new List<String>();
                    rows3 = Medatabase.fetchRecords("SELECT sub_code,grade FROM student_marks WHERE exam_id=" + (long)row.column["exam_id"]);
                    foreach(Row row2 in rows3)
                    {
                        objTable.Cell(xrow, 3).Range.Font.Bold = 1;
                        objTable.Cell(xrow, 1).Range.Text = System.Text.RegularExpressions.Regex.Replace((string)row2.column["sub_code"], @"[_]", " ");
                        objTable.Cell(xrow, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        objTable.Cell(xrow, 4).Range.Text = (string)row2.column["grade"];
                        objTable.Cell(xrow, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        objTable.Cell(xrow, 4).Range.Font.Bold = 1;
                        sub_codes.Add((String)row2.column["sub_code"]);
                        App.Current.Dispatcher.Invoke(new System.Action(() =>
                        {
                            progress_two.Value = (Convert.ToInt16(((prog_two_value * 100) / rows3.Count)));
                        }));
                        xrow++; prog_two_value++;
                    }

                    xrow = 2;
                    String elective = "";
                    foreach (String code in sub_codes)
                    {
                        rows3 = Medatabase.fetchRecords("SELECT name,credits FROM subject_master WHERE sub_code='" + code + "'");
                        elective = objTable.Cell(xrow, 1).Range.Text; //subject code
                        //if subject is elective
                        if (elective.IndexOf(".") != -1)
                            elective = "ELECTIVE - " + semester + " ";
                        else
                            elective = ""; //make empty
                        objTable.Cell(xrow, 2).Range.Text = elective + ((String)rows3[0].column["name"]).ToUpper().Replace('_', '-');
                        objTable.Cell(xrow, 3).Range.Text = rows3[0].column["credits"].ToString();
                        objTable.Cell(xrow, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        App.Current.Dispatcher.Invoke(new System.Action(() =>
                        {
                            progress_two.Value = (Convert.ToInt16(((prog_two_value * 100) / sub_codes.Count)));
                        }));
                        xrow++; prog_two_value++;
                    }

                    rows3 = Medatabase.fetchRecords("SELECT total_credits FROM course_master WHERE course_code='" + course_code + "'");
                    objTable.Cell(xrow, 1).Range.Font.Bold = 1;
                    objTable.Cell(xrow, 1).RightPadding = 0.0f;
                    objTable.Cell(xrow, 1).Range.Text = "TOTAL  : ";
                    objTable.Cell(xrow, 2).Range.Text = rows3[0].column["total_credits"].ToString();
                    objTable.Cell(xrow, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    objTable.Cell(xrow, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    rows3 = Medatabase.fetchRecords("SELECT gpa FROM student_cgpa WHERE exam_id=" + (long)row.column["exam_id"]);
                    objPara6 = objDoc.Content.Paragraphs.Add(ref oMissing);
                    pRng6 = objPara4.Range;
                    pRng6.Font.Size = 10;
                    pRng6.Font.Name = "Arial";
                    pRng6.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    objPara6.Range.Font.Bold = 1;
                    objPara6.Range.Text = " GPA : " + String.Format("{0:0.00}", Convert.ToDouble(rows3[0].column["gpa"]));
                    objPara6.Range.InsertParagraphAfter();

                    output_filename = output_folder + "\\" + course_code + "\\" + row.column["registration_id"].ToString() + "_" + rows2[0].column["name"].ToString() + "_" + row.column["year"].ToString() + ".docx";
                    objDoc.SaveAs(ref output_filename);
                    objDoc.Close(ref oMissing, ref oMissing, ref oMissing);
                    objWord.Quit(ref oMissing, ref oMissing, ref oMissing);
                    objDoc = null; objWord = null;

                    //Update output textbox
                    App.Current.Dispatcher.Invoke(new System.Action(() =>
                    {
                        txtOutput.AppendText(tempProcessLog);
                        txtOutput.ScrollToEnd();
                    }));

                    bgwGenGradeSheet.ReportProgress(Convert.ToInt16(((prog_one_value * 100) / rows.Count)));
                    tempProcessLog = "";
                    prog_one_value++;
                }
            }
            catch (databaseException)
            {
                tempProcessLog += ln + "SOMETHING WENT WRONG IN DATABASE :(" + ln;
                error_occured = true;
            }
            catch (IOException ioe)
            {
                tempProcessLog += ln + "ERROR HANDLING FILE :(" + ln;
                error_occured = true;
                LogWriter.WriteError("While generating mark sheet", ioe.Message);
            }
            catch (UnauthorizedAccessException uae)
            {
                tempProcessLog += ln + "[Access denied] Please select a different path :(" + ln;
                error_occured = true;
                LogWriter.WriteError("While generating mark sheet", uae.Message);
            }
            catch (Exception ex)
            {
                tempProcessLog += ln + "UNEXPECTED ERROR OCCURED :(" + ln + "Please select a diffrent output folder and retry" + ln;
                error_occured = true;
                LogWriter.WriteError("While generating mark sheet", ex.Message);
            }
            finally
            {
                tempProcessLog += ln + ln + "Task completed...";
                tempProcessLog += ln + "Took " + String.Format("{0:0.00}",(DateTime.Now - begin).TotalMinutes) + " minutes";
                App.Current.Dispatcher.Invoke(new System.Action(() =>
                {
                    txtOutput.AppendText(tempProcessLog);
                    txtOutput.ScrollToEnd();
                    progress_one.Value = 100;
                    progress_two.Value = 100;
                }));
            }
        }

        private void bgwGenResultSheet_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //for one progress bar
            progress_one.Value = e.ProgressPercentage;
        }

        private void bgwGenResultSheet_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //task completed
            if (error_occured)
                Windows.generatorWindow.exportCompleted("An error occured while generating result sheet!" + ln + "Please see the logs", true);
            else
                Windows.generatorWindow.exportCompleted("Result sheet has been generated!" + ln + ln + "Please find the file in" + ln + output_folder, false);
        }

        private void bgwGenResultSheet_DoWork(object sender, DoWorkEventArgs e)
        {
            DateTime begin = DateTime.Now; //to calculate how much time it took to complete operation
            try
            {
                excelApp = new Excel.Application();
                int xrow, xcolumn, elective_count = 0;
                List<Row> rows,rows2,row3;
                String[] parts;

                Excel.Workbook gradeWorkbook = (Excel.Workbook)(excelApp.Workbooks.Add());
                Excel.Worksheet gradeWorksheet = (Excel.Worksheet)gradeWorkbook.ActiveSheet;
                gradeWorksheet.Name = course_code;

                tempProcessLog = ln + "======= GENERATING RESULT SHEET ======" + ln;

                gradeWorksheet.Cells[1, 1] = "S.NO";
                gradeWorksheet.Cells[1, 2] = "REG NO";

                //fetch subject_codes and fill in cells
                rows = Medatabase.fetchRecords("SELECT DISTINCT sub_code FROM subject_master WHERE course_code='" + course_code + "' ORDER BY sub_code ASC");
                xrow = 2; xcolumn = 3;
                foreach(Row row in rows)
                {
                    //check if subject is elective
                    if(((String)row.column["sub_code"]).IndexOf(".") != -1)
                    {
                        elective_count++;
                        if (elective_count > 1)
                        {
                            parts = System.Text.RegularExpressions.Regex.Split((String)row.column["sub_code"], @"[.]");
                            gradeWorksheet.Cells[1, xcolumn] = parts[0];
                            
                            gradeWorksheet.Range[gradeWorksheet.Cells[1, xcolumn], gradeWorksheet.Cells[1, xcolumn + 1]].Merge();
                            xcolumn++;
                        }
                        else
                            continue;
                    }
                    else
                        gradeWorksheet.Cells[1, xcolumn] = (String)row.column["sub_code"];
                    xcolumn++;
                }

                gradeWorksheet.Cells[1, xcolumn++] = "GPA";

                if (regIdList == null)
                {
                    //fetch registration id of students'
                    rows = Medatabase.fetchRecords("SELECT registration_id FROM exam_master WHERE course_code='" + course_code + "' AND semester=" + semester + " AND year=" + year + " AND month='" + month + "'");
                }
                else //for custom list of students
                {
                    rows = new List<Row>();
                    Row trow = new Row();
                    foreach (long regid in regIdList) {
                        trow.column.Add("registration_id", regid);
                        rows.Add(trow);
                        trow = new Row();
                    }
                }

                //fill grades and gpa of each students'
                xrow = xcolumn = 2;
                foreach (Row row in rows)
                {
                    tempProcessLog += ln + "Processing row " + xrow;

                    //fill serial no and registration id in cell
                    gradeWorksheet.Cells[xrow, xcolumn] = (long)row.column["registration_id"];
                    gradeWorksheet.Cells[xrow, xcolumn - 1] = xrow - 1;
                    gradeWorksheet.Columns.AutoFit(); //fits the column width according to size

                    //fetch grade of subjects from student_mark table
                    rows2 = Medatabase.fetchRecords("SELECT grade,sub_code FROM student_marks AS s,exam_master AS e WHERE s.exam_id = e.exam_id AND e.registration_id=" + (long)row.column["registration_id"] + " ORDER BY sub_code ASC");

                    //fill subject grades
                    foreach (Row row2 in rows2)
                    {
                        xcolumn++; //begin from 3rd column
                        parts = System.Text.RegularExpressions.Regex.Split((String)row2.column["sub_code"], @"[.]");

                        if (parts.Length > 1)
                        {
                            gradeWorksheet.Cells[xrow, xcolumn] = parts[1];
                            xcolumn++;
                            gradeWorksheet.Cells[xrow, xcolumn] = (String)row2.column["grade"];
                        }
                        else
                            gradeWorksheet.Cells[xrow, xcolumn] = (String)row2.column["grade"];

                        //update 2nd progressbar
                        App.Current.Dispatcher.Invoke(new System.Action(() =>
                        {
                            progress_two.Value = (Convert.ToInt16(((xcolumn * 100) / rows2.Count)));
                        }));
                    }
                    //fill gpa in cell
                    xcolumn++;
                    row3 = Medatabase.fetchRecords("SELECT gpa FROM student_cgpa AS s,exam_master e WHERE s.exam_id = e.exam_id AND e.registration_id=" + (long)row.column["registration_id"]);
                    gradeWorksheet.Cells[xrow, xcolumn] = String.Format("{0:0.00}",(float)row3[0].column["gpa"]);

                    xrow++; xcolumn = 2;

                    //update output text
                    App.Current.Dispatcher.Invoke(new System.Action(() =>
                    {
                        txtOutput.AppendText(tempProcessLog);
                        txtOutput.ScrollToEnd();
                    }));

                    bgwGenResultSheet.ReportProgress(Convert.ToInt16(((xrow * 100) / rows.Count)));
                    tempProcessLog = "";
                }

                //strip end slash if present
                output_folder = (output_folder[output_folder.Length - 1] == '\\') ? output_folder.Substring(0,output_folder.Length - 1) : output_folder;

                //create directory
                if (!Directory.Exists(output_folder + "\\Results"))
                    Directory.CreateDirectory(output_folder + "\\Results");

                output_folder = output_folder + "\\Results";

                //create directory with course_code inside Results folder
                if (!Directory.Exists(output_folder + "\\" + course_code))
                    Directory.CreateDirectory(output_folder + "\\" + course_code);

                //save the excel file
                gradeWorkbook.SaveAs(output_folder + "\\" + course_code + "_" + month + "_" + year +".xlsx");
                gradeWorkbook.Close();
                error_occured = false;
            }
            catch(databaseException de)
            {
                tempProcessLog += ln + "SOMETHING WENT WRONG IN DATABASE :(" + ln;
                error_occured = true;
                LogWriter.WriteError("While generating grade sheet", de.Message);

            }
            catch(IOException ioe)
            {
                tempProcessLog += ln + "ERROR HANDLING FILE :(" + ln;
                error_occured = true;
                LogWriter.WriteError("While generating grade sheet", ioe.Message);
            }
            catch(UnauthorizedAccessException uae)
            {
                tempProcessLog += ln + "[Access denied] Please select a different path :(" + ln;
                error_occured = true;
                LogWriter.WriteError("While generating grade sheet", uae.Message);
            }
            catch (Exception ex)
            {
                tempProcessLog += ln + "UNEXPECTED ERROR OCCURED :(" + ln + "If the file is open then close and retry or try a different path." + ln;
                error_occured = true;
                LogWriter.WriteError("While generating grade sheet", ex.Message);
            }
            finally
            {
                tempProcessLog += ln + ln + "Task completed...";
                tempProcessLog += ln + "Took " + String.Format("{0:0.00}", (DateTime.Now - begin).TotalSeconds) + " seconds";
                App.Current.Dispatcher.Invoke(new System.Action(() =>
                {
                    txtOutput.AppendText(tempProcessLog);
                    txtOutput.ScrollToEnd();
                    progress_one.Value = 100;
                    progress_two.Value = 100;
                }));
            }
        }

        public void getRegIdFromSheet(String course_code, int semester, String month, int year, String fileName)
        {
            this.course_code = course_code;
            this.semester = semester;
            this.month = month;
            this.year = year;
            this.fileName = fileName;

            bgwFetchRegid = new BackgroundWorker();
            bgwFetchRegid.DoWork += BgwFetchRegid_DoWork;
            bgwFetchRegid.ProgressChanged += BgwFetchRegid_ProgressChanged;
            bgwFetchRegid.RunWorkerCompleted += BgwFetchRegid_RunWorkerCompleted;

            bgwFetchRegid.RunWorkerAsync();
            bgwFetchRegid.WorkerReportsProgress = true;
        }

        private void BgwFetchRegid_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (!isValid)
                Windows.generatorWindow.regIdCollected("Your file contains some error! Please see the logs", false, false,null);
            else
            {
                if (regIdList.Count == 0)
                    Windows.generatorWindow.regIdCollected("There are no registration id(s) in the excel sheet", false, false, null);
                else if (incorrect_reg_ids > 0)
                    Windows.generatorWindow.regIdCollected("There are some incorrect registration id's which will be skipped", true,false,regIdList);
                else
                    Windows.generatorWindow.regIdCollected("ALL_OK", true, true,regIdList);
            }
        }

        private void BgwFetchRegid_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progress_one.Value = e.ProgressPercentage;
            progress_two.Value = e.ProgressPercentage;
        }

        private void BgwFetchRegid_DoWork(object sender, DoWorkEventArgs e)
        {
            DateTime begin = DateTime.Now;
            try
            {
                int rowCount = 0, columnCount = 0, i;
                long reg_id;
                isValid = true;
                incorrect_reg_ids = 0;
                regIdList = new List<long>();

                excelApp = new Excel.Application();
                Excel.Workbook cslWorkbook = excelApp.Workbooks.Open(fileName, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel.Worksheet cslWorksheet = (Excel.Worksheet)cslWorkbook.Sheets.get_Item(1);
                Excel.Range cslWorksheetRange = cslWorksheet.UsedRange;

                tempProcessLog = ln + "Collecting registration ids's..." + ln;

                rowCount = cslWorksheetRange.Rows.Count;
                columnCount = cslWorksheetRange.Columns.Count;

                tempProcessLog += ln + "Rows " + rowCount + " | Columns " + columnCount + ln;

                if (rowCount <= 1)
                {
                    tempProcessLog += ln + "There are no rows!";

                    App.Current.Dispatcher.Invoke(new System.Action(() =>
                    {
                        txtOutput.AppendText(tempProcessLog);
                        txtOutput.ScrollToEnd();
                        progress_one.Value = progress_two.Value = 100;
                    }));

                    isValid = false;
                    return;
                }

                if (columnCount < 2 )
                {
                    tempProcessLog += ln + "Expecting columns to be more than 2, " + columnCount + " is present." + ln + "This excelsheet is not meeting the desired format.";

                    App.Current.Dispatcher.Invoke(new System.Action(() =>
                    {
                        txtOutput.AppendText(tempProcessLog);
                        txtOutput.ScrollToEnd();
                        progress_one.Value = progress_two.Value = 100;
                    }));

                    isValid = false;
                    return;
                }

                tempProcessLog += ln + "Expecting first row to be column names. " + ln + "reading from 2nd row...";

                List<long> tmpRegId = new List<long>();
                List<Row> rows;

                for (i = 2; i <= rowCount; i++)
                {
                    try
                    {
                        tempProcessLog += ln + "> Processing Row: " + i + ln;
                        if (cslWorksheet.Cells[i, 2].Value != null)
                        {
                            //check if reg_id is present in database or not for given exam parameters
                            reg_id = Convert.ToInt64(cslWorksheet.Cells[i, 2].Value);
                            rows = Medatabase.fetchRecords("SELECT * FROM exam_master WHERE course_code='" + course_code + "' AND semester=" + semester + " AND month='" + month + "' AND registration_id=" + reg_id);
                            if (rows.Count > 0)
                            {
                                //Check for duplicate reg_id in the same excel sheet
                                if (tmpRegId.IndexOf(reg_id) != -1)
                                {
                                    tempProcessLog += ln + "duplicate reg_id '" + reg_id + "'. Skiped";
                                    incorrect_reg_ids++;
                                }
                                else
                                {
                                    regIdList.Add(reg_id);
                                    tmpRegId.Add(reg_id);
                                }
                            }
                            else
                            {
                                tempProcessLog += ln + reg_id + " no records present for the given exam details. Skipping";
                                incorrect_reg_ids++;
                            }
                        }
                        else
                        {
                            tempProcessLog += ln + "empty value given. Skiping to next record! " + ln;
                            incorrect_reg_ids++;
                        }
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException be)
                    {
                        invalidRows++;
                        isValid = false;
                        tempProcessLog += ln + "In row [" + i + "] expecting numeric value, string given!";
                        LogWriter.WriteError("While fetching reg_id from sheet", be.Message);
                        incorrect_reg_ids++;
                    }

                    //Update output textbox
                    App.Current.Dispatcher.Invoke(new System.Action(() =>
                    {
                        txtOutput.AppendText(tempProcessLog);
                        txtOutput.ScrollToEnd();
                    }));

                    bgwFetchRegid.ReportProgress(Convert.ToInt16(((i * 100) / rowCount)));
                    tempProcessLog = "";
                }
            }
            catch (System.IO.IOException ioe)
            {
                tempProcessLog += ln + "ERROR HANDLING FILE :(" + ln + "Please retry";
                LogWriter.WriteError("While fetching reg_id from sheet", ioe.Message);
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
                LogWriter.WriteError("While fetching reg_id from sheet", ex.Message);
                isValid = false;
                return;
            }
            finally
            {
                tempProcessLog += ln + "Took " + String.Format("{0:0.00}", (DateTime.Now - begin).TotalSeconds) + " seconds" + ln + ln;
                App.Current.Dispatcher.Invoke(new System.Action(() =>
                {
                    txtOutput.AppendText(tempProcessLog);
                    txtOutput.ScrollToEnd();
                    progress_one.Value = progress_two.Value = 100;
                }));

                excelApp.Workbooks.Close();
            }
        }

        /// <summary>
        /// Generates a result sheet containing list of failed or absent students.
        /// </summary>
        /// <param name="course_code">course code</param>
        /// <param name="semester">semsester</param>
        /// <param name="year">year</param>
        /// <param name="month">month</param>
        /// <param name="output_folder"> Where to save the excel file</param>
        /// <param name="stype">Student type failed or absent</param>
        public void generateAnalysisResultSheet(String course_code, int semester, String month, int year, String output_folder,StudentType stype)
        {
            this.course_code = course_code;
            this.semester = semester;
            this.month = month;
            this.year = year;
            this.output_folder = output_folder;
            this.stype = stype;

            bgwGenAnalysisSheet = new BackgroundWorker();
            bgwGenAnalysisSheet.DoWork += BgwGenAnalysisSheet_DoWork;
            bgwGenAnalysisSheet.ProgressChanged += BgwGenAnalysisSheet_ProgressChanged;
            bgwGenAnalysisSheet.RunWorkerCompleted += BgwGenAnalysisSheet_RunWorkerCompleted;

            bgwGenAnalysisSheet.RunWorkerAsync();
            bgwGenAnalysisSheet.WorkerReportsProgress = true;
        }

        private void BgwGenAnalysisSheet_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (!error_occured)
                Windows.adminWindow.analysisReportResult("Analysis result sheet has been generated." + ln + ln + "Please find the file in" + ln + output_folder, true);
            else
                Windows.adminWindow.analysisReportResult("There was an error generating analysis excel sheet :(", false);
        }

        private void BgwGenAnalysisSheet_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            Windows.adminWindow.prgsbarAnalysis.Value = e.ProgressPercentage;
        }

        private void BgwGenAnalysisSheet_DoWork(object sender, DoWorkEventArgs e)
        {
            DateTime begin = DateTime.Now; //to calculate how much time it took to complete operation
            try
            {
                excelApp = new Excel.Application();
                int xrow;
                List<Row> rows;

                Excel.Workbook gradeWorkbook = (Excel.Workbook)(excelApp.Workbooks.Add());
                Excel.Worksheet gradeWorksheet = (Excel.Worksheet)gradeWorkbook.ActiveSheet;
                gradeWorksheet.Name = course_code;

                gradeWorksheet.Cells[1, 1] = "S.NO";
                gradeWorksheet.Cells[1, 2] = "REG NO";
                gradeWorksheet.Cells[1, 3] = "Name";
                gradeWorksheet.Cells[1, 4] = "Subject code";
                gradeWorksheet.Cells[1, 5] = "Subject name";

                rows = Medatabase.fetchRecords("SELECT em.registration_id AS reg_id,sd.name As std_name,smr.sub_code AS sub_code,sm.name AS sub_name FROM exam_master AS em,student_marks AS smr,student_details AS sd, subject_master AS sm WHERE em.exam_id=smr.exam_id AND sm.sub_code=smr.sub_code AND em.course_code='" + course_code + "' AND em.semester=" + semester + " AND em.year=" + year + " AND em.month='" + month + "' AND sd.registration_id=em.registration_id AND smr.status='" + ((stype == StudentType.FAILURE) ? "Fail" : "Absent") + "'");

                xrow = 2;
                foreach (Row row in rows)
                {
                    gradeWorksheet.Cells[xrow, 1] = xrow - 1;
                    gradeWorksheet.Cells[xrow, 2] = Convert.ToInt64(row.column["reg_id"]);
                    gradeWorksheet.Cells[xrow, 3] = (String)row.column["std_name"];
                    gradeWorksheet.Cells[xrow, 4] = (String)row.column["sub_code"];
                    gradeWorksheet.Cells[xrow, 5] = (String)row.column["sub_name"];
                    gradeWorksheet.Columns.AutoFit();
                    xrow++;
                    bgwGenAnalysisSheet.ReportProgress(Convert.ToInt16(((xrow * 100) / rows.Count)));
                }

                //strip end slash if present
                output_folder = (output_folder[output_folder.Length - 1] == '\\') ? output_folder.Substring(0, output_folder.Length - 1) : output_folder;

                //create directory
                if (!Directory.Exists(output_folder + "\\Results"))
                    Directory.CreateDirectory(output_folder + "\\Results");

                output_folder = output_folder + "\\Results";

                //save the excel file
                gradeWorkbook.SaveAs(output_folder + "\\" + course_code + "_" + year + "_" + month + "_" + ((stype == StudentType.FAILURE) ? "Failure" : "Absent") + ".xlsx");
                gradeWorkbook.Close();
                error_occured = false;
            }
            catch (databaseException de)
            {
                error_occured = true;
                LogWriter.WriteError("While generating analysis result sheet", de.Message);

            }
            catch (IOException ioe)
            {
                error_occured = true;
                LogWriter.WriteError("While generating analysis result sheet", ioe.Message);
            }
            catch (UnauthorizedAccessException uae)
            {
                error_occured = true;
                LogWriter.WriteError("While generating analysis result sheet", uae.Message);
            }
            catch (Exception ex)
            {
                error_occured = true;
                LogWriter.WriteError("While generating analysis result sheet", ex.Message);
            }
        }

        String NumberInWord(int number)
        {
            String word;
            switch (number)
            {
                case 1:
                    word = "First";
                    break;
                case 2:
                    word = "Second";
                    break;
                case 3:
                    word = "Third";
                    break;
                case 4:
                    word = "Fourth";
                    break;
                case 5:
                    word = "Fifth";
                    break;
                case 6:
                    word = "Sixth";
                    break;
                case 7:
                    word = "Seventh";
                    break;
                case 8:
                    word = "Eighth";
                    break;
                default:
                    word = "Number";
                    break;
            }
            return word;
        }

        /// <summary>
        /// Returns list of years
        /// </summary>
        /// <returns>integer collection [years] </returns>
        public static List<int> getExamYearsList()
        {
            List<int> years = new List<int>();
            List<Row> rows = Medatabase.fetchRecords("SELECT DISTINCT year FROM exam_master ORDER BY year");

            if (rows.Count > 0)
                foreach (Row row in rows)
                    years.Add(Convert.ToInt16(row.column["year"]));

            int year = DateTime.Today.Year;
            if (years.Count >= 1)
            {
                year = years[years.Count - 1] + 1;
                for (int i = year; i < year + 4; i++)
                    years.Add(i);
            }
            else
            {
                year -= 2;
                for (int i = year; i < year + 4; i++)
                    years.Add(i);
            }
            return years;
        }

        /// <summary>
        /// Returns a list of student who have either failed or were absent during exam
        /// </summary>
        /// <param name="course"></param>
        /// <param name="semester"></param>
        /// <param name="year"></param>
        /// <param name="month"></param>
        /// <returns>item collection</returns>
        public static List<NonregularCols> getNonregularStudentList(String course,int semester,int year,String month,StudentType stype)
        {
            List<NonregularCols> items = new List<NonregularCols>();
            try
            {
                foreach (Row row in Medatabase.fetchRecords("SELECT em.registration_id AS reg_id,sd.name As std_name,smr.sub_code,sm.name AS sub_name FROM exam_master AS em,student_marks AS smr,student_details AS sd, subject_master AS sm WHERE em.exam_id=smr.exam_id AND sm.sub_code=smr.sub_code AND em.course_code='" + course + "' AND em.semester=" + semester + " AND em.year=" + year + " AND em.month='" + month + "' AND sd.registration_id=em.registration_id AND smr.status='" + ((stype == StudentType.FAILURE) ? "Fail" : "Absent") + "'"))
                    items.Add(new NonregularCols() {
                        reg_id = Convert.ToInt64(row.column["reg_id"]),
                        std_name = (String)row.column["std_name"],
                        sub_code = (String)row.column["sub_code"],
                        sub_name = (String)row.column["sub_name"],
                    });
            }
            catch (Exception ex)
            {
                LogWriter.WriteError("Fetching exam result of student", ex.Message);
                items = null;
            }
            return items;
        }
    }

    //To display in listview for marks
    class ResultCols
    {
        public string mrk_s_code { get; set; }
        public string mrk_s_grade { get; set; }
        public float mrk_s_iamark { get; set; }
        public float mrk_s_eamark { get; set; }
        public float mrk_s_fmark { get; set; }
    }

    //To display in Listview of absent/failed students
    class NonregularCols
    {
        public long reg_id { get; set; }
        public string std_name { get; set; }
        public string sub_code { get; set; }
        public string sub_name { get; set; }
    }
}
