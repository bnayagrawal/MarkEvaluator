using System;
using System.Windows.Controls;
using System.Collections.Generic;
using System.ComponentModel;

namespace markevaluator
{
    class Students
    {
        /// <summary>
        /// Returns student details list
        /// </summary>
        /// <param name="course_code">Course code</param>
        /// <param name="year">Batch or year</param>
        /// <returns>Studentcols collection</returns>
        public static List<StudentCols> getStudentList(String course_code,int year)
        {
            List<StudentCols> items = new List<StudentCols>();
            try
            {
                List<Row> rows = Medatabase.fetchRecords("SELECT * FROM student_details WHERE course_code='" + course_code + "' AND year=" + year);

                foreach (Row row in rows)
                    items.Add(new StudentCols
                    {
                        s_reg_id = (long)row.column["registration_id"],
                        s_name = (string)row.column["name"],
                        s_month = (string)row.column["month"]
                    });
            }
            catch (Exception ex)
            {
                LogWriter.WriteError("fetching student details", ex.Message);
            }
            return items;
        }

        /// <summary>
        /// Returns batch years of a course
        /// </summary>
        /// <param name="course_code">course code</param>
        /// <returns>int collection</returns>
        public static List<int> getBatchList(String course_code)
        {
            List<int> list = new List<int>();
            try
            {
                List<Row> rows = Medatabase.fetchRecords("SELECT DISTINCT year FROM student_details WHERE course_code='" + course_code + "'");
                foreach (Row row in rows)
                    list.Add((int)row.column["year"]);
            }
            catch(Exception)
            {
                //Something went wrong
            }
            return list;
        }
    }

    class StudentCols
    {
        public long s_reg_id { get; set; }
        public string s_name { get; set; }
        public string s_month { get; set; }
    }

    public enum StudentType
    {
        REGULAR,
        ABSENT,
        FAILURE
    }
}
