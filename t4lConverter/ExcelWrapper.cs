using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Excel;
using System.Text.RegularExpressions;
using System.Collections;

namespace t4lConverter
{
    public class ExcelWrapper
    {
        public event EventHandler ThreadDone;

        public void Convert(FileInfo filename)
        {
            var csvData = ExtractRowColumns(filename);
            CreateNewFormat(csvData, filename.DirectoryName);
            if (ThreadDone != null)
                ThreadDone(this, EventArgs.Empty);
        }

        private void CreateNewFormat(string csvdata, string dirName)
        {
            //extract all lines until the next block
            string block = null;
            var output = new StringBuilder();
            using (StringReader reader = new StringReader(csvdata))
            {
                string line;
                string weekFileName = null;
                while ((line = reader.ReadLine()) != null)
                {
                    if (line.Contains("Week of:") && !string.IsNullOrEmpty(block))
                    {
                        output.Append(CreateWeek(block));
                        block = string.Empty;
                        weekFileName = string.Format(@"{0}.csv", output.ToString().Substring(9, 12));
                        weekFileName = weekFileName.Replace(",", "");
                        using (var sw = new StreamWriter(dirName + @"\" + weekFileName, true))
                        {
                            sw.Write(output);
                        }
                        output.Clear();
                    }
                    block += line + "\n";
                }
            }            
        }

        private string CreateWeek(string block)
        {
            //header
            //day of week, subject, lesson, LA#, Lesson, LA#, Quiz, Chapter Test
            StringBuilder sb = new StringBuilder();
            var week = Regex.Match(block, @"Week of:.*,").Value;
            sb.AppendLine(week);            
            var subjects = new ArrayList();
            //get the subject and the index where it starts
            foreach (Match m in Regex.Matches(block, @".*,\n\nGrade Level:"))
            {                
                var s = Regex.Split(m.Value, ",");                
                var subObj = new Subject();
                subObj.Name = s[0];
                subObj.Index = m.Index;
                subjects.Add(subObj);
            }
            //for each subject, rip the lessons, quizes and tests out
            for (int i = 0; i < subjects.Count; i++)
            {
                //subject
                var next = 0;
                var obj = subjects[i] as Subject;
                if (i + 1 < subjects.Count)
                {
                    var tObj = subjects[i + 1] as Subject;
                    next = tObj.Index;
                }
                else
                {
                    next = block.Length;
                }
                var day = block.Substring(obj.Index, next - obj.Index);                
                sb.AppendLine("," + obj.Name + ",,LA #,");

                string line;
                using (StringReader reader = new StringReader(day))
                {
                    while ((line = reader.ReadLine()) != null)
                    {
                        //lessons
                        if (Regex.IsMatch(line, @".+,\d.*,|.+,\w\d.*|Lesson Quiz:.*|Chapter Test:.*"))
                        {
                            sb.AppendLine(",," + line);
                        }
                    }
                }                
            }
            return sb.ToString();
        }

        private string ExtractRowColumns(FileInfo file)
        {
            string csvData = null;
            
            var stream = File.Open(file.FullName, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelDataReader = null;

            if (file.Extension.ToLower().Equals(".xls"))
            {
                // Reading from a binary Excel file ('97-2003 format; *.xls)
                excelDataReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else if (file.Extension.ToLower().Equals(".xlsx"))
            {
                // Reading from a OpenXml Excel file (2007 format; *.xlsx)
                excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }

            // DataSet - The result of each spreadsheet will be created in the result.Tables
            if (excelDataReader != null)
            {
                var result = excelDataReader.AsDataSet();
                // Free resources (IExcelDataReader is IDisposable)
                excelDataReader.Close();

                var rowNo = 0;
                //CoreLog.Info("extracting rows and columns...");

                while (rowNo < result.Tables[0].Rows.Count)
                {
                    for (var colNo = 0; colNo < result.Tables[0].Columns.Count; colNo++)
                    {
                        var str = result.Tables[0].Rows[rowNo][colNo].ToString();
                        if (str != string.Empty)
                        {
                            csvData += str + ",";
                        }
                    }

                    rowNo++;
                    csvData += "\n";
                }
            }

            return csvData;
        }
    }
}
