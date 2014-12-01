using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ConsoleApplication3
{
    class Program
    {
        static void Main(string[] args)
        {
            DateTime curTime = DateTime.Now;
            Program prg = new Program();
            prg.App();
            DateTime endTime = DateTime.Now;
            TimeSpan duration = endTime - curTime;
            Console.WriteLine(String.Format("Продолжительность: {0}", duration));
            Console.Read();

        }

        public void App()
        {
            Workbook wrkbook;
            IEnumerable<Sheet> wrksheets;
            SharedStringTable sharedStrings;
            WorksheetPart jobSheet;

            string sheetId;

            using (SpreadsheetDocument document = 
                SpreadsheetDocument.Open(@"C:\Users\alexb_000\Desktop\1\Freelance Jobs_EN_RU_1-Jan_17 Nov.xlsx", true))
            {
                wrkbook = document.WorkbookPart.Workbook;
                wrksheets = wrkbook.Descendants<Sheet>();
                sharedStrings = document.WorkbookPart.SharedStringTablePart.SharedStringTable;

                sheetId = wrksheets.First(s => s.Name == "Freelance Jobs").Id;
                jobSheet = (WorksheetPart)document.WorkbookPart.GetPartById(sheetId);

                List<Project> projects = Project.LoadProjects(jobSheet, sharedStrings);

                IEnumerable<Project> groupProjects =
                    from proj in projects
                    where (
                        (proj.manager == "Irina Zakharova" ||
                        proj.manager == "Nikita Maynagashev" ||
                        proj.manager == "Karolina Shelepova" ||
                        proj.manager == "Margarita Kovalevskaya" ||
                        proj.manager == "Irina Min") &&
                        proj.groupServices == "Editing" &&
                        proj.unit == "words"
                    )
                    select proj;

                IEnumerable<DateTime> assignedDates =
                    from proj in groupProjects
                    select proj.assigned;

                DateTime minAssigned = assignedDates.Min();

                IEnumerable<DateTime> deadlines =
                    from proj in groupProjects
                    select proj.deadline;

                DateTime maxDeadline = deadlines.Max();

                for (DateTime i = minAssigned; i < maxDeadline; i.AddDays(1)) 
                {
                       
                }
                        
            }
        }
    }

    class Project
    {
        public string jobCode { get; set; }
        public string jobName { get; set; }
        public string expert { get; set; }
        public string projectCode { get; set; }
        public DateTime assigned { get; set; }
        public DateTime deadline { get; set; }
        public DateTime completed { get; set; }
        public string groupServices { get; set; }
        public double volume { get; set; }
        public string unit { get; set; }
        public string client { get; set; }
        public string manager { get; set; }
        public TimeSpan duration { get; set; }
        public double days { get; set; }
        public double productivity { get; set; }

        public static List<Project> LoadProjects(WorksheetPart wrksheetPart, SharedStringTable sharedStrings)
        {
            List<Project> projects = new List<Project>();
            Worksheet wrksheet = wrksheetPart.Worksheet;

            IEnumerable<Row> dataRows =
                from row in wrksheet.Descendants<Row>()
                select row;

            Row headingsRow = dataRows.First();

            IEnumerable<string> headings = Project.ParseRow(headingsRow, sharedStrings);
            Array headingsArray = headings.ToArray();

           
            IEnumerable<Row> projectRows =
                from dataRow in dataRows
                where dataRow.RowIndex > 1
                select dataRow;

            foreach (Row dataRow in projectRows)
            {
                Project prj = new Project();
                IEnumerable<string> textValues = Project.ParseRow(dataRow, sharedStrings);

                try
                {
                    prj.jobCode = textValues.ElementAt(Array.IndexOf(headingsArray, "Job Code"));
                    prj.jobName = textValues.ElementAt(Array.IndexOf(headingsArray, "Job Name"));
                    prj.expert = textValues.ElementAt(Array.IndexOf(headingsArray, "Expert"));
                    prj.projectCode = textValues.ElementAt(Array.IndexOf(headingsArray, "Project Code"));
                    prj.assigned = DateTime.FromOADate(double.Parse(textValues.ElementAt(Array.IndexOf(headingsArray, "Assigned")), System.Globalization.CultureInfo.InvariantCulture));
                    prj.deadline = DateTime.FromOADate(double.Parse(textValues.ElementAt(Array.IndexOf(headingsArray, "Deadline")), System.Globalization.CultureInfo.InvariantCulture));
                    prj.completed = DateTime.FromOADate(double.Parse(textValues.ElementAt(Array.IndexOf(headingsArray, "Completed")), System.Globalization.CultureInfo.InvariantCulture));
                    prj.groupServices = textValues.ElementAt(Array.IndexOf(headingsArray, "Group of Services"));
                    prj.volume = Double.Parse(textValues.ElementAt(Array.IndexOf(headingsArray, "Volume")), System.Globalization.CultureInfo.InvariantCulture);
                    prj.unit = textValues.ElementAt(Array.IndexOf(headingsArray, "Units"));
                    prj.manager = textValues.ElementAt(Array.IndexOf(headingsArray, "Project Manager"));
                    prj.duration = prj.deadline - prj.assigned;
                    prj.days = prj.duration.TotalDays;
                    prj.productivity = prj.volume / prj.days;
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
                
                projects.Add(prj);
            }

            return projects;

        }

        private static IEnumerable<string> ParseRow(Row parseRow, SharedStringTable sharedStrings)
        {
            IEnumerable<string> result;
            result = from dataCell in parseRow.Descendants<Cell>()
                     where dataCell.CellValue != null
                     select (dataCell.DataType != null
                     && dataCell.DataType.HasValue
                     && dataCell.DataType == CellValues.SharedString
                     ? sharedStrings.ChildElements[int.Parse(dataCell.CellValue.InnerText)].InnerText :
                     dataCell.CellValue.InnerText);
            return result;
        }
        
    }
}
