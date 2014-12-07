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

               DateTime i = minAssigned;

               List<ProjectCalendar> calend = new List<ProjectCalendar>();

               while (i < maxDeadline)
               {
                   ProjectCalendar day = new ProjectCalendar();
                   day.day = i;

                   IEnumerable<Project> dayProjects =
                       from proj in groupProjects
                       where ((proj.assigned <= i) &&
                       (proj.deadline >= i))
                       select proj;

                   day.projects = dayProjects.ToList<Project>();
                   calend.Add(day);
                   i = i.AddDays(1);
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

            ProjectStrings projectStrings = new ProjectStrings();
            projectStrings.headings = headings.ToArray();

           
           projectStrings.projectRows =
                from dataRow in dataRows
                where dataRow.RowIndex > 1
                select dataRow;

            foreach (Row dataRow in projectStrings.projectRows)
            {
                Project prj = new Project();
                projectStrings.textValues = Project.ParseRow(dataRow, sharedStrings);

                try
                {
                    prj.jobCode = projectStrings.ParseText("Job Code");
                    prj.jobName = projectStrings.ParseText("Job Name");
                    prj.expert = projectStrings.ParseText("Expert");
                    prj.projectCode = projectStrings.ParseText("Project Code");
                    prj.assigned = projectStrings.ParseDate("Assigned");
                    prj.deadline = projectStrings.ParseDate("Deadline");
                    prj.completed = projectStrings.ParseDate("Completed");
                    prj.groupServices = projectStrings.ParseText("Group of Services");
                    prj.volume = projectStrings.ParseDouble("Volume");
                    prj.unit = projectStrings.ParseText("Units");
                    prj.manager = projectStrings.ParseText("Project Manager");
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

    //Used to parse row data
    class ProjectStrings
    {
        public Array headings { get; set; }
        public IEnumerable<Row> projectRows { get; set; }
        public IEnumerable<string> textValues { get; set; }

        public string ParseText (string arg)
        {
            string result = textValues.ElementAt(Array.IndexOf(headings, arg));
            return result;
        }

        public DateTime ParseDate (string arg)
        {
            DateTime result = DateTime.FromOADate(ParseDouble(arg));
            return result;
        }

        public double ParseDouble (string arg)
        {
            double result = Double.Parse(textValues.ElementAt(Array.IndexOf(headings, arg)), System.Globalization.CultureInfo.InvariantCulture);
            return result;
        }

    }


    class ProjectCalendar
    {
        public DateTime day { get; set; }
        private List<Project> _projects;
        public List<Project> projects 
        { get
            {
                return _projects;
            }
          set
            {
                _projects = value;
                workload = 0;
                foreach (Project proj in _projects)
                {
                    workload += proj.productivity;
                }
            }
        }
        public double workload
        {
            get;
            private set;
        }

        public void AddProj (Project proj)
        {
            projects.Add(proj);
            workload += proj.productivity;
        }


        
    }
}
