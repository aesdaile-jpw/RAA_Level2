#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;


#endregion

namespace RAA_Level2
{
    [Transaction(TransactionMode.Manual)]

    public class Command : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // put any code needed for the form here

            // open form
            MyForm currentForm = new MyForm()
            {
                Width = 800,
                Height = 500,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                Topmost = true,
            };

            currentForm.ShowDialog();

            // if form is cancelled, exit command

            if (currentForm.DialogResult == false)
            {
                TaskDialog.Show("Project Setup", "Project Setup cancelled");
                return Result.Cancelled;
            }

            // get data from form

            string filePath = currentForm.GetFilePath();
            string projectUnits = currentForm.GetUnits();
            bool floorPlans = currentForm.GetPlansCreate();
            bool ceilingPlans = currentForm.GetCeilingPlansCreate();
            bool createSheets = currentForm.GetSheetsCreate();


            // test output

            //string msg = "File Path: " + filePath + "\n" +
            //             "Project Units: " + projectUnits + "\n" +
            //             "Create Floor Plans: " + floorPlans + "\n" +
            //             "Create Ceiling Plans: " + ceilingPlans + "\n" +
            //             "Create Sheets: " + createSheets + "\n";
            //TaskDialog.Show("Project Setup", msg);

            // set project units

            using (Transaction tx = new Transaction(doc, "Set Project Units"))
            {
                tx.Start();

                // Get current units
                Units units = doc.GetUnits();

                // Create new FormatOptions for project units
                if (projectUnits == "mm")
                {
                    FormatOptions mmFormatOptions = new FormatOptions(UnitTypeId.Millimeters);
                    units.SetFormatOptions(SpecTypeId.Length, mmFormatOptions);
                }
                else
                {
                    FormatOptions mmFormatOptions = new FormatOptions(UnitTypeId.Feet);
                    units.SetFormatOptions(SpecTypeId.Length, mmFormatOptions);
                }
                // Apply the updated units to the document
                doc.SetUnits(units);

                tx.Commit();
                tx.Dispose();
            }

            // Replace this line:
            // List<string>[] csvData = ReadCSVFile(filePath);
            // With this line
            //List<string[]> csvData = ReadCSVFile(filePath);

            List<List<string>> csvData = ExcelReader.ReadExcelToList(filePath);
            csvData.RemoveAt(0); // Remove header row

            //int rowCount = csvData.Count;
            //TaskDialog.Show("CSV Data", $"Number of data rows: {rowCount}");

            // create lists from csvData

            List<string> levelNames = ListFromCSV(csvData, 0);
            List<string> levelElevationsFT = ListFromCSV(csvData, 1);
            List<string> levelElevationsM = ListFromCSV(csvData, 2);

            //TaskDialog.Show("Lists", GetMethod() + "\n" +
            //    $"Level Names: {string.Join(", ", levelNames)}\n" +
            //    $"Level Elevations (ft): {string.Join(", ", levelElevationsFT)}\n" +
            //    $"Level Elevations (m): {string.Join(", ", levelElevationsM)}");

            // convert string lists to double lists for heights

            List<double> levelElevsFT = ListStringToDouble(levelElevationsFT);
            List<double> levelElevsM = ListStringToDouble(levelElevationsM);

            // Create levels

            List<Level> createdLevels = new List<Level>();


            using (Transaction tx = new Transaction(doc, "Create Levels"))
            {
                tx.Start();
                for (int i = 0; i < levelNames.Count; i++)
                {
                    string lvlName = levelNames[i];
                    double lvlElev = levelElevsFT[i];
                    double lvlElevM = levelElevsM[i];
                    double lvlElevFT = lvlElev; // default to feet
                    if (projectUnits == "mm")
                    {
                        lvlElevFT = lvlElevM * 3.2808399; // convert m to feet to create levels
                    }
                    Level newLevel = Level.Create(doc, lvlElevFT);
                    newLevel.Name = lvlName;
                    createdLevels.Add(newLevel);
                }
                tx.Commit();
                tx.Dispose();
            }

            TaskDialog.Show("Levels Created", $"Created {createdLevels.Count} levels.");

            // Create views from Levels

            CreatePlanViewsAndSheets.CreateViewsAndSheets(doc, createdLevels, floorPlans, ceilingPlans, createSheets);

            return Result.Succeeded;
        }

        //private List<string> ListFromCSV(List<List<string>> csvData, int v)
        //{
        //    throw new NotImplementedException();
        //}

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }

        // Replace the ListFromCSV method with the correct implementation.
        // The original method signature and implementation are incorrect for the intended use.
        // The method should return List<string> instead of List<string[]> and accept List<string[]> as input.

        public static List<string> ListFromCSV(List<List<string>> csvList, int n)
        {
            var nthElements = csvList
                .Where(sublist => sublist.Count > n)
                .Select(sublist => sublist[n])
                .ToList();
            return nthElements;
        }

        public static List<double> ListStringToDouble(List<string> stringList)
        {
            List<double> doubleList = new List<double>();
            foreach (string str in stringList)
            {
                if (double.TryParse(str, out double value))
                {
                    doubleList.Add(value);
                }
                else
                {
                    // Handle parsing error if necessary
                    doubleList.Add(0); // or throw an exception, or skip, etc.
                }
            }
            return doubleList;
        }

        public static List<string[]> ReadCSVFile(string filepath)
        {
            //create list for sheet information
            List<string[]> dataList = new List<string[]>();

            using (var myReader = new StreamReader(filepath))
            {
                while (!myReader.EndOfStream)
                {
                    var line = myReader.ReadLine();
                    var values = line.Split(',');

                    dataList.Add(values);
                }

                myReader.Close();
            }

            return dataList;
        }

        //public static List<string[]> ReadExcelFile(string filepath)
        //{
        //    // Placeholder for Excel reading logic
        //    // Implement Excel reading using a library like EPPlus, ClosedXML, or Interop
        //    List<string[]> dataList = new List<string[]>();
        //    // Example: dataList.Add(new string[] { "Level 1", "0", "0" });
        //    return dataList;
        //}

        public class ExcelReader
        {
            public static List<List<string>> ReadExcelToList(string filePath)
            {
                var data = new List<List<string>>();
                ExcelPackage.License.SetNonCommercialOrganization("JPW");

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0]; // First worksheet
                    int rows = worksheet.Dimension.Rows;
                    int columns = worksheet.Dimension.Columns;

                    for (int i = 1; i <= rows; i++)
                    {
                        var row = new List<string>();
                        for (int j = 1; j <= columns; j++)
                        {
                            row.Add(worksheet.Cells[i, j].Text);
                        }
                        data.Add(row);
                    }
                }
                return data;
            }
        }


        public class CreatePlanViewsAndSheets
        {
            public static void CreateViewsAndSheets(Document doc, List<Level> levels, bool floorPlan, bool ceilingPlan, bool createSheets)
            {
                List<ViewPlan> createdPlanViews = new List<ViewPlan>();
                List<ViewPlan> createdCeilingViews = new List<ViewPlan>();

                using (Transaction tx = new Transaction(doc, "Create Plan Views"))
                {
                    tx.Start();

                    foreach (Level level in levels)
                    {
                        // Check if a plan view already exists for the level
                        bool viewExists = false;
                        FilteredElementCollector collector = new FilteredElementCollector(doc)
                            .OfClass(typeof(ViewPlan));

                        foreach (ViewPlan view in collector)
                        {
                            if (view.GenLevel != null && view.GenLevel.Id == level.Id)
                            {
                                viewExists = true;
                                break;
                            }
                        }

                        // If no plan view exists, create one
                        if (!viewExists)
                        {
                            if (floorPlan)
                            {
                                ViewFamilyType viewFamilyType = GetViewFamilyType(doc, ViewFamily.FloorPlan);
                                if (viewFamilyType != null)
                                {
                                    ViewPlan newViewPlan = ViewPlan.Create(doc, viewFamilyType.Id, level.Id);
                                    createdPlanViews.Add(newViewPlan);
                                }
                            }
                            if (ceilingPlan)
                            {
                                ViewFamilyType viewFamilyType = GetViewFamilyType(doc, ViewFamily.CeilingPlan);
                                if (viewFamilyType != null)
                                {
                                    ViewPlan newViewCeiling = ViewPlan.Create(doc, viewFamilyType.Id, level.Id);
                                    createdCeilingViews.Add(newViewCeiling);
                                }
                            }
                        }
                    }

                    TaskDialog.Show("Plan Views Created", $"Created {createdPlanViews.Count} floor plans and {createdCeilingViews.Count} ceiling plans.");

                    tx.Commit();
                    tx.Dispose();
                }

                if (createSheets)
                {
                    using (Transaction tx = new Transaction(doc, "Create Sheets"))
                    {
                        tx.Start();

                        // Get title block family symbol
                        FamilySymbol titleBlock = new FilteredElementCollector(doc)
                            .OfClass(typeof(FamilySymbol))
                            .OfCategory(BuiltInCategory.OST_TitleBlocks)
                            .FirstOrDefault() as FamilySymbol;
                        if (titleBlock == null)
                        {
                            TaskDialog.Show("Error", "No title block found in the project.");
                            return;
                        }
                        // Ensure the title block is activated
                        if (!titleBlock.IsActive)
                        {
                            titleBlock.Activate();
                            doc.Regenerate();
                        }
                        // Create sheets for each created plan view
                        foreach (ViewPlan view in createdPlanViews)
                        {
                            ViewSheet newSheet = ViewSheet.Create(doc, titleBlock.Id);
                            newSheet.Name = view.Name + " Sheet";
                            newSheet.SheetNumber = "A-" + view.GenLevel.Name; // Example sheet number format
                            // Place the view on the sheet
                            XYZ insertPoint = new XYZ(0, 0, 0); // Adjust as needed
                            Viewport.Create(doc, newSheet.Id, view.Id, insertPoint);
                        }
                        // Create sheets for each created ceiling view
                        foreach (ViewPlan view in createdCeilingViews)
                        {
                            ViewSheet newSheet = ViewSheet.Create(doc, titleBlock.Id);
                            newSheet.Name = view.Name + " Ceiling Sheet";
                            newSheet.SheetNumber = "A-" + view.GenLevel.Name + "C"; // Example sheet number format
                            // Place the view on the sheet
                            XYZ insertPoint = new XYZ(0, 0, 0); // Adjust as needed
                            Viewport.Create(doc, newSheet.Id, view.Id, insertPoint);
                        }
                        TaskDialog.Show("Sheets Created", $"Created sheets for {createdPlanViews.Count} floor plans and {createdCeilingViews.Count} ceiling plans.");

                        tx.Commit();
                        tx.Dispose();
                    }
                }
            }


            private static ViewFamilyType GetViewFamilyType(Document doc, ViewFamily viewFamily)
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewFamilyType));

                foreach (ViewFamilyType viewFamilyType in collector)
                {
                    if (viewFamilyType.ViewFamily == viewFamily)
                    {
                        return viewFamilyType;
                    }
                }

                return null;
            }
        }
    }
}
