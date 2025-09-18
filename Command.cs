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

            // test output

            string msg = "File Path: " + filePath + "\n" +
                         "Project Units: " + projectUnits + "\n" +
                         "Create Floor Plans: " + floorPlans + "\n" +
                         "Create Ceiling Plans: " + ceilingPlans + "\n";

            TaskDialog.Show("Project Setup", msg);

            // set project units

            if (projectUnits == "mm")
            {
                using (Transaction tx = new Transaction(doc, "Set Project Units to mm"))
                {
                    tx.Start();

                    // Get current units
                    Units units = doc.GetUnits();

                    // Create new FormatOptions for millimeters
                    FormatOptions mmFormatOptions = new FormatOptions(UnitTypeId.Millimeters);

                    // Set the format options for length to millimeters
                    units.SetFormatOptions(SpecTypeId.Length, mmFormatOptions);

                    // Apply the updated units to the document
                    doc.SetUnits(units);

                    tx.Commit();
                    tx.Dispose();
                }
            }
            else if (projectUnits == "feet")
            {
                using (Transaction tx = new Transaction(doc, "Set Project Units to feet"))
                {
                    tx.Start();

                    Units units = doc.GetUnits();
                    FormatOptions cmFormatOptions = new FormatOptions(UnitTypeId.Feet);
                    units.SetFormatOptions(SpecTypeId.Length, cmFormatOptions);
                    doc.SetUnits(units);

                    tx.Commit();
                    tx.Dispose();
                }
            }

            // Replace this line:
            // List<string>[] csvData = ReadCSVFile(filePath);
            // With this line:

            List<string[]> csvData = ReadCSVFile(filePath);

            csvData.RemoveAt(0); // Remove header row

            int rowCount = csvData.Count;

            TaskDialog.Show("CSV Data", $"Number of data rows: {rowCount}");

            // create lists from csvData

            List<string> levelNames = ListFromCSV(csvData, 0);

            List<string> levelElevationsFT = ListFromCSV(csvData, 1); 

            List<string> levelElevationsM = ListFromCSV(csvData, 2);

            TaskDialog.Show("Lists", GetMethod() + "\n" +
                $"Level Names: {string.Join(", ", levelNames)}\n" +
                $"Level Elevations (ft): {string.Join(", ", levelElevationsFT)}\n" +
                $"Level Elevations (m): {string.Join(", ", levelElevationsM)}");

            // convert string lists to double lists for heights

            List<double> levelElevsFT = ListStringToDouble(levelElevationsFT);

            List<double> levelElevsM = ListStringToDouble(levelElevationsM);

            // Create levels FT

            if (projectUnits == "feet")
            {
                using (Transaction tx = new Transaction(doc, "Create Levels in feet"))
                {
                    tx.Start();
                    for (int i = 0; i < levelNames.Count; i++)
                    {
                        string lvlName = levelNames[i];
                        double lvlElev = levelElevsFT[i];
                        Level newLevel = Level.Create(doc, lvlElev);
                        newLevel.Name = lvlName;
                    }
                    tx.Commit();
                    tx.Dispose();
                }
            }
            else if (projectUnits == "mm")
            {
                using (Transaction tx = new Transaction(doc, "Create Levels in mm"))
                {
                    tx.Start();
                    for (int i = 0; i < levelNames.Count; i++)
                    {
                        string lvlName = levelNames[i];
                        double lvlElevM = levelElevsM[i];
                        double lvlElevFT = lvlElevM * 3.2808399; // convert m to feet to create levels
                        Level newLevel = Level.Create(doc, lvlElevFT);
                        newLevel.Name = lvlName;
                    }
                    tx.Commit();
                    tx.Dispose();
                }
            }






            return Result.Succeeded;
        }

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }

        // Replace the ListFromCSV method with the correct implementation.
        // The original method signature and implementation are incorrect for the intended use.
        // The method should return List<string> instead of List<string[]> and accept List<string[]> as input.

        public static List<string> ListFromCSV(List<string[]> csvList, int n)
        {
            var nthElements = csvList
                .Where(sublist => sublist.Length > n)
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
    }
}
