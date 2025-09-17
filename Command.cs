#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
                }
            }


            return Result.Succeeded;
        }

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
