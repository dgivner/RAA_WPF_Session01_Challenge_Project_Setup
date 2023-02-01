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
using System.IO;
using Microsoft.Win32;

#endregion

namespace RAA_WPF_Session01_Challenge_Project_Setup
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
            FilteredElementCollector tblockCollector = new FilteredElementCollector(doc);
            tblockCollector.OfCategory(BuiltInCategory.OST_TitleBlocks);
            ElementId tblockId = tblockCollector.FirstElementId();

            //var fileName = OpenFile();

            //FilteredElementCollector viewTemplateCollector = new FilteredElementCollector(doc);
            //viewTemplateCollector.OfClass(typeof(ViewTemplateApplicationOption));

            // open form
            MyWindow currentWindow = new MyWindow()
            {
                Width = 800,
                Height = 450,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                Topmost = true,
            };

            currentWindow.ShowDialog();

            if (currentWindow.DialogResult == false)
            {
                return Result.Cancelled;
            }

            //do something
            List<string[]> dataList = new List<string[]>();
            string textboxresult = currentWindow.GetTextBoxValue();

            // get form data and do something
            string[] dataArray = System.IO.File.ReadAllLines(textboxresult);
            foreach (string data in dataArray)
            {
                string[] cellString = data.Split(',');
                dataList.Add(cellString);
            }

            dataList.RemoveAt(0);

            bool checkBox1Value = currentWindow.GetCheckbox1();

            string radioButtonValue = currentWindow.GetGroup1();

            //TaskDialog.Show("Test", "text box result is " + textboxresult);

            if (checkBox1Value == true)
            {
                TaskDialog.Show("Test", "Check box 1 was selected");
            }

            TaskDialog.Show("Test", radioButtonValue);


            //go through csv data and do something


            FilteredElementCollector vftCollector = new FilteredElementCollector(doc);
            vftCollector.OfClass(typeof(ViewFamilyType));

            ViewFamilyType planVFT = null;
            ViewFamilyType rcpVFT = null;

            foreach (ViewFamilyType vft in vftCollector)
            {
                if (vft.ViewFamily == ViewFamily.FloorPlan) planVFT = vft;

                if (vft.ViewFamily == ViewFamily.CeilingPlan) rcpVFT = vft;
            }
            using (Transaction tx = new Transaction(doc))
            {

                tx.Start("Create Level");
                foreach (string[] currentArray in dataList)
                {
                    string text = currentArray[0];
                    string number = currentArray[1];

                    double actualNumber;
                    bool convertNumber = double.TryParse(number, out actualNumber);
                    if (convertNumber == false)
                    {
                        continue;
                    }

                    double actualNumber2 = 0;
                    try
                    {
                        actualNumber2 = double.Parse(number);
                    }
                    catch (Exception e)
                    {
                        TaskDialog.Show("Error", "This item in the number column is not a number.");
                    }

                    if (convertNumber == false)
                    {
                        TaskDialog.Show("Error", "This item in the number column is not a number.");
                    }

                    double metricConvert = actualNumber * 3.28084;
                    Level currentLevel = Level.Create(doc, metricConvert);
                    currentLevel.Name = text;

                    ViewFamilyType floorPlanVFT = GetViewFamilyTypeByName(doc, "Floor Plan", ViewFamily.FloorPlan);
                    ViewFamilyType ceilingPlanVFT =
                        GetViewFamilyTypeByName(doc, "Ceiling Plan", ViewFamily.CeilingPlan);

                    ViewPlan plan = ViewPlan.Create(doc, planVFT.Id,currentLevel.Id);
                    ViewPlan ceilingPlan = ViewPlan.Create(doc, rcpVFT.Id, currentLevel.Id);


                }
                //foreach (var level in LevelsList())
                //{
                //    Level newLevel = Level.Create(doc, level.Elevation);
                //    newLevel.Name = level.Name;

                //    ViewPlan newPlanVIew = ViewPlan.Create(doc, planVFT.Id, newLevel.Id);
                //    ViewPlan newCeilingPlan = ViewPlan.Create(doc, rcpVFT.Id, newLevel.Id);

                //    newPlanVIew.Name = newPlanVIew.Name + "Floor Plan";
                //    newCeilingPlan.Name = newCeilingPlan.Name = "RCP";

                    //ViewSheet newSheet = ViewSheet.Create(doc, tblockId);
                    //ViewSheet newCeilingSheet = ViewSheet.Create(doc, tblockId);

                    //XYZ insertPoint = new XYZ(2, 1, 0);
                    //XYZ secondInsertPoint = new XYZ(0, 1, 0);

                    //Viewport newViewport = Viewport.Create(doc, newSheet.Id, newPlanVIew.Id, insertPoint);
                    //Viewport newCeilingViewport = Viewport.Create(doc, newCeilingSheet.Id, newCeilingPlan.Id, secondInsertPoint);
                //}

                //foreach (var sheet in SheetList())
                //{
                //    ViewSheet newSheet = ViewSheet.Create(doc, tblockId);
                //    newSheet.Name = sheet.Name;
                //    newSheet.SheetNumber = sheet.Number;
                //}

                tx.Commit();
                tx.Dispose();

            }
            return Result.Succeeded;
        }
        private static string OpenFile()
        {
            OpenFileDialog selectFile = new OpenFileDialog();
            selectFile.InitialDirectory = "C:\\";
            selectFile.Filter = "CSV Files|*.csv";
            selectFile.Multiselect = false;

            string fileName = "";
            if (selectFile.ShowDialog() == true)
            {
                fileName = selectFile.FileName;
            }

            return fileName;
        }

        //Read Sheets Text file for data
        private static List<Classes.dSheets> SheetList()
        {
            string sheetsFilePath = OpenFile();

            List<Classes.dSheets> sheets = new List<Classes.dSheets>();
            string[] sheetsArray = File.ReadAllLines(sheetsFilePath);
            foreach (var sheetsRowString in sheetsArray)
            {
                string[] sheetsCellString = sheetsRowString.Split(',');
                var sheet = new Classes.dSheets()
                {
                    Number = sheetsCellString[0],
                    Name = sheetsCellString[1]
                };

                sheets.Add(sheet);
            }

            return sheets;
        }

        //Read Level Text file for data
        private static List<Classes.dLevels> LevelsList()
        {
            string levelsFilePath = OpenFile();

            List<Classes.dLevels> levels = new List<Classes.dLevels>();
            string[] levelsArray = File.ReadAllLines(levelsFilePath);
            foreach (var levelsRowString in levelsArray)
            {
                string[] levelsCellString = levelsRowString.Split(',');
                var level = new Classes.dLevels()
                {
                    Name = levelsCellString[0]
                };

                bool didItParse = double.TryParse(levelsCellString[1], out level.Elevation);

                levels.Add(level);
            }

            return levels;
        }

        public static List<View> GetAllViews(Document curDoc)
        {
            FilteredElementCollector allViews = new FilteredElementCollector(curDoc);
            allViews.OfCategory(BuiltInCategory.OST_Views);

            List<View> multiViews = new List<View>();
            foreach (View av in allViews.ToElements())
            {
                multiViews.Add(av);
            }

            return multiViews;
        }

        public static List<View> GetAllViewTemplates(Document curDoc)
        {
            List<View> returnList = new List<View>();
            List<View> viewList = GetAllViews(curDoc);
            foreach (View v in viewList)
            {
                if (v.IsTemplate == true)
                {
                    returnList.Add(v);
                }
            }

            return returnList;
        }

        private View GetViewByName(Document doc, string name)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Views);

            foreach (View currentView in collector)
            {
                if (currentView.Name == name)
                {
                    return currentView;
                }
            }

            return null;
        }

        private ViewFamilyType GetViewFamilyTypeByName(Document doc, string typeName, ViewFamily viewFamily)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewFamilyType));

            foreach (ViewFamilyType currentVFT in collector)
            {
                if (currentVFT.Name == typeName && currentVFT.ViewFamily == viewFamily)
                {
                    return currentVFT;
                }
            }

            return null;
        }

        internal Element GetTitleBlockByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
            foreach (Element currentTblock in collector)
            {
                if (currentTblock.Name == typeName)
                {
                    return currentTblock;
                }
            }

            return null;
        }
        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
    
    internal class ViewStruct
    {
        public string Name;
        public string Discipline;
        public string Level;
        public ViewStruct(string name, string discipline, string level)
        {
            Name = name;
            Discipline = discipline;
            Level = level;
        }
    }
}
