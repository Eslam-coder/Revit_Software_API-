using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SelectElement
{
    [Transaction(TransactionMode.Manual)]
    public class SelectElement : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get Document and UI Document 
            UIDocument UIDoc = commandData.Application.ActiveUIDocument;
            Document doc = commandData.Application.ActiveUIDocument.Document;

            #region Create Message appear when you use this AddIn

            //TaskDialog.Show("Message", "Helloooooo :)");
            //return Result.Succeeded;
            #endregion

            #region Select Element
            /*try
            {
                Reference SelectedElement = UIDoc.Selection.PickObject(ObjectType.Element);
                Element SelectedElementData = doc.GetElement(SelectedElement.ElementId);
                ElementId SelectedElementTypeId = SelectedElementData.GetTypeId();
                ElementType SelectedElementTypeData = doc.GetElement(SelectedElementTypeId) as ElementType;
                
                TaskDialog.Show("Element Information\n", 
                                "Element ID: " + SelectedElementData.Id + "\n" +
                                "Element Name: " + SelectedElementData.Name + "\n" +
                                "Element Category: " + SelectedElementTypeData.Name + "\n" +
                                "Element Type Name: " + SelectedElementTypeData.FamilyName + "\n");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }*/

            #endregion

            #region Select Elements
            /*try
            {
                IList<Reference> SelectedElements = UIDoc.Selection.PickObjects(ObjectType.Element);
                string ElementsInformation = " ";
                foreach (var item in SelectedElements)
                {
                    Element SelectedElementData = doc.GetElement(item.ElementId);
                    ElementId SelectedElementTypeId = SelectedElementData.GetTypeId();
                    ElementType SelectedElementTypeData = doc.GetElement(SelectedElementTypeId) as ElementType;
                    ElementsInformation += "Element ID: " + SelectedElementData.Id + "\n" +
                                          "Element Name: " + SelectedElementData.Name + "\n" +
                                          "Element Category: " + SelectedElementTypeData.Name + "\n" +
                                          "Element Type Name: " + SelectedElementTypeData.FamilyName + "\n";

                }
                TaskDialog.Show("Element Information\n", ElementsInformation);
                return Result.Succeeded;
             }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }*/

            #endregion

            #region Select Many Elements And Filter Them According To Category For Example (Columns)
            /*try
            {
                FilteredElementCollector collecter = new FilteredElementCollector(doc);
                ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_StructuralColumns);
                
                //Select Columns had been drawn only 
                IList<Element> SelectedColumnsAccordingToFilteration = collecter.OfClass(typeof(FamilyInstance)).WherePasses(filter).ToElements();
                
                //Select Columns had been drawn and hadn't been drawn 
                //IList<Element> SelectedColumnsAccordingToFilteration = collecter.OfClass(typeof(FamilySymbol)).WherePasses(filter).ToElements();
                
                string ElementsInformation =" ";
                foreach (var item in SelectedColumnsAccordingToFilteration)
                {
                    ElementsInformation +="Element Name: "+item.Name+Environment.NewLine;

                }
                TaskDialog.Show("Element Information\n","Number of Structural Columns = "+ SelectedColumnsAccordingToFilteration.Count() + Environment.NewLine+ElementsInformation);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }*/
            #endregion

            #region Select Many Elements And Filter Them According To Category For Example (Sheets)
            /*try
            {
                FilteredElementCollector collecter = new FilteredElementCollector(doc);
                ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_Sheets);

                //Select Sheets had been drawn only 
                IList<Element> SelectedSheetsAccordingToFilteration = collecter.OfClass(typeof(View)).WherePasses(filter).ToElements();

                string ElementsInformation = " ";
                foreach (var item in SelectedSheetsAccordingToFilteration)
                {
                    ElementsInformation += "Element Name: " + item.Name + Environment.NewLine;

                }
                TaskDialog.Show("Element Information\n", "Number of Structural Columns = " + SelectedSheetsAccordingToFilteration.Count() + Environment.NewLine + ElementsInformation);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }*/
            #endregion

            #region Select Many Elements And Filter Them According To Two Category For Example (Columns And Beams)
            /*try
            {
                FilteredElementCollector collecter = new FilteredElementCollector(doc);
                ElementCategoryFilter StrColumns = new ElementCategoryFilter(BuiltInCategory.OST_StructuralColumns);
                ElementCategoryFilter StrBeams = new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming);

                LogicalOrFilter ColumnsAndBeamsFilter = new LogicalOrFilter(StrColumns, StrBeams);

                //Select Sheets had been drawn only 
                IList<Element> ColumnsAndBeams = collecter.OfClass(typeof(FamilyInstance)).WherePasses(ColumnsAndBeamsFilter).ToElements();

                string ElementsInformation = " ";
                foreach (var item in ColumnsAndBeams)
                {
                    ElementsInformation += "Element Name: " + item.Name + Environment.NewLine;

                }
                TaskDialog.Show("Element Information\n", "Number of Structural Columns and Structural Beams= " + ColumnsAndBeams.Count() + Environment.NewLine + ElementsInformation);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }*/
            #endregion

            #region Select Many Elements And Filter Them According To Category For Example (Doors)
            /*try
            {
                FilteredElementCollector collecter = new FilteredElementCollector(doc);
                ElementCategoryFilter Doors = new ElementCategoryFilter(BuiltInCategory.OST_Doors);
                ElementCategoryFilter Windows = new ElementCategoryFilter(BuiltInCategory.OST_Windows);

                LogicalOrFilter DoorsAndWindowsFilter = new LogicalOrFilter(Doors, Windows);

                //Select Sheets had been drawn only 
                IList<Element> DoorsAndWindows = collecter.OfClass(typeof(FamilyInstance)).WherePasses(DoorsAndWindowsFilter).ToElements();

                string ElementsInformation = " ";
                foreach (var item in DoorsAndWindows)
                {
                    ElementsInformation += "Element Name: " + item.Name + Environment.NewLine;

                }
                TaskDialog.Show("Element Information\n", "Number of Doors and Windows= " + DoorsAndWindows.Count() + Environment.NewLine + ElementsInformation);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }*/
            #endregion

            #region Change Parameters of Any Element For Example (TopOffest,MoveWithGrid,Level,.....)
            /*try
            {
                IList<Reference> SelectedElements = UIDoc.Selection.PickObjects(ObjectType.Element);

                FilteredElementCollector collecter = new FilteredElementCollector(doc);
                ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_Levels);
                IList<Level> SelectedLevelsAccordingToFilteration = collecter.OfClass(typeof(Level)).WherePasses(filter).ToElements().Cast<Level>().ToList();
                
                ElementId roofId = null;
                //Get Roof Id 
                foreach (var item in SelectedLevelsAccordingToFilteration)
                {
                    if (item.Name == "Roof")
                    {
                        roofId = item.Id;
                    }
                }

                //Get All Materials In Project
                FilteredElementCollector MaterialsInProject = new FilteredElementCollector(doc).OfClass(typeof(Material));
                
                ElementId MaterialWoodId = null;

                //Get Wood Material Id 
                foreach (var item in MaterialsInProject)
                {
                    if(item.Name == "Wood - stained")
                    {
                        MaterialWoodId = item.Id; ;
                    }
                }
                foreach (var item in SelectedElements)
                {
                    Element SelectedElementData = doc.GetElement(item.ElementId);

                    //Start Transaction 
                    using (Transaction trans = new Transaction(doc, "Edit Parameter"))
                    {
                        trans.Start();

                        //Change Comment
                        Parameter ElementComment = SelectedElementData.LookupParameter("Comments");
                        ElementComment.Set("This Comment From Revit");

                        //Change MoveWithGrid
                        Parameter ElementMoveWithGrid = SelectedElementData.LookupParameter("Moves With Grids");
                        ElementComment.Set(0);

                        //Change TopOffest
                        Parameter ElementTopOffest = SelectedElementData.LookupParameter("Top Offest");
                        ElementComment.Set(10);

                        //Change Column Height from Base Level To Top Level
                        Parameter ElementTopLevel = SelectedElementData.LookupParameter("Top Level");
                        ElementComment.Set(roofId);

                        //Change Column Material
                        Parameter Material = SelectedElementData.LookupParameter("Structural Material");
                        ElementComment.Set(MaterialWoodId);

                        trans.Commit();
                    }
                }
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }*/
            #endregion

            #region Excel

            //Appear Dialog , you can choose Excel file from this dialog
            try
            {
                OpenFileDialog DialogRead = new OpenFileDialog();
                DialogRead.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                DialogRead.ShowDialog();
                string ExcelFileName = DialogRead.FileName;
                using (ExcelPackage ExcelPack = new ExcelPackage(new FileInfo(ExcelFileName)))
                {
                    //Number between WorkSheet is the number of sheet in file excel
                    ExcelWorksheet Worksheets = ExcelPack.Workbook.Worksheets[1];
                    //Get Value from Cell in row = 1 , column=1
                    //var CellValue = Worksheets.Cells[1, 1].Value;
                    //Get Value from Cell in all rows , column=1
                    string CellsValues = null;
                    for (int row = 1; row < 99999; row++)
                    {
                        var CellValue = Worksheets.Cells[row, 1].Value;
                       
                        if(CellValue==null)
                        {
                            break;
                        }
                        CellsValues += CellValue + Environment.NewLine;
                    }
                    
                    //Appear Value of cell from excel in Revit in dialog box 
                    TaskDialog.Show("Cell Value:", CellsValues.ToString());

                }
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
            #endregion
        }
    }
}
