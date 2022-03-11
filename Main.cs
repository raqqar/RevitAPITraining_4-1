using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPITraining_4
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            string wallInfo = string.Empty;

            var Pipes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeCurves)
                .ToList();

            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Pipes.xlsx");

            FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write);

            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Лист! с трубами");

            int rowIndex = 0;
            foreach (var pipe in Pipes)
            {
                Pipe p = doc.GetElement(pipe.Id) as Pipe;
                if (p != null)
                {
                    Parameter p1 = p.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                    Parameter p2 = p.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                    Parameter p3 = p.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM);
                    double size1 = UnitUtils.ConvertFromInternalUnits(p1.AsDouble(), UnitTypeId.Millimeters);
                    double size2 = UnitUtils.ConvertFromInternalUnits(p2.AsDouble(), UnitTypeId.Millimeters);
                    double size3 = UnitUtils.ConvertFromInternalUnits(p3.AsDouble(), UnitTypeId.Millimeters);
                    sheet.SetCellValue(rowIndex, 0, p.Id.GetHashCode());
                    sheet.SetCellValue(rowIndex, 1, p.Name);
                    sheet.SetCellValue(rowIndex, 2, size1);
                    sheet.SetCellValue(rowIndex, 3, size2);
                    sheet.SetCellValue(rowIndex, 4, size3);
                    rowIndex++;

                }
            }
            workbook.Write(stream);
            workbook.Close();

            return Result.Succeeded;
        }
    }

  
}
