using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace CreateModelPlugin_Part2
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class CreateModelPlugin_Part2 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            double width = UnitUtils.ConvertToInternalUnits(10000, UnitTypeId.Millimeters);
            double depth = UnitUtils.ConvertToInternalUnits(5000, UnitTypeId.Millimeters);

            List<Level> levels = GetLevelList(doc);
            Level level1 = GetLevelByName(levels, "Уровень 1");
            Level level2 = GetLevelByName(levels, "Уровень 2");
            List<XYZ> points = GetPointsByWidthAndDepth(width, depth);

            List<Wall> wallList = CreateWall(doc, points, level1, level2);

            AddDoor(doc, level1, wallList[0]);
            AddWindow(doc, level1, wallList[1]);
            AddWindow(doc, level1, wallList[2]);
            AddWindow(doc, level1, wallList[3]);
            return Result.Succeeded;
        }

        private void AddWindow(Document doc, Level level, Wall wall)
        {
            FamilySymbol windowType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0915 x 1830 мм"))
                .Where(x => x.FamilyName.Equals("Фиксированные"))
                .FirstOrDefault();
            if (windowType != null)
            {
                using (var ts = new Transaction(doc, "create doors"))
                {
                    ts.Start();

                    LocationCurve hostCurve = wall.Location as LocationCurve;
                    XYZ midPoint = (hostCurve.Curve.GetEndPoint(0) + hostCurve.Curve.GetEndPoint(1)) / 2;

                    if (!windowType.IsActive)
                    {
                        windowType.Activate();
                    }
                    FamilyInstance window = doc.Create.NewFamilyInstance(midPoint, windowType, wall, level, StructuralType.NonStructural);

                    double windowLevelOffset = UnitUtils.ConvertToInternalUnits(800, UnitTypeId.Millimeters); //смещение созданного окна от уровня
                    window.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM).Set(windowLevelOffset);
                    ts.Commit();
                }
            }
            else
            {
                TaskDialog.Show("Info", "Не найден типоразмер");
            }
        }

        public void AddDoor(Document doc, Level level, Wall wall)
        {
            FamilySymbol doorType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0915 x 2032 мм"))
                .Where(x => x.FamilyName.Equals("Одиночные-Щитовые"))
                .FirstOrDefault();
            if (doorType != null)
            {
                using (var ts = new Transaction(doc, "create doors"))
                {
                    ts.Start();

                    LocationCurve hostCurve = wall.Location as LocationCurve;
                    XYZ midPoint = (hostCurve.Curve.GetEndPoint(0) + hostCurve.Curve.GetEndPoint(1)) / 2;

                    if (!doorType.IsActive)
                    {
                        doorType.Activate();
                    }
                    doc.Create.NewFamilyInstance(midPoint, doorType, wall, level, StructuralType.NonStructural);

                    ts.Commit();
                }
            }
            else
            {
                TaskDialog.Show("Info", "Не найден типоразмер");
            }
        }

        public List<Level> GetLevelList(Document doc)
        {
            List<Level> listLevel = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .OfType<Level>()
                .ToList();
            return listLevel;
        }
        public Level GetLevelByName(List<Level> levelList, string levelName)
        {
            return levelList
                .Where(x => x.Name.Equals(levelName))
                .FirstOrDefault();

        }
        public List<Wall> CreateWall(Document doc, List<XYZ> points, Level bottomLevel, Level topLevel)
        {

            using (var ts = new Transaction(doc, "Create walls"))
            {
                List<Wall> walls = new List<Wall>();
                ts.Start();
                for (int i = 0; i < 4; i++)
                {
                    Line line = Line.CreateBound(points[i], points[i + 1]);
                    Wall wall = Wall.Create(doc, line, bottomLevel.Id, false);
                    walls.Add(wall);
                    wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(topLevel.Id);

                }
                ts.Commit();
                return walls;
            }
        }
        public List<XYZ> GetPointsByWidthAndDepth(double width, double depth)
        {
            double dx = width / 2;
            double dy = depth / 2;

            List<XYZ> points = new List<XYZ>();
            points.Add(new XYZ(-dx, -dy, 0));
            points.Add(new XYZ(dx, -dy, 0));
            points.Add(new XYZ(dx, dy, 0));
            points.Add(new XYZ(-dx, dy, 0));
            points.Add(new XYZ(-dx, -dy, 0));
            return points;
        }
    }
}
