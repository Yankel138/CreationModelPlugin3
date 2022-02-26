using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreationModelPlugin
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class CreationModel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            List<Level> listLevel = GetLevels(doc);
            Level level1 = GetLevel(listLevel, "Level 1");
            Level level2 = GetLevel(listLevel, "Level 2");

            double width = UnitUtils.ConvertToInternalUnits(10000, UnitTypeId.Millimeters);
            double depth = UnitUtils.ConvertToInternalUnits(5000, UnitTypeId.Millimeters);

            List<Wall> walls = AddWalls(doc, width, depth, level1, level2);
            AddDoor(doc, level1, walls[0]);
            AddWindows(doc, level1, walls);
            
            return Result.Succeeded;
        }

        private void AddWindows(Document doc, Level level1, List<Wall> walls)
        {
            FamilySymbol windowType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("700 x 1200mm"))
                .Where(x => x.FamilyName.Equals("M_Window-Double-Hung"))
                .FirstOrDefault();



            Transaction ts = new Transaction(doc, "Создание окна");
            ts.Start();

            if (!windowType.IsActive)
                windowType.Activate();
            for (int i = 1; i < 4; i++)
            {
                Wall wall = walls[i];
                XYZ point = GetElementCenter(wall);
                doc.Create.NewFamilyInstance(point, windowType, wall, level1, StructuralType.NonStructural);
            }

            ts.Commit();
        }

        private void AddDoor(Document doc, Level level, Wall wall)
        {

            FamilySymbol doorType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0762 x 2134mm"))
                .Where(x => x.FamilyName.Equals("M_Single-Flush"))
                .FirstOrDefault();

            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ point1 = hostCurve.Curve.GetEndPoint(0);
            XYZ point2 = hostCurve.Curve.GetEndPoint(1);
            XYZ point = (point1 + point2) / 2;

            Transaction ts = new Transaction(doc, "Создание двери");
            ts.Start();

            if (!doorType.IsActive)
                doorType.Activate();

            doc.Create.NewFamilyInstance(point, doorType, wall, level, StructuralType.NonStructural);
            ts.Commit();
        }

        public List<Wall> AddWalls(Document doc, double width, double depth, Level level1, Level level2)
        {
            double dx = width / 2;
            double dy = depth / 2;

            List<XYZ> points = new List<XYZ>();
            points.Add(new XYZ(-dx, -dy, 0));
            points.Add(new XYZ(dx, -dy, 0));
            points.Add(new XYZ(dx, dy, 0));
            points.Add(new XYZ(-dx, dy, 0));
            points.Add(new XYZ(-dx, -dy, 0));

            List<Wall> walls = new List<Wall>();

            Transaction ts = new Transaction(doc, "Построение стен");
            ts.Start();
            for (int i = 0; i < 4; i++)
            {
                Line line = Line.CreateBound(points[i], points[i + 1]);
                Wall wall = Wall.Create(doc, line, level1.Id, false);
                walls.Add(wall);
                wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(level2.Id);
            }
            ts.Commit();

            return walls;
        }

        public List<Level> GetLevels(Document doc)
        {
            List<Level> listLevel = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .OfType<Level>()
                .ToList();

            return listLevel;
        }

        public Level GetLevel(List<Level> listLevel, string levelName)
        {
            Level level = listLevel
                .Where(x => x.Name.Equals(levelName))
                .OfType<Level>()
                .FirstOrDefault();

            return level;
        }

        public XYZ GetElementCenter(Element element)
        {
            BoundingBoxXYZ bounding = element.get_BoundingBox(null);
            return (bounding.Max + bounding.Min) / 2;
        }
    }
}
