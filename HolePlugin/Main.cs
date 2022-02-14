using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HolePlugin
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document;
            Document ovDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ОВ")).FirstOrDefault();
            if (ovDoc==null)
            {
                TaskDialog.Show("Ошибка", "Ошибка");
                return Result.Succeeded;
            }
            



            FamilySymbol familySymbol = new FilteredElementCollector(arDoc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_GenericModel).OfType<FamilySymbol>().Where(x => x.FamilyName.Equals("Отверстия")).FirstOrDefault();

            if (familySymbol == null)
            {
                TaskDialog.Show("Ошибка", "Не найдено семейство");
                return Result.Succeeded;
            }

            List<Duct> ducts = new FilteredElementCollector(ovDoc).OfClass(typeof(Duct)).OfType<Duct>().ToList();
            List<Pipe> pipes = new FilteredElementCollector(ovDoc).OfClass(typeof(Pipe)).OfType<Pipe>().ToList();

            View3D view3d = new FilteredElementCollector(arDoc).OfClass(typeof(View3D)).OfType<View3D>().Where(x => !x.IsTemplate).FirstOrDefault();
            if (view3d == null)
            {
                TaskDialog.Show("Ошибка", "Ошибка");
                return Result.Succeeded;
            }

            ReferenceIntersector referenceIn = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3d);
            using (var ts = new Transaction(arDoc, "ts start"))
            {

                ts.Start();
                if (!familySymbol.IsActive)
                {
                    familySymbol.Activate();
                }
                ts.Commit();
            }
                using (var ts= new Transaction(arDoc, "ts start"))
            {

                ts.Start();
                

                foreach (Duct d in ducts)
                {
                    Line curve = (d.Location as LocationCurve).Curve as Line;

                    XYZ point = curve.GetEndPoint(0);

                    XYZ direction = curve.Direction;

                    List<ReferenceWithContext> refа = referenceIn.Find(point, direction).Where(x => x.Proximity <= curve.Length)
                     .Distinct(new ReferenceWithContextElementEqualityComparer())
                     .ToList();
                    foreach (ReferenceWithContext refer in refа)
                    {
                        double proximity = refer.Proximity;
                        Reference reference = refer.GetReference();
                        Wall wall = arDoc.GetElement(reference.ElementId) as Wall;

                        Level level = arDoc.GetElement(wall.LevelId) as Level;
                        XYZ pointHole = point + (direction * proximity);
                        FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                        Parameter width = hole.LookupParameter("ADSK_Размер_Ширина");
                        Parameter height = hole.LookupParameter("ADSK_Размер_Высота");
                        width.Set(d.Width+1);
                        height.Set(d.Height+1);
                    }
                }
                foreach (Duct d in ducts)
                {
                    Line curve = (d.Location as LocationCurve).Curve as Line;

                    XYZ point = curve.GetEndPoint(0);

                    XYZ direction = curve.Direction;

                    List<ReferenceWithContext> refа = referenceIn.Find(point, direction).Where(x => x.Proximity <= curve.Length)
                     .Distinct(new ReferenceWithContextElementEqualityComparer())
                     .ToList();
                    foreach (ReferenceWithContext refer in refа)
                    {
                        double proximity = refer.Proximity;
                        Reference reference = refer.GetReference();
                        Wall wall = arDoc.GetElement(reference.ElementId) as Wall;

                        Level level = arDoc.GetElement(wall.LevelId) as Level;
                        XYZ pointHole = point + (direction * proximity);
                        FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                        Parameter width = hole.LookupParameter("ADSK_Размер_Ширина");
                        Parameter height = hole.LookupParameter("ADSK_Размер_Высота");
                        width.Set(d.Width + 1);
                        height.Set(d.Height + 1);
                    }
                }

                foreach (Pipe  p in pipes)
                {
                    Line curve = (p.Location as LocationCurve).Curve as Line;

                    XYZ point = curve.GetEndPoint(0);

                    XYZ direction = curve.Direction;

                    List<ReferenceWithContext> refа = referenceIn.Find(point, direction).Where(x => x.Proximity <= curve.Length)
                     .Distinct(new ReferenceWithContextElementEqualityComparer())
                     .ToList();
                    foreach (ReferenceWithContext refer in refа)
                    {
                        double proximity = refer.Proximity;
                        Reference reference = refer.GetReference();
                        Wall wall = arDoc.GetElement(reference.ElementId) as Wall;

                        Level level = arDoc.GetElement(wall.LevelId) as Level;
                        XYZ pointHole = point + (direction * proximity);
                        FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                        Parameter width = hole.LookupParameter("ADSK_Размер_Ширина");
                        Parameter height = hole.LookupParameter("ADSK_Размер_Высота");
                        width.Set(p.Diameter + 1);
                        height.Set(p.Diameter + 1);
                    }

                }
                ts.Commit();
                return Result.Succeeded;
                
            }
        }
        public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
        {
            public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
            {
                if (ReferenceEquals(x, y)) return true;
                if (ReferenceEquals(null, x)) return false;
                if (ReferenceEquals(null, y)) return false;

                var xReference = x.GetReference();

                var yReference = y.GetReference();

                return xReference.LinkedElementId == yReference.LinkedElementId
                           && xReference.ElementId == yReference.ElementId;
            }

            public int GetHashCode(ReferenceWithContext obj)
            {
                var reference = obj.GetReference();

                unchecked
                {
                    return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
                }
            }
        }
    }
}
