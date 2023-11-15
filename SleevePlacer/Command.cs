using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using System.Windows.Forms;
using Application = Autodesk.Revit.ApplicationServices.Application;
using Autodesk.Revit.DB.Plumbing;

namespace SleevePlacer
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            using (UIApplication uiApp = commandData.Application)
            {
                DateTime start = DateTime.Now;

                Application application = uiApp.Application;
                Document mainDocument = uiApp.ActiveUIDocument.Document;

                FamilySymbol symbol = new FilteredElementCollector(mainDocument)
                    .OfClass(typeof(FamilySymbol))
                    .Where(s => s != null)
                    .FirstOrDefault(e => e.Name == "00_Гильза") as FamilySymbol;

                List<ElementId> wallIds = new FilteredElementCollector(mainDocument)
                    .OfClass(typeof(Wall))
                    .ToElementIds()
                    .Where(w => w != null)
                    .ToList();

                List<Document> links = new FilteredElementCollector(mainDocument)
                    .OfClass(typeof(RevitLinkInstance))
                    .Select(l => l as RevitLinkInstance)
                    .Select(l => l.GetLinkDocument())
                    .Where(d => d != null)
                    .ToList();

                Dictionary<Wall, List<Element>> wallPipesDict = new Dictionary<Wall, List<Element>>();
                Dictionary<Element, List<Wall>> pipeWallsDict = new Dictionary<Element, List<Wall>>();

                double offset = 100;

                foreach (Document linkedDocument in links)
                {
                    if (linkedDocument == null)
                    {
                        continue;
                    }

                    IteratePipesAndWalls(mainDocument, linkedDocument, offset, symbol);
                }

                MessageBox.Show($"{DateTime.Now - start}");
                return Result.Succeeded;
            }
        }
        private void IteratePipesAndWalls(Document mainDocument, Document document, double offset, FamilySymbol symbol)
        {
            using (FilteredElementCollector collector = new FilteredElementCollector(document))
            {
                FilteredElementCollector pipes = collector.OfClass(typeof(Pipe));

                ReferenceIntersector referenceIntersector = new ReferenceIntersector(
                    new ElementClassFilter(typeof(Wall)),
                    FindReferenceTarget.Element,
                    (View3D)mainDocument.ActiveView)
                {
                    FindReferencesInRevitLinks = true
                };

                foreach (Element pipe in pipes)
                {
                    Curve pipeCurve = GetTheCurve(pipe);

                    if (pipeCurve == null
                        || !pipeCurve.IsBound
                        || IsAxisZ(pipeCurve))
                    {
                        continue;
                    }

                    double diameter = pipe
                        .get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)
                        .AsDouble() + (offset / 304.8);
                    Line pipeLine = pipeCurve as Line;

                    XYZ origin = pipeLine.GetEndPoint(0);

                    List<ReferenceWithContext> intersections = referenceIntersector
                        .Find(origin, pipeLine.Direction)
                        .Where(x => x.Proximity <= pipeLine.Length)
                        .Distinct(new ReferenceWithContextElementEqualityComparer())
                        .ToList();

                    foreach (ReferenceWithContext intersection in intersections)
                    {
                        ElementId wallId = intersection.GetReference().ElementId;
                        Wall wall = mainDocument.GetElement(wallId) as Wall;
                        Curve wallCurve = GetTheCurve(wall);
                        Solid solid = GetTheSolid(wall);

                        if (!IsPerpendicular(pipeCurve, wallCurve))
                        {
                            continue;
                        }

                        SolidCurveIntersection wallPipeIntersection;
                        wallPipeIntersection = solid.IntersectWithCurve(pipeCurve, null);
                        //try
                        //{
                        //    wallPipeIntersection = solid.IntersectWithCurve(pipeCurve, null);
                        //}
                        //catch
                        //{
                        //    continue;
                        //}

                        if (wallPipeIntersection == null
                            || wallPipeIntersection == default
                            || !wallPipeIntersection.IsValidObject
                            || wallPipeIntersection.SegmentCount < 1)
                        {
                            continue;
                        }

                        Curve line = wallPipeIntersection.GetCurveSegment(0);

                        if (line == null)
                        {
                            continue;
                        }

                        XYZ startPoint = line.GetEndPoint(0);
                        XYZ endPoint = line.GetEndPoint(1);

                        XYZ center = (startPoint + endPoint) / 2;

                        using (Transaction transaction = new Transaction(mainDocument))
                        {
                            transaction.Start("Добавление гильзы");

                            FamilyInstance insertNew = mainDocument.Create
                                .NewFamilyInstance(
                                center,
                                symbol,
                                wall,
                                mainDocument.GetElement(wall.LevelId) as Level,
                                StructuralType.NonStructural);

                            insertNew.LookupParameter("Диаметр").Set(diameter);
                            insertNew.LookupParameter("Id").Set(pipe.UniqueId);

                            transaction.Commit();
                        }
                    }
                }
            }
        }

        private Solid GetTheSolid(Wall wall)
        {
            GeometryElement geometry = wall.get_Geometry(new Options());

            if (geometry is null)
            {
                return null;
            }

            Solid solid = geometry
                .Select(e => e as GeometryObject)
                .Where(e => e != null)
                .Select(e => e as Solid)
                .FirstOrDefault(e => e != null);

            return solid;
        }

        private Curve GetTheCurve(Element element)
        {
            Curve curve = (element.Location as LocationCurve).Curve;

            return curve;
        }

        private bool IsPerpendicular(Curve curve1, Curve curve2)
        {
            XYZ direction1 = (curve1 as Line).Direction;
            XYZ direction2 = (curve2 as Line).Direction;

            double angle = direction1.AngleTo(direction2);

            if (Math.Round(angle, 4) == Math.Round(Math.PI / 2, 4))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private bool IsAxisZ(Curve curve)
        {
            double startZ = curve.GetEndPoint(0).Z;
            double endZ = curve.GetEndPoint(1).Z;

            if (Math.Round(startZ - endZ) != 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }

    public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
    {
        public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
        {
            if (ReferenceEquals(x, y)) return true;
            if (x is null || y is null) return false;

            Reference xReference = x.GetReference();

            Reference yReference = y.GetReference();

            return xReference.LinkedElementId == yReference.LinkedElementId
                       && xReference.ElementId == yReference.ElementId;
        }

        public int GetHashCode(ReferenceWithContext obj)
        {
            Reference reference = obj.GetReference();

            unchecked
            {
                return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
            }
        }
    }
}
