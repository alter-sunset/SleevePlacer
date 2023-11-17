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
                UIDocument uiDocument = uiApp.ActiveUIDocument;
                Document mainDocument = uiDocument.Document;

                FamilySymbol symbol;
                FamilySymbol symbolWall = FamSymbol(mainDocument, "00_Гильза_Стена");
                FamilySymbol symbolFloor = FamSymbol(mainDocument, "00_Гильза_Плита");

                if (symbolWall is null || symbolFloor is null)
                {
                    return Result.Failed;
                }

                List<Document> links = new FilteredElementCollector(mainDocument)
                    .OfClass(typeof(RevitLinkInstance))
                    .Select(l => l as RevitLinkInstance)
                    .Select(l => l.GetLinkDocument())
                    .Where(d => d != null)
                    .ToList();

                ReferenceIntersector referenceIntersector;
                ReferenceIntersector referenceIntersectorWalls = RefIntersector<Wall>(uiDocument);
                ReferenceIntersector referenceIntersectorFloors = RefIntersector<Floor>(uiDocument);

                double offset = 100;

                using (Transaction transaction = new Transaction(mainDocument))
                {
                    transaction.Start("Добавление гильз");

                    foreach (Document linkedDocument in links)
                    {
                        if (linkedDocument == null)
                        {
                            continue;
                        }

                        using (FilteredElementCollector collector = new FilteredElementCollector(linkedDocument))
                        {
                            IQueryable<Element> pipes = collector
                                .OfClass(typeof(Pipe))
                                .Where(d => d != null)
                                .AsQueryable();

                            foreach (Element pipe in pipes)
                            {
                                Curve pipeCurve = GetTheCurve(pipe);

                                if (pipeCurve is null
                                    || !pipeCurve.IsBound)
                                {
                                    continue;
                                }

                                double diameter = pipe
                                    .get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)
                                    .AsDouble() + (offset / 304.8);
                                Line pipeLine = pipeCurve as Line;
                                XYZ origin = pipeLine.GetEndPoint(0);

                                if (IsVertical(pipeCurve))
                                {
                                    referenceIntersector = referenceIntersectorFloors;
                                    symbol = symbolFloor;
                                }
                                else if (IsHorizontal(pipeCurve))
                                {
                                    referenceIntersector = referenceIntersectorWalls;
                                    symbol = symbolWall;
                                }
                                else
                                {
                                    continue;
                                }

                                IterateStructuralElements(mainDocument, referenceIntersector, pipe, pipeCurve, diameter, origin, pipeLine, symbol);
                            }
                        }
                    }

                    transaction.Commit();
                }

                MessageBox.Show($"{DateTime.Now - start}");
                return Result.Succeeded;
            }
        }

        public ReferenceIntersector RefIntersector<T>(UIDocument uiDocument)
        {
            ReferenceIntersector referenceIntersector = new ReferenceIntersector(
                            new ElementClassFilter(typeof(T)),
                            FindReferenceTarget.Element,
                            (View3D)uiDocument.ActiveGraphicalView)
            {
                FindReferencesInRevitLinks = false
            };

            return referenceIntersector;
        }

        public FamilySymbol FamSymbol(Document document, string name)
        {
            FamilySymbol symbol = new FilteredElementCollector(document)
                    .OfClass(typeof(FamilySymbol))
                    .Where(s => s != null)
                    .FirstOrDefault(e => e.Name == name) as FamilySymbol;

            return symbol;
        }

        private Solid GetTheSolid(Element element)
        {
            GeometryElement geometry = element.get_Geometry(new Options());

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
            Curve curve;

            try
            {
                curve = (element.Location as LocationCurve).Curve;
            }
            catch
            {
                return null;
            }
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

        private bool IsHorizontal(Curve curve)
        {
            double startZ = curve.GetEndPoint(0).Z;
            double endZ = curve.GetEndPoint(1).Z;

            if (Math.Round(startZ - endZ) != 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        private bool IsVertical(Curve curve)
        {
            XYZ start = curve.GetEndPoint(0);
            XYZ end = curve.GetEndPoint(1);

            double startX = start.X;
            double endX = end.X;
            double startY = start.Y;
            double endY = end.Y;
            double startZ = start.Z;
            double endZ = end.Z;

            if (Math.Round(startX - endX) == 0
                && Math.Round(startY - endY) == 0
                && Math.Round(startZ - endZ) != 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void IterateStructuralElements(Document mainDocument,
            ReferenceIntersector referenceIntersector,
            Element pipe,
            Curve pipeCurve,
            double diameter,
            XYZ origin,
            Line pipeLine,
            FamilySymbol symbol)
        {
            IEnumerable<ReferenceWithContext> intersections = referenceIntersector
                       .Find(origin, pipeLine.Direction)
                       .Where(x => x.Proximity <= pipeLine.Length)
                       .Distinct(new ReferenceWithContextElementEqualityComparer());

            foreach (ReferenceWithContext intersection in intersections)
            {
                ElementId elementId = intersection.GetReference().ElementId;
                Element element = mainDocument.GetElement(elementId);
                Solid solid = GetTheSolid(element);
                Curve elementCurve = GetTheCurve(element);

                if (elementCurve != null
                    && !IsPerpendicular(pipeCurve, elementCurve))
                {
                    continue;
                }

                SolidCurveIntersection elementPipeIntersection = solid.IntersectWithCurve(pipeCurve, null);

                if (elementPipeIntersection == null
                    || elementPipeIntersection == default
                    || !elementPipeIntersection.IsValidObject
                    || elementPipeIntersection.SegmentCount < 1)
                {
                    continue;
                }

                Curve line = elementPipeIntersection.GetCurveSegment(0);

                if (line == null)
                {
                    continue;
                }

                XYZ center = (line.GetEndPoint(0) + line.GetEndPoint(1)) / 2;

                FamilyInstance insertNew = mainDocument.Create
                    .NewFamilyInstance(
                    center,
                    symbol,
                    element,
                    mainDocument.GetElement(element.LevelId) as Level,
                    StructuralType.NonStructural);

                insertNew.LookupParameter("Диаметр").Set(diameter);
                insertNew.LookupParameter("Id").Set(pipe.UniqueId);
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
