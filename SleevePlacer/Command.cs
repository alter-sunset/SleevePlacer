using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Events;
using System.Windows.Forms;
using Application = Autodesk.Revit.ApplicationServices.Application;

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

                double offset = 100;

                foreach (Document linkedDocument in links)
                {
                    if (linkedDocument == null)
                    {
                        continue;
                    }

                    PlaceTheSleeve(mainDocument, linkedDocument, wallIds, symbol, offset);
                }

                return Result.Succeeded;
            }
        }

        private void PlaceTheSleeve(Document mainDocument,
            Document linkedDocument,
            List<ElementId> wallIds,
            FamilySymbol symbol,
            double offset)
        {
            foreach (ElementId wallId in wallIds)
            {
                if (!(mainDocument.GetElement(wallId) is Wall wall))
                {
                    continue;
                }

                GeometryElement geometry = wall.get_Geometry(new Options());

                BoundingBoxIntersectsFilter bbFilter = GetOutlineFromWall(mainDocument, wall);

                if (geometry is null)
                {
                    continue;
                }

                Curve wallCurve = (wall.Location as LocationCurve).Curve;

                Solid solid = geometry
                    .Select(e => e as GeometryObject)
                    .Where(e => e != null)
                    .Select(e => e as Solid)
                    .FirstOrDefault(e => e != null);

                if (solid == null || solid == default)
                {
                    continue;
                }

                ElementIntersectsSolidFilter filter = new ElementIntersectsSolidFilter(solid);

                List<Element> pipes = new FilteredElementCollector(linkedDocument)
                    .OfCategory(BuiltInCategory.OST_PipeCurves)
                    .WhereElementIsNotElementType()
                    .WherePasses(bbFilter)
                    .WherePasses(filter)
                    .Where(p => p != null)
                    .ToList();

                bbFilter.Dispose();
                filter.Dispose();

                foreach (Element pipe in pipes)
                {
                    double diameter = pipe
                        .get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)
                        .AsDouble() + (offset / 304.8);

                    Curve pipeCurve = (pipe.Location as LocationCurve).Curve;

                    if (pipeCurve == null
                        || !pipeCurve.IsBound
                        || IsAxisZ(pipeCurve)
                        || !IsPerpendicular(pipeCurve, wallCurve))
                    {
                        continue;
                    }

                    //EXCEPTION MAFAKA
                    //reversion of curve doesn't work
                    //It seems that catching an exception causes all algo go to shit with a need for rerun
                    SolidCurveIntersection intersection;
                    try
                    {
                        intersection = solid.IntersectWithCurve(pipeCurve, null);
                    }
                    catch
                    {
                        continue;
                    }

                    //intersection = solid.IntersectWithCurve(pipeCurve, options);

                    if (intersection == null
                        || intersection == default
                        || !intersection.IsValidObject
                        || intersection.SegmentCount < 1)
                    {
                        continue;
                    }

                    Curve line = intersection.GetCurveSegment(0);

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

                solid.Dispose();
            }
        }

        private BoundingBoxIntersectsFilter GetOutlineFromWall(Document document, Wall wall)
        {
            BoundingBoxXYZ bb = wall.get_BoundingBox(document.ActiveView);

            Outline outline = new Outline(bb.Min, bb.Max);

            BoundingBoxIntersectsFilter bbFilter = new BoundingBoxIntersectsFilter(outline);

            return bbFilter;
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
}
