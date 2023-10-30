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
                    .FirstOrDefault(e => e.Name == "00_Гильза") as FamilySymbol;

                List<ElementId> wallIds = new FilteredElementCollector(mainDocument)
                    .OfClass(typeof(Wall))
                    .ToElementIds()
                    .ToList();

                List<Document> links = new FilteredElementCollector(mainDocument)
                    .OfClass(typeof(RevitLinkInstance))
                    .Select(l => l as RevitLinkInstance)
                    .Select(l => l.GetLinkDocument())
                    .ToList();

                double offset = 100;

                foreach (Document linkedDocument in links)
                {
                    foreach (ElementId wallId in wallIds)
                    {
                        if (!(mainDocument.GetElement(wallId) is Wall wall))
                        {
                            continue;
                        }

                        GeometryElement geometry = wall.get_Geometry(new Options());

                        if (geometry is null)
                        {
                            continue;
                        }

                        Solid solid = geometry
                            .Select(e => e as GeometryObject)
                            .Where(e => e != null)
                            .Select(e => e as Solid)
                            .FirstOrDefault(e => e != null);

                        List<Element> pipes = new FilteredElementCollector(linkedDocument)
                            .OfCategory(BuiltInCategory.OST_PipeCurves)
                            .WhereElementIsNotElementType()
                            .WherePasses(new ElementIntersectsSolidFilter(solid))
                            .ToList();

                        foreach (Element pipe in pipes)
                        {
                            double diameter = pipe
                                .get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)
                                .AsDouble() + (offset / 304.8);

                            Curve pipeCurve = (pipe.Location as LocationCurve).Curve;

                            Curve line = solid
                                .IntersectWithCurve(pipeCurve, new SolidCurveIntersectionOptions())
                                .GetCurveSegment(0);

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

                                transaction.Commit();
                            }
                        }
                    }
                }

                return Result.Succeeded;
            }
        }
    }
}
