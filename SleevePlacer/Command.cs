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
                FamilySymbol symbolWall = Utils.GetFamilySymbol(mainDocument, "00_Гильза_Стена");
                FamilySymbol symbolFloor = Utils.GetFamilySymbol(mainDocument, "00_Гильза_Плита");

                if (symbolWall is null || symbolFloor is null)
                {
                    MessageBox.Show("Требуемые семейства гильз не обнаружены");
                    return Result.Failed;
                }

                List<Document> links = new FilteredElementCollector(mainDocument)
                    .OfClass(typeof(RevitLinkInstance))
                    .Select(l => l as RevitLinkInstance)
                    .Select(l => l.GetLinkDocument())
                    .Where(d => d != null)
                    .ToList();

                ReferenceIntersector referenceIntersector;
                ReferenceIntersector referenceIntersectorWalls = Utils.ReferenceIntersector<Wall>(uiDocument);
                ReferenceIntersector referenceIntersectorFloors = Utils.ReferenceIntersector<Floor>(uiDocument);

                double offset = 100;

                using (Transaction transaction = new Transaction(mainDocument))
                {
                    transaction.Start("Добавление гильз");

                    symbolWall.Activate();
                    symbolFloor.Activate();

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
                                Curve pipeCurve = Utils.GetTheCurve(pipe);
                                if (pipeCurve is null
                                    || !pipeCurve.IsBound)
                                {
                                    continue;
                                }

                                //engineering department said that this method would only create problems
                                //so no need for this
                                //CheckExistingSleeves(mainDocument, pipe, pipeCurve);

                                if (Utils.IsCurveVertical(pipeCurve))
                                {
                                    referenceIntersector = referenceIntersectorFloors;
                                    symbol = symbolFloor;
                                }
                                else if (Utils.IsCurveHorizontal(pipeCurve))
                                {
                                    referenceIntersector = referenceIntersectorWalls;
                                    symbol = symbolWall;
                                }
                                else
                                {
                                    continue;
                                }

                                try
                                {
                                    IterateStructuralElements(mainDocument, referenceIntersector, pipe, pipeCurve, offset, symbol);
                                }
                                catch (Exception e)
                                {
                                    if (e is NullReferenceException)
                                    {
                                        MessageBox.Show("Проверьте параметры семейств.");
                                        return Result.Failed;
                                    }
                                }
                            }
                        }
                    }

                    transaction.Commit();
                }

                MessageBox.Show($"{DateTime.Now - start}");
                return Result.Succeeded;
            }
        }

        private void IterateStructuralElements(Document mainDocument,
            ReferenceIntersector referenceIntersector,
            Element pipe,
            Curve pipeCurve,
            double offset,
            FamilySymbol symbol)
        {
            double diameter = pipe
                .get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)
                .AsDouble() + (offset / 304.8);
            Line pipeLine = pipeCurve as Line;
            XYZ origin = pipeLine.GetEndPoint(0);

            IEnumerable<Element> elements = referenceIntersector
                .Find(origin, pipeLine.Direction)
                .Where(x => x.Proximity <= pipeLine.Length)
                .Distinct(new ReferenceWithContextElementEqualityComparer())
                .Where(r => r != null)
                .Select(r => r.GetReference().ElementId)
                .Select(e => mainDocument.GetElement(e));

            foreach (Element element in elements)
            {
                Solid solid = Utils.GetTheSolid(element);
                if (solid is null)
                {
                    continue;
                }

                Curve elementCurve = Utils.GetTheCurve(element);
                if (elementCurve != null
                    && !(Utils.IsCurvePerpendicular(pipeCurve, elementCurve)))
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
                insertNew.LookupParameter("XYZ").Set(XYZParser.Insert(center));
            }
        }

        private void CheckExistingSleeves(Document mainDocument, Element pipe, Curve pipeCurve)
        {
            List<FamilyInstance> sleeves = new FilteredElementCollector(mainDocument)
                                    .OfClass(typeof(FamilyInstance))
                                    .Where(i => i.Name.StartsWith("00_Гильза_"))
                                    .Where(i => i.LookupParameter("Id").AsString() == pipe.UniqueId)
                                    .Select(i => i as FamilyInstance)
                                    .Where(i => Math.Abs(
                                        pipeCurve.Distance(
                                            XYZParser.Extract(i))) > 1e-6)
                                    .Where(d => d != null)
                                    .ToList();

            foreach (FamilyInstance sleeve in sleeves)
            {
                //not an optimal solution if we do everything in one operation
                //cuz new instances won't be created
                mainDocument.Delete(sleeve.Id);
            }

            // get pipes collection
            // iterate collection
            // find all family instances that relate to this pipe
            // check if an instance is centered
            // if not, move it to designated position
            // update coordinates
        }
    }
}
