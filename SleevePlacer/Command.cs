using System;
using System.Linq;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace SleevePlacer
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            using UIApplication uiApp = commandData.Application;
            DateTime start = DateTime.Now;
            UIDocument uiDocument = uiApp.ActiveUIDocument;
            Document mainDocument = uiDocument.Document;

            FamilySymbol symbolWall = Utils.GetFamilySymbol(mainDocument, "00_Гильза_Стена");
            FamilySymbol symbolFloor = Utils.GetFamilySymbol(mainDocument, "00_Гильза_Плита");

            if (symbolWall is null || symbolFloor is null)
            {
                TaskDialog.Show("Нет семейств", "Требуемые семейства гильз не обнаружены");
                return Result.Failed;
            }

            Document[] links = new FilteredElementCollector(mainDocument)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Select(l => l.GetLinkDocument())
                .Where(d => d is not null)
                .ToArray();

            using Transaction transaction = new(mainDocument);
            transaction.Start("Добавление гильз");
            symbolWall?.Activate();
            symbolFloor?.Activate();

            foreach (Document linkedDocument in links)
            {
                IEnumerable<Pipe> pipes = new FilteredElementCollector(linkedDocument)
                    .OfClass(typeof(Pipe))
                    .Cast<Pipe>()
                    .Where(p => p is not null);

                foreach (Pipe pipe in pipes)
                {
                    Curve pipeCurve = Utils.GetTheCurve(pipe);
                    if (pipeCurve is null || !pipeCurve.IsBound) continue;

                    (ReferenceIntersector referenceIntersector, FamilySymbol symbol) = pipeCurve.GetReferenceIntersectorAndSymbol(uiDocument, symbolWall, symbolFloor);

                    if (referenceIntersector == null || symbol == null) continue;

                    try
                    {
                        IterateStructuralElements(mainDocument, referenceIntersector, pipe, pipeCurve, 100, symbol);
                    }
                    catch (NullReferenceException)
                    {
                        TaskDialog.Show("Ошибка", "Проверьте параметры семейств.");
                        return Result.Failed;
                    }
                }
            }

            transaction.Commit();
            TaskDialog.Show("Готово!", $"{DateTime.Now - start}");
            return Result.Succeeded;
        }

        private static void IterateStructuralElements(Document mainDocument,
            ReferenceIntersector referenceIntersector,
            Pipe pipe,
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
                .Distinct(new ReferenceWithContextComparer())
                .Select(r => mainDocument.GetElement(r.GetReference().ElementId));

            foreach (Element element in elements)
            {
                Solid solid = Utils.GetTheSolid(element);
                if (solid is null) continue;

                Curve elementCurve = Utils.GetTheCurve(element);
                if (elementCurve is not null
                    && !Utils.IsCurvePerpendicular(pipeCurve, elementCurve)) continue;

                SolidCurveIntersection elementPipeIntersection = solid.IntersectWithCurve(pipeCurve, null);
                if (elementPipeIntersection == null
                    || !elementPipeIntersection.IsValidObject
                    || elementPipeIntersection.SegmentCount < 1) continue;

                Curve line = elementPipeIntersection.GetCurveSegment(0);
                if (line is null) continue;

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
    }
}