using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Linq;

namespace SleevePlacer
{
    public static class Utils
    {
        public static ReferenceIntersector ReferenceIntersector<T>(UIDocument uiDocument)
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
        public static FamilySymbol GetFamilySymbol(Document document, string name)
        {
            FamilySymbol symbol = new FilteredElementCollector(document)
                    .OfClass(typeof(FamilySymbol))
                    .Where(s => s != null)
                    .FirstOrDefault(e => e.Name == name) as FamilySymbol;

            return symbol;
        }

        public static Solid GetTheSolid(Element element)
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

        public static Curve GetTheCurve(Element element)
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

        public static bool IsCurvePerpendicular(Curve curve1, Curve curve2)
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

        public static bool IsCurveHorizontal(Curve curve)
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

        public static bool IsCurveVertical(Curve curve)
        {
            XYZ start = curve.GetEndPoint(0);
            XYZ end = curve.GetEndPoint(1);

            if (Math.Round(start.X - end.X) == 0
                && Math.Round(start.Y - end.Y) == 0
                && Math.Round(start.Z - end.Z) != 0)
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
