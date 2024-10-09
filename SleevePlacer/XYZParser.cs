using Autodesk.Revit.DB;

namespace SleevePlacer
{
    public static class XYZParser
    {
        public static string Insert(XYZ xyz) => $"{xyz.X};{xyz.Y};{xyz.Z}";
        public static XYZ Extract(FamilyInstance familyInstance)
        {
            string extraction = familyInstance.LookupParameter("XYZ").AsString();
            string[] split = extraction.Split(';');
            double x = double.Parse(split[0]);
            double y = double.Parse(split[1]);
            double z = double.Parse(split[2]);
            return new XYZ(x, y, z);
        }
    }
}