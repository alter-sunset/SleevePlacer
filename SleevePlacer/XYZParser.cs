using Autodesk.Revit.DB;
using System.Windows.Forms;

namespace SleevePlacer
{
    public static class XYZParser
    {
        public static string Insert(XYZ xyz)
        {
            string insertion = $"{xyz.X};{xyz.Y};{xyz.Z}";

            return insertion;
        }

        public static XYZ Extract(FamilyInstance familyInstance)
        {
            string extraction = familyInstance.LookupParameter("XYZ").AsString();
            string[] split = extraction.Split(';');

            double x = double.Parse(split[0]);
            double y = double.Parse(split[1]);
            double z = double.Parse(split[2]);

            XYZ xyz = new XYZ(x, y, z);
            return xyz;
        }
    }
}
