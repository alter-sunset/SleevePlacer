using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace SleevePlacer
{
    public class ReferenceWithContextComparer : IEqualityComparer<ReferenceWithContext>
    {
        public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
        {
            if (x is null || y is null) return false;
            if (ReferenceEquals(x, y)) return true;

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