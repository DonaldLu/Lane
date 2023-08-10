using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System;

namespace RevitFamilyInstanceLock
{
    public class InsulationSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element element)
        {
            return element.Category.Id.IntegerValue != -2008123 /*-2008123.GetHashCode()*/;
        }

        public bool AllowReference(Reference refer, XYZ point)
        {
            return false;
        }
    }
}
