using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using Form = System.Windows.Forms.Form;

namespace RevitFamilyInstanceLock
{
    public partial class FamilyPicker : Form
    {
        public FamilyPicker(ExternalEvent externalEvent, AddShareParameterExternalEvent addShareParameterExternalEvent, Document document, SortedList<string, FamilySymbol> sortedList)
        {
            InitializeComponent();
        }
    }
}
