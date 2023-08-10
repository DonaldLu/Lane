using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Threading;
using System.Windows.Forms;

namespace RevitFamilyInstanceLock
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    public class FamilyConversionToDirectShape : IExternalCommand
    {
        public const string _appId = "5517523135976590772";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document document = commandData.Application.ActiveUIDocument.Document;
            if (Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName == "zh")
            {
                new ResourceManager("RevitFamilyInstanceLock.Properties.Message_zh-Hant", Assembly.GetExecutingAssembly());
            }
            else
            {
                new ResourceManager("RevitFamilyInstanceLock.Properties.Message_en-US", Assembly.GetExecutingAssembly());
            }
            FilteredElementCollector expr_54 = new FilteredElementCollector(document);
            expr_54.OfClass(typeof(Family));
            SortedList<string, FamilySymbol> sortedList = new SortedList<string, FamilySymbol>();
            int num = 0;
            int num2 = 0;
            int num3 = 0;
            using (IEnumerator<Element> enumerator = expr_54.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    Element arg_8E_0 = (Family)enumerator.Current;
                    num++;
                    FamilySymbolFilter familySymbolFilter = new FamilySymbolFilter(arg_8E_0.Id);
                    FilteredElementCollector expr_A0 = new FilteredElementCollector(document);
                    expr_A0.WherePasses(familySymbolFilter);
                    using (IEnumerator<Element> enumerator2 = expr_A0.GetEnumerator())
                    {
                        while (enumerator2.MoveNext())
                        {
                            FamilySymbol familySymbol = (FamilySymbol)enumerator2.Current;
                            num2++;
                            FamilyInstanceFilter familyInstanceFilter = new FamilyInstanceFilter(document, familySymbol.Id);
                            IEnumerable<FamilyInstance> enumerable = new FilteredElementCollector(document, document.ActiveView.Id).WherePasses(familyInstanceFilter).OfType<FamilyInstance>();
                            int num4 = enumerable.Count<FamilyInstance>();
                            num3 += num4;
                            if (num4 > 0)
                            {
                                bool flag = true;
                                using (IEnumerator<FamilyInstance> enumerator3 = enumerable.GetEnumerator())
                                {
                                    while (enumerator3.MoveNext())
                                    {
                                        flag = (enumerator3.Current.SuperComponent != null);
                                    }
                                }
                                if (!flag)
                                {
                                    sortedList.Add(familySymbol.FamilyName + "::" + familySymbol.Name, familySymbol);
                                }
                            }
                        }
                    }
                }
            }
            AddShareParameterExternalEvent addShareParameterExternalEvent = new AddShareParameterExternalEvent();
            FamilyPicker familyPicker = new FamilyPicker(ExternalEvent.Create(addShareParameterExternalEvent), addShareParameterExternalEvent, document, sortedList);
            familyPicker.Show();
            this.UpdateFamilySymbolListView(ref familyPicker, sortedList);
            return 0;
        }

        public void UpdateFamilySymbolListView(ref FamilyPicker cpForm, SortedList<string, FamilySymbol> sortedFamilySymbols)
        {
            if (cpForm.familySymbolListView.Items.Count > 0)
            {
                cpForm.familySymbolListView.Clear();
            }
            foreach (FamilySymbol current in sortedFamilySymbols.Values.ToList<FamilySymbol>())
            {
                ListViewItem value = new ListViewItem(new string[]
                {
                    current.FamilyName,
                    current.Name
                });
                cpForm.familySymbolListView.Items.Add(value);
            }
            for (int i = 0; i < cpForm.familySymbolListView.Columns.Count; i++)
            {
                cpForm.familySymbolListView.Columns[i].Width = -1;
            }
        }
    }
}