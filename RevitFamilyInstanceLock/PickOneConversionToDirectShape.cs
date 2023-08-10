using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Threading;

namespace RevitFamilyInstanceLock
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    public class PickOneConversionToDirectShape : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            ResourceManager resourceManager;
            if (Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName == "zh")
            {
                resourceManager = new ResourceManager("RevitFamilyInstanceLock.Properties.Message_zh-Hant", Assembly.GetExecutingAssembly());
            }
            else
            {
                resourceManager = new ResourceManager("RevitFamilyInstanceLock.Properties.Message_en-US", Assembly.GetExecutingAssembly());
            }
            UIApplication expr_43 = commandData.Application;
            Document document = expr_43.ActiveUIDocument.Document;
            Selection selection = expr_43.ActiveUIDocument.Selection;
            StringBuilder stringBuilder = new StringBuilder();
            HashSet<string> hashSet = new HashSet<string>();
            int num = 0;
            Result result;
            try
            {
                Reference reference = selection.PickObject(ObjectType.Element, new InsulationSelectionFilter());
                Element element = document.GetElement(reference);
                FamilySymbol familySymbol = document.GetElement(element.GetTypeId()) as FamilySymbol;
                if (familySymbol == null)
                {
                    TaskDialog.Show(resourceManager.GetString("pickOne.NoFamily.Title"), resourceManager.GetString("pickOne.NoFamily.Text"));
                    result = Result.Cancelled;
                }
                else
                {
                    CommonTools.GetRelatedDim(document, element);
                    List<Element> allInstancesOfType = CommonTools.GetAllInstancesOfType(familySymbol, document);
                    List<Parameter> typeParameters = CommonTools.GetTypeParameters(familySymbol);
                    if (!CommonTools.AddParameterToSharedParameter(familySymbol, typeParameters, document))
                    {
                        result = Result.Failed;
                    }
                    else
                    {
                        stringBuilder.AppendLine("1.新增共用參數成功");
                        using (Transaction transaction = new Transaction(document, "Create direct shape"))
                        {
                            transaction.Start();
                            foreach (Element current in allInstancesOfType)
                            {
                                ElementClassFilter elementClassFilter = new ElementClassFilter(typeof(InsulationLiningBase));
                                List<ElementId> list = element.GetDependentElements(elementClassFilter).ToList<ElementId>();
                                if (list.Count > 0)
                                {
                                    document.GetElement(list[0]);
                                }
                                FamilySymbol familySymbol2 = document.GetElement(current.GetTypeId()) as FamilySymbol;
                                typeParameters = CommonTools.GetTypeParameters(familySymbol2);
                                Transform arg_182_0 = Transform.Identity;
                                if (current is FamilyInstance)
                                {
                                    (current as FamilyInstance).GetTotalTransform();
                                }
                                List<GeometryObject> list2 = CommonTools.GetElementSolids(current, document);
                                DirectShape directShape = DirectShape.CreateElement(document, current.Category.Id);
                                directShape.ApplicationId = "鉤逸科技";
                                directShape.ApplicationDataId = "sugoiitech.com";
                                if (directShape.IsValidShape(list2) && list2.Count != 0)
                                {
                                    directShape.SetShape(list2);
                                }
                                else
                                {
                                    List<GeometryObject> list3 = new List<GeometryObject>();
                                    foreach (GeometryObject current2 in list2)
                                    {
                                        if (directShape.IsValidShape(new List<GeometryObject>
                                        {
                                            current2
                                        }))
                                        {
                                            list3.Add(current2);
                                        }
                                    }
                                    list2 = list3;
                                    directShape.SetShape(list2);
                                    hashSet.Add(document.GetElement(current.GetTypeId()).Name);
                                }
                                directShape.Name = document.GetElement(current.GetTypeId()).Name;
                                stringBuilder.AppendLine("2.生成非參數物件成功");
                                if (!CommonTools.SetPropertyValueFromOriginalElement(directShape, current))
                                {
                                    result = Result.Failed;
                                    return result;
                                }
                                stringBuilder.AppendLine("3.修改普通參數成功");
                                if (!CommonTools.SetPropertyValueFromParameters(directShape, familySymbol2, typeParameters))
                                {
                                    result = Result.Failed;
                                    return result;
                                }
                                stringBuilder.AppendLine("4.修改族群參數成功");
                                document.Delete(current.Id);
                                stringBuilder.AppendLine("5.刪除原始物件成功");
                                num++;
                            }
                            document.Delete(familySymbol.Id);
                            List<ElementId> list4 = new List<ElementId>();
                            List<int> list5 = new List<int>();
                            list5.Add(14);
                            PerformanceAdviser.GetPerformanceAdviser();
                            using (IEnumerator<FailureMessage> enumerator3 = PerformanceAdviser.GetPerformanceAdviser().ExecuteRules(document, list5).GetEnumerator())
                            {
                                while (enumerator3.MoveNext())
                                {
                                    list4 = enumerator3.Current.GetFailingElements().ToList<ElementId>();
                                    if (document.Delete(list4).Count == 0)
                                    {
                                        stringBuilder.AppendLine("6.族群元件尚未移除");
                                    }
                                    else
                                    {
                                        stringBuilder.AppendLine("6.移除族群元件成功");
                                    }
                                }
                            }
                            transaction.Commit();
                            stringBuilder.AppendLine("7.模型變更成功");
                        }
                        //if (num > 1)
                        //{
                        //    TaskDialog.Show(resourceManager.GetString("pickOne.Result.Title"), string.Format("{0}{1}{2}", resourceManager.GetString("pickOne.Results.Text.First"), num, resourceManager.GetString("pickOne.Results.Text.Second")));
                        //}
                        //else if (num == 1)
                        //{
                        //    TaskDialog.Show(resourceManager.GetString("pickOne.Result.Title"), string.Format("{0}{1}{2}", resourceManager.GetString("pickOne.Result.Text.First"), num, resourceManager.GetString("pickOne.Result.Text.Second")));
                        //}
                        //else
                        //{
                        //    TaskDialog.Show(resourceManager.GetString("pickOne.Result.Title"), resourceManager.GetString("pickOne.Result.Fail") ?? "");
                        //}
                        if (hashSet.Count > 0)
                        {
                            StringBuilder stringBuilder2 = new StringBuilder();
                            foreach (string current3 in hashSet)
                            {
                                stringBuilder2.AppendLine(current3);
                            }
                            //TaskDialog.Show(resourceManager.GetString("pickOne.Warning.GeometryCheck.Title"), string.Format("{0}\n\n{1}", resourceManager.GetString("pickOne.Warning.GeometryCheck.Text"), stringBuilder2));
                        }
                        result = 0;
                    }
                }
            }
            catch (Exception)
            {
                TaskDialog.Show(resourceManager.GetString("pickOne.Error.ExecutionError.Title"), resourceManager.GetString("pickOne.Error.ExecutionError.Text"));
                result = Result.Failed;
            }
            return result;
        }
    }
}