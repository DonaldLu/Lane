using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Threading;

namespace RevitFamilyInstanceLock
{
    public class AddShareParameterExternalEvent : RevitEventWrapper<List<FamilySymbol>>
    {
        public override void Execute(UIApplication uiApp, List<FamilySymbol> familySymbols)
        {
            Document document = uiApp.ActiveUIDocument.Document;
            this.LockFamilySymbol(document, familySymbols);
        }

        public new string GetName()
        {
            return base.GetType().Name;
        }

        private void LockFamilySymbol(Document doc, List<FamilySymbol> familySymbols)
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
            using (Transaction transaction = new Transaction(doc, "Create direct shape"))
            {
                StringBuilder stringBuilder = new StringBuilder();
                HashSet<string> hashSet = new HashSet<string>();
                int num = 0;
                foreach (FamilySymbol current in familySymbols)
                {
                    try
                    {
                        List<Element> arg_9D_0 = CommonTools.GetAllInstancesOfType(current, doc);
                        List<Parameter> typeParameters = CommonTools.GetTypeParameters(current);
                        CommonTools.AddParameterToSharedParameter(current, typeParameters, doc);
                        stringBuilder.AppendLine("1.新增共用參數成功");
                        transaction.Start();
                        foreach (Element current2 in arg_9D_0)
                        {
                            FamilySymbol familySymbol = doc.GetElement(current2.GetTypeId()) as FamilySymbol;
                            typeParameters = CommonTools.GetTypeParameters(familySymbol);
                            Transform arg_D4_0 = Transform.Identity;
                            if (current2 is FamilyInstance)
                            {
                                (current2 as FamilyInstance).GetTotalTransform();
                            }
                            List<GeometryObject> list = CommonTools.GetElementSolids(current2, doc);
                            DirectShape directShape = DirectShape.CreateElement(doc, current2.Category.Id);
                            directShape.ApplicationId = "鉤逸科技";
                            directShape.ApplicationDataId = "sugoiitech.com";
                            if (directShape.IsValidShape(list) && list.Count != 0)
                            {
                                directShape.SetShape(list);
                            }
                            else
                            {
                                List<GeometryObject> list2 = new List<GeometryObject>();
                                foreach (GeometryObject current3 in list)
                                {
                                    if (directShape.IsValidShape(new List<GeometryObject>
                                    {
                                        current3
                                    }))
                                    {
                                        list2.Add(current3);
                                    }
                                }
                                list = list2;
                                directShape.SetShape(list);
                                hashSet.Add(doc.GetElement(current2.GetTypeId()).Name);
                            }
                            directShape.Name = doc.GetElement(current2.GetTypeId()).Name;
                            stringBuilder.AppendLine("2.生成非參數物件成功");
                            CommonTools.SetPropertyValueFromOriginalElement(directShape, current2);
                            stringBuilder.AppendLine("3.修改普通參數成功");
                            CommonTools.SetPropertyValueFromParameters(directShape, familySymbol, typeParameters);
                            stringBuilder.AppendLine("4.修改族群參數成功");
                            doc.Delete(current2.Id);
                            stringBuilder.AppendLine("5.刪除原始物件成功");
                            num++;
                        }
                        doc.Delete(current.Id);
                        List<ElementId> list3 = new List<ElementId>();
                        List<int> list4 = new List<int>();
                        list4.Add(14);
                        PerformanceAdviser.GetPerformanceAdviser();
                        using (IEnumerator<FailureMessage> enumerator4 = PerformanceAdviser.GetPerformanceAdviser().ExecuteRules(doc, list4).GetEnumerator())
                        {
                            while (enumerator4.MoveNext())
                            {
                                list3 = enumerator4.Current.GetFailingElements().ToList<ElementId>();
                                if (doc.Delete(list3).Count != 0)
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
                    catch (Exception arg_2F6_0)
                    {
                        throw arg_2F6_0;
                    }
                }
                if (num > 1)
                {
                    TaskDialog.Show(resourceManager.GetString("pickOne.Result.Title"), string.Format("{0}{1}{2}", resourceManager.GetString("pickOne.Results.Text.First"), num, resourceManager.GetString("pickOne.Results.Text.Second")));
                }
                else if (num == 1)
                {
                    TaskDialog.Show(resourceManager.GetString("pickOne.Result.Title"), string.Format("{0}{1}{2}", resourceManager.GetString("pickOne.Result.Text.First"), num, resourceManager.GetString("pickOne.Result.Text.Second")));
                }
                else
                {
                    TaskDialog.Show(resourceManager.GetString("pickOne.Result.Title"), resourceManager.GetString("pickOne.Result.Fail") ?? "");
                }
                if (hashSet.Count > 0)
                {
                    StringBuilder stringBuilder2 = new StringBuilder();
                    foreach (string current4 in hashSet)
                    {
                        stringBuilder2.AppendLine(current4);
                    }
                    TaskDialog.Show(resourceManager.GetString("pickOne.Warning.GeometryCheck.Title"), string.Format("{0}\n\n{1}", resourceManager.GetString("pickOne.Warning.GeometryCheck.Text"), stringBuilder2));
                }
            }
        }
    }
}