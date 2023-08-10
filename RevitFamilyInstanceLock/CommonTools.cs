using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RevitFamilyInstanceLock
{
    public class CommonTools
    {
        public static List<GeometryObject> GetElementSolids2(Solid s, Document doc)
        {
            return new List<GeometryObject>();
        }

        public static List<GeometryObject> GetElementSolids(Element e, Document doc)
        {
            List<GeometryObject> list = new List<GeometryObject>();
            Options options = new Options();
            options.ComputeReferences = true;
            options.DetailLevel = ((doc.ActiveView != null) ? doc.ActiveView.DetailLevel : ViewDetailLevel.Medium);
            list = CommonTools.GetSolidsAndCurves(e as FamilyInstance, list, options);
            if (e is FamilyInstance)
            {
                list = CommonTools.GetSubElementSolids(e as FamilyInstance, list, options);
            }
            return list;
        }

        public static List<GeometryObject> GetSubElementSolids(FamilyInstance fi, List<GeometryObject> resultList, Options geoOpt)
        {
            List<ElementId> list = fi.GetSubComponentIds().ToList<ElementId>();
            if (list.Count == 0)
            {
                return resultList;
            }
            foreach (ElementId current in list)
            {
                FamilyInstance expr_38 = fi.Document.GetElement(current) as FamilyInstance;
                resultList = CommonTools.GetSolidsAndCurves(expr_38, resultList, geoOpt);
                CommonTools.GetSubElementSolids(expr_38, resultList, geoOpt);
            }
            return resultList;
        }

        public static List<GeometryObject> GetSolidsAndCurves(FamilyInstance fi, List<GeometryObject> resultList, Options geoOpt)
        {
            using (IEnumerator<GeometryObject> enumerator = fi.get_Geometry(geoOpt).GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    GeometryInstance geometryInstance = enumerator.Current as GeometryInstance;
                    if (null != geometryInstance)
                    {
                        Transform transform = geometryInstance.Transform;
                        foreach (GeometryObject expr_45 in geometryInstance.GetSymbolGeometry())
                        {
                            Solid solid = expr_45 as Solid;
                            Curve curve = expr_45 as Curve;
                            if (solid != null && solid.Volume != 0.0 && solid.SurfaceArea != 0.0)
                            {
                                resultList.Add(SolidUtils.CreateTransformed(solid, transform));
                            }
                            if (curve != null)
                            {
                                Curve item = NurbSpline.CreateCurve(HermiteSpline.Create(curve.Tessellate(), false)).CreateTransformed(transform);
                                resultList.Add(item);
                            }
                        }
                    }
                }
            }
            return resultList;
        }

        public static bool AddParameterToSharedParameter(FamilySymbol fs, List<Parameter> pList, Document doc)
        {
            bool result;
            try
            {
                string directoryName = Path.GetDirectoryName(doc.PathName);
                doc.Application.SharedParametersFilename = Path.Combine(directoryName, "sugoiitechSharedParameters.txt");
                using (Transaction transaction = new Transaction(doc, "新增共用參數"))
                {
                    transaction.Start();
                    Application application = doc.Application;
                    DefinitionFile definitionFile = application.OpenSharedParameterFile();
                    if (definitionFile == null)
                    {
                        if (File.Exists(doc.Application.SharedParametersFilename))
                        {
                            File.Delete(doc.Application.SharedParametersFilename);
                        }
                        File.Create(doc.Application.SharedParametersFilename).Close();
                        definitionFile = application.OpenSharedParameterFile();
                    }
                    DefinitionGroup definitionGroup = definitionFile.Groups.get_Item("鉤逸科技");
                    if (definitionGroup == null)
                    {
                        definitionGroup = definitionFile.Groups.Create("鉤逸科技");
                    }
                    foreach (Parameter current in pList)
                    {
                        string text = current.Definition.Name.Contains("Type") ? ("_" + current.Definition.Name) : (fs.FamilyName + "_" + current.Definition.Name);
                        if (definitionGroup.Definitions.get_Item(text) == null)
                        {
                            ExternalDefinitionCreationOptions externalDefinitionCreationOptions = new ExternalDefinitionCreationOptions(text, current.Definition.ParameterType);
                            definitionGroup.Definitions.Create(externalDefinitionCreationOptions);
                        }
                        Definition definition = definitionGroup.Definitions.get_Item(text);
                        Category category = fs.Category;
                        CategorySet categorySet = application.Create.NewCategorySet();
                        DefinitionBindingMapIterator definitionBindingMapIterator = doc.ParameterBindings.ForwardIterator();
                        while (definitionBindingMapIterator.MoveNext())
                        {
                            Definition key = definitionBindingMapIterator.Key;
                            ElementBinding elementBinding = (ElementBinding)definitionBindingMapIterator.Current;
                            if (text == key.Name)
                            {
                                IEnumerator enumerator2 = elementBinding.Categories.GetEnumerator();
                                //using (IEnumerator enumerator2 = elementBinding.Categories.GetEnumerator())
                                //{
                                    while (enumerator2.MoveNext())
                                    {
                                        Category category2 = (Category)enumerator2.Current;
                                        categorySet.Insert(category2);
                                    }
                                    break;
                                //}
                            }
                        }
                        categorySet.Insert(category);
                        InstanceBinding instanceBinding = application.Create.NewInstanceBinding(categorySet);
                        if (categorySet.Size > 1)
                        {
                            doc.ParameterBindings.ReInsert(definition, instanceBinding);
                        }
                        else
                        {
                            doc.ParameterBindings.Insert(definition, instanceBinding);
                        }
                    }
                    transaction.Commit();
                }
                result = true;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("錯誤訊息", ex.Message);
                result = false;
            }
            return result;
        }

        public static bool SetPropertyValueFromOriginalElement(Element newElement, Element originalElement)
        {
            bool result;
            try
            {
                foreach (Parameter parameter in originalElement.Parameters)
                {
                    Parameter parameter2 = newElement.LookupParameter(parameter.Definition.Name);
                    if (parameter2 != null && !parameter2.IsReadOnly)
                    {
                        if (parameter.StorageType == StorageType.Double)
                        {
                            parameter2.Set(parameter.AsDouble());
                        }
                        else if (parameter.StorageType == StorageType.ElementId)
                        {
                            parameter2.Set(parameter.AsElementId());
                        }
                        else if (parameter.StorageType == StorageType.Integer)
                        {
                            parameter2.Set(parameter.AsInteger());
                        }
                        else if (parameter.StorageType == StorageType.String)
                        {
                            parameter2.Set(parameter.AsString());
                        }
                    }
                }
                result = true;
            }
            catch (Exception)
            {
                result = false;
            }
            return result;
        }

        public static List<Element> GetAllInstancesOfType(FamilySymbol fs, Document doc)
        {
            List<Element> result = new List<Element>();
            try
            {
                result = (from FamilyInstance x in new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).WhereElementIsNotElementType()
                          where x.Symbol.Family.Id.IntegerValue == fs.Family.Id.IntegerValue && x.Symbol.Id == fs.Id
                          select x).ToList<Element>();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("獲取同族群元件失敗", string.Concat(new string[]
                {
                    ex.Message,
                    "\n",
                    fs.FamilyName,
                    "::",
                    fs.Name
                }));
            }
            return result;
        }

        public static List<Dimension> GetRelatedDim(Document doc, Element selectedElem)
        {
            List<Dimension> list = new List<Dimension>();
            try
            {
                foreach (Dimension current in new FilteredElementCollector(doc).OfClass(typeof(Dimension)).Cast<Dimension>().ToList<Dimension>())
                {
                    IEnumerator enumerator2 = current.References.GetEnumerator();
                    //using (IEnumerator enumerator2 = current.References.GetEnumerator())
                    //{
                        while (enumerator2.MoveNext())
                        {
                            if (((Reference)enumerator2.Current).ElementId.IntegerValue == selectedElem.Id.IntegerValue)
                            {
                                list.Add(current);
                                break;
                            }
                        }
                    //}
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("抓取關聯尺寸標註失敗", ex.Message);
            }
            return list;
        }

        public static List<Parameter> GetTypeParameters(FamilySymbol fs)
        {
            List<Parameter> list = new List<Parameter>();
            foreach (Parameter parameter in fs.Parameters)
            {
                if (parameter.Definition.Name.Contains("Type") || parameter.Definition.ParameterType == ParameterType.Length || parameter.Definition.ParameterType == ParameterType.Area || parameter.Definition.ParameterType == ParameterType.Volume)
                {
                    list.Add(parameter);
                }
            }
            return list;
        }

        public static bool SetPropertyValueFromParameters(Element newElement, FamilySymbol symbol, List<Parameter> pList)
        {
            bool result;
            try
            {
                foreach (Parameter current in pList)
                {
                    Parameter parameter = newElement.LookupParameter(current.Definition.Name.Contains("Type") ? ("_" + current.Definition.Name) : (symbol.FamilyName + "_" + current.Definition.Name));
                    if (parameter != null && !parameter.IsReadOnly)
                    {
                        if (current.StorageType == StorageType.Double)
                        {
                            parameter.Set(current.AsDouble());
                        }
                        else if (current.StorageType == StorageType.ElementId)
                        {
                            parameter.Set(current.AsElementId());
                        }
                        else if (current.StorageType == StorageType.Integer)
                        {
                            parameter.Set(current.AsInteger());
                        }
                        else if (current.StorageType == StorageType.String)
                        {
                            parameter.Set(current.AsString());
                        }
                    }
                }
                result = true;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("修改族群參數失敗", ex.Message);
                result = false;
            }
            return result;
        }
    }
}