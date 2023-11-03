using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Lane
{
    [Transaction(TransactionMode.Manual)]
    public class LaneCrush : IExternalCommand
    {
        public class StartEndPoint
        {
            public XYZ start = new XYZ();
            public XYZ end = new XYZ();
            public List<XYZ> xyzs = new List<XYZ>();
            public XYZ other = new XYZ();
            public bool isPeriodic = false;
            public List<XYZ> tangents = new List<XYZ>();
            public Double perimeter = 0.0;
            public string type = string.Empty;
        }
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            IList<ElementFilter> elementFilters = new List<ElementFilter>(); // 清空過濾器
            ElementCategoryFilter roomFilter = new ElementCategoryFilter(BuiltInCategory.OST_Floors); // 房間
            elementFilters.Add(roomFilter);
            LogicalOrFilter logicalOrFilter = new LogicalOrFilter(elementFilters);
            List<Floor> floors = new FilteredElementCollector(doc).WherePasses(logicalOrFilter).WhereElementIsNotElementType().Where(x => x is Floor)
                                    .Where(x => x.LookupParameter("樓板高程") != null).Where(x => x.LookupParameter("樓板高程").AsValueString().Equals("車道板")).Cast<Floor>().ToList();
            //floors = floors.Where(x => x.Id.IntegerValue.Equals(2250673)).Cast<Floor>().ToList();
            double height = UnitUtils.ConvertToInternalUnits(210, UnitTypeId.Centimeters);
            TransactionGroup tranGrp = new TransactionGroup(doc, "車道板校核");
            tranGrp.Start();
            using (Transaction trans = new Transaction(doc, "生成干涉模型"))
            {
                trans.Start();
                // 取得樓板的頂面輪廓
                foreach (Floor floor in floors)
                {
                    List<Solid> solidList = GetRoadwayPlankSolids(doc, floor); // 儲存所有車道板的Solid
                    List<Face> topFaces = GetTopFaces(solidList); // 取得車道的頂面
                    foreach (Face topFace in topFaces)
                    {
                        try
                        {
                            ElementId materialId = floor.Category.Id;
                            CreateCrushFaces(doc, topFace, materialId); // 建立Face的封閉曲線(座標點不用重複)
                            //break;
                        }
                        catch (Exception ex)
                        {
                            string msg = ex.Message + "\n" + ex.ToString();
                        }
                    }
                }
                trans.Commit();
                uidoc.RefreshActiveView();
            }
            tranGrp.Assimilate();

            return Result.Succeeded;
        }
        // 儲存所有車道板的Solid
        private List<Solid> GetRoadwayPlankSolids(Document doc, Floor floor)
        {
            List<Solid> solidList = new List<Solid>();

            // 1.讀取Geometry Option
            Options options = new Options();
            //options.View = doc.GetElement(room.Level.FindAssociatedPlanViewId()) as Autodesk.Revit.DB.View;
            options.DetailLevel = ((doc.ActiveView != null) ? doc.ActiveView.DetailLevel : ViewDetailLevel.Medium);
            options.ComputeReferences = true;
            options.IncludeNonVisibleObjects = true;
            // 得到幾何元素
            GeometryElement geomElem = floor.get_Geometry(options);
            List<Solid> solids = GeometrySolids(geomElem);
            foreach (Solid solid in solids)
            {
                solidList.Add(solid);
            }

            return solidList;
        }
        // 取得車道的Solid
        private List<Solid> GeometrySolids(GeometryObject geoObj)
        {
            List<Solid> solids = new List<Solid>();
            if (geoObj is Solid)
            {
                Solid solid = (Solid)geoObj;
                if (solid.Faces.Size > 0)
                {
                    solids.Add(solid);
                }
            }
            if (geoObj is GeometryInstance)
            {
                GeometryInstance geoIns = geoObj as GeometryInstance;
                GeometryElement geometryElement = (geoObj as GeometryInstance).GetSymbolGeometry(geoIns.Transform); // 座標轉換
                foreach (GeometryObject o in geometryElement)
                {
                    solids.AddRange(GeometrySolids(o));
                }
            }
            else if (geoObj is GeometryElement)
            {
                GeometryElement geometryElement2 = (GeometryElement)geoObj;
                foreach (GeometryObject o in geometryElement2)
                {
                    solids.AddRange(GeometrySolids(o));
                }
            }
            return solids;
        }
        // 取得車道的頂面
        private List<Face> GetTopFaces(List<Solid> solidList)
        {
            List<Face> topFaces = new List<Face>();
            foreach (Solid solid in solidList)
            {
                foreach (Face face in solid.Faces)
                {
                    double faceTZ = face.ComputeNormal(new UV(0.5, 0.5)).Z;
                    if (faceTZ > 0.0) // 頂面
                    {
                        topFaces.Add(face);
                    }
                }
            }
            return topFaces;
        }
        // 建立Face的封閉曲線(座標點不用重複)
        private void CreateCrushFaces(Document doc, Face topFace, ElementId materialId)
        {
            IList<XYZ> curveXYZs = new List<XYZ>();
            TessellatedShapeBuilder builder = new TessellatedShapeBuilder();
            builder.OpenConnectedFaceSet(true);

            // 干涉元件底面
            curveXYZs.Clear();
            foreach (EdgeArray edgeArray in topFace.EdgeLoops)
            {
                foreach (Edge edge in edgeArray)
                {
                    Curve curve = edge.AsCurveFollowingFace(topFace);
                    for(int i = 1; i < curve.Tessellate().Count; i++)
                    {
                        curveXYZs.Add(curve.Tessellate()[i]);
                    }
                }
            }
            builder.AddFace(new TessellatedFace(curveXYZs, materialId));
            // 複製偏移的頂面
            double height = UnitUtils.ConvertToInternalUnits(210, UnitTypeId.Centimeters);
            // 干涉元件頂面
            curveXYZs.Clear();
            foreach (EdgeArray edgeArray in topFace.EdgeLoops)
            {
                foreach (Edge edge in edgeArray)
                {
                    Curve curve = edge.AsCurveFollowingFace(topFace);
                    for (int i = 1; i < curve.Tessellate().Count; i++)
                    {
                        curveXYZs.Add(new XYZ(curve.Tessellate()[i].X, curve.Tessellate()[i].Y, curve.Tessellate()[i].Z + height));
                    }
                }
            }
            builder.AddFace(new TessellatedFace(curveXYZs, materialId));
            // 干涉元件側面
            foreach (EdgeArray edgeArray in topFace.EdgeLoops)
            {
                foreach (Edge edge in edgeArray)
                {
                    curveXYZs.Clear();
                    Curve curve = edge.AsCurveFollowingFace(topFace);
                    curveXYZs.Add(curve.Tessellate()[0]);
                    curveXYZs.Add(curve.Tessellate()[curve.Tessellate().Count - 1]);
                    curveXYZs.Add(new XYZ(curve.Tessellate()[curve.Tessellate().Count - 1].X, curve.Tessellate()[curve.Tessellate().Count - 1].Y, curve.Tessellate()[curve.Tessellate().Count - 1].Z + height));
                    curveXYZs.Add(new XYZ(curve.Tessellate()[0].X, curve.Tessellate()[0].Y, curve.Tessellate()[0].Z + height));
                    builder.AddFace(new TessellatedFace(curveXYZs, materialId));
                }
            }

            try
            {
                builder.CloseConnectedFaceSet();
                builder.Target = TessellatedShapeBuilderTarget.Solid;
                builder.Fallback = TessellatedShapeBuilderFallback.Abort;
                builder.Build();

                TessellatedShapeBuilderResult result = builder.GetBuildResult();
                DirectShape ds = DirectShape.CreateElement(doc, materialId);
                ds.ApplicationId = "Application id";
                ds.ApplicationDataId = "Geometry object id";
                ds.SetShape(result.GetGeometricalObjects());
            }
            catch(Exception ex) { }
        }
        // 儲存所有Face的封閉曲線
        private List<Curve> SaveFaceCurveLoop(Face topFace)
        {
            Transform transform = new Transform(Transform.CreateTranslation(new XYZ(0, 0, 0)));
            List<Curve> resultList = new List<Curve>();
            IList<CurveLoop> curveLoops = topFace.GetEdgesAsCurveLoops(); // 頂面的封閉曲線
            // 複製偏移的頂面
            double height = UnitUtils.ConvertToInternalUnits(210, UnitTypeId.Centimeters);
            // 底面
            foreach (EdgeArray edgeArray in topFace.EdgeLoops)
            {
                foreach (Edge edge in edgeArray)
                {
                    Curve curve = edge.AsCurveFollowingFace(topFace);
                    IList<XYZ> curveXYZs = new List<XYZ>();
                    foreach (XYZ xyz in curve.Tessellate())
                    {
                        curveXYZs.Add(new XYZ(xyz.X, xyz.Y, xyz.Z + height));
                    }
                    curve = NurbSpline.CreateCurve(HermiteSpline.Create(curveXYZs, false)).CreateTransformed(transform);
                    resultList.Add(curve);
                }
            }
            // 頂面
            foreach (EdgeArray edgeArray in topFace.EdgeLoops)
            {
                foreach (Edge edge in edgeArray)
                {
                    Curve curve = edge.AsCurveFollowingFace(topFace);
                    IList<XYZ> curveXYZs = new List<XYZ>();
                    foreach (XYZ xyz in curve.Tessellate())
                    {
                        curveXYZs.Add(new XYZ(xyz.X, xyz.Y, xyz.Z));
                    }
                    curve = NurbSpline.CreateCurve(HermiteSpline.Create(curveXYZs, false)).CreateTransformed(transform);
                    resultList.Add(curve);
                }
            }
            // 找到底邊, 連結成面
            foreach (Curve bottomCurve in curveLoops.First())
            {
                //resultList.Add(bottomCurve); // 儲存底邊
                XYZ startXYZ = new XYZ(bottomCurve.Tessellate()[bottomCurve.Tessellate().Count - 1].X, bottomCurve.Tessellate()[bottomCurve.Tessellate().Count - 1].Y, bottomCurve.Tessellate()[bottomCurve.Tessellate().Count - 1].Z);
                XYZ endXYZ = new XYZ(bottomCurve.Tessellate()[bottomCurve.Tessellate().Count - 1].X, bottomCurve.Tessellate()[bottomCurve.Tessellate().Count - 1].Y, bottomCurve.Tessellate()[bottomCurve.Tessellate().Count - 1].Z + height);
                IList<XYZ> curveXYZs = new List<XYZ>();
                curveXYZs.Add(startXYZ);
                curveXYZs.Add(endXYZ);
                Curve curve = NurbSpline.CreateCurve(HermiteSpline.Create(curveXYZs, false)).CreateTransformed(transform);
                resultList.Add(curve); // 儲存底邊連結頂邊
                // 反轉底面的座標
                curveXYZs = new List<XYZ>();
                IList<XYZ> reversedXYZs = bottomCurve.CreateReversed().Tessellate();
                foreach (XYZ xyz in reversedXYZs)
                {
                    curveXYZs.Add(new XYZ(xyz.X, xyz.Y, xyz.Z + height));
                }
                curve = NurbSpline.CreateCurve(HermiteSpline.Create(curveXYZs, false)).CreateTransformed(transform);
                //resultList.Add(curve); // 儲存頂邊
                startXYZ = new XYZ(curve.Tessellate()[curve.Tessellate().Count - 1].X, curve.Tessellate()[curve.Tessellate().Count - 1].Y, curve.Tessellate()[curve.Tessellate().Count - 1].Z);
                endXYZ = new XYZ(curve.Tessellate()[curve.Tessellate().Count - 1].X, curve.Tessellate()[curve.Tessellate().Count - 1].Y, curve.Tessellate()[curve.Tessellate().Count - 1].Z - height);
                curveXYZs = new List<XYZ>();
                curveXYZs.Add(startXYZ);
                curveXYZs.Add(endXYZ);
                curve = NurbSpline.CreateCurve(HermiteSpline.Create(curveXYZs, false)).CreateTransformed(transform);
                //resultList.Add(curve); // 儲存頂邊連結底邊
            }

            return resultList;
        }
        public static List<GeometryObject> GetElementSolids(Floor floor, Document doc)
        {
            List<GeometryObject> list = new List<GeometryObject>();
            Options options = new Options();
            options.ComputeReferences = true;
            options.DetailLevel = ((doc.ActiveView != null) ? doc.ActiveView.DetailLevel : ViewDetailLevel.Medium);
            list = GetSolidsAndCurves(floor, list, options);
            return list;
        }
        public static List<GeometryObject> GetSolidsAndCurves(Floor floor, List<GeometryObject> resultList, Options geoOpt)
        {
            using (IEnumerator<GeometryObject> enumerator = floor.get_Geometry(geoOpt).GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    List<Face> topFaces = new List<Face>();
                    List<CurveLoop> curveLoop = new List<CurveLoop>();
                    double height = UnitUtils.ConvertToInternalUnits(210, UnitTypeId.Centimeters);

                    GeometryInstance geometryInstance = enumerator.Current as GeometryInstance;
                    if (null != geometryInstance)
                    {
                        Transform transform = geometryInstance.Transform;
                        foreach (GeometryObject geometryObject in geometryInstance.GetSymbolGeometry())
                        {
                            Solid solid = geometryObject as Solid;
                            Curve curve = geometryObject as Curve;
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
                    if(enumerator.Current is Solid)
                    {
                        Solid solid = enumerator.Current as Solid;
                        if (solid != null && solid.Volume != 0.0 && solid.SurfaceArea != 0.0)
                        {
                            // 找到Solid所有的頂面
                            foreach (Face face in solid.Faces)
                            {
                                CurveLoop baseLoop = new CurveLoop();
                                List <Curve> curves = new List<Curve>();
                                Transform transform = new Transform(Transform.CreateTranslation(new XYZ(0, 0, 0)));
                                double faceTZ = face.ComputeNormal(new UV(0.5, 0.5)).Z;
                                if (faceTZ > 0.0) // 頂面
                                {
                                    topFaces.Add(face);
                                    foreach (EdgeArray edgeArray in face.EdgeLoops)
                                    {
                                        foreach (Edge edge in edgeArray)
                                        {
                                            IList<XYZ> curveXYZs = new List<XYZ>();
                                            foreach (XYZ xyz in edge.Tessellate())
                                            {
                                                curveXYZs.Add(new XYZ(xyz.X, xyz.Y, xyz.Z + height));
                                            }
                                            Curve curve = NurbSpline.CreateCurve(HermiteSpline.Create(curveXYZs, false)).CreateTransformed(transform);
                                            curves.Add(curve);
                                            //resultList.Add(item);
                                        }
                                    }
                                    baseLoop = CurveLoop.Create(curves);
                                    curveLoop.Add(baseLoop);
                                }
                            }
                            Solid preTransformBox = GeometryCreationUtilities.CreateExtrusionGeometry(curveLoop, XYZ.BasisZ, height);
                            Solid transformBox = SolidUtils.CreateTransformed(preTransformBox, new Transform(Transform.CreateTranslation(new XYZ(0, 0, 0))));
                            resultList.Add(transformBox);
                            //resultList.Add(SolidUtils.CreateTransformed(solid, new Transform(Transform.CreateTranslation(new XYZ(0, 0, 10)))));
                            //resultList.Add(solid);
                        }
                    }
                }
            }
            return resultList;
        }
    }
}