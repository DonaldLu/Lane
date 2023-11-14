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
    [Regeneration(RegenerationOption.Manual)]
    public class CreateCrushElems : IExternalEventHandler
    {
        public class CrushElemInfo
        {
            public string hostName { get; set; }
            public List<string> crushElemName = new List<string>();
        }
        public void Execute(UIApplication uiapp)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            IList<ElementFilter> elementFilters = new List<ElementFilter>(); // 清空過濾器
            ElementCategoryFilter roomFilter = new ElementCategoryFilter(BuiltInCategory.OST_Floors); // 樓板
            elementFilters.Add(roomFilter);
            LogicalOrFilter logicalOrFilter = new LogicalOrFilter(elementFilters);
            RemoveElems(uidoc, doc, logicalOrFilter); // 移除干涉元件
            List<Floor> floors = new List<Floor>();
            try
            {
                floors = new FilteredElementCollector(doc).WherePasses(logicalOrFilter).WhereElementIsNotElementType().Where(x => x is Floor)
                         .Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString() != null)
                         .Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Contains("車道板")).Cast<Floor>().ToList();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", "樓板參數錯誤！\n" + ex.Message + "\n" + ex.ToString());
            }
            //floors = floors.Where(x => x.Id.IntegerValue.Equals(2250673)).Cast<Floor>().ToList(); // 測試
            double height = UnitUtils.ConvertToInternalUnits(LaneCrushForm.height, UnitTypeId.Centimeters); // 2022
            //double height = UnitUtils.ConvertToInternalUnits(LaneCrushForm.height, DisplayUnitType.DUT_CENTIMETERS); // 2020
            TransactionGroup tranGrp = new TransactionGroup(doc, "車道板校核");
            tranGrp.Start();
            using (Transaction trans = new Transaction(doc, "生成干涉模型"))
            {
                // 取得樓板的頂面輪廓
                foreach (Floor floor in floors)
                {
                    ElementId materialId = floor.Category.Id;
                    string number = string.Empty;
                    //try { number = floor.LookupParameter("編號").AsString(); }
                    //catch (Exception) { }
                    List<Solid> crushSolids = new List<Solid>();
                    List<Solid> solidList = GetSolids(doc, floor); // 儲存所有車道板的Solid
                    List<Face> topFaces = GetTopFaces(solidList); // 取得車道的頂面
                    foreach (Face topFace in topFaces)
                    {
                        try
                        {
                            trans.Start();
                            Solid crushSolid = CreateCrushSolids(doc, topFace, height, materialId, floor.Id, number); // 使用CreateLoftGeometry擠出車道板頂面Solid
                            crushSolids.Add(crushSolid);
                            //// 建立Face的封閉曲線(座標點不用重複)
                            //if (topFace is RuledFace) { GetTessellatedSolid(doc, topFace, height, materialId); }
                            //else { CreateCrushFaces(doc, topFace, height, materialId); }
                            trans.Commit();
                            uidoc.RefreshActiveView();
                        }
                        catch (Exception ex)
                        {
                            string msg = ex.Message + "\n" + ex.ToString();
                        }
                    }
                }
            }
            tranGrp.Assimilate();

            List<CrushElemInfo> crushElemInfos = CrushReport(doc, logicalOrFilter, floors); // 出衝突報告
            if (crushElemInfos.Count > 0)
            {

            }
            else
            {
                TaskDialog.Show("Revit", "車道板無干涉到的元件");
            }
        }
        // 移除干涉元件
        private void RemoveElems(UIDocument uidoc, Document doc, LogicalOrFilter logicalOrFilter)
        {
            List<ElementId> crushFloorIds = new List<ElementId>();
            try
            {
                crushFloorIds = new FilteredElementCollector(doc).WherePasses(logicalOrFilter).WhereElementIsNotElementType().Where(x => x is DirectShape).Cast<DirectShape>()
                                .Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString() != null)
                                .Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Contains("車道板干涉元件")).Select(x => x.Id).ToList();
            }
            catch (Exception ex)
            {
                string error = "樓板參數錯誤！\n" + ex.Message + "\n" + ex.ToString();
            }
            using (Transaction trans = new Transaction(doc, "移除干涉元件"))
            {
                trans.Start();
                doc.Delete(crushFloorIds);
                trans.Commit();
                uidoc.RefreshActiveView();
            }
        }
        // 儲存所有車道板的Solid
        private List<Solid> GetSolids(Document doc, Element elem)
        {
            List<Solid> solidList = new List<Solid>();

            // 1.讀取Geometry Option
            Options options = new Options();
            //options.View = doc.GetElement(room.Level.FindAssociatedPlanViewId()) as Autodesk.Revit.DB.View;
            options.DetailLevel = ((doc.ActiveView != null) ? doc.ActiveView.DetailLevel : ViewDetailLevel.Medium);
            options.ComputeReferences = true;
            options.IncludeNonVisibleObjects = true;
            // 得到幾何元素
            GeometryElement geomElem = elem.get_Geometry(options);
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
        // 使用CreateLoftGeometry擠出車道板頂面Solid
        private Solid CreateCrushSolids(Document doc, Face topFace, double height, ElementId materialId, ElementId floorId, string number)
        {
            CurveLoop buttomCurves = new CurveLoop();
            CurveLoop topCurves = new CurveLoop();
            // 干涉元件底面
            foreach (EdgeArray edgeArray in topFace.EdgeLoops)
            {
                List<Curve> curves = new List<Curve>();
                foreach (Edge edge in edgeArray)
                {
                    Curve curve = edge.AsCurveFollowingFace(topFace);
                    curves.Add(curve);
                }
                buttomCurves = CurveLoop.Create(curves);
            }
            // 干涉元件頂面
            foreach (EdgeArray edgeArray in topFace.EdgeLoops)
            {
                List<Curve> curves = new List<Curve>();
                foreach (Edge edge in edgeArray)
                {
                    Curve curve = edge.AsCurveFollowingFace(topFace);
                    IList<XYZ> curveXYZs = new List<XYZ>();
                    foreach (XYZ curveXYZ in curve.Tessellate())
                    {
                        XYZ xyz = new XYZ(curveXYZ.X, curveXYZ.Y, curveXYZ.Z + height);
                        curveXYZs.Add(xyz);
                    }
                    curve = NurbSpline.CreateCurve(HermiteSpline.Create(curveXYZs, false));
                    curves.Add(curve);
                }
                topCurves = CurveLoop.Create(curves);
            }

            IList<CurveLoop> curveLoops = new List<CurveLoop>();
            curveLoops.Add(buttomCurves);
            curveLoops.Add(topCurves);
            Solid solid = null;
            try
            {
                SolidOptions options = new SolidOptions(materialId, materialId);
                solid = GeometryCreationUtilities.CreateLoftGeometry(curveLoops, options);
                DirectShape ds = DirectShape.CreateElement(doc, materialId);
                ds.ApplicationId = "Application id";
                ds.ApplicationDataId = "Geometry object id";
                //try { ds.LookupParameter("編號").Set(number); }
                //catch (Exception ex) { string error = floorId + " 無「編號」參數\n" + ex.Message + "\n" + ex.ToString(); }
                ds.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(floorId + "：車道板干涉元件");
                ds.SetShape(new List<GeometryObject>() { solid });
            }
            catch (Exception e) { string error = e.Message + "\n" + e.ToString(); }

            return solid;
        }
        // RuledFace建立Face的封閉曲線(座標點不用重複)
        private void GetTessellatedSolid(Document doc, Face topFace, double height, ElementId materialId)
        {
            TessellatedShapeBuilder builder = new TessellatedShapeBuilder();
            builder.OpenConnectedFaceSet(true);

            // 干涉元件底面
            List<XYZ> bottomTriFace = new List<XYZ>(3);
            Mesh bottomMesh = topFace.Triangulate();
            int bottomTriCount = bottomMesh.NumTriangles;
            for (int i = 0; i < bottomTriCount; i++)
            {
                bottomTriFace.Clear();
                for (int n = 0; n < 3; n++)
                {
                    bottomTriFace.Add(bottomMesh.get_Triangle(i).get_Vertex(n));
                }
                builder.AddFace(new TessellatedFace(bottomTriFace, materialId));
            }
            // 干涉元件頂面
            List<XYZ> topTriFace = new List<XYZ>(3);
            Mesh topMesh = topFace.Triangulate();
            int topTriCount = topMesh.NumTriangles;
            for (int i = 0; i < topTriCount; i++)
            {
                topTriFace.Clear();
                for (int n = 0; n < 3; n++)
                {
                    XYZ xyz = topMesh.get_Triangle(i).get_Vertex(n);
                    XYZ topXYZ = new XYZ(xyz.X, xyz.Y, xyz.Z + height);
                    topTriFace.Add(topXYZ);
                }
                builder.AddFace(new TessellatedFace(topTriFace, materialId));
            }
            //// 干涉元件側面
            //IList<XYZ> sideXYZs = new List<XYZ>();
            //List<XYZ> sideTriFace = new List<XYZ>(3);
            //foreach (EdgeArray edgeArray in topFace.EdgeLoops)
            //{
            //    foreach (Edge edge in edgeArray)
            //    {
            //        sideXYZs.Clear();
            //        Curve curve = edge.AsCurveFollowingFace(topFace);
            //        int curveTessellateCount = curve.Tessellate().Count;
            //        for (int i = 0; i < curveTessellateCount; i++)
            //        {
            //            XYZ xyz = curve.Tessellate()[i];
            //            sideXYZs.Add(xyz);
            //        }
            //        for (int i = curveTessellateCount - 1; i >= 0 ; i--)
            //        {
            //            XYZ xyz = new XYZ(curve.Tessellate()[i].X, curve.Tessellate()[i].Y, curve.Tessellate()[i].Z + height);
            //            sideXYZs.Add(xyz);
            //        }
            //        for(int i = 0; i < sideXYZs.Count; i++)
            //        {
            //            //for (int n = 0; n < 3; n++)
            //            //{
            //                XYZ xyz = sideXYZs[i];
            //                sideTriFace.Add(xyz);
            //            //}
            //        }
            //        builder.AddFace(new TessellatedFace(sideTriFace, materialId));
            //    }
            //}

            builder.CloseConnectedFaceSet();

            //return builder.Build(TessellatedShapeBuilderTarget.Solid, TessellatedShapeBuilderFallback.Abort, materialId); // 2016

            //builder.Fallback = TessellatedShapeBuilderFallback.Abort;
            //builder.Target = TessellatedShapeBuilderTarget.Solid;

            builder.Build(); // 2020
            TessellatedShapeBuilderResult result = builder.GetBuildResult();
            DirectShape ds = DirectShape.CreateElement(doc, materialId);
            ds.ApplicationId = "Application id";
            ds.ApplicationDataId = "Geometry object id";
            ds.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("車道板干涉元件");
            ds.SetShape(result.GetGeometricalObjects());
        }
        // PlanarFace建立Face的封閉曲線(座標點不用重複)
        private void CreateCrushFaces(Document doc, Face topFace, double height, ElementId materialId)
        {
            List<XYZ> edgeStartPoints = new List<XYZ>();
            IList<XYZ> curveXYZs = new List<XYZ>();
            TessellatedShapeBuilder builder = new TessellatedShapeBuilder();
            builder.OpenConnectedFaceSet(true);

            // 干涉元件底面
            foreach (EdgeArray edgeArray in topFace.EdgeLoops)
            {
                curveXYZs.Clear();
                foreach (Edge edge in edgeArray)
                {
                    Curve curve = edge.AsCurveFollowingFace(topFace);
                    int curveTessellateCount = curve.Tessellate().Count;
                    for (int i = 0; i < curveTessellateCount - 1; i++)
                    {
                        XYZ xyz = ChangeXYZ(curve, i, 0);
                        curveXYZs.Add(xyz);
                        if (i.Equals(0)) { edgeStartPoints.Add(xyz); } // 儲存各線段的起點
                    }
                }
                builder.AddFace(new TessellatedFace(curveXYZs, materialId));
            }
            // 干涉元件頂面
            foreach (EdgeArray edgeArray in topFace.EdgeLoops)
            {
                curveXYZs.Clear();
                foreach (Edge edge in edgeArray)
                {
                    Curve curve = edge.AsCurveFollowingFace(topFace);
                    int curveTessellateCount = curve.Tessellate().Count;
                    for (int i = 0; i < curveTessellateCount - 1; i++)
                    {
                        XYZ xyz = ChangeXYZ(curve, i, height);
                        curveXYZs.Add(xyz);
                    }
                }
                builder.AddFace(new TessellatedFace(curveXYZs, materialId));
            }
            // 干涉元件側面
            foreach (EdgeArray edgeArray in topFace.EdgeLoops)
            {
                foreach (Edge edge in edgeArray)
                {
                    List<XYZ> saveXYZs = new List<XYZ>();
                    curveXYZs.Clear();
                    Curve curve = edge.AsCurveFollowingFace(topFace);
                    int curveTessellateCount = curve.Tessellate().Count;
                    XYZ startXYZ = ChangeXYZ(curve, 0, 0);
                    for (int i = 0; i < curveTessellateCount - 1; i++)
                    {
                        XYZ xyz = ChangeXYZ(curve, i, 0);
                        curveXYZs.Add(xyz);
                        saveXYZs.Add(xyz);
                    }
                    // 新增終點
                    int index = edgeStartPoints.FindIndex(x => x.X.Equals(startXYZ.X) && x.Y.Equals(startXYZ.Y) && x.Z.Equals(startXYZ.Z));
                    if (index == edgeStartPoints.Count - 1) { curveXYZs.Add(edgeStartPoints[0]); saveXYZs.Add(edgeStartPoints[0]); }
                    else { curveXYZs.Add(edgeStartPoints[index + 1]); saveXYZs.Add(edgeStartPoints[index + 1]); }
                    for (int i = saveXYZs.Count - 1; i >= 0; i--)
                    {
                        XYZ xyz = new XYZ(Math.Round(saveXYZs[i].X, 8, MidpointRounding.AwayFromZero),
                                          Math.Round(saveXYZs[i].Y, 8, MidpointRounding.AwayFromZero),
                                          Math.Round(saveXYZs[i].Z + height, 8, MidpointRounding.AwayFromZero));
                        curveXYZs.Add(xyz);
                    }
                    builder.AddFace(new TessellatedFace(curveXYZs, materialId));
                }
            }
            try
            {
                builder.CloseConnectedFaceSet();
                //builder.Target = TessellatedShapeBuilderTarget.Solid;
                //builder.Fallback = TessellatedShapeBuilderFallback.Abort;
                //builder.Target = TessellatedShapeBuilderTarget.AnyGeometry;
                //builder.Fallback = TessellatedShapeBuilderFallback.Mesh;
                builder.Build();

                TessellatedShapeBuilderResult result = builder.GetBuildResult();
                DirectShape ds = DirectShape.CreateElement(doc, materialId);
                ds.ApplicationId = "Application id";
                ds.ApplicationDataId = "Geometry object id";
                ds.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("車道板干涉元件");
                ds.SetShape(result.GetGeometricalObjects());
            }
            catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
        }
        // 回傳Curve修正的座標
        private XYZ ChangeXYZ(Curve curve, int i, double height)
        {
            XYZ xyz = new XYZ(Math.Round(curve.Tessellate()[i].X, 8, MidpointRounding.AwayFromZero),
                              Math.Round(curve.Tessellate()[i].Y, 8, MidpointRounding.AwayFromZero),
                              Math.Round(curve.Tessellate()[i].Z + height, 8, MidpointRounding.AwayFromZero));
            return xyz;
        }
        // 聯集所有Crush Solid
        private Solid UnionSolids(IList<Solid> solids, Solid hostSolid)
        {
            Solid unionSolid = null;
            foreach (Solid subSolid in solids)
            {
                if (subSolid.Volume != 0)
                {
                    try
                    {
                        unionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(hostSolid, subSolid, BooleanOperationsType.Union);
                        hostSolid = unionSolid;
                    }
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException) { }
                }
            }

            return hostSolid;
        }
        // 出衝突報告
        private List<CrushElemInfo> CrushReport(Document doc, LogicalOrFilter logicalOrFilter, List<Floor> floors)
        {
            List<CrushElemInfo> crushElemInfos = new List<CrushElemInfo>();
            List<DirectShape> crushFloors = new FilteredElementCollector(doc).WherePasses(logicalOrFilter).WhereElementIsNotElementType().Where(x => x is DirectShape).Cast<DirectShape>()
                                            .Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString() != null)
                                            .Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Contains("車道板干涉元件")).ToList();
            List<ElementId> crushFloorIds = crushFloors.Select(x => x.Id).ToList();
            foreach (Floor floor in floors) { crushFloorIds.Add(floor.Id); }
            foreach (DirectShape crushFloor in crushFloors)
            {
                try
                {
                    int floorId = Convert.ToInt32(crushFloor.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split('：')[0]);
                    List<Solid> solidList = GetSolids(doc, crushFloor); // 儲存所有車道板的Solid
                    foreach (Solid solid in solidList)
                    {
                        IList<Element> elems = new FilteredElementCollector(doc).WherePasses(new ElementIntersectsSolidFilter(solid)).Excluding(crushFloorIds).ToList();
                        if (elems.Count > 0)
                        {
                            string hostName = doc.GetElement(new ElementId(floorId)).Name + "：" + floorId;
                            CrushElemInfo crushElemInfo = new CrushElemInfo();
                            if (crushElemInfos.Where(x => x.hostName.Equals(hostName)).FirstOrDefault() != null)
                            {
                                crushElemInfo = crushElemInfos.Where(x => x.hostName.Equals(hostName)).FirstOrDefault();
                            }
                            else
                            {
                                crushElemInfo.hostName = hostName;
                            }
                            foreach (Element elem in elems)
                            {
                                crushElemInfo.crushElemName.Add(elem.Name + "：" + elem.Id + "、干涉車道板：" + crushFloor.Id);
                            }
                            if (crushElemInfos.Where(x => x.hostName.Equals(hostName)).FirstOrDefault() == null)
                            {
                                crushElemInfos.Add(crushElemInfo);
                            }
                        }
                    }
                }
                catch (Exception) { }
            }

            return crushElemInfos;
        }
        // 儲存所有Face的封閉曲線
        private List<Curve> SaveFaceCurveLoop(Face topFace, double height)
        {
            Transform transform = new Transform(Transform.CreateTranslation(new XYZ(0, 0, 0)));
            List<Curve> resultList = new List<Curve>();
            IList<CurveLoop> curveLoops = topFace.GetEdgesAsCurveLoops(); // 頂面的封閉曲線
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
        public static List<GeometryObject> GetElementSolids(Floor floor, Document doc, double height)
        {
            List<GeometryObject> list = new List<GeometryObject>();
            Options options = new Options();
            options.ComputeReferences = true;
            options.DetailLevel = ((doc.ActiveView != null) ? doc.ActiveView.DetailLevel : ViewDetailLevel.Medium);
            list = GetSolidsAndCurves(floor, height, list, options);
            return list;
        }
        public static List<GeometryObject> GetSolidsAndCurves(Floor floor, double height, List<GeometryObject> resultList, Options geoOpt)
        {
            using (IEnumerator<GeometryObject> enumerator = floor.get_Geometry(geoOpt).GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    List<Face> topFaces = new List<Face>();
                    List<CurveLoop> curveLoop = new List<CurveLoop>();

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
                    if (enumerator.Current is Solid)
                    {
                        Solid solid = enumerator.Current as Solid;
                        if (solid != null && solid.Volume != 0.0 && solid.SurfaceArea != 0.0)
                        {
                            // 找到Solid所有的頂面
                            foreach (Face face in solid.Faces)
                            {
                                CurveLoop baseLoop = new CurveLoop();
                                List<Curve> curves = new List<Curve>();
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

        public string GetName()
        {
            throw new NotImplementedException();
        }
    }
}
