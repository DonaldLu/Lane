using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Visual;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Media.Media3D;
using Material = Autodesk.Revit.DB.Material;

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
            List<Floor> floors = new FilteredElementCollector(doc).WherePasses(logicalOrFilter).WhereElementIsNotElementType()
                                .Where(x => x.LookupParameter("樓板高程") != null)
                                .Where(x => x.LookupParameter("樓板高程").AsValueString().Equals("車道板")).Cast<Floor>().ToList();
            foreach(Floor floor in floors)
            {
                List<GeometryObject> list = GetElementSolids(floor, doc);
                DirectShape directShape = DirectShape.CreateElement(doc, floor.Category.Id);
                directShape.ApplicationId = "Test";
                directShape.ApplicationDataId = "Test.com";
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
                }

                //// 取得樓板的頂面輪廓
                //List<Solid> solidList = GetRoadwayPlankSolids(floor); // 儲存所有車道板的Solid
                //List<Face> topFaces = GetTopFaces(solidList); // 回傳Solid的頂面
                //try
                //{
                //    ElementId materialId = floor.FloorType.StructuralMaterialId;
                //    CreateTessellatedShape(doc, materialId, topFaces);
                //}
                //catch (Exception)
                //{
                //    throw new Exception("創建幾何元件失敗.");
                //}
            }

            //TransactionGroup tranGrp = new TransactionGroup(doc, "路徑篩選");
            //tranGrp.Start();
            //using (Transaction trans = new Transaction(doc, "生成干涉模型"))
            //{
            //    trans.Start();
            //    // 取得樓板的頂面輪廓
            //    foreach (Floor floor in floors)
            //    {
            //        List<Solid> solidList = GetRoadwayPlankSolids(floor); // 儲存所有車道板的Solid
            //        List<Face> topFaces = GetTopFaces(solidList); // 回傳Solid的頂面
            //        try
            //        {
            //            // 建立擠出樓板
            //            CreateCubicDirectShape(doc, topFaces);
            //        }
            //        catch (Exception)
            //        {
            //            throw new Exception("創建幾何元件失敗.");
            //        }
            //    }
            //    trans.Commit();
            //}
            //tranGrp.Assimilate();

            return Result.Succeeded;
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
                        Transform transform = geometryInstance.Transform;
                        Solid solid = enumerator.Current as Solid;
                        if (solid != null && solid.Volume != 0.0 && solid.SurfaceArea != 0.0)
                        {
                            resultList.Add(SolidUtils.CreateTransformed(solid, transform));
                        }
                    }
                }
            }
            return resultList;
        }
        // 儲存所有車道板的Solid
        private List<Solid> GetRoadwayPlankSolids(Floor floor)
        {
            List<Solid> solidList = new List<Solid>();

            // 1.讀取Geometry Option
            Options options = new Options();
            //options.View = doc.GetElement(room.Level.FindAssociatedPlanViewId()) as Autodesk.Revit.DB.View;
            options.DetailLevel = ViewDetailLevel.Medium;
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
        // 取得頂面
        private List<Face> GetTopFaces(List<Solid> solidList)
        {
            List<Face> topFaces = new List<Face>();
            foreach(Solid solid in solidList)
            {
                foreach (Face face in solid.Faces)
                {
                    if(face is PlanarFace)
                    {
                        PlanarFace planarFace = face as PlanarFace;
                        if (planarFace != null)
                        {
                            //if (planarFace.FaceNormal.Z > 0.0)
                            //{
                                topFaces.Add(face);
                            //}
                        }
                    }
                    //else if (face is CylindricalFace)
                    //{
                    //    CylindricalFace cylindricalFace = face as CylindricalFace;
                    //    if (cylindricalFace != null)
                    //    {
                    //        double faceTZ = face.ComputeNormal(new UV(0.5, 0.5)).Z;
                    //        if (cylindricalFace.Axis.Z > 0.0)
                    //        {
                    //            topFaces.Add(face);
                    //        }
                    //    }
                    //}
                    //else if(face is RuledFace)
                    //{
                    //    RuledFace ruledFace = face as RuledFace;
                    //    if(ruledFace != null)
                    //    {
                    //        //double faceTZ = face.ComputeNormal(new UV(0.5, 0.5)).Z;
                    //        //if (faceTZ > 0.0)
                    //        //{
                    //            topFaces.Add(face);
                    //        //}
                    //    }
                    //}
                }
            }
            return topFaces;
        }
        // 使用TessellatedShapeBuilder建立面
        // Create a pyramid-shaped DirectShape using given material for the faces
        private void CreateTessellatedShape(Document doc, ElementId materialId, List<Face> topFaces)
        {
            List<XYZ> loopVertices = new List<XYZ>(4);

            TessellatedShapeBuilder builder = new TessellatedShapeBuilder();

            builder.OpenConnectedFaceSet(true);
            double length = 4.0; 
            double height = UnitUtils.ConvertToInternalUnits(210, UnitTypeId.Centimeters);

            foreach (Face topFace in topFaces)
            {
                loopVertices.Clear();
                if (topFace is PlanarFace)
                {
                    List<StartEndPoint> startEndPointList = new List<StartEndPoint>();
                    EdgeArrayArray edgeArrayArray = topFace.EdgeLoops;
                    foreach (EdgeArray edgeArray in edgeArrayArray)
                    {
                        foreach (Edge edge in edgeArray)
                        {
                            XYZ startPoint = new XYZ(edge.Tessellate()[0].X, edge.Tessellate()[0].Y, edge.Tessellate()[0].Z + height);
                            XYZ endPoint = new XYZ(edge.Tessellate()[edge.Tessellate().Count - 1].X, edge.Tessellate()[edge.Tessellate().Count - 1].Y, edge.Tessellate()[edge.Tessellate().Count - 1].Z + height);
                            StartEndPoint startEndPoint = new StartEndPoint();
                            startEndPoint.start = startPoint;
                            startEndPoint.end = endPoint;
                            startEndPoint.type = "Line";
                            startEndPointList.Add(startEndPoint);
                            //foreach (XYZ xyz in edge.Tessellate())
                            //{

                            //    //XYZ romoveXYZ = new XYZ(xyz.X, xyz.Y, xyz.Z + height);
                            //    //// 重複的座標點則不儲存
                            //    //XYZ test = loopVertices.Where(x => Math.Round(x.X, 8, MidpointRounding.AwayFromZero).Equals(Math.Round(romoveXYZ.X, 8, MidpointRounding.AwayFromZero)) && 
                            //    //                                   Math.Round(x.Y, 8, MidpointRounding.AwayFromZero).Equals(Math.Round(romoveXYZ.Y, 8, MidpointRounding.AwayFromZero)) &&
                            //    //                                   Math.Round(x.Z, 8, MidpointRounding.AwayFromZero).Equals(Math.Round(romoveXYZ.Z, 8, MidpointRounding.AwayFromZero))).FirstOrDefault();
                            //    //if (test == null)
                            //    //{
                            //    //    loopVertices.Add(romoveXYZ);
                            //    //}
                            //}
                        }
                        PointSort(startEndPointList); // 排序座標點, 終點接起點
                        foreach (StartEndPoint startEndPoint in startEndPointList)
                        {
                            XYZ startPoint = new XYZ(startEndPoint.start.X, startEndPoint.start.Y, startEndPoint.start.Z);
                            loopVertices.Add(startPoint);
                        }
                        builder.AddFace(new TessellatedFace(loopVertices, materialId));
                    }
                }
                //else if (topFace is RuledFace)
                //{
                //    EdgeArrayArray edgeArrayArray = topFace.EdgeLoops;
                //    foreach(EdgeArray edgeArray in edgeArrayArray)
                //    {
                //        foreach(Edge edge in edgeArray)
                //        {
                //            foreach(XYZ xyz in edge.Tessellate())
                //            {
                //                XYZ romoveXYZ = new XYZ(xyz.X, xyz.Y, xyz.Z + height);
                //                loopVertices.Add(xyz);
                //            }
                //            //if (edge.AsCurve() is HermiteSpline)
                //            //{

                //            //}
                //            //else if (edge.AsCurve() is Line)
                //            //{

                //            //}
                //        }
                //        builder.AddFace(new TessellatedFace(loopVertices, materialId));
                //    }
                //}
            }

            //XYZ basePt1 = XYZ.Zero;
            //XYZ basePt2 = new XYZ(length, 0, 0);
            //XYZ basePt3 = new XYZ(length, length, 0);
            //XYZ basePt4 = new XYZ(0, length, 0);
            //XYZ apex = new XYZ(length / 2, length / 2, height);

            //loopVertices.Add(basePt1);
            //loopVertices.Add(basePt2);
            //loopVertices.Add(basePt3);
            //loopVertices.Add(basePt4);
            //builder.AddFace(new TessellatedFace(loopVertices, materialId));

            //loopVertices.Clear();
            //loopVertices.Add(basePt1);
            //loopVertices.Add(apex);
            //loopVertices.Add(basePt2);
            //builder.AddFace(new TessellatedFace(loopVertices, materialId));

            //loopVertices.Clear();
            //loopVertices.Add(basePt2);
            //loopVertices.Add(apex);
            //loopVertices.Add(basePt3);
            //builder.AddFace(new TessellatedFace(loopVertices, materialId));

            //loopVertices.Clear();
            //loopVertices.Add(basePt3);
            //loopVertices.Add(apex);
            //loopVertices.Add(basePt4);
            //builder.AddFace(new TessellatedFace(loopVertices, materialId));

            //loopVertices.Clear();
            //loopVertices.Add(basePt4);
            //loopVertices.Add(apex);
            //loopVertices.Add(basePt1);
            //builder.AddFace(new TessellatedFace(loopVertices, materialId));

            builder.CloseConnectedFaceSet();
            builder.Target = TessellatedShapeBuilderTarget.Solid;
            builder.Fallback = TessellatedShapeBuilderFallback.Abort;
            builder.Build();

            TessellatedShapeBuilderResult result = builder.GetBuildResult();

            using (Transaction t = new Transaction(doc, "Create tessellated direct shape"))
            {
                t.Start();

                DirectShape ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
                ds.ApplicationId = "Application id";
                ds.ApplicationDataId = "Geometry object id";

                ds.SetShape(result.GetGeometricalObjects());
                t.Commit();
            }
        }
        // 擠出面
        public void CreateCubicDirectShape(Document doc, List<Face> topFaces)
        {
            foreach (Face topFace in topFaces)
            {
                List<StartEndPoint> startEndPointList = new List<StartEndPoint>();
                List<Curve> profile = new List<Curve>();
                foreach (EdgeArray edgeArray in topFace.EdgeLoops)
                {
                    foreach (Edge edge in edgeArray)
                    {
                        if (edge.AsCurve() is Line)
                        {
                            XYZ startPoint = new XYZ(edge.Tessellate()[0].X, edge.Tessellate()[0].Y, edge.Tessellate()[0].Z);
                            XYZ endPoint = new XYZ(edge.Tessellate()[edge.Tessellate().Count - 1].X, edge.Tessellate()[edge.Tessellate().Count - 1].Y, edge.Tessellate()[edge.Tessellate().Count - 1].Z);
                            StartEndPoint startEndPoint = new StartEndPoint();
                            startEndPoint.start = startPoint;
                            startEndPoint.end = endPoint;
                            startEndPoint.type = "Line";
                            startEndPointList.Add(startEndPoint);
                        }
                        else if (edge.AsCurve() is HermiteSpline)
                        {
                            HermiteSpline hermiteSpline = edge.AsCurve() as HermiteSpline;
                            XYZ startPoint = new XYZ(edge.Tessellate()[0].X, edge.Tessellate()[0].Y, edge.Tessellate()[0].Z);
                            XYZ endPoint = new XYZ(edge.Tessellate()[edge.Tessellate().Count - 1].X, edge.Tessellate()[edge.Tessellate().Count - 1].Y, edge.Tessellate()[edge.Tessellate().Count - 1].Z);
                            StartEndPoint startEndPoint = new StartEndPoint();
                            startEndPoint.start = startPoint;
                            startEndPoint.end = endPoint;
                            // 儲存中間的座標點
                            for(int i = 1; i < edge.Tessellate().Count - 1; i++)
                            {
                                startEndPoint.xyzs.Add(edge.Tessellate()[i]);
                            }
                            startEndPoint.type = "HermiteSpline";
                            startEndPoint.isPeriodic = hermiteSpline.IsPeriodic;
                            startEndPoint.tangents = hermiteSpline.Tangents.ToList();
                            startEndPointList.Add(startEndPoint);
                        }
                    }
                }
                PointSort(startEndPointList); // 排序座標點, 終點接起點
                foreach (StartEndPoint startEndPoint in startEndPointList)
                {
                    XYZ startPoint = new XYZ(startEndPoint.start.X, startEndPoint.start.Y, startEndPoint.start.Z);
                    XYZ endPoint = new XYZ(startEndPoint.end.X, startEndPoint.end.Y, startEndPoint.end.Z);
                    if (startEndPoint.type.Equals("Line"))
                    {
                        Line line = Line.CreateBound(startPoint, endPoint);
                        profile.Add(line);
                    }
                    else if (startEndPoint.type.Equals("HermiteSpline"))
                    {
                        startEndPoint.xyzs.Insert(0, startPoint);
                        startEndPoint.xyzs.Insert(startEndPoint.xyzs.Count, endPoint);
                        HermiteSpline hermiteSpline = HermiteSpline.Create(startEndPoint.xyzs, startEndPoint.isPeriodic);
                        profile.Add(hermiteSpline);
                        //for (int i = 0; i < startEndPoint.xyzs.Count - 1; i++)
                        //{
                        //    Line line = Line.CreateBound(startEndPoint.xyzs[i], startEndPoint.xyzs[i + 1]);
                        //    profile.Add(line);
                        //}
                    }
                }
                try
                {
                    CurveLoop curveLoop = new CurveLoop();
                    curveLoop = CurveLoop.Create(profile);
                    SolidOptions options = new SolidOptions(ElementId.InvalidElementId, ElementId.InvalidElementId);
                    double height = UnitUtils.ConvertToInternalUnits(210, UnitTypeId.Centimeters);
                    Solid cubic = GeometryCreationUtilities.CreateExtrusionGeometry(new CurveLoop[] { curveLoop }, XYZ.BasisZ, height);
                    DirectShape ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
                    ds.ApplicationId = "Application id";
                    ds.ApplicationDataId = "Geometry object id";
                    ds.SetShape(new GeometryObject[] { cubic });
                }
                catch (Exception ex)
                {
                    string error = ex.Message + "\n" + ex.ToString();
                }
            }
        }
        // 排序座標點, 終點接起點
        private void PointSort(List<StartEndPoint> sepList)
        {
            XYZ newStart = new XYZ();
            XYZ newEnd = new XYZ();

            if (Math.Round(sepList[0].start.X, 4, MidpointRounding.AwayFromZero) == Math.Round(sepList[1].start.X, 4, MidpointRounding.AwayFromZero) &&
                Math.Round(sepList[0].start.Y, 4, MidpointRounding.AwayFromZero) == Math.Round(sepList[1].start.Y, 4, MidpointRounding.AwayFromZero) &&
                Math.Round(sepList[0].start.Z, 4, MidpointRounding.AwayFromZero) == Math.Round(sepList[1].start.Z, 4, MidpointRounding.AwayFromZero))
            {
                newStart = sepList[0].end;
                newEnd = sepList[0].start;
                sepList[0].start = newStart;
                sepList[0].end = newEnd;
                if (sepList[0].type.Equals("HermiteSpline"))
                {
                    sepList[0].xyzs.Reverse();
                }
            }
            else if (Math.Round(sepList[0].start.X, 4, MidpointRounding.AwayFromZero) == Math.Round(sepList[1].end.X, 4, MidpointRounding.AwayFromZero) &&
                     Math.Round(sepList[0].start.Y, 4, MidpointRounding.AwayFromZero) == Math.Round(sepList[1].end.Y, 4, MidpointRounding.AwayFromZero) &&
                     Math.Round(sepList[0].start.Z, 4, MidpointRounding.AwayFromZero) == Math.Round(sepList[1].end.Z, 4, MidpointRounding.AwayFromZero))
            {
                newStart = sepList[0].end;
                newEnd = sepList[0].start;
                sepList[0].start = newStart;
                sepList[0].end = newEnd;
                if (sepList[0].type.Equals("HermiteSpline"))
                {
                    sepList[0].xyzs.Reverse();
                }
                newStart = sepList[1].end;
                newEnd = sepList[1].start;
                sepList[1].start = newStart;
                sepList[1].end = newEnd;
                if (sepList[1].type.Equals("HermiteSpline"))
                {
                    sepList[1].xyzs.Reverse();
                }
            }

            for (int i = 1; i < sepList.Count; i++)
            {
                newStart = new XYZ();
                newEnd = new XYZ();

                if (Math.Round(sepList[i].end.X, 4, MidpointRounding.AwayFromZero) == Math.Round(sepList[i - 1].end.X, 4, MidpointRounding.AwayFromZero) &&
                    Math.Round(sepList[i].end.Y, 4, MidpointRounding.AwayFromZero) == Math.Round(sepList[i - 1].end.Y, 4, MidpointRounding.AwayFromZero) &&
                    Math.Round(sepList[i].end.Z, 4, MidpointRounding.AwayFromZero) == Math.Round(sepList[i - 1].end.Z, 4, MidpointRounding.AwayFromZero))
                {
                    newStart = sepList[i].end;
                    newEnd = sepList[i].start;
                    sepList[i].start = newStart;
                    sepList[i].end = newEnd;
                    if(sepList[i].type.Equals("HermiteSpline"))
                    {
                        sepList[i].xyzs.Reverse();
                    }
                }
            }
        }
        // 將所有Solid聯集
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
                    catch (Exception)
                    {

                    }
                }
            }

            return hostSolid;
        }
    }
}