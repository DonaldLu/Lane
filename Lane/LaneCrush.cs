using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using Create = Autodesk.Revit.Creation;

namespace Lane
{
    [Transaction(TransactionMode.Manual)]
    public class LaneCrush : IExternalCommand
    {
        public class StartEndPoint
        {
            public XYZ start = new XYZ();
            public XYZ end = new XYZ();
            public XYZ other = new XYZ();
            public Double perimeter = 0.0;
            public string type = string.Empty;
        }
        public static List<StartEndPoint> StartEndList = null;
        List<List<StartEndPoint>> list = new List<List<StartEndPoint>>();
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            List<FamilySymbol> familySymbols = new List<FamilySymbol>();
            FloorType floorType = new FilteredElementCollector(doc).OfClass(typeof(FloorType)).Cast<FloorType>().FirstOrDefault();
            List<Level> levels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList();
            Level level = null;
            foreach(Level lvl in levels)
            {
                if (lvl.Name.Equals("B2FL"))
                {
                    level = lvl;
                    break;
                }
            }

            IList<ElementFilter> elementFilters = new List<ElementFilter>(); // 清空過濾器
            ElementCategoryFilter roomFilter = new ElementCategoryFilter(BuiltInCategory.OST_Floors); // 房間
            elementFilters.Add(roomFilter);
            LogicalOrFilter logicalOrFilter = new LogicalOrFilter(elementFilters);
            List<Floor> floors = new FilteredElementCollector(doc).WherePasses(logicalOrFilter).WhereElementIsNotElementType()
                                .Where(x => x.LookupParameter("樓板高程") != null)
                                .Where(x => x.LookupParameter("樓板高程").AsValueString().Equals("車道板")).Cast<Floor>().ToList();
            // 儲存所有車道板的Solid
            List<Solid> solidList = new List<Solid>();
            foreach (Floor floor in floors)
            {
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
            }
            // 將所有Solid聯集
            Solid unionSolid = UnionSolids(solidList, solidList[0]);
            List<CurveArray> curvesList = new List<CurveArray>();
            if (unionSolid != null)
            {
                foreach (Face face in unionSolid.Faces)
                {
                    PlanarFace pf = face as PlanarFace;
                    if (pf != null)
                    {
                        //if (pf.FaceNormal.IsAlmostEqualTo(XYZ.BasisZ))
                        //{
                            EdgeArrayArray loops = pf.EdgeLoops;
                            foreach (EdgeArray loop in loops)
                            {
                                CurveArray curves = app.Create.NewCurveArray();
                                foreach (Edge edge in loop)
                                {
                                    IList<XYZ> edgePoints = edge.Tessellate();
                                    Arc arc = null;
                                    if (edgePoints.Count > 2)
                                    {
                                        arc = Arc.Create(edgePoints[0], edgePoints[edgePoints.Count - 1], edgePoints[1]);
                                        if (arc != null)
                                        {
                                            curves.Append(arc);
                                        }
                                    }
                                    else
                                    {
                                        Line line = Line.CreateBound(edgePoints[0], edgePoints[1]);
                                        curves.Append(line);
                                    }
                                }
                                curvesList.Add(curves);
                            }
                        //}
                    }
                }
                // 找到Solid所有封閉迴圈邊界的座標點            
                foreach (CurveArray curves in curvesList)
                {
                    StartEndList = new List<StartEndPoint>();
                    foreach (Curve curve in curves)
                    {
                        string type = curve.GetType().Name;
                        StartEndPoint startEndPoint = new StartEndPoint();
                        startEndPoint.type = type;

                        // 將X、Y、Z先轉成double後在存

                        if (type == "Arc")
                        {
                            Arc arc = curve as Arc;
                            double startX = Math.Round(arc.Tessellate()[0].X, 8);
                            double startY = Math.Round(arc.Tessellate()[0].Y, 8);
                            double startZ = Math.Round(arc.Tessellate()[0].Z, 8);
                            double endX = Math.Round(arc.Tessellate()[arc.Tessellate().Count - 1].X, 8);
                            double endY = Math.Round(arc.Tessellate()[arc.Tessellate().Count - 1].Y, 8);
                            double endZ = Math.Round(arc.Tessellate()[arc.Tessellate().Count - 1].Z, 8);
                            double otherX = Math.Round(arc.Tessellate()[1].X, 8);
                            double otherY = Math.Round(arc.Tessellate()[1].Y, 8);
                            double otherZ = Math.Round(arc.Tessellate()[1].Z, 8);

                            startEndPoint.start = new XYZ(startX, startY, startZ);
                            startEndPoint.end = new XYZ(endX, endY, endZ);
                            startEndPoint.other = new XYZ(otherX, otherY, otherZ);
                            startEndPoint.perimeter = curve.Length;
                            StartEndList.Add(startEndPoint);
                        }
                        else
                        {
                            double startX = Math.Round(curve.Tessellate()[0].X, 8);
                            double startY = Math.Round(curve.Tessellate()[0].Y, 8);
                            double startZ = Math.Round(curve.Tessellate()[0].Z, 8);
                            double endX = Math.Round(curve.Tessellate()[curve.Tessellate().Count - 1].X, 8);
                            double endY = Math.Round(curve.Tessellate()[curve.Tessellate().Count - 1].Y, 8);
                            double endZ = Math.Round(curve.Tessellate()[curve.Tessellate().Count - 1].Z, 8);

                            startEndPoint.start = new XYZ(startX, startY, startZ);
                            startEndPoint.end = new XYZ(endX, endY, endZ);
                            startEndPoint.perimeter = curve.Length;
                            StartEndList.Add(startEndPoint);
                        }
                    }
                    list.Add(StartEndList);
                }

                // 移除list內最長的周長
                int removeItem = RemoveLogestPerimeter(list);
                list.RemoveAt(removeItem);
            }

            // 新增樓板
            using (Transaction trans = new Transaction(doc, "建立樓板"))
            {
                try
                {
                    trans.Start();
                    foreach (List<StartEndPoint> sepList in list)
                    {
                        // 排序座標點, 終點接起點
                        PointSort(sepList);
                        // 儲存要建樓板的封閉區域
                        CurveArray curves = SaveCurves(sepList);
                        List<CurveLoop> curveLoops = new List<CurveLoop>();
                        CurveLoop curveLoop = new CurveLoop();
                        foreach (Curve curve in curves)
                        {
                            curveLoop.Append(curve);
                        }
                        curveLoops.Add(curveLoop);
                        try
                        {
                            // 新增樓板
                            Floor createFloor = Floor.Create(doc, curveLoops, floorType.Id, level.Id);
                        }
                        catch (Exception)
                        {

                        }
                    }
                    trans.Commit();
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Revit", ex.ToString());
                }
            }
            return Result.Succeeded;
        }
        private List<List<Curve>> GetBoundarySegment(Floor floor)
        {
            List<List<Curve>> boundarySegment = new List<List<Curve>>();
            try
            {
                // 1.讀取Geometry Option
                Options options = new Options();
                //options.View = doc.GetElement(room.Level.FindAssociatedPlanViewId()) as Autodesk.Revit.DB.View;
                options.DetailLevel = ViewDetailLevel.Medium;
                options.ComputeReferences = true;
                options.IncludeNonVisibleObjects = true;
                // 得到幾何元素
                GeometryElement geomElem = floor.get_Geometry(options);
                List<Solid> solids = GeometrySolids(geomElem);
                Solid solid = solids.FirstOrDefault();
                List<Face> planarFaces = new List<Face>();
                foreach (Face face in solid.Faces)
                {
                    if (face is PlanarFace)
                    {
                        PlanarFace planarFace = (PlanarFace)face;
                        if (planarFace.FaceNormal.Z > 0)
                        {
                            planarFaces.Add(planarFace);
                        }
                    }
                }
                foreach (PlanarFace planarFace in planarFaces)
                {
                    foreach (EdgeArray edgeArray in planarFace.EdgeLoops)
                    {
                        // 先將所有的邊儲存起來
                        List<Curve> curveLoop = new List<Curve>();
                        foreach (Edge edge in edgeArray)
                        {
                            Curve curve = edge.AsCurve();
                            curveLoop.Add(curve);
                        }
                        boundarySegment.Add(curveLoop);
                    }
                }
                // 從Solid查詢干涉到的元件, 找到樓梯、電扶梯
                //foreach (Solid roomSolid in solids)
                //{
                //    FindTheIntersectElems(doc, roomSolid, elementInfo);
                //}
            }
            catch (Exception ex)
            {
                string error = ex.Message + "\n" + ex.ToString();
            }

            return boundarySegment;
        }
        // 取得車道的Solid
        private static List<Solid> GeometrySolids(GeometryObject geoObj)
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
        // 移除周長最長的sepList
        private int RemoveLogestPerimeter(List<List<StartEndPoint>> list)
        {
            double longest = 0.0;
            int count = 0;
            int removeItem = 0;
            foreach (List<StartEndPoint> sepList in list)
            {
                double perimeter = 0.0;
                foreach (StartEndPoint item in sepList)
                {
                    perimeter += item.perimeter;
                }
                if (perimeter > longest)
                {
                    longest = perimeter;
                    removeItem = count;
                }
                count++;
            }
            return removeItem;
        }
        // 排序座標點, 終點接起點
        private void PointSort(List<StartEndPoint> sepList)
        {
            XYZ newStart = new XYZ();
            XYZ newEnd = new XYZ();

            if (sepList[0].start.X.ToString() == sepList[1].start.X.ToString() &&
               sepList[0].start.Y.ToString() == sepList[1].start.Y.ToString() &&
               sepList[0].start.Z.ToString() == sepList[1].start.Z.ToString())
            {
                newStart = sepList[0].end;
                newEnd = sepList[0].start;
                sepList[0].start = newStart;
                sepList[0].end = newEnd;
            }
            else if (sepList[0].start.X.ToString() == sepList[1].end.X.ToString() &&
                    sepList[0].start.Y.ToString() == sepList[1].end.Y.ToString() &&
                    sepList[0].start.Z.ToString() == sepList[1].end.Z.ToString())
            {
                newStart = sepList[0].end;
                newEnd = sepList[0].start;
                sepList[0].start = newStart;
                sepList[0].end = newEnd;
            }

            for (int i = 1; i < sepList.Count; i++)
            {
                newStart = new XYZ();
                newEnd = new XYZ();

                if (sepList[i].end.X.ToString() == sepList[i - 1].end.X.ToString() &&
                    sepList[i].end.Y.ToString() == sepList[i - 1].end.Y.ToString() &&
                    sepList[i].end.Z.ToString() == sepList[i - 1].end.Z.ToString())
                {
                    newStart = sepList[i].end;
                    newEnd = sepList[i].start;
                    sepList[i].start = newStart;
                    sepList[i].end = newEnd;
                }
            }
        }
        // 儲存要建樓板的封閉區域
        private CurveArray SaveCurves(List<StartEndPoint> sepList)
        {
            CurveArray curves = new CurveArray();
            foreach (StartEndPoint edge in sepList)
            {
                if (edge.type == "Arc")
                {
                    Arc arc = Arc.Create(edge.start, edge.end, edge.other);
                    curves.Append(arc);
                }
                else
                {
                    Line line = Line.CreateBound(edge.start, edge.end);
                    curves.Append(line);
                }
            }
            return curves;
        }
    }
}
