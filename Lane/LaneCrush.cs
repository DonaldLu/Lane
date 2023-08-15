using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Lane
{
    [Transaction(TransactionMode.Manual)]
    public class LaneCrush : IExternalCommand
    {
        private string rfaPath = @"D:\車道干涉\";
        private string rftPath = @"C:\ProgramData\Autodesk\RVT 2022\Family Templates\Traditional Chinese\公制通用模型.rft";
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

            FloorType floorType = new FilteredElementCollector(doc).OfClass(typeof(FloorType)).Cast<FloorType>().FirstOrDefault();
            Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().Where(x => x.Name.Equals("B2FL")).FirstOrDefault();

            IList<ElementFilter> elementFilters = new List<ElementFilter>(); // 清空過濾器
            ElementCategoryFilter roomFilter = new ElementCategoryFilter(BuiltInCategory.OST_Floors); // 房間
            elementFilters.Add(roomFilter);
            LogicalOrFilter logicalOrFilter = new LogicalOrFilter(elementFilters);
            List<Floor> floors = new FilteredElementCollector(doc).WherePasses(logicalOrFilter).WhereElementIsNotElementType()
                                .Where(x => x.LookupParameter("樓板高程") != null)
                                .Where(x => x.LookupParameter("樓板高程").AsValueString().Equals("車道板")).Cast<Floor>().ToList();

            TransactionGroup tranGrp = new TransactionGroup(doc, "路徑篩選");
            tranGrp.Start();

            int count = DeleteDirectory(rfaPath); // 刪除資料夾檔案
            using (Transaction trans = new Transaction(doc, "生成干涉模型"))
            {
                trans.Start();
                Document rftDoc = app.OpenDocumentFile(rftPath); // 讀取檔案到內存

                // 取得樓板的頂面輪廓
                foreach (Floor floor in floors)
                {
                    List<Solid> solidList = GetRoadwayPlankSolids(floor); // 儲存所有車道板的Solid
                    List<Face> topFaces = GetTopFaces(solidList); // 回傳Solid的頂面
                    try
                    {
                        // 建立擠出樓板
                        Extrusion rectExtrusion = ExtrusionFloors(rftDoc, topFaces); // 擠出面
                        // 創建成功的話則儲存族群元件, 並載入到專案中
                        if (null != rectExtrusion)
                        {
                            try
                            {
                                // 儲存干涉元件
                                rftDoc.SaveAs(rfaPath + "干涉元件" + count + ".rfa");
                                // 重新載入族群
                                doc.LoadFamilySymbol(rfaPath + "干涉元件" + count + ".rfa", "干涉元件" + count, new JtFamilyLoadOptions(), out FamilySymbol outFS);
                                count++;
                            }
                            catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                            {

                            }
                        }
                    }
                    catch (Exception)
                    {
                        throw new Exception("創建幾何元件失敗.");
                    }
                }
                trans.Commit();
            }
            tranGrp.Assimilate();

            return Result.Succeeded;
        }
        // 刪除資料夾檔案
        private static int DeleteDirectory(string target_dir)
        {
            int maxCount = 1; // 紀錄資料夾內火源最大的值
            List<int> sort = new List<int>();
            string[] nameSplit = new string[] { };
            string[] files = Directory.GetFiles(target_dir); // 找到所有檔案
            string[] dirs = Directory.GetDirectories(target_dir); // 找到所有子資料夾
            foreach (string file in files)
            {
                try
                {
                    File.SetAttributes(file, FileAttributes.Normal);
                    File.Delete(file);
                }
                catch (System.IO.IOException)
                {
                    try
                    {
                        nameSplit = file.Replace(target_dir + "干涉元件", "").Split('.');
                        int count = Convert.ToInt32(nameSplit[0]);
                        sort.Add(count);
                    }
                    catch (Exception) // 避免資料夾中有"干涉元件"外的其他檔案名稱
                    {

                    }
                }
            }
            foreach (string dir in dirs)
            {
                try
                {
                    DeleteDirectory(dir);
                }
                catch (System.IO.IOException)
                {

                }
            }
            //Directory.Delete(target_dir, false);
            try
            {
                maxCount = (from x in sort
                            select x).Max() + 1;
            }
            catch (System.InvalidOperationException)
            {

            }
            return maxCount;
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
                    PlanarFace pf = face as PlanarFace;
                    if (pf != null)
                    {
                        if(pf.FaceNormal.Z > 0.0)
                        {
                            topFaces.Add(face);
                        }
                        //double faceTZ = face.ComputeNormal(new UV(0.5, 0.5)).Z;
                        //if (faceTZ > 0.0)
                        //{
                        //    topFaces.Add(face);
                        //}
                    }
                }
            }
            return topFaces;
        }
        // 擠出面
        private Extrusion ExtrusionFloors(Document rftDoc, List<Face> topFaces)
        {
            Extrusion rectExtrusion = null;

            // 確認開啟族群編輯器
            if (true == rftDoc.IsFamilyDocument)
            {
                // 定義擠出的輪廓線
                foreach (Face topFace in topFaces)
                {
                    using (Transaction rftTrans = new Transaction(rftDoc, "擠出樓板"))
                    {
                        rftTrans.Start();
                        CurveArrArray curveArrArray = new CurveArrArray();
                        CurveArray curveArray = new CurveArray();
                        XYZ normal = XYZ.BasisZ;
                        XYZ origin = new XYZ();
                        if (topFace is PlanarFace)
                        {
                            PlanarFace planarFace = topFace as PlanarFace;
                            //normal = planarFace.FaceNormal;
                            origin = planarFace.Origin;
                        }
                        else if (topFace is CylindricalFace)
                        {
                            CylindricalFace cylindricalFace = topFace as CylindricalFace;
                            //normal = cylindricalFace.Axis;
                            origin = cylindricalFace.Origin;
                        }
                        //建立一個草繪平面
                        SketchPlane sketchPlane = CreateSketchPlane(rftDoc, normal, origin);
                        //sketchPlane = SketchPlane.Create(doc, topFace.Reference);
                        List<StartEndPoint> startEndPointList = new List<StartEndPoint>();
                        foreach (EdgeArray edgeArray in topFace.EdgeLoops)
                        {
                            foreach (Edge edge in edgeArray)
                            {
                                XYZ startPoint = new XYZ(edge.Tessellate()[0].X, edge.Tessellate()[0].Y, 0);
                                XYZ endPoint = new XYZ(edge.Tessellate()[1].X, edge.Tessellate()[1].Y, 0);
                                StartEndPoint startEndPoint = new StartEndPoint();
                                startEndPoint.start = startPoint;
                                startEndPoint.end = endPoint;
                                startEndPointList.Add(startEndPoint);
                            }
                        }
                        PointSort(startEndPointList); // 排序座標點, 終點接起點
                        foreach (StartEndPoint startEndPoint in startEndPointList)
                        {
                            XYZ startPoint = new XYZ(startEndPoint.start.X, startEndPoint.start.Y, startEndPoint.start.Z);
                            XYZ endPoint = new XYZ(startEndPoint.end.X, startEndPoint.end.Y, startEndPoint.end.Z);
                            Line line = Line.CreateBound(startPoint, endPoint);
                            curveArray.Append(line);
                        }
                        curveArrArray.Append(curveArray);

                        try
                        {
                            rectExtrusion = rftDoc.FamilyCreate.NewExtrusion(true, curveArrArray, sketchPlane, 10); // 創建矩形的擠出元件
                        }
                        catch(Exception)
                        {

                        }

                        rftTrans.Commit();
                    }
                }
            }
            else
            {
                throw new Exception("請開啟族群編輯器.");
            }

            return rectExtrusion;
        }
        // 草繪平面
        private static SketchPlane CreateSketchPlane(Document rftDoc, XYZ normal, XYZ origin)
        {
            Plane geometryPlane = Plane.CreateByNormalAndOrigin(normal, origin); // 草繪平面
            if (null == geometryPlane)
            {
                throw new Exception("創建幾何平面失敗.");
            }
            SketchPlane plane = SketchPlane.Create(rftDoc, geometryPlane); // 草繪平面
            if (null == plane)
            {
                throw new Exception("創建幾何平面失敗.");
            }
            return plane;
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
            }
            else if (Math.Round(sepList[0].start.X, 4, MidpointRounding.AwayFromZero) == Math.Round(sepList[1].end.X, 4, MidpointRounding.AwayFromZero) &&
                     Math.Round(sepList[0].start.Y, 4, MidpointRounding.AwayFromZero) == Math.Round(sepList[1].end.Y, 4, MidpointRounding.AwayFromZero) &&
                     Math.Round(sepList[0].start.Z, 4, MidpointRounding.AwayFromZero) == Math.Round(sepList[1].end.Z, 4, MidpointRounding.AwayFromZero))
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

                if (Math.Round(sepList[i].end.X, 4, MidpointRounding.AwayFromZero) == Math.Round(sepList[i - 1].end.X, 4, MidpointRounding.AwayFromZero) &&
                    Math.Round(sepList[i].end.Y, 4, MidpointRounding.AwayFromZero) == Math.Round(sepList[i - 1].end.Y, 4, MidpointRounding.AwayFromZero) &&
                    Math.Round(sepList[i].end.Z, 4, MidpointRounding.AwayFromZero) == Math.Round(sepList[i - 1].end.Z, 4, MidpointRounding.AwayFromZero))
                {
                    newStart = sepList[i].end;
                    newEnd = sepList[i].start;
                    sepList[i].start = newStart;
                    sepList[i].end = newEnd;
                }
            }
        }
        // 重新載入FamilySymbol
        private class JtFamilyLoadOptions : IFamilyLoadOptions
        {
            public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
            {
                overwriteParameterValues = true;
                return true;
            }
            public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
            {
                source = FamilySource.Family;
                overwriteParameterValues = true;
                return true;
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
