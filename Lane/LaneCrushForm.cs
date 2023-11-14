using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Media;
using static Lane.CreateCrushElems;
using Form = System.Windows.Forms.Form;

namespace Lane
{
    public partial class LaneCrushForm : Form
    {
        UIApplication revitUIApp = null;
        UIDocument revitUIDoc = null;
        Document revitDoc = null;
        public static double height {  get; set; }
        private List<string> errorMessages = new List<string>();

        public LaneCrushForm(UIApplication uiapp, double inputHeight)
        {
            revitUIApp = uiapp;
            revitUIDoc = uiapp.ActiveUIDocument;
            revitDoc = uiapp.ActiveUIDocument.Document;
            height = inputHeight;
            List<CrushElemInfo> crushElemInfos = CreateCrushElems(revitUIDoc, revitDoc);
            errorMessages = errorMessages.Distinct().ToList();
            string error = "";
            for (int i = 0; i < errorMessages.Count; i++)
            {
                error += i + ". " + errorMessages[i] + "\n";
            }
            if(error != "")
            {
                TaskDialog.Show("Error", error);
            }

            InitializeComponent();
            CreateNodes(crushElemInfos); // 新增節點
            CenterToParent();
        }
        // 建立模型干涉檢查
        private List<CrushElemInfo> CreateCrushElems(UIDocument uidoc, Document doc)
        {
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
            catch(Exception ex)
            {
                errorMessages.Add("樓板參數錯誤！\n" + ex.Message + "\n" + ex.ToString());
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
                    string number = "";
                    //try { number = floor.LookupParameter("編號").AsString(); }
                    //catch (Exception ex) { errorMessages.Add(floor.Id + " 無「編號」參數\n" + ex.Message + "\n" + ex.ToString()); }
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

            return crushElemInfos;
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
                errorMessages.Add("樓板參數錯誤！\n" + ex.Message + "\n" + ex.ToString());
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
                //catch(Exception ex) { errorMessages.Add(floorId + " 無「編號」參數\n" + ex.Message + "\n" + ex.ToString()); }
                ds.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(floorId + "：車道板干涉元件");
                ds.SetShape(new List<GeometryObject>() { solid });
            }
            catch (Exception ex)
            {
                errorMessages.Add(floorId + "建立干涉元件錯誤\n" + ex.Message + "\n" + ex.ToString());
            }

            return solid;
        }
        // 出衝突報告
        private List<CrushElemInfo> CrushReport(Document doc, LogicalOrFilter logicalOrFilter, List<Floor> floors)
        {
            List<CrushElemInfo> crushElemInfos = new List<CrushElemInfo>();
            List<DirectShape> crushFloors = new List<DirectShape>();
            try
            {
                crushFloors = new FilteredElementCollector(doc).WherePasses(logicalOrFilter).WhereElementIsNotElementType().Where(x => x is DirectShape).Cast<DirectShape>()
                              .Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString() != null)
                              .Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Contains("車道板干涉元件")).ToList();
            }
            catch (Exception ex)
            {
                errorMessages.Add("DirectShape參數錯誤\n" + ex.Message + "\n" + ex.ToString());
            }
            List<ElementId> crushFloorIds = crushFloors.Select(x => x.Id).ToList();
            foreach (Floor floor in floors) { crushFloorIds.Add(floor.Id); }
            foreach(Floor floor in floors)
            {
                List<DirectShape> directShapes = new List<DirectShape>();
                try
                {
                    directShapes = crushFloors.Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split('：')[0].Equals(floor.Id.ToString())).ToList();
                }
                catch (Exception ex)
                {
                    errorMessages.Add("DirectShape參數解析錯誤\n" + ex.Message + "\n" + ex.ToString());
                }
                CrushElemInfo crushElemInfo = new CrushElemInfo();
                string hostName = doc.GetElement(floor.Id).Name + "：" + floor.Id.ToString();
                crushElemInfo.hostName = hostName;
                foreach (DirectShape directShape in directShapes)
                {
                    try
                    {
                        List<Solid> solidList = GetSolids(doc, directShape); // 儲存所有車道板的Solid
                        foreach (Solid solid in solidList)
                        {
                            IList<Element> elems = new FilteredElementCollector(doc).WherePasses(new ElementIntersectsSolidFilter(solid)).Excluding(crushFloorIds).ToList();
                            if (elems.Count > 0)
                            {
                                foreach (Element elem in elems)
                                {
                                    crushElemInfo.crushElemName.Add(elem.Name + "：" + elem.Id + "、干涉車道板：" + directShape.Id);
                                }
                                crushElemInfo.crushElemName = crushElemInfo.crushElemName.OrderBy(x => x).ToList();
                            }
                        }
                    }
                    catch (Exception) { }
                }
                if(crushElemInfo.crushElemName.Count > 0)
                {
                    crushElemInfos.Add(crushElemInfo);
                }
            }

            return crushElemInfos;
        }
        // 新增節點
        private void CreateNodes(List<CrushElemInfo> crushElemInfos)
        {
            treeView1.Nodes.Clear(); // 清空節點
            List<string> hostNames = crushElemInfos.Select(x => x.hostName).Distinct().ToList();
            int hostNameCount = 0;
            foreach(string hostName in hostNames)
            {
                treeView1.Nodes.Add(hostName);
                treeView1.Nodes[0].Checked = true;
                List<List<string>> crushElemNameList = (from x in crushElemInfos
                                                        where x.hostName.Equals(hostName)
                                                        select x).Select(x => x.crushElemName).ToList();
                int nodeCount = 0;
                foreach (List<string> crushElemNames in crushElemNameList)
                {
                    foreach (string crushElemName in crushElemNames)
                    {
                        treeView1.Nodes[hostNameCount].Nodes.Add(crushElemName);
                        // 預設勾選
                        treeView1.Nodes[hostNameCount].Nodes[nodeCount].Checked = true;
                        nodeCount++;
                    }
                }
                hostNameCount++;
            }
        }
        // 重新檢核
        private List<CrushElemInfo> ReviewReport(Document doc)
        {
            IList<ElementFilter> elementFilters = new List<ElementFilter>(); // 清空過濾器
            ElementCategoryFilter roomFilter = new ElementCategoryFilter(BuiltInCategory.OST_Floors); // 樓板
            elementFilters.Add(roomFilter);
            LogicalOrFilter logicalOrFilter = new LogicalOrFilter(elementFilters);
            List<Floor> floors = new List<Floor>();
            try
            {
                floors = new FilteredElementCollector(doc).WherePasses(logicalOrFilter).WhereElementIsNotElementType().Where(x => x is Floor)
                         .Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString() != null)
                         .Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Contains("車道板")).Cast<Floor>().ToList();
            }
            catch (Exception ex)
            {
                errorMessages.Add("樓板參數錯誤\n" + ex.Message + "\n" + ex.ToString());
            }
            List<CrushElemInfo> crushElemInfo = CrushReport(doc, logicalOrFilter, floors);

            return crushElemInfo;
        }
        // 亮顯元件
        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            // 檢查狀態變更時, 才會執行
            if (e.Action != TreeViewAction.Unknown)
            {
                if (e.Node.IsSelected)
                {
                    try
                    {
                        if (e.Node.Level.Equals(0))
                        {
                            string[] nodeTest = e.Node.Text.Split('：');
                            ElementId hostId = new ElementId(Convert.ToInt32(nodeTest[1]));
                            // 亮顯id元件
                            try
                            {
                                IList<ElementId> highlightElems = new List<ElementId>();
                                highlightElems.Add(hostId);
                                revitUIApp.ActiveUIDocument.ShowElements(highlightElems);
                                revitUIApp.ActiveUIDocument.Selection.SetElementIds(highlightElems);
                            }
                            catch (Exception)
                            {

                            }
                        }
                        else if (e.Node.Level.Equals(1))
                        {
                            string[] nodeTest = e.Node.Text.Split('、');
                            ElementId hostId = new ElementId(Convert.ToInt32(nodeTest[1].Split('：')[1]));
                            ElementId crushElemId = new ElementId(Convert.ToInt32(nodeTest[0].Split('：')[1]));
                            // 亮顯id元件
                            try
                            {
                                IList<ElementId> highlightElems = new List<ElementId>();
                                highlightElems.Add(hostId);
                                highlightElems.Add(crushElemId);
                                revitUIApp.ActiveUIDocument.ShowElements(highlightElems);
                                revitUIApp.ActiveUIDocument.Selection.SetElementIds(highlightElems);
                            }
                            catch (Exception)
                            {

                            }
                        }
                    }
                    catch(Exception) { }
                }
            }
        }
        // 匯出報告
        private void exportBtn_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog path = new FolderBrowserDialog();
            path.ShowDialog();
            string filePath = Path.Combine(path.SelectedPath, "匯出報告.txt");
            string content = string.Empty;
            if (treeView1.Nodes.Count.Equals(0))
            {
                content = "車道板無干涉到的元件";
            }
            else
            {
                int i = 1;
                foreach (TreeNode treeNode in treeView1.Nodes)
                {
                    content += i + "、" + treeNode.Text + "\n";
                    List<string> subNodes = new List<string>();
                    foreach (TreeNode subNode in treeNode.Nodes)
                    {
                        subNodes.Add(subNode.Text.Split('、')[0]);
                    }
                    subNodes = subNodes.Distinct().OrderBy(x => x.Split('：')[0]).ToList();
                    foreach (string subNode in subNodes)
                    {
                        content += subNode + "\n";
                    }
                    i++;
                    content += "\n";
                }
            }
            // 先檢查是否有此檔案, 沒有的話則新增
            if (!File.Exists(filePath))
            {
                using (FileStream fs = File.Create(filePath))
                {

                }
            }
            using (StreamWriter sw = new StreamWriter(filePath))
            {
                sw.WriteLine(content);
                sw.Close();
            }
            TaskDialog.Show("Revit", "匯出完成。");
        }
        // 重新檢核
        private void reviewBtn_Click(object sender, EventArgs e)
        {
            CreateNodes(ReviewReport(revitDoc)); // 新增節點
        }
        // 關閉
        private void colseBtn_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}