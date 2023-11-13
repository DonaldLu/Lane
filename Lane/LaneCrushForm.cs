using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using static Lane.LaneCrush;
using Form = System.Windows.Forms.Form;

namespace Lane
{
    public partial class LaneCrushForm : Form
    {
        UIApplication revitUIApp = null;
        UIDocument revitUIDoc = null;
        Document revitDoc = null;

        public LaneCrushForm(UIApplication uiapp, List<CrushElemInfo> crushElemInfos)
        {
            revitUIApp = uiapp;
            revitUIDoc = uiapp.ActiveUIDocument;
            revitDoc = uiapp.ActiveUIDocument.Document;

            InitializeComponent();
            CreateNodes(crushElemInfos);
            CenterToParent();
        }
        // 新增節點
        private void CreateNodes(List<CrushElemInfo> crushElemInfos)
        {
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
        // 亮顯元件
        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            // 檢查狀態變更時, 才會執行
            if (e.Action != TreeViewAction.Unknown)
            {
                if (e.Node.Checked)
                {
                    try
                    {
                        string[] nodeTest = e.Node.Text.Split('、');
                        ElementId hostId = new ElementId(Convert.ToInt32(nodeTest[1].Split('：')[1]));
                        ElementId crushElemId = new ElementId(Convert.ToInt32(nodeTest[0].Split('：')[1]));
                        if (e.Node.Level.Equals(1))
                        {
                            try
                            {
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
                            catch (Exception) { }
                        }

                    }
                    catch(Exception) { }
                }
            }
        }
        // 確定
        private void sureBtn_Click(object sender, EventArgs e)
        {
            Close();
        }
        // 取消
        private void cancelBtn_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
