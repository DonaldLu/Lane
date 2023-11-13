using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Windows.Media.Imaging;

namespace Lane
{
    public class LaneCrush_Button : IExternalApplication
    {
        // 封包版路徑位址
        public class PacketPathName
        {
            //// 2020
            //public string assembly = @"C:\ProgramData\Autodesk\Revit\Addins\2020\Lane\"; // .dll檔案存取路徑
            //public string picPath = @"C:\ProgramData\Autodesk\Revit\Addins\2020\Lane\Pic\"; // Button圖片存取路徑
            // 2022
            public string assembly = @"C:\ProgramData\Autodesk\Revit\Addins\2022\Lane\"; // .dll檔案存取路徑
            public string picPath = @"C:\ProgramData\Autodesk\Revit\Addins\2022\Lane\Pic\"; // Button圖片存取路徑
        }
        public Result OnStartup(UIControlledApplication application)
        {
            // 封包檔案
            PacketPathName packetPathName = new PacketPathName();
            string sinoPathAsb = packetPathName.assembly + "Lane.dll";
            string picPath = packetPathName.picPath;

            // 創建一個新的選單
            string tabName = "干涉報告";
            string panelName = "干涉報告";
            RibbonPanel sinoPathRP = null;
            try
            { application.CreateRibbonTab(tabName); }
            catch(Exception) { }
            try
            { sinoPathRP = application.CreateRibbonPanel(tabName, panelName); }
            catch
            {
                List<RibbonPanel> panelList = application.GetRibbonPanels(tabName);
                foreach (RibbonPanel rp in panelList)
                {
                    if (rp.Name == panelName)
                    {
                        sinoPathRP = rp;
                        break;
                    }
                }
            }
            // 在面板上添加一個按鈕, 點擊此按鈕觸動Lane.LaneCrush
            PushButton laneCrushBtn = sinoPathRP.AddItem(new PushButtonData("LaneCrush", "干涉檢查", sinoPathAsb, "Lane.LaneCrush")) as PushButton;
            laneCrushBtn.LargeImage = new BitmapImage(new Uri(picPath + "干涉檢查.png")); // 給按鈕添加圖片
            // 在面板上添加一個按鈕, 點擊此按鈕觸動Lane.RemoveElem
            PushButton removeElemBtn = sinoPathRP.AddItem(new PushButtonData("RemoveElem", "移除干涉元件", sinoPathAsb, "Lane.RemoveElem")) as PushButton;            
            removeElemBtn.LargeImage = new BitmapImage(new Uri(picPath + "移除干涉元件.png")); // 給按鈕添加圖片

            return Result.Succeeded;
        }
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}