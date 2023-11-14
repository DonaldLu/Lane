using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows;
using System.Windows.Media.Imaging;

namespace Lane
{
    // 封包版路徑位址
    public class PacketPathName
    {
        // 2022
        public string assembly = @"C:\ProgramData\Autodesk\Revit\Addins\2022\"; // .dll檔案存取路徑
    }
    public class LaneCrush_Button : IExternalApplication
    {
        static string addinAssmeblyPath = Assembly.GetExecutingAssembly().Location; // 封包版路徑位址
        public Result OnStartup(UIControlledApplication a)
        {            
            // 封包檔案
            PacketPathName packetPathName = new PacketPathName();
            addinAssmeblyPath = packetPathName.assembly + "Lane.dll";

            RibbonPanel ribbonPanel = null;
            try { a.CreateRibbonTab("干涉報告"); } catch { }
            try { ribbonPanel = a.CreateRibbonPanel("干涉報告", "干涉報告"); }
            catch
            {
                List<RibbonPanel> panel_list = new List<RibbonPanel>();
                panel_list = a.GetRibbonPanels("干涉報告");
                foreach (RibbonPanel rp in panel_list)
                {
                    if (rp.Name == "干涉報告")
                    {
                        ribbonPanel = rp;
                    }
                }
            }
            // 在面板上添加一個按鈕, 點擊此按鈕觸動Lane.LaneCrush
            PushButton laneCrushBtn = ribbonPanel.AddItem(new PushButtonData("LaneCrush", "干涉檢查", addinAssmeblyPath, "Lane.LaneCrush")) as PushButton;
            laneCrushBtn.LargeImage = convertFromBitmap(Properties.Resources.干涉檢查);
            // 在面板上添加一個按鈕, 點擊此按鈕觸動Lane.RemoveElem
            PushButton removeElemBtn = ribbonPanel.AddItem(new PushButtonData("RemoveElem", "移除干涉元件", addinAssmeblyPath, "Lane.RemoveCrushElems")) as PushButton;
            removeElemBtn.LargeImage = convertFromBitmap(Properties.Resources.移除干涉元件);

            return Result.Succeeded;
        }

        BitmapSource convertFromBitmap(System.Drawing.Bitmap bitmap)
        {
            return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                bitmap.GetHbitmap(),
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}