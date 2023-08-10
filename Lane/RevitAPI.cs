using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows;
using System.Windows.Media.Imaging;

namespace Lane
{
    public class RevitAPI : IExternalApplication
    {
        static string addinAssmeblyPath = Assembly.GetExecutingAssembly().Location;
        //static string checkPlatformPath = @"C:\ProgramData\Autodesk\Revit\Addins\2018\AutoCreateModel\AutoCreateModel.dll";
        public Result OnStartup(UIControlledApplication a)
        {
            RibbonPanel ribbonPanel = null;
            try { a.CreateRibbonTab("檢核工具"); } catch { }
            try { ribbonPanel = a.CreateRibbonPanel("檢核工具", "車道檢核"); }
            catch
            {
                List<RibbonPanel> panel_list = new List<RibbonPanel>();
                panel_list = a.GetRibbonPanels("檢核工具");
                foreach (RibbonPanel rp in panel_list)
                {
                    if (rp.Name == "車道檢核")
                    {
                        ribbonPanel = rp;
                    }
                }
            }

            PushButton pushbutton1 = ribbonPanel.AddItem(new PushButtonData("Lane", "車道檢核", addinAssmeblyPath, "Lane.LaneCrush")) as PushButton;
            pushbutton1.LargeImage = convertFromBitmap(Properties.Resources.Lane);

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
