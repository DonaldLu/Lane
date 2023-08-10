using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Reflection;
using System.Resources;
using System.Threading;
using System.Windows.Media.Imaging;

namespace RevitFamilyInstanceLock
{
    public class AddPanel : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            ResourceManager resourceManager;
            if (Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName == "zh")
            {
                resourceManager = new ResourceManager("RevitFamilyInstanceLock.Properties.Localization_zh-Hant", Assembly.GetExecutingAssembly());
            }
            else
            {
                resourceManager = new ResourceManager("RevitFamilyInstanceLock.Properties.Localization_en-US", Assembly.GetExecutingAssembly());
            }
            application.CreateRibbonTab(resourceManager.GetString("tab.Text"));
            RibbonPanel arg_DF_0 = application.CreateRibbonPanel(resourceManager.GetString("tab.Text"), resourceManager.GetString("panel.Text"));
            string location = Assembly.GetExecutingAssembly().Location;
            PushButtonData pushButtonData = new PushButtonData("pickOneConversion", resourceManager.GetString("pickOneButton.Text"), location, "RevitFamilyInstanceLock.PickOneConversionToDirectShape");
            PushButtonData pushButtonData2 = new PushButtonData("familyConversion", resourceManager.GetString("pickByFamilyButton.Text"), location, "RevitFamilyInstanceLock.FamilyConversionToDirectShape");
            string expr_BC = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            Uri uriSource = new Uri(Path.Combine(expr_BC, "../Resources/moai.png"));
            Uri uriSource2 = new Uri(Path.Combine(expr_BC, "../Resources/moais.png"));
            PushButton expr_EB = arg_DF_0.AddItem(pushButtonData) as PushButton;
            expr_EB.LargeImage = (new BitmapImage(uriSource));
            expr_EB.ToolTip = (resourceManager.GetString("pickOneButton.ToolTip"));
            expr_EB.LongDescription = (resourceManager.GetString("pickOneButton.LongDescription"));
            PushButton expr_124 = arg_DF_0.AddItem(pushButtonData2) as PushButton;
            expr_124.LargeImage = (new BitmapImage(uriSource2));
            expr_124.ToolTip = (resourceManager.GetString("pickByFamilyButton.ToolTip"));
            expr_124.LongDescription = (resourceManager.GetString("pickByFamilyButton.LongDescription"));
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
