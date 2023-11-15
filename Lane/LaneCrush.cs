using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;

namespace Lane
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    [Journaling(JournalingMode.NoCommandData)]
    public class LaneCrush : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //IExternalEventHandler handler_CreateCrushElems = new CreateCrushElems();
            //ExternalEvent externalEvent_CreateCrushElems = ExternalEvent.Create(handler_CreateCrushElems);
            //commandData.Application.Idling += Application_Idling;
            RevitDocument m_connect = new RevitDocument(commandData.Application);

            InputHeightForm inputHeightForm = new InputHeightForm();
            inputHeightForm.ShowDialog();
            if (inputHeightForm.trueOrFalse == true)
            {
                LaneCrushForm laneCrushForm = new LaneCrushForm(commandData.Application, inputHeightForm.height);
                laneCrushForm.Show();
            }

            return Result.Succeeded;
        }
        private void Application_Idling(object sender, IdlingEventArgs e)
        {
            throw new NotImplementedException();
        }
    }
}