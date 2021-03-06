﻿using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId(
        "fitToHeightShape", 
        "fitToHeightFreeform", 
        "fitToHeightPicture",
        "fitToHeightChart", 
        "fitToHeightTable")]
    class FitToHeightActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();
            var selectedShape = this.GetCurrentSelection().ShapeRange[1];
            var pres = this.GetCurrentPresentation();
            FitToSlide.FitToHeight(selectedShape, pres.SlideWidth, pres.SlideHeight);
        }
    }
}
