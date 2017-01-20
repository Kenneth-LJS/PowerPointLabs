using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using PPExtraEventHelper;
using PowerPointLabs.SyncLab;
using PowerPointLabs.SyncLab.ObjectFormats;
using System.Drawing;

namespace PowerPointLabs
{
    public partial class SyncLabPane : UserControl
    {
#pragma warning disable 0618

        private bool _firstTimeLoading = true;

        # region Constructors
        public SyncLabPane(string syncRootFolderPath, string defaultSyncCategoryName)
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.DoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
            InitializeComponent();
        }
        # endregion

        # region API
        public void PaneReload(bool forceReload = false)
        {
            if (!_firstTimeLoading && !forceReload)
            {
                return;
            }

            _firstTimeLoading = false;
        }

        public void CopyFormat()
        {
            ShapeRange selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
            CopyFormat(selectedShapes);
        }

        public void CopyFormat(ShapeRange shapes)
        {
            List<ObjectFormat> newFormats = new List<ObjectFormat>();
            foreach (Shape shape in shapes)
            {
                newFormats.Add(new SyncLab.ObjectFormats.FillFormat(shape));
            }
            syncLabListBox.AddFormat(newFormats);
        }

        public void PasteFormat()
        {
            ShapeRange selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
            PasteFormat(selectedShapes);
        }

        public void PasteFormat(ShapeRange shapes)
        {
            foreach (Shape shape in shapes)
            {
                PasteFormat(shape);
            }
        }

        public void PasteFormat(Shape shape)
        {
            List<int> checkedIndices = syncLabListBox.CheckedIndices.Cast<int>().ToList<int>();
            checkedIndices.Sort();
            checkedIndices.Reverse();
            for (int i = 0; i < checkedIndices.Count; i++)
            {
                ObjectFormat format = syncLabListBox.GetFormat(checkedIndices[i]);
                format.ApplyTo(shape);
            }
        }

        #endregion

        #region Functional Test APIs

        public void AddStyleToList(ObjectFormat format)
        {
            syncLabListBox.AddFormat(format);
        }

        #endregion

        #region GUI Handlers
        private void CopyButton_Click(object sender, EventArgs e)
        {
            CopyFormat();
        }

        private void PasteButton_Click(object sender, EventArgs e)
        {
            PasteFormat();
        }
        #endregion
    }
}
