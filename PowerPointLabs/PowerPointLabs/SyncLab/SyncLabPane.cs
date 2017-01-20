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
        //private const string ImportLibraryFileDialogFilter =
        //    "PowerPointLabs Shapes File|*.pptlabsshapes;*.pptx";
        //private const string ImportShapesFileDialogFilter =
        //    "PowerPointLabs Shape File|*.pptlabsshape;*.pptx";
        //private const string ImportFileNameNoExtension = "import";
        //private const string ImportFileCopyName = ImportFileNameNoExtension + ".pptx";

        private readonly int _doubleClickTimeSpan = SystemInformation.DoubleClickTime;
        private int _clicks;

        private bool _firstTimeLoading = true;
        //private bool _firstClick = true;
        private bool _clickOnSelected;
        private bool _isLeftButton;

        //private bool _isPanelMouseDown;
        //private bool _isPanelDrawingFinish;
        //private Point _startPosition;
        //private Point _curPosition;

        private readonly Timer _timer;
        
        # region Properties
       
        protected override CreateParams CreateParams
        {
            get
            {
                var createParams = base.CreateParams;

                // do this optimization only for office 2010 since painting speed on 2013 is
                // really slow
                if (Globals.ThisAddIn.Application.Version == Globals.ThisAddIn.OfficeVersion2010)
                {
                    createParams.ExStyle |= (int)Native.Message.WS_EX_COMPOSITED;  // Turn on WS_EX_COMPOSITED
                }

                return createParams;
            }
        }
        # endregion

        # region Constructors
        public SyncLabPane(string syncRootFolderPath, string defaultSyncCategoryName)
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.DoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
            InitializeComponent();

            _timer = new Timer { Interval = _doubleClickTimeSpan };
            _timer.Tick += TimerTickHandler;

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
        
        public void AddStyleToList(ObjectFormat format)
        {
            syncLabListBox.AddFormat(format);
        }
        # endregion

        # region Helper Functions
        private void ClickTimerReset()
        {
            _clicks = 0;
            _clickOnSelected = false;
            // _firstClick = true;
            _isLeftButton = false;
        }

        private void TimerTickHandler(object sender, EventArgs args)
        {
            _timer.Stop();

            // if we got only 1 click in a threshold value, we take it as a single click
            if (_clicks == 1 &&
                _isLeftButton &&
                _clickOnSelected)
            {
                // _selectedThumbnail[0].StartNameEdit();
            }

            ClickTimerReset();
        }
        #endregion

        #region Functional Test APIs

        //public LabeledThumbnail GetLabeledThumbnail(string labelName)
        //{
        //    return FindLabeledThumbnail(labelName);
        //}

        //public void ImportLibrary(string pathToLibrary)
        //{
        //    ImportShapes(pathToLibrary, fromLibrary: true);
        //}

        //public void ImportShape(string pathToShape)
        //{
        //    ImportShapes(pathToShape, fromLibrary: false);
        //}

        //public Presentation GetShapeGallery()
        //{
        //    return Globals.ThisAddIn.ShapePresentation.Presentation;
        //}

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
