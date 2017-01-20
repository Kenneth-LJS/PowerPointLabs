using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PowerPointLabs.SyncLab
{
    public partial class SyncLabListBox : ListView
    {

        public static readonly int MAX_LIST_SIZE = 50;
        LinkedList<ListViewItem> formatList = new LinkedList<ListViewItem>();

        public SyncLabListBox()
        {
            InitializeComponent();
            this.View = View.LargeIcon;
            this.ArrangeIcons(ListViewAlignment.Left);
            this.LargeImageList = new ImageList();
        }

        protected override void OnPaint(PaintEventArgs pe)
        {
            base.OnPaint(pe);
        }

        public void AddFormat(ObjectFormat format)
        {
            AddFormatNoUpdate(format);
            UpdateList();
        }

        public void AddFormat(ICollection<ObjectFormat> formats)
        {
            foreach (ObjectFormat format in formats)
            {
                AddFormatNoUpdate(format);
            }
            UpdateList();
        }

        private void AddFormatNoUpdate(ObjectFormat format)
        {
            string imageKey = GetNextImageKey();
            ListViewItem newItem = new ListViewItem(format.DisplayText, imageKey);
            newItem.Tag = format;
            this.LargeImageList.Images.Add(imageKey, format.DisplayImage);
            formatList.AddFirst(newItem);
            while (formatList.Count > MAX_LIST_SIZE)
            {
                formatList.RemoveLast();
            }
        }

        private void UpdateList()
        {
            this.BeginUpdate();
            this.Items.Clear();
            this.Items.AddRange(formatList.ToArray());
            this.EndUpdate();
        }

        public ObjectFormat GetFormat(int index)
        {
            return (ObjectFormat)this.Items[index].Tag;
        }

        public void RemoveFormat(int index)
        {
            string imageKey = this.Items[index].ImageKey;
            this.Items.RemoveAt(index);
            this.LargeImageList.Images.RemoveByKey(imageKey);
        }

        private int curImageIndex = 0;
        private string GetNextImageKey()
        {
            string key;
            do
            { // Find new index for the image
                key = (curImageIndex++).ToString();
            }
            while (this.LargeImageList.Images.ContainsKey(key));
            return key;
        }
    }
}
