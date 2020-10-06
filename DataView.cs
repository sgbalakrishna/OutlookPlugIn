using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace mailBoxWizard
{
    public partial class DataView : Form
    {
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.Run(new Form1());
        }
        public List<string> tagslist { get; private set; }

        public DataView()
        {
            Application.EnableVisualStyles();
            InitializeComponent();
        }

        
        private void DrawDynamicCheckBoxes()
        {
            var result = Program.PrepareData();
            var res = result.OrderByDescending(o => o.Counts);
            int i = 0;
            foreach (var item in res)
            {
                CheckBox chk = new CheckBox();
                string textString = item.Counts + " Emails from " + item.EmailSender + " With an average frequency of " + item.freq + " days.";
                chk.Text = textString.ToString();
                chk.Name = item.EmailSender.ToString();
                chk.Tag = item.EmailSender.ToString();
                chk.AutoSize = true;
                chk.CheckedChanged += new EventHandler(updateCheckedList);
                chk.Location = new Point(10, i * 10);
                i += 1;
                flowLayoutPanel1.Controls.Add(chk);
                flowLayoutPanel1.WrapContents = true;
                flowLayoutPanel1.AutoScroll = false;
                flowLayoutPanel1.HorizontalScroll.Enabled = false;
                flowLayoutPanel1.HorizontalScroll.Visible = false;
                flowLayoutPanel1.HorizontalScroll.Maximum = 0;
                flowLayoutPanel1.AutoScroll = true;
                flowLayoutPanel1.FlowDirection = FlowDirection.TopDown;
                flowLayoutPanel1.WrapContents = false;
                flowLayoutPanel1.Dock = DockStyle.Fill;
            }
        }

        private void updateCheckedList(object sender, EventArgs e)
        {
            CheckBox cb = sender as CheckBox;
            Button btn = new Button();
            btn.Text = "Delete selected domains";
            btn.Width = 175;
            flowLayoutPanel2.Controls.Clear();
            var checkedItems = showClickedBoxes();
            TextBox tb = new TextBox();
            string joined = string.Join(",\n", checkedItems);
            tb.Text = joined;
            tb.Font = new Font(tb.Font.FontFamily, 12);
            tb.AutoSize = true;
            tb.Multiline = true;
            tb.WordWrap = true;
            tb.Width = flowLayoutPanel2.Width-5; ;
            tb.Height = flowLayoutPanel2.Height-50;
            tb.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
            flowLayoutPanel2.Controls.Add(tb);
            if(tb.Text.Length > 0)
            {
                flowLayoutPanel2.Controls.Add(btn);
                btn.Click += new EventHandler(MyButton_Click);
                //DrawDynamicCheckBoxes();
            }
            
        }

        private void MyButton_Click(object sender, EventArgs e)
        {

            var checkedItems = showClickedBoxes();
            
            DialogResult dialogResult = MessageBox.Show("Are you sure?", "Confirmation", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                RemoveMailItems.deleteMailItems(checkedItems);
                MessageBox.Show("Succesfully deleted!");

            }
            else if (dialogResult == DialogResult.No)
            {
                GC.Collect();
            }
            
            GC.Collect();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DrawDynamicCheckBoxes();
            this.StartPosition = FormStartPosition.CenterScreen;
        }


        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {
            this.Anchor = AnchorStyles.Top;

        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private List<string> showClickedBoxes()
        {
            var tagsList = new List<string>();
            foreach (Control c in flowLayoutPanel1.Controls)
            {
                if (c is CheckBox && ((CheckBox)c).Checked)
                {
                    tagsList.Add(c.Tag.ToString());
                }
            }
            return tagsList;

        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var lines = BuildLines();
            var path = getFilePathToSave();
            try
            {
                if (lines != null)
                {
                    System.IO.File.WriteAllLines(path.ToString(), (string[])lines);
                }
                
            }
            catch
            {
            }

        }

        private object BuildLines()
        {
            var res = Program.PrepareData();
            string[] lines = new string[res.Count()];
            int i = 0;
            foreach (var item in res)
            {
                string buildLine = item.Counts + " Emails from " + item.EmailSender + " With frequency of average " + item.freq + " days.";

                lines[i++] = buildLine.ToString();
            }
            return lines;
        }

        private object getFilePathToSave()
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Text File|*.txt";
            saveFileDialog1.Title = "Save an Image File";
            saveFileDialog1.ShowDialog();
            return saveFileDialog1.FileName;
        }

        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void fileToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void flowLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox1 aboutbox = new AboutBox1();
            aboutbox.Developer = "Balakrishna";
             aboutbox.Show();
        }
    }
}
