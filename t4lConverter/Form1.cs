using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.IO;

namespace t4lConverter
{
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            InitializeComponent();
        }

        private void OnBtnOpenFileClicked(object sender, EventArgs e)
        {
            //browse for excel file
            
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Excel Files (*.xls)|*.xls|(*.xlsx)|*.xlsx";
            
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                var fileName = new FileInfo(dlg.FileName);            
                if (fileName.Extension.ToLower().Equals(".xls") ||
                    fileName.Extension.ToLower().Equals(".xlsx"))
                {
                    var wrapper = new ExcelWrapper();
                    wrapper.ThreadDone += HandleThreadDone;

                    var convertThread = new Thread(() => wrapper.Convert(fileName));
                    convertThread.Start();
                }
                else
                {
                    MessageBox.Show("File is not xls or xlsx format");
                }
            }
        }

        public void HandleThreadDone(object sender, EventArgs e)
        {
            MessageBox.Show("Done Converting!");
        }       
    }
}
