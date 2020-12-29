using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office;
using Microsoft.Office.Interop;
using System.Diagnostics;
namespace TBFSReport
{
    public partial class Form1 : Form
    {
        public DataTable dt = new DataTable();

        public Form1()
        {
            InitializeComponent();
            dt.Columns.Add(new DataColumn("Files", typeof(string)));
            dt.Columns.Add(new DataColumn("Size", typeof(string)));
            dt.Columns.Add(new DataColumn("Path", typeof(string)));
            dt.Columns.Add(new DataColumn("Open File", typeof(string)));
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = true;
            // Show the FolderBrowserDialog.  
            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox1.Text = folderDlg.SelectedPath;
                Environment.SpecialFolder root = folderDlg.RootFolder;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
          //  dt.Columns.Add(new DataColumn("Files"),typeof(string));
           

            List<string> files = Directory.GetFiles(textBox1.Text,"*",SearchOption.AllDirectories).ToList();

            // Display all the files.
            foreach (string file in files)
            {
                DataRow dr = dt.NewRow();
                FileInfo fi = new FileInfo(file);
                dr[0] = fi.Name;
                dr[1] = fi.Length.ToString();
                dr[2] = fi.FullName.ToString();
                Label lblopenfile = new Label();
                dr[3] = lblopenfile.Text="Click";
                //dr[3] = new Label().Text="Click";
                dt.Rows.Add(dr);
                //   Console.WriteLine(file);

            }
            dataGridView1.DataSource = dt;
          
        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            //(dataGridView1.DataSource as DataTable).DefaultView.RowFilter = string.Format("Files LIKE '{0}%'", textBox2.Text);
            (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = $"[Files] LIKE '%{textBox2.Text}%'";
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            // Creating a Excel object.
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(Type.Missing);
                Microsoft.Office.Interop.Excel.Worksheet worksheet = null;

                try
                {

                    worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;

                    worksheet.Name = "ExportedFromDatGrid";

                    int cellRowIndex = 1;
                    int cellColumnIndex = 1;

                    //Loop through each row and read value from each column.
                    for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < 3; j++)
                        {
                            // Excel index starts from 1,1. As first Row would have the Column headers, adding a condition check.
                            if (cellRowIndex == 1)
                            {
                                worksheet.Cells[cellRowIndex, cellColumnIndex] = dataGridView1.Columns[j].HeaderText;
                            }
                            else
                            {
                                worksheet.Cells[cellRowIndex, cellColumnIndex] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                            }
                            cellColumnIndex++;
                        }
                        cellColumnIndex = 1;
                        cellRowIndex++;
                    }

                    //Getting the location and file name of the excel to save from user.
                    SaveFileDialog saveDialog = new SaveFileDialog();
                    saveDialog.Filter = "Excel files (.xlsx)|*.xlsx|All files (.*)|*.*";
                    saveDialog.FilterIndex = 2;

                    if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        workbook.SaveAs(saveDialog.FileName);
                        MessageBox.Show("Export Successful");
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    excel.Quit();
                    workbook = null;
                    excel = null;
                }

            }

       
    }
    public class filerecords
    {
        public string filename { get; set; }

    }

}
