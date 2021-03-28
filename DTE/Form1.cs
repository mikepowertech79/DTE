using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
namespace DTE
{
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = ExportToExcel();
        }

     public System.Data.DataTable ExportToExcel()
        {
            System.Data.DataTable table = new System.Data.DataTable();
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Sex", typeof(string));
            table.Columns.Add("Subject1", typeof(int));
            table.Columns.Add("Subject2", typeof(int));
            table.Columns.Add("Subject3", typeof(int));
            table.Columns.Add("Subject4", typeof(int));
            table.Columns.Add("Subject5", typeof(int));
            table.Columns.Add("Subject6", typeof(int));
            table.Rows.Add(1, "Amar", "M",78,59,72,95,83,77);
            table.Rows.Add(2, "Mohit", "M", 76,65,85,87,72,90);
            table.Rows.Add(3, "Garima", "F", 77,73,83,64,86,63);
            table.Rows.Add(4, "jyoti", "F", 55,77,85,69,70,86);
            table.Rows.Add(5, "Avinash", "M", 87,73,69,75,67,81);
            table.Rows.Add(6, "Devesh", "M", 92,87,78,73,75,72);
            return table;
        }

       private void button2_Click(object sender, EventArgs e)
       {

           SaveFileDialog saveFileDialog1 = new SaveFileDialog();
           saveFileDialog1.InitialDirectory = @"C:\";
           saveFileDialog1.Title = "Save Excel Files";

           saveFileDialog1.DefaultExt = "xlsx";
           saveFileDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
           saveFileDialog1.FilterIndex = 1;
           saveFileDialog1.RestoreDirectory = true;

           if (saveFileDialog1.ShowDialog() == DialogResult.OK)
           {
               textBox1.Text = saveFileDialog1.FileName;

           }
           else
           {
               return;
           }
       }

       private void button1_Click(object sender, EventArgs e)
       {
           Microsoft.Office.Interop.Excel.Application excel;
           Microsoft.Office.Interop.Excel.Workbook worKbooK;
           Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
           Microsoft.Office.Interop.Excel.Range celLrangE;

           try
           {
               excel = new Microsoft.Office.Interop.Excel.Application();
               excel.Visible = false;
               excel.DisplayAlerts = false;
              worKbooK = excel.Workbooks.Add(Type.Missing);

               
               worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
               worKsheeT.Name = "StudentRepoertCard";

               worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[1, 8]].Merge();
               worKsheeT.Cells[1, 1] = "Student Report Card";
               worKsheeT.Cells.Font.Size = 15;

               
               int rowcount = 2;

               foreach (DataRow datarow in ExportToExcel().Rows)
               {
                   rowcount += 1;
                   for (int i = 1; i <= ExportToExcel().Columns.Count; i++)
                   {
                      
                       if (rowcount == 3)
                       {
                           worKsheeT.Cells[2, i] = ExportToExcel().Columns[i - 1].ColumnName;
                           worKsheeT.Cells.Font.Color = System.Drawing.Color.Black;

                       }

                       worKsheeT.Cells[rowcount, i] = datarow[i - 1].ToString();

                      
                       if (rowcount > 3)
                       {
                           if (i == ExportToExcel().Columns.Count)
                           {
                               if (rowcount % 2 == 0)
                               {
                                   celLrangE = worKsheeT.Range[worKsheeT.Cells[rowcount, 1], worKsheeT.Cells[rowcount, ExportToExcel().Columns.Count]];
                               }

                           }
                       }

                   }

               }

               
               celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[rowcount, ExportToExcel().Columns.Count]];
               celLrangE.EntireColumn.AutoFit();
               Microsoft.Office.Interop.Excel.Borders border = celLrangE.Borders;
               border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
               border.Weight = 2d;


               celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[2, ExportToExcel().Columns.Count]];


               


               worKbooK.SaveAs(textBox1.Text); ;
               worKbooK.Close();
               excel.Quit();
               MessageBox.Show("Successfully Create Excel File");
               
           }
           catch (Exception ex)
           {
               MessageBox.Show(ex.Message);
               
           }
           finally
           {
               worKsheeT = null;
               celLrangE = null;
               worKbooK = null;
           }

       }
    }
}
