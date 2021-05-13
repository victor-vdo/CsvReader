using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CsvMake
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd1 = new OpenFileDialog();
            ofd1.Multiselect = true;
            ofd1.Title = "Select a CSV FIle";
            ofd1.InitialDirectory = @"C:\Users\victor\Desktop\";
            ofd1.Filter = "Files (*.CSV;)|*.CSV;|" + "All files (*.*)|*.*";
            ofd1.CheckFileExists = true;
            ofd1.CheckPathExists = true;
            ofd1.FilterIndex = 2;
            ofd1.RestoreDirectory = true;
            ofd1.ReadOnlyChecked = true;
            ofd1.ShowReadOnly = true;

            dgvCsv.Columns.Clear();

            DialogResult dr = ofd1.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                // Le os arquivos selecionados 
                foreach (String fileName in ofd1.FileNames)
                {
                    try
                    {
                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                        using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
                        {
                            using (var reader = ExcelReaderFactory.CreateCsvReader(stream))
                            {
                                var result = reader.AsDataSet();
                                var tables = result.Tables.Cast<DataTable>();

                                foreach (var table in tables)
                                { 
                                    var rows = table.Rows.Cast<DataRow>();
                                    var firstRow = rows.ToList().FirstOrDefault().ItemArray;
                                    List<string> listColumns = new List<string>();
                                   
                                    foreach (var item in firstRow) listColumns.Add(item.ToString());
                                    var disColumns = listColumns.Distinct().ToList();

                                    foreach (var stg in disColumns.Select((x, i) => new { Index = i, Value = x }))
                                        dgvCsv.Columns.Add(string.Format("col{0}", stg.Index), stg.Value);   
                                    
                                    foreach (var item in disColumns.Select((x, i) => new { Index = i, Value = x }))
                                    {
                                        if (item.ToString().StartsWith("Column"))
                                            dgvCsv.Columns.RemoveAt(item.Index);
                                    }
                                    dgvCsv.DataSource = table;
                                }
                            }
                        }
                    }
                    catch (SecurityException ex)
                    {
                        MessageBox.Show("Security Error! Contact your network security administrator.\n\n" +
                                                    "Message : " + ex.Message + "\n\n" +
                                                    "Details (send to support):\n\n" + ex.StackTrace);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Cannot display the file: " + fileName.Substring(fileName.LastIndexOf('\\'))
                                                   + ". You may not be allowed to read the file, or " +
                                                   " or it may be corrupted.\n\nError Message : " + ex.Message);
                    }
                }
            }
        }

    }
}
