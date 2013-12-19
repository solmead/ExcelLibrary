using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Dynamic;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelLibrary;

namespace TestApplication
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var list = new List<AClass>();
            for (int a = 0; a < 100; a++)
            {
                list.Add(new AClass() { ID = a.ToString(), Name = "Name" + a });
            }
            ExportListToCSV(list);
            //var dt = list.ToDataTable(true);
            //Debug.WriteLine(dt.Rows.Count);
        }
        private void ExportListToCSV(IEnumerable list, string fileName = "export.csv")
        {
            var csvFile = CSVFile.LoadFromIEnumerable(list, useDisplayName: true);
            var file = new FileInfo("c:\\development\\test.csv");
            if (file.Exists)
            {
                file.Delete();
            }
            if (!file.Directory.Exists)
            {
                file.Directory.Create();
            }
            var output = file.OpenWrite(); // new MemoryStream();
            var csvWriter = new StreamWriter(output, Encoding.UTF8);
            csvWriter.Write(csvFile.GetAsCSV());
            csvWriter.Flush();
            csvWriter.BaseStream.Position = 0;
            output.Close();

            var lst = CSVImporter.ImportFromFile(file, new Dictionary<string, string>(), () => new AClass());
            Debug.WriteLine(lst.Count());
        }
    }
    
    public class BClass
    {
        public string ID { get; set; }
    }

    public class AClass : BClass
    {
        [Display(Name = "MainName")]
        [Column("MainName")]
        public string Name { get; set; }
    }
}
