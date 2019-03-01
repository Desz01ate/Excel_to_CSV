using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excsv
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.Text = "Convert Excel To CSV";
            this.MaximizeBox = false;
            this.AllowDrop = true;
            this.DragDrop += (s, e) =>
            {
            };
            this.DragEnter += (s, e) =>
            {
                var effects = DragDropEffects.None;
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    var path = ((string[])e.Data.GetData(DataFormats.FileDrop)).First();
                    if (Directory.Exists(path))
                    {

                        var excelFiles = Directory.EnumerateFiles(path).Where(file => file.EndsWith(".xls") || file.EndsWith(".xlsx") || file.EndsWith(".xlsm"));
                        if (excelFiles.Count() > 0)
                        {
                            effects = DragDropEffects.Move;
                            var dir = Path.Combine(path, "Excsv");
                            Directory.CreateDirectory(dir);
                            foreach (var excelFile in excelFiles)
                            {
                                var extIndex = excelFile.LastIndexOf('\\');
                                var fileName = Path.GetFileNameWithoutExtension(excelFile);
                                var csv = Path.Combine(dir, fileName) + ".csv";
                                ExcelFileExtension.SaveAsCSV(excelFile, csv);

                            }
                            Process.Start(dir);
                        }
                    }
                    else if (File.Exists(path))
                    {
                        var extension = Path.GetExtension(path).ToLower();
                        if (extension == ".xls" || extension == ".xlsx" || extension == ".xlsm")
                        {

                            effects = DragDropEffects.Move;
                            var extIndex = path.LastIndexOf('.');
                            var csv = path.Replace(path.Substring(extIndex, path.Length - extIndex), ".csv");
                            ExcelFileExtension.SaveAsCSV(path, csv);
                            Process.Start(csv);
                        }
                    }
                }
                e.Effect = effects;
            };
            this.DragLeave += (s, e) =>
            {
                Console.WriteLine("Leave");
            };
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
