using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text; 
using Newtonsoft.Json; 
using Microsoft.Win32;
using Newtonsoft.Json.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel; 

namespace JsonExcel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string json;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnLoad_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            
            dialog.Filter = "Excel 2007|*.xlsx|Excel 2003|*.xls|All files (*.*)|*.*";
            
            if (dialog.ShowDialog() == true)
            {
                string file = dialog.FileName; 

                try
                {
                    json = this.ExcelTable(file); 
                
                    this.richtext.AppendText("load ok." + "\n\r");
                    this.richtext.AppendText(json + "\n\r");
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        } 
       
        private void btnExport_Click(object sender, RoutedEventArgs e)
        { 
            try
            { 
                SaveFileDialog dialog = new SaveFileDialog();

                dialog.Filter = "Json text (*.json)|*.txt|All files (*.*)|*.*";

                if (dialog.ShowDialog() == true) {
                    
                    string filetxt = dialog.FileName; 

                    File.WriteAllText(filetxt, this.json);

                    this.richtext.AppendText("export ok.");
                } 
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        } 
        
        public string ExcelTable(string file)
        { 
            var list = new JArray() as dynamic;

            using (FileStream fstream = new FileStream(file, FileMode.Open))
            {
                IWorkbook wbook = new XSSFWorkbook(fstream);  
      
                //only 1 sheet
                ISheet sheet = wbook.GetSheetAt(0);

                var header = sheet.GetRow(0).Cells;            

                for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);

                    if (row == null) continue;

                    dynamic obj = new JObject();
 
                    for (int j = 0; j < header.Count; j++)
                    {
                        ICell cell = row.GetCell(j);

                        if (cell != null)
                        {                   
                            obj.Add(header[j].ToString(), cell.ToString()); 
                        }
                    }

                    obj.Add("CampaignId", "EVMK");
                    obj.Add("Status", "pending");

                    list.Add(obj);
                }
            } 

            return list.ToString();
        }  
    }   
}
