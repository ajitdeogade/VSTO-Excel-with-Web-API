using System;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System.Net;
using Newtonsoft.Json;


namespace ExcelAddInTest2
{
    public partial class Ribbon1
    {

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

            Microsoft.Office.Interop.Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;

            Range actCell = Globals.ThisAddIn.Application.ActiveCell;
            int row = actCell.Row;
            int col = actCell.Column;
            int dataIndex = -1;


            if (actCell.Value2 != null)
            {
                string strValue2 = actCell.Value2;
                string strText = actCell.Text;

                string countryName = actCell.Text;
                activeWorksheet.Cells[row, col + 1] = "";
                string tmpcountryName = "";

                //API URL
                //string strURI = "https://restcountries.eu/rest/v2/all";  //All Countries
                string strURI = "https://restcountries.eu/rest/v2/name/" + countryName.Trim() + "?fulltext=true"; //Selected Country

                using (var client = new WebClient()) //WebClient  
                {
                    client.Headers.Add("Content-Type:application/json"); //Content-Type  
                    client.Headers.Add("Accept:application/json");

                    try
                    {
                        var result = client.DownloadString(strURI);

                        dynamic data = JsonConvert.DeserializeObject<dynamic>(result);

                        //Some countries has multiple data like China (i.e. China, Taiwan and Macao)
                        //down and dirty code - rectify in leisure with linq
                        for (int i = 0; i < data.Count; i++)
                        {
                            tmpcountryName = data[i]["name"];
                            if (tmpcountryName.ToLower() == countryName.ToLower())
                            {
                                dataIndex = i;
                                break;
                            }
                        }

                        if (dataIndex==-1)
                        {
                            activeWorksheet.Cells[row, col + 1] = "No Data Found";
                        }
                        else
                        { 
                            activeWorksheet.Cells[row, col + 1] = data[dataIndex]["capital"];
                        }
                        //System.Windows.Forms.MessageBox.Show(result);

                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(ex.ToString());
                        //activeWorksheet.Cells[row, col + 1] = "No Data Found - " + ex.ToString();
                    }
                }

            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Select Cell with Data");
            }

        }
    }
}
