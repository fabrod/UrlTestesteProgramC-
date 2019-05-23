using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;           
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab

namespace ReadExcelFileApp 
{
    public class Read_From_Excel

    {
       // public static void Main() => getExcelFile();

        public static void getExcelFile()
        {

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Data\Sundown2019\public\SdFileProducts.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string currentStatus = "";
            string currentCellvalue = "";
            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    //new line
                    if (j == 1)
                        Console.Write("\r\n");

                    //write the value to the console
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        currentCellvalue = xlRange.Cells[i, j].Value2.ToString();
                    if (currentCellvalue.Contains("http")) {
                        currentStatus = GetPage(xlRange.Cells[i, j].Value2.ToString());
                        //   Console.Write(xlRange.Cells[i, j].Value2.ToString() + " " + currentStatus + "\t");
                        // WriteToFile(currentCellvalue, currentStatus);
                        xlRange.Cells[i, j + 1] = currentStatus;
                    
                        }
                }
            }
            xlWorksheet.SaveAs(@"C:\Data\Sundown2019\public\SdFileProducts.xlsx");
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

          

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
        //add on to this program , two variables and a for loop for iteration
        private static string GetPage(string url)
        {
            string statusCode = "failed";
            try
            {
                // Creates an HttpWebRequest for the specified URL. 
                HttpWebRequest myHttpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                // Sends the HttpWebRequest and waits for a response.
                HttpWebResponse myHttpWebResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();
                statusCode = myHttpWebResponse.StatusCode.ToString();
                if (myHttpWebResponse.ResponseUri.ToString() != url) {
                    statusCode = myHttpWebResponse.ResponseUri.ToString();
                }
                if (myHttpWebResponse.StatusCode == HttpStatusCode.OK) 
                   // Console.WriteLine("\r\nResponse Status Code is OK and StatusDescription is: {0}",
                      //                   myHttpWebResponse.StatusDescription);
                       
                // Releases the resources of the response.
                myHttpWebResponse.Close();

            }
            catch (WebException e)
            {
                statusCode = e.Status.ToString();
                Console.WriteLine("\r\nWebException Raised. The following error occured : {0}", e.Status);
            }
            catch (Exception e)
            {
                statusCode = e.Message;
                Console.WriteLine("\nThe following Exception was raised : {0}", e.Message);
            }
            return statusCode;
        }
       
    }


} 






            
          


