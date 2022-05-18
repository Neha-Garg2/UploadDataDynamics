using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Xrm.Tooling.Connector;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.ServiceModel.Description;
using System.Net;
using System.Configuration;

namespace UploadDataDynamics
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Class1 getService = new Class1();
            CrmServiceClient service = getService.connect();
            if (service != null)
            {
                Console.WriteLine("Connection Established with crm");

            }
            Console.WriteLine("Connection Created Successfully...!");
            Console.WriteLine("Counting Records...!");



            //Create COM Objects. Create a COM object for everything that is referenced
                Excel.Application xlApp = new Excel.Application();
                Console.WriteLine("Start Reading Excel file!!.");
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\hp\Documents\DataForUploading.xlsx");
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                Console.WriteLine("No. Of Rows " + rowCount);
                Console.WriteLine("No. of Columns " +colCount);

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i <= rowCount; i++)
            {
                //for (int j = 1; j <= colCount; j++)
               // {
                    //String str = (String)(xlRange.Cells[rowCount, colCount] as Excel.Range).Value2;
                    string a = xlRange.Cells[i, 1].Value2.ToString();
                    string b = xlRange.Cells[i, 2].Value2.ToString();
                    string c = xlRange.Cells[i, 3].Value2.ToString();

                    Entity localentity = new Entity("new_loginwithcrm");
                    localentity.Attributes["new_id"] = a;
                    localentity.Attributes["new_username"] = b;
                    localentity.Attributes["new_mobileno"] = c;
                    Console.Write("Updated Successfully");
                    service.Create(localentity);
                //}
                
            }
            

            Console.Write("Updated!!!!!!");
            Console.ReadLine();
        }

    }
}

// var a = xlRange.Cells[i, j].Value2.ToString();
//string a = xlRange.Cells[i, 1].Value2.ToString();
//string b = xlRange.Cells[i, 2].Value2.ToString();
//string c = xlRange.Cells[i, 3].Value2.ToString();
// Console.Write(a );


//Entity localentity = new Entity("new_loginwithcrm");
// Entity updatedTeam = new Entity("crc66_datausingexcel");
/*
    if (j == 1)
    {
        localentity.Attributes["new_name"] = xlRange.Cells[i, j].Value2.ToString();

    }
    if (j == 2)
    {
        localentity.Attributes["new_id"] = xlRange.Cells[i, j + 1].Value2.ToString();
    }
    if (j == 3)
    {
        localentity.Attributes["new_mobileno"] = xlRange.Cells[i, j + 2].Value2.ToString();                      

    }
*/

// Console.Write(xlRange.Cells[i, 1].Value2.ToString() + " ");
//var firstrow = xlRange.Cells[i, 1].Value2.ToString();

// Console.Write(xlRange.Cells[i, 2].Value2.ToString() + " ");
//var secRow = xlRange.Cells[i, 2].Value2.ToString();

//Console.Write(xlRange.Cells[i, 3].Value2.ToString() + "  ");
//var thirdRow = xlRange.Cells[i, 3].Value2.ToString();
//Console.Write("\n");

//Create local entity
// var Entity = "new_loginwithcrm";
// EntityCollection entityCollection = service.Retrieve("new_loginwithcrm", new ColumnSet(true));

// Entity createlocal = new Entity("new_loginwithcrm");
//EntityCollection ec = service.Create("new_loginwithcrm");


//createlocal.Attributes["new_name"] = xlRange.Cells[i, 1].Value2.ToString();
//createlocal.Attributes["new_name"] = xlRange.Cells[i, 1].Value2.ToString();
//createlocal.Attributes["new_name"] = xlRange.Cells[i, 1].Value2.ToString();

//service.Update(createlocal);



//      }
//    Console.ReadLine();
//}

//}
//}
