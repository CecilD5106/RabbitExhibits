using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace RabbitExhibits
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = "E:\\Excel\\2019OnlineExhibits.xlsx";
            Application excel = new Application();
            Workbook wb = excel.Workbooks.Open(path);

            try
            {
                Worksheet wsRabbits = wb.Worksheets["Rabbits"];
                Worksheet wsExhibitors = wb.Worksheets["Exhibitors"];
                Worksheet wsDepartment = wb.Worksheets["Department"];
                Worksheet wsDivision = wb.Worksheets["Division"];
                Console.WriteLine("Starting Worksheet Update");
                for (int i = 2; i < 124; i++)
                {
                    string sKey = wsRabbits.Cells[i, 2].Value;
                    for (int j = 2; j < 2828; j++)
                    {
                        if (wsExhibitors.Cells[j, 2].Value == sKey)
                        {
                            wsRabbits.Cells[i, 3].Value = wsExhibitors.Cells[j, 3].Value + " " + wsExhibitors.Cells[j, 4].Value;
                        }
                    }

                    int DeptID = wsRabbits.Cells[i, 4].Value;
                    for (int k = 2; k < 12; k++)
                    {
                        if (wsDepartment.Cells[k, 1].Value == DeptID)
                        {
                            wsRabbits.Cells[i, 4].Value = wsDepartment.Cells[k, 2].Value;
                        }
                    }

                    int DivID = wsRabbits.Cells[i, 5].Value;
                    for (int l = 2; l < 17; l++)
                    {
                        if (wsDivision.Cells[l, 1].Value == DivID)
                        {

                        }
                    }

                    Console.WriteLine("On row number " + i.ToString());
                }
                wb.Save();
                excel.Quit();
            }
            catch (Exception e)
            {
                excel.Quit();
                Console.WriteLine(e.ToString());
                throw;
            }
            finally
            {
                excel.Quit();
            }
        }
    }
}
