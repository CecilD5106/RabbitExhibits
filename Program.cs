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
                Worksheet wsSection = wb.Worksheets["Section"];
                Worksheet wsClass = wb.Worksheets["Class"];

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

                    double DeptID = wsRabbits.Cells[i, 4].Value;
                    for (int k = 2; k < 12; k++)
                    {
                        if (wsDepartment.Cells[k, 1].Value == DeptID)
                        {
                            wsRabbits.Cells[i, 4].Value = wsDepartment.Cells[k, 2].Value;
                        }
                    }

                    double DivID = wsRabbits.Cells[i, 5].Value;
                    for (int l = 2; l < 17; l++)
                    {
                        if (wsDivision.Cells[l, 1].Value == DivID)
                        {
                            wsRabbits.Cells[i, 5].Value = wsDivision.Cells[l, 3].Value;
                        }
                    }

                    double SecID = wsRabbits.Cells[i, 6].Value;
                    for (int m = 2; m < 120; m++)
                    {
                        if (wsSection.Cells[m, 1].Value == SecID)
                        {
                            wsRabbits.Cells[i, 6].Value = wsSection.Cells[m, 3].Value;
                        }
                    }

                    double ClassID = wsRabbits.Cells[i, 7].Value;
                    for (int n = 2; n < 705; n++)
                    {
                        if (wsClass.Cells[n, 1].Value == ClassID)
                        {
                            wsRabbits.Cells[i, 7].Value = wsClass.Cells[n, 6].Value;
                        }
                    }

                    Console.WriteLine("Updated row number " + i.ToString());
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
