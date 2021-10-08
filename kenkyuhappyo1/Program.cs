using System;
using Excel = Microsoft.Office.Interop.Excel;



namespace kenkyuhappyo1
{
    class Program
    {
        static void Main(string[] args)
        {
            
         

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbooks excelBooks = excelApp.Workbooks;
            Excel.Workbook excelBook = excelBooks.Open("C:\\Users\\USER\\Desktop\\kenkyu\\testb"); //指定された場所のExcelファイルを開く
            Excel.Worksheet sheet = excelApp.Worksheets["sheet1"];

            Console.Write("学籍番号を入力してください：");
            var x = double.Parse(Console.ReadLine()); //学籍番号の入力
            Console.Write("月曜日～金曜日を1～5としたとき、今日の曜日にあたる数字を入力してください：");
            var y = double.Parse(Console.ReadLine()); //曜日の入力
            Console.Write("体温を入力してください[℃]：");
            var z = double.Parse(Console.ReadLine()); //体温の入力

            try
            {
                excelApp.Visible = false;
                sheet.Cells[x , y] = z; //指定されたセルに体温データを保存する

                excelBook.Save(); //Excelファイルの上書き保存
            }
            catch
            {
                throw;
            }

            finally
            {
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
                }
        }
    }
