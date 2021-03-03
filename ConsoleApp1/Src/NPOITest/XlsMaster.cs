using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    public class XlsMaster
    {        
        public void StartWork()
        {
            string path = @"E:\code\vs\ConsoleApp1\Temp\CaliReport.xls";
            string outPath = @"E:\code\vs\ConsoleApp1\Temp\TestOut.xls";

            //创建Excel文件的对象             
            FileStream fs = new FileStream(path, FileMode.Open);
            HSSFWorkbook workbook = new HSSFWorkbook(fs);

            ISheet sheet = (HSSFSheet)workbook.GetSheetAt(0);

            int startRow = 1;//开始插入行索引
            //模拟数据导入
            List<string> toAdd = new List<string>()
            {
                "1,2",
                "3,4",
                "5,6"
            };
           
            if(toAdd.Count == 1)
            {
                string[] values = toAdd[0].Split(',');
                sheet.GetRow(startRow).GetCell(1).SetCellValue(values[0]);
                sheet.GetRow(startRow).GetCell(2).SetCellValue(values[2]);

            }
            else
            {
                sheet.ShiftRows(startRow + 1, sheet.LastRowNum, toAdd.Count - 1, true, false);
                var rowSource = sheet.GetRow(startRow);
                var rowStyle = rowSource.RowStyle;

                for(int i = 0; i< toAdd.Count; i++)
                {
                    string[] values = toAdd[i].Split(',');
                    if(i == 0)
                    {
                        sheet.GetRow(startRow + i).GetCell(1).SetCellValue(values[0]);
                        sheet.GetRow(startRow + i).GetCell(2).SetCellValue(values[2]);
                    }
                    else
                    {
                        var row = sheet.CreateRow(startRow + i);
                        row.RowStyle = rowStyle;
                        row.CreateCell(1).SetCellValue(values[0]);
                        row.CreateCell(2).SetCellValue(values[1]);
                    }
                    
                }

            }           

            //绑定数据
            //for (int j = 0; j < toAdd.Count; j++)
            //{
            //    //单元格赋值等其他代码
            //    IRow r = sheet.GetRow(j);
            //    r.Cells[0].SetCellValue(j + 1);
            //}
            //后续操作。。。。。。。。。。

            //输出
            FileStream fso = new FileStream(outPath, FileMode.Create);
            workbook.Write(fso);
            fso.Flush();
            fso.Close();

        }
    }
}
