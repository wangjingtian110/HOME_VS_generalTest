using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
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

            //E:\code\vs\Console\ConsoleApp1\bin\Debug
            string path = @"..\..\Temp\CaliReport.xls";
            string outPath = @"..\..\Temp\TestOut.xls";            

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

            List<string> exAdd = new List<string>()
            {
                "4","6","7"
            };
           
            if(toAdd.Count == 1)
            {
                string[] values = toAdd[0].Split(',');
                sheet.GetRow(startRow).GetCell(1).SetCellValue(values[0]);
                sheet.GetRow(startRow).GetCell(2).SetCellValue(values[2]);

            }
            else
            {
                //行移动
                sheet.ShiftRows(startRow + 1, sheet.LastRowNum, toAdd.Count - 1, true, false);
                var rowSource = sheet.GetRow(startRow);
                var rowStyle = rowSource.RowStyle;

                for(int i = 0; i< toAdd.Count; i++)
                {
                    string[] values = toAdd[i].Split(',');
                    if(i == 0)
                    {
                        sheet.GetRow(startRow + i).GetCell(1).SetCellValue(values[0]);
                        sheet.GetRow(startRow + i).GetCell(2).SetCellValue(values[1]);
                    }
                    else
                    {
                        var row = sheet.CreateRow(startRow + i);
                        //行的格式，一般为null
                        if (rowStyle != null)
                            row.RowStyle = rowStyle;
                        row.Height = rowSource.Height;
                        
                        //对于本行的每个单元格，赋予和源行同样的格式
                        for (int col = 0; col < rowSource.LastCellNum; col++)
                        {
                            var cellsource = rowSource.GetCell(col);
                            var cellInsert = row.CreateCell(col);
                            var cellStyle = cellsource.CellStyle;
                            //设置单元格样式　　　　
                            if (cellStyle != null)
                                cellInsert.CellStyle = cellsource.CellStyle;
                        }
                        row.Cells[1].SetCellValue(values[0]);
                        row.Cells[2].SetCellValue(values[1]);
                    }
                }
                if(exAdd.Count > 0)
                {

                }

                //合并单元格
                CellRangeAddress region = new CellRangeAddress(startRow, toAdd.Count + startRow - 1, 3, 4);
                sheet.AddMergedRegion(region);
                sheet.GetRow(startRow).GetCell(3).SetCellValue("误差");

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
