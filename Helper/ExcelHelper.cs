using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;

namespace Helper
{
    public class ExcelHelper
    {
        public static void Export<T>(string path, List<T> data)
        {
            using (var fstream = System.IO.File.Create(path))
            {

                IWorkbook workbook = new HSSFWorkbook();
                //创建一个 sheet 表
                ISheet sheet = workbook.CreateSheet(typeof(T).Name);

                //创建单元格样式
                ICellStyle cellStyle = workbook.CreateCellStyle();

                //创建格式
                IDataFormat dataFormat = workbook.CreateDataFormat();

                //设置为文本格式，也可以为 text，即 dataFormat.GetFormat("text");
                cellStyle.DataFormat = dataFormat.GetFormat("@");
                var props = typeof(T).GetProperties();

                //创建一个单元格
                ICell cell = null;
                var header = sheet.CreateRow(0);
                //设置列名
                for (int j = 0; j < props.Length; j++)
                {
                    cell = header.CreateCell(j);
                    //创建单元格并设置单元格内容
                    cell.SetCellValue(props[j].Name);
                    //设置单元格格式
                    cell.CellStyle = cellStyle;
                }

                //写入数据
                for (int i = 0; i < data.Count; i++)
                {
                    //跳过第一行，第一行为列名
                    IRow row = sheet.CreateRow(i + 1);

                    for (int j = 0; j < props.Length; j++)
                    {
                        cell = row.CreateCell(j);
                        var p = props[j];
                        var obj = p.GetValue(data[i]);
                        cell.CellStyle = cellStyle;
                        if (IsNumber(p.PropertyType) && obj != null)
                        {
                            cell.SetCellValue(double.Parse(obj.ToString()));
                            cell.SetCellType(CellType.Numeric);
                        }
                        else
                            cell.SetCellValue(obj?.ToString());

                    }
                }

                workbook.Write(fstream);
                workbook.Close();
            }
        }

        private static bool IsNumber(Type type)
        {
            return type.IsValueType && type.IsPrimitive && type != typeof(bool) && type != typeof(char);
        }
        public static IEnumerable<dynamic> ReadExcel(string path, int sheetIndex = 0, bool skipHeader = true)
        {
            using (var fstream = System.IO.File.OpenRead(path))
            {

                IWorkbook workbook = new HSSFWorkbook(fstream);

                var sheet = workbook.GetSheetAt(sheetIndex);
                var header = sheet.GetRow(0);
                var hsells = header.GetEnumerator();
                Dictionary<int, string> dic = new Dictionary<int, string>();
                int i = -1;
                while (hsells.MoveNext())
                {
                    i++;
                    string v;
                    if (skipHeader)
                    {
                        if (hsells.Current.CellType != CellType.String) continue;
                        v = hsells.Current.StringCellValue;
                        if (dic.ContainsValue(v))
                            v += i;
                    }
                    else
                        v = "C" + i;
                    dic.Add(i, v);
                }
                var rows = sheet.GetRowEnumerator();
                if (skipHeader)
                    rows.MoveNext();//跳过row0
                return Read();

                IEnumerable<dynamic> Read()
                {
                    while (rows.MoveNext())
                    {
                        var row = rows.Current as IRow;
                        if (row == null) continue;
                        yield return ReadRowAsDynamic(dic, row);

                    }
                }
            }

        }

        private static dynamic ReadRowAsDynamic(Dictionary<int, string> dic, IRow row)
        {
            ICell cell;
            var dy = new System.Dynamic.ExpandoObject();
            var p = (IDictionary<string, object>)dy;
            foreach (var item in dic)
            {
                cell = row.GetCell(item.Key);
                switch (cell.CellType)
                {
                    case CellType.Unknown:
                        break;
                    case CellType.Numeric:
                        p.Add(item.Value, cell.NumericCellValue);
                        break;
                    case CellType.String:
                        p.Add(item.Value, cell.StringCellValue);
                        break;
                    case CellType.Formula:
                        p.Add(item.Value, cell.NumericCellValue);
                        break;
                    case CellType.Blank:
                        p.Add(item.Value, null);
                        break;
                    case CellType.Boolean:
                        p.Add(item.Value, cell.BooleanCellValue);
                        break;
                    case CellType.Error:
                        p.Add(item.Value, cell.ErrorCellValue);

                        break;
                }

            }

            return dy;
        }

    }
    
}
