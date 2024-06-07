using MiniExcelLibs.Utils;
using System;
using System.Collections.Generic;
using System.IO;

namespace MiniExcelLibs
{
    /// <summary>
    /// 
    /// </summary>
    /// <param name="cellIndex">单元格索引</param>
    /// <param name="columnInfo">列信息</param>
    /// <param name="cellValue">单元格值</param>
    /// <returns></returns>
    public delegate (object, NewColumnInfo) CellDataGeterDelegate(CellIndex cellIndex, ExcelColumnInfo columnInfo, object cellValue);

    public static partial class MiniExcel
    {

        /// <summary>
        /// 保存Excel, 并修改列数据类型
        /// </summary>
        /// <param name="path">目录不存在时，字段创建</param>
        /// <param name="value"></param>
        /// <param name="printHeader"></param>
        /// <param name="sheetName"></param>
        /// <param name="configuration"></param>
        public static SaveExceleNext SaveFileAndChangeColumn(string path, object value, bool printHeader = true, string sheetName = "Sheet1", IConfiguration configuration = null, bool overwriteFile = false, CellDataGeterDelegate cellDataGeter = null)
        {
            var dir = Path.GetDirectoryName(path);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            using (var stream = overwriteFile ? File.Create(path) : new FileStream(path, FileMode.CreateNew))
            {
                ExcelWriterFactory.GetXLSXProvider(stream, value, sheetName, configuration, printHeader, cellDataGeter).SaveAs();
            }

            var saveExceleNext = new SaveExceleNext { Path = path, PrintHeader = printHeader };
            return saveExceleNext;
        }
    }

    public class SaveExceleNext
    {
        internal string Path { get; set; }
        internal bool PrintHeader { get; set; }

        public void ChangeColumn(Dictionary<string, Dictionary<int, NewColumnInfo>> newSheetColumnDic, Func<string, NewColumnInfo, bool, string> reWriteValue = null)
        {
            MiniExcelChangeDataType.ChangeDataType(Path, PrintHeader, newSheetColumnDic, reWriteValue);
        }
    }

    /// <summary>
    /// 单元格索引
    /// </summary>
    public class CellIndex
    {
        public CellIndex(int rowIndex, int columnIndex, int sheetIndex, string sheetName)
        {
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
            SheetIndex = sheetIndex;
            SheetName = sheetName;
        }


        /// <summary>
        /// 行索引，从1开始
        /// </summary>
        public int RowIndex { get; }

        /// <summary>
        /// 列索引，从1开始
        /// </summary>
        public int ColumnIndex { get; }

        /// <summary>
        /// sheet索引，从1开始
        /// </summary>
        public int SheetIndex { get; }

        public string SheetName { get; }
    }


    public class NewColumnInfo
    {
        /// <summary>
        /// 字段名称
        /// </summary>
        public string FieldName { get; set; }

        /// <summary>
        /// Excel列索引, 从1开始
        /// </summary>
        public int ColumnIndex { get; set; }

        /// <summary>
        /// 来源类型
        /// </summary>
        public CellDataType SourceDataType { get; set; } = CellDataType.Other;

        /// <summary>
        /// 目标类型
        /// </summary>
        public CellDataType TargetDataType { get; set; }

        /// <summary>
        /// TargetDataType为DateTime时生效
        /// </summary>
        public string FormatStr { get; set; }

        /// <summary>
        /// 提前创建好的样式索引
        /// </summary>
        internal int NewStyleIndex { get; set; } = 2;

        /// <summary>
        /// 存储临时值，不参与Excel的构造
        /// </summary>
        public object TempColumnValue { get; set; }

    }

    public enum CellDataType
    {
        DateTime,
        Number,
        String,
        Other
    }
}
