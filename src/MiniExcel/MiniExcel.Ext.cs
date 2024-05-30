using MiniExcelLibs.Utils;
using System;
using System.Collections.Generic;
using System.IO;

namespace MiniExcelLibs
{
    /// <summary>
    /// 
    /// </summary>
    /// <param name="dataRowIdx">行索引，从1开始</param>
    /// <param name="columnIdx">列索引，从1开始</param>
    /// <param name="columnInfo">列信息</param>
    /// <param name="cellValue">单元格值</param>
    /// <returns></returns>
    public delegate (object, NewColumnInfo) CellDataGeterDelegate(int dataRowIdx,int columnIdx, ExcelColumnInfo columnInfo, object cellValue);

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
            if(!Directory.Exists(dir))
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

        public void ChangeColumn(Dictionary<int, NewColumnInfo> newColumns, Func<string, NewColumnInfo, bool, string> reWriteValue = null)
        {
            MiniExcelChangeDataType.ChangeDataType(Path, PrintHeader, newColumns, reWriteValue);
        }
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

    }

    public enum CellDataType
    {
        DateTime,
        Number,
        String,
        Other
    }
}
