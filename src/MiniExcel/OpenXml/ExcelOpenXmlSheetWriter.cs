﻿using MiniExcelLibs.Utils;
using MiniExcelLibs.Zip;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using static MiniExcelLibs.Utils.ImageHelper;

namespace MiniExcelLibs.OpenXml
{
    internal class FileDto
    {
        public string ID { get; set; } = $"R{Guid.NewGuid():N}";
        public string Extension { get; set; }
        public string Path { get { return $"xl/media/{ID}.{Extension}"; } }
        public string Path2 { get { return $"/xl/media/{ID}.{Extension}"; } }
        public Byte[] Byte { get; set; }
        public int RowIndex { get; set; }
        public int CellIndex { get; set; }
        public bool IsImage { get; set; } = false;
        public int SheetId { get; set; }
    }
    internal class SheetDto
    {
        public string ID { get; set; } = $"R{Guid.NewGuid():N}";
        public string Name { get; set; }
        public int SheetIdx { get; set; }
        public string Path { get { return $"xl/worksheets/sheet{SheetIdx}.xml"; } }

        public string State { get; set; }
    }
    internal class DrawingDto
    {
        public string ID { get; set; } = $"R{Guid.NewGuid():N}";
    }
    internal partial class ExcelOpenXmlSheetWriter : IExcelWriter
    {
        private readonly MiniExcelZipArchive _archive;
        private readonly static UTF8Encoding _utf8WithBom = new System.Text.UTF8Encoding(true);
        private readonly OpenXmlConfiguration _configuration;
        private readonly Stream _stream;
        private readonly bool _printHeader;
        private readonly object _value;
        private readonly List<SheetDto> _sheets = new List<SheetDto>();
        private readonly List<FileDto> _files = new List<FileDto>();
        private int currentSheetIndex = 0;

        private readonly CellDataGeterDelegate _cellDataGeter;


        public ExcelOpenXmlSheetWriter(Stream stream, object value, string sheetName, IConfiguration configuration, bool printHeader, CellDataGeterDelegate cellDataGeter = null)
        {
            this._stream = stream;
            // Why ZipArchiveMode.Update not ZipArchiveMode.Create?
            // R : Mode create - ZipArchiveEntry does not support seeking.'
            this._configuration = configuration as OpenXmlConfiguration ?? OpenXmlConfiguration.DefaultConfig;
            if (_configuration.FastMode)
                this._archive = new MiniExcelZipArchive(_stream, ZipArchiveMode.Update, true, _utf8WithBom);
            else
                this._archive = new MiniExcelZipArchive(_stream, ZipArchiveMode.Create, true, _utf8WithBom);
            this._printHeader = printHeader;
            this._value = value;
            _cellDataGeter = cellDataGeter;
            var defaultSheetInfo = GetSheetInfos(sheetName);
            _sheets.Add(defaultSheetInfo.ToDto(1)); //TODO:remove
        }

        public ExcelOpenXmlSheetWriter()
        {
        }

        public void SaveAs()
        {
            GenerateDefaultOpenXml();
            {
                if (_value is IDictionary<string, object>)
                {
                    var sheetId = 0;
                    var sheets = _value as IDictionary<string, object>;
                    _sheets.RemoveAt(0);//TODO:remove
                    foreach (var sheet in sheets)
                    {
                        sheetId++;
                        var sheetInfos = GetSheetInfos(sheet.Key);
                        var sheetDto = sheetInfos.ToDto(sheetId);
                        _sheets.Add(sheetDto); //TODO:remove

                        currentSheetIndex = sheetId;

                        CreateSheetXml(sheet.Value, sheetDto.Path);
                    }
                }
                else if (_value is DataSet)
                {
                    var sheetId = 0;
                    var sheets = _value as DataSet;
                    _sheets.RemoveAt(0);//TODO:remove
                    foreach (DataTable dt in sheets.Tables)
                    {
                        sheetId++;
                        var sheetInfos = GetSheetInfos(dt.TableName);
                        var sheetDto = sheetInfos.ToDto(sheetId);
                        _sheets.Add(sheetDto); //TODO:remove

                        currentSheetIndex = sheetId;

                        CreateSheetXml(dt, sheetDto.Path);
                    }
                }
                else
                {
                    //Single sheet export.
                    currentSheetIndex++;

                    CreateSheetXml(_value, _sheets[0].Path);
                }
            }
            GenerateEndXml();
            _archive.Dispose();
        }

        internal void GenerateDefaultOpenXml()
        {
            CreateZipEntry("_rels/.rels", "application/vnd.openxmlformats-package.relationships+xml", _defaultRels);
            CreateZipEntry("xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", _defaultSharedString);
        }

        private void GenerateSheetByEnumerable(MiniExcelStreamWriter writer, IEnumerable values)
        {
            var maxColumnIndex = 0;
            var maxRowIndex = 0;
            List<ExcelColumnInfo> props = null;
            string mode = null;

            int? rowCount = null;
            var collection = values as ICollection;
            if (collection != null)
            {
                rowCount = collection.Count;
            }
            else if (!_configuration.FastMode)
            {
                // The row count is only required up front when not in fastmode
                collection = new List<object>(values.Cast<object>());
                rowCount = collection.Count;
            }

            // Get the enumerator once to ensure deferred linq execution
            var enumerator = (collection ?? values).GetEnumerator();

            // Move to the first item in order to inspect the value type and determine whether it is empty
            var empty = !enumerator.MoveNext();

            if (empty)
            {
                // only when empty IEnumerable need to check this issue #133  https://github.com/shps951023/MiniExcel/issues/133
                var genericType = TypeHelper.GetGenericIEnumerables(values).FirstOrDefault();
                if (genericType == null || genericType == typeof(object) // sometime generic type will be object, e.g: https://user-images.githubusercontent.com/12729184/132812859-52984314-44d1-4ee8-9487-2d1da159f1f0.png
                    || typeof(IDictionary<string, object>).IsAssignableFrom(genericType)
                    || typeof(IDictionary).IsAssignableFrom(genericType))
                {
                    WriteEmptySheet(writer);
                    return;
                }
                else
                {
                    SetGenericTypePropertiesMode(genericType, ref mode, out maxColumnIndex, out props);
                }
            }
            else
            {
                var firstItem = enumerator.Current;
                if (firstItem is IDictionary<string, object> genericDic)
                {
                    mode = "IDictionary<string, object>";
                    props = CustomPropertyHelper.GetDictionaryColumnInfo(genericDic, null, _configuration);
                    maxColumnIndex = props.Count;
                }
                else if (firstItem is IDictionary dic)
                {
                    mode = "IDictionary";
                    props = CustomPropertyHelper.GetDictionaryColumnInfo(null, dic, _configuration);
                    //maxColumnIndex = dic.Keys.Count;
                    maxColumnIndex = props.Count; // why not using keys, because ignore attribute ![image](https://user-images.githubusercontent.com/12729184/163686902-286abb70-877b-4e84-bd3b-001ad339a84a.png)
                }
                else
                {
                    SetGenericTypePropertiesMode(firstItem.GetType(), ref mode, out maxColumnIndex, out props);
                }
            }

            writer.Write($@"<?xml version=""1.0"" encoding=""utf-8""?><x:worksheet xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" >");

            long dimensionWritePosition = 0;

            // We can write the dimensions directly if the row count is known
            if (_configuration.FastMode && rowCount == null)
            {
                // Write a placeholder for the table dimensions and save thee position for later
                dimensionWritePosition = writer.WriteAndFlush("<x:dimension ref=\"");
                writer.Write("                              />");
            }
            else
            {
                maxRowIndex = rowCount.Value + (_printHeader && rowCount > 0 ? 1 : 0);
                writer.Write($@"<x:dimension ref=""{GetDimensionRef(maxRowIndex, maxColumnIndex)}""/>");
            }

            //cols:width
            WriteColumnsWidths(writer, props);

            //header
            writer.Write($@"<x:sheetData>");
            var yIndex = 1;
            var xIndex = 1;
            if (_printHeader)
            {
                PrintHeader(writer, props);
                yIndex++;
            }

            if (!empty)
            {
                // body
                if (mode == "IDictionary<string, object>") //Dapper Row
                    maxRowIndex = GenerateSheetByColumnInfo<IDictionary<string, object>>(writer, enumerator, props, xIndex, yIndex);
                else if (mode == "IDictionary") //IDictionary
                    maxRowIndex = GenerateSheetByColumnInfo<IDictionary>(writer, enumerator, props, xIndex, yIndex);
                else if (mode == "Properties")
                    maxRowIndex = GenerateSheetByColumnInfo<object>(writer, enumerator, props, xIndex, yIndex);
                else
                    throw new NotImplementedException($"Type {values.GetType().FullName} is not implemented. Please open an issue.");
            }

            writer.Write("</x:sheetData>");
            if (_configuration.AutoFilter)
                writer.Write($"<x:autoFilter ref=\"{GetDimensionRef(maxRowIndex, maxColumnIndex)}\" />");

            // The dimension has already been written if row count is defined
            if (_configuration.FastMode && rowCount == null)
            {
                // Flush and save position so that we can get back again.
                var pos = writer.Flush();

                // Seek back and write the dimensions of the table
                writer.SetPosition(dimensionWritePosition);
                writer.WriteAndFlush($@"{GetDimensionRef(maxRowIndex, maxColumnIndex)}""");
                writer.SetPosition(pos);
            }

            writer.Write("<x:drawing  r:id=\"drawing" + currentSheetIndex + "\" /></x:worksheet>");
        }

        private static void PrintHeader(MiniExcelStreamWriter writer, List<ExcelColumnInfo> props)
        {
            var xIndex = 1;
            var yIndex = 1;
            writer.Write($"<x:row r=\"{yIndex}\">");

            foreach (var p in props)
            {
                if (p == null)
                {
                    xIndex++; //reason : https://github.com/shps951023/MiniExcel/issues/142
                    continue;
                }

                var r = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                WriteC(writer, r, columnName: p.ExcelColumnName);
                xIndex++;
            }

            writer.Write("</x:row>");
        }

        private void CreateSheetXml(object value, string sheetPath)
        {
            ZipArchiveEntry entry = _archive.CreateEntry(sheetPath, CompressionLevel.Fastest);
            using (var zipStream = entry.Open())
            using (MiniExcelStreamWriter writer = new MiniExcelStreamWriter(zipStream, _utf8WithBom, _configuration.BufferSize))
            {
                if (value == null)
                {
                    WriteEmptySheet(writer);
                    goto End; //for re-using code
                }

                //DapperRow

                if (value is IDataReader)
                {
                    GenerateSheetByIDataReader(writer, value as IDataReader);
                }
                else if (value is IEnumerable)
                {
                    GenerateSheetByEnumerable(writer, value as IEnumerable);
                }
                else if (value is DataTable)
                {
                    GenerateSheetByDataTable(writer, value as DataTable);
                }
                else
                {
                    throw new NotImplementedException($"Type {value.GetType().FullName} is not implemented. Please open an issue.");
                }
            }
        End: //for re-using code
            _zipDictionary.Add(sheetPath, new ZipPackageInfo(entry, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
        }

        private void SetGenericTypePropertiesMode(Type genericType, ref string mode, out int maxColumnIndex, out List<ExcelColumnInfo> props)
        {
            mode = "Properties";
            if (genericType.IsValueType)
                throw new NotImplementedException($"MiniExcel not support only {genericType.Name} value generic type");
            else if (genericType == typeof(string) || genericType == typeof(DateTime) || genericType == typeof(Guid))
                throw new NotImplementedException($"MiniExcel not support only {genericType.Name} generic type");
            props = CustomPropertyHelper.GetSaveAsProperties(genericType, _configuration);

            maxColumnIndex = props.Count;
        }

        private void WriteEmptySheet(MiniExcelStreamWriter writer)
        {
            writer.Write($@"<?xml version=""1.0"" encoding=""utf-8""?><x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><x:dimension ref=""A1""/><x:sheetData></x:sheetData></x:worksheet>");
        }

        private int GenerateSheetByColumnInfo<T>(MiniExcelStreamWriter writer, IEnumerator value, List<ExcelColumnInfo> props, int xIndex = 1, int yIndex = 1)
        {
            var isDic = typeof(T) == typeof(IDictionary);
            var isDapperRow = typeof(T) == typeof(IDictionary<string, object>);
            do
            {
                // The enumerator has already moved to the first item
                T v = (T)value.Current;

                writer.Write($"<x:row r=\"{yIndex}\">");
                var cellIndex = xIndex;
                foreach (var columnInfo in props)
                {
                    if (columnInfo == null) //reason:https://github.com/shps951023/MiniExcel/issues/142
                    {
                        cellIndex++;
                        continue;
                    }
                    object cellValue = null;
                    if (isDic)
                    {
                        cellValue = ((IDictionary)v)[columnInfo.Key];
                        //WriteCell(writer, yIndex, cellIndex, cellValue, null); // why null because dictionary that needs to check type every time
                        //TODO: user can specefic type to optimize efficiency
                    }
                    else if (isDapperRow)
                    {
                        cellValue = ((IDictionary<string, object>)v)[columnInfo.Key.ToString()];
                    }
                    else
                    {
                        cellValue = columnInfo.Property.GetValue(v);
                    }
                    if (_cellDataGeter != null)
                        (cellValue, _) = _cellDataGeter(yIndex, cellIndex, columnInfo, cellValue);

                    WriteCell(writer, yIndex, cellIndex, cellValue, columnInfo);

                    cellIndex++;
                }
                writer.Write($"</x:row>");
                yIndex++;
            } while (value.MoveNext());

            return yIndex - 1;
        }

        private void WriteCell(MiniExcelStreamWriter writer, int rowIndex, int cellIndex, object value, ExcelColumnInfo columnInfo)
        {
            var columname = ExcelOpenXmlUtils.ConvertXyToCell(cellIndex, rowIndex);
            var valueIsNull = value is null || value is DBNull;

            if (_configuration.EnableWriteNullValueCell && valueIsNull)
            {
                writer.Write($"<x:c r=\"{columname}\" s=\"2\"></x:c>"); // s: style index
                return;
            }

            var tuple = GetCellValue(rowIndex, cellIndex, value, columnInfo, valueIsNull);

            var styleIndex = tuple.Item1; // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cell?view=openxml-3.0.1
            var dataType = tuple.Item2; // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-3.0.1
            var cellValue = tuple.Item3;

            if (cellValue != null && (cellValue.StartsWith(" ", StringComparison.Ordinal) || cellValue.EndsWith(" ", StringComparison.Ordinal))) /*Prefix and suffix blank space will lost after SaveAs #294*/
            {
                writer.Write($"<x:c r=\"{columname}\" {(dataType == null ? "" : $"t =\"{dataType}\"")} s=\"{styleIndex}\" xml:space=\"preserve\"><x:v>{cellValue}</x:v></x:c>");
            }
            else
            {
                //t check avoid format error ![image](https://user-images.githubusercontent.com/12729184/118770190-9eee3480-b8b3-11eb-9f5a-87a439f5e320.png)
                writer.Write($"<x:c r=\"{columname}\" {(dataType == null ? "" : $"t =\"{dataType}\"")} s=\"{styleIndex}\"><x:v>{cellValue}</x:v></x:c>");
            }
        }

        private Tuple<string, string, string> GetCellValue(int rowIndex, int cellIndex, object value, ExcelColumnInfo columnInfo, bool valueIsNull)
        {
            var styleIndex = "2"; // format code: 0.00
            var cellValue = string.Empty;
            var dataType = "str";

            if (valueIsNull)
            {
                // use defaults
            }
            else if (value is string str)
            {
                cellValue = ExcelOpenXmlUtils.EncodeXML(str);
            }
            else if (columnInfo?.ExcelFormat != null && value is IFormattable formattableValue)
            {
                var formattedStr = formattableValue.ToString(columnInfo.ExcelFormat, _configuration.Culture);
                cellValue = ExcelOpenXmlUtils.EncodeXML(formattedStr);
            }
            else
            {
                Type type;
                if (columnInfo == null || columnInfo.Key != null)
                {
                    // TODO: need to optimize
                    // Dictionary need to check type every time, so it's slow..
                    type = value.GetType();
                    type = Nullable.GetUnderlyingType(type) ?? type;
                }
                else
                {
                    type = columnInfo.ExcludeNullableType; //sometime it doesn't need to re-get type like prop
                }

                if (type.IsEnum)
                {
                    dataType = "str";
                    var description = CustomPropertyHelper.DescriptionAttr(type, value); //TODO: need to optimze
                    if (description != null)
                        cellValue = description;
                    else
                        cellValue = value.ToString();
                }
                else if (TypeHelper.IsNumericType(type))
                {
                    if (_configuration.Culture != CultureInfo.InvariantCulture)
                        dataType = "str"; //TODO: add style format
                    else
                        dataType = "n";

                    if (type.IsAssignableFrom(typeof(decimal)))
                        cellValue = ((decimal)value).ToString(_configuration.Culture);
                    else if (type.IsAssignableFrom(typeof(Int32)))
                        cellValue = ((Int32)value).ToString(_configuration.Culture);
                    else if (type.IsAssignableFrom(typeof(Double)))
                        cellValue = ((Double)value).ToString(_configuration.Culture);
                    else if (type.IsAssignableFrom(typeof(Int64)))
                        cellValue = ((Int64)value).ToString(_configuration.Culture);
                    else if (type.IsAssignableFrom(typeof(UInt32)))
                        cellValue = ((UInt32)value).ToString(_configuration.Culture);
                    else if (type.IsAssignableFrom(typeof(UInt16)))
                        cellValue = ((UInt16)value).ToString(_configuration.Culture);
                    else if (type.IsAssignableFrom(typeof(UInt64)))
                        cellValue = ((UInt64)value).ToString(_configuration.Culture);
                    else if (type.IsAssignableFrom(typeof(Int16)))
                        cellValue = ((Int16)value).ToString(_configuration.Culture);
                    else if (type.IsAssignableFrom(typeof(Single)))
                        cellValue = ((Single)value).ToString(_configuration.Culture);
                    else if (type.IsAssignableFrom(typeof(Single)))
                        cellValue = ((Single)value).ToString(_configuration.Culture);
                    else
                        cellValue = (decimal.Parse(value.ToString())).ToString(_configuration.Culture);
                }
                else if (type == typeof(bool))
                {
                    dataType = "b";
                    cellValue = (bool)value ? "1" : "0";
                }
                else if (type == typeof(byte[]) && _configuration.EnableConvertByteArray)
                {
                    var bytes = (byte[])value;
                    if (bytes != null)
                    {
                        // TODO: Setting configuration because it might have high cost?
                        var format = ImageHelper.GetImageFormat(bytes);
                        //it can't insert to zip first to avoid cache image to memory
                        //because sheet xml is opening.. https://github.com/shps951023/MiniExcel/issues/304#issuecomment-1017031691
                        //int rowIndex, int cellIndex
                        var file = new FileDto()
                        {
                            Byte = bytes,
                            RowIndex = rowIndex,
                            CellIndex = cellIndex,
                            SheetId = currentSheetIndex
                        };
                        if (format != ImageFormat.unknown)
                        {
                            file.Extension = format.ToString();
                            file.IsImage = true;
                        }
                        else
                        {
                            file.Extension = "bin";
                        }
                        _files.Add(file);

                        //TODO:Convert to base64
                        var base64 = $"@@@fileid@@@,{file.Path}";
                        cellValue = ExcelOpenXmlUtils.EncodeXML(base64);
                        styleIndex = "4";
                    }
                }
                else if (type == typeof(DateTime))
                {
                    if (_configuration.Culture != CultureInfo.InvariantCulture)
                    {
                        dataType = "str";
                        cellValue = ((DateTime)value).ToString(_configuration.Culture);
                    }
                    else if (columnInfo == null || columnInfo.ExcelFormat == null)
                    {
                        var oaDate = CorrectDateTimeValue((DateTime)value);

                        dataType = null;
                        styleIndex = "3";
                        cellValue = oaDate.ToString(CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        // TODO: now it'll lose date type information
                        dataType = "str";
                        cellValue = ((DateTime)value).ToString(columnInfo.ExcelFormat, _configuration.Culture);
                    }
                }
#if NET6_0_OR_GREATER
                else if (type == typeof(DateOnly))
                {
                    if (_configuration.Culture != CultureInfo.InvariantCulture)
                    {
                        dataType = "str";
                        cellValue = ((DateOnly)value).ToString(_configuration.Culture);
                    }
                    else if (columnInfo == null || columnInfo.ExcelFormat == null)
                    {
                        var day = (DateOnly)value;
                        var oaDate = CorrectDateTimeValue(day.ToDateTime(TimeOnly.MinValue));

                        dataType = null;
                        styleIndex = "3";
                        cellValue = oaDate.ToString(CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        // TODO: now it'll lose date type information
                        dataType = "str";
                        cellValue = ((DateOnly)value).ToString(columnInfo.ExcelFormat, _configuration.Culture);
                    }
                }
#endif
                else
                {
                    //TODO: _configuration.Culture
                    cellValue = ExcelOpenXmlUtils.EncodeXML(value.ToString());
                }
            }

            return Tuple.Create(styleIndex, dataType, cellValue);
        }

        private static double CorrectDateTimeValue(DateTime value)
        {
            var oaDate = value.ToOADate();

            // Excel says 1900 was a leap year  :( Replicate an incorrect behavior thanks
            // to Lotus 1-2-3 decision from 1983...
            // https://github.com/ClosedXML/ClosedXML/blob/develop/ClosedXML/Extensions/DateTimeExtensions.cs#L45
            const int nonExistent1900Feb29SerialDate = 60;
            if (oaDate <= nonExistent1900Feb29SerialDate)
            {
                oaDate -= 1;
            }

            return oaDate;
        }

        private void GenerateSheetByDataTable(MiniExcelStreamWriter writer, DataTable value)
        {
            var xy = ExcelOpenXmlUtils.ConvertCellToXY("A1");

            //GOTO Top Write:
            writer.Write($@"<?xml version=""1.0"" encoding=""utf-8""?><x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
            {
                var yIndex = xy.Item2;

                // dimension
                var maxRowIndex = value.Rows.Count + (_printHeader && value.Rows.Count > 0 ? 1 : 0);
                var maxColumnIndex = value.Columns.Count;
                writer.Write($@"<x:dimension ref=""{GetDimensionRef(maxRowIndex, maxColumnIndex)}""/>");

                var props = new List<ExcelColumnInfo>();
                for (var i = 0; i < value.Columns.Count; i++)
                {
                    var columnName = value.Columns[i].Caption ?? value.Columns[i].ColumnName;
                    var columnType = value.Columns[i].DataType;
                    var prop = GetColumnInfosFromDynamicConfiguration(columnName, columnType, i);
                    props.Add(prop);
                }

                WriteColumnsWidths(writer, props);

                writer.Write("<x:sheetData>");
                if (_printHeader)
                {
                    writer.Write($"<x:row r=\"{yIndex}\">");
                    var xIndex = xy.Item1;
                    foreach (var p in props)
                    {
                        var r = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                        WriteC(writer, r, columnName: p.ExcelColumnName);
                        xIndex++;
                    }

                    writer.Write($"</x:row>");
                    yIndex++;
                }

                for (int i = 0; i < value.Rows.Count; i++)
                {
                    writer.Write($"<x:row r=\"{yIndex}\">");
                    var xIndex = xy.Item1;

                    for (int j = 0; j < value.Columns.Count; j++)
                    {
                        var cellValue = value.Rows[i][j];
                        if (_cellDataGeter != null)
                            (cellValue, _) = _cellDataGeter(yIndex, xIndex, props[j], cellValue);
                        WriteCell(writer, yIndex, xIndex, cellValue, columnInfo: props[j]);
                        xIndex++;
                    }
                    writer.Write($"</x:row>");
                    yIndex++;
                }
                writer.Write("</x:sheetData>");
                if (_configuration.AutoFilter)
                    writer.Write($"<x:autoFilter ref=\"{GetDimensionRef(maxRowIndex, maxColumnIndex)}\" />");
                writer.WriteAndFlush("</x:worksheet>");
            }
        }

        private void GenerateSheetByIDataReader(MiniExcelStreamWriter writer, IDataReader reader)
        {
            long dimensionWritePosition = 0;
            writer.Write($@"<?xml version=""1.0"" encoding=""utf-8""?><x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
            var xIndex = 1;
            var yIndex = 1;
            var maxColumnIndex = 0;
            var maxRowIndex = 0;
            {

                if (_configuration.FastMode)
                {
                    dimensionWritePosition = writer.WriteAndFlush($@"<x:dimension ref=""");
                    writer.Write("                              />"); // end of code will be replaced
                }

                var props = new List<ExcelColumnInfo>();
                for (var i = 0; i < reader.FieldCount; i++)
                {
                    var columnName = reader.GetName(i);
                    var columnType = reader.GetFieldType(i);
                    var prop = GetColumnInfosFromDynamicConfiguration(columnName, columnType, i);
                    props.Add(prop);
                }
                maxColumnIndex = props.Count;

                WriteColumnsWidths(writer, props);

                writer.Write("<x:sheetData>");
                int fieldCount = reader.FieldCount;
                if (_printHeader)
                {
                    PrintHeader(writer, props);
                    yIndex++;
                }

                while (reader.Read())
                {
                    writer.Write($"<x:row r=\"{yIndex}\">");
                    xIndex = 1;
                    for (int i = 0; i < fieldCount; i++)
                    {
                        var cellValue = reader.GetValue(i);
                        if (_cellDataGeter != null)
                            (cellValue, _) = _cellDataGeter(yIndex, xIndex, props[i], cellValue);
                        WriteCell(writer, yIndex, xIndex, cellValue, columnInfo: props[i]);
                        xIndex++;
                    }
                    writer.Write($"</x:row>");
                    yIndex++;
                }

                // Subtract 1 because cell indexing starts with 1
                maxRowIndex = yIndex - 1;
            }
            writer.Write("</x:sheetData>");
            if (_configuration.AutoFilter)
                writer.Write($"<x:autoFilter ref=\"{GetDimensionRef(maxRowIndex, maxColumnIndex)}\" />");
            writer.WriteAndFlush("</x:worksheet>");

            if (_configuration.FastMode)
            {
                writer.SetPosition(dimensionWritePosition);
                writer.WriteAndFlush($@"{GetDimensionRef(maxRowIndex, maxColumnIndex)}""");
            }
        }

        /// <summary>
        /// 构建动态
        /// </summary>
        /// <param name="columnName">列字名称</param>
        /// <param name="columnType">列类型</param>
        /// <param name="excelColumnIndex">列索引,从0开始</param>
        /// <returns></returns>
        private ExcelColumnInfo GetColumnInfosFromDynamicConfiguration(string columnName, Type columnType, int excelColumnIndex)
        {
            var prop = new ExcelColumnInfo
            {
                ExcelColumnName = columnName,
                Key = columnName,
                ExcludeNullableType = Nullable.GetUnderlyingType(columnType) ?? columnType,
                ExcelColumnIndex = excelColumnIndex,
            };

            if (_configuration.DynamicColumns == null || _configuration.DynamicColumns.Length <= 0)
                return prop;

            var dynamicColumn = _configuration.DynamicColumns.SingleOrDefault(_ => _.Key == columnName);
            if (dynamicColumn == null || dynamicColumn.Ignore)
            {
                return prop;
            }

            prop.Nullable = true;
            //prop.ExcludeNullableType = item2[key]?.GetType();
            if (dynamicColumn.Format != null)
                prop.ExcelFormat = dynamicColumn.Format;
            if (dynamicColumn.Aliases != null)
                prop.ExcelColumnAliases = dynamicColumn.Aliases;
            if (dynamicColumn.IndexName != null)
                prop.ExcelIndexName = dynamicColumn.IndexName;
            if (dynamicColumn.Index >= 0)
                prop.ExcelColumnIndex = dynamicColumn.Index;
            if (dynamicColumn.Name != null)
                prop.ExcelColumnName = dynamicColumn.Name;
            prop.ExcelColumnWidth = dynamicColumn.Width;

            return prop;
        }

        private ExcellSheetInfo GetSheetInfos(string sheetName)
        {
            var info = new ExcellSheetInfo
            {
                ExcelSheetName = sheetName,
                Key = sheetName,
                ExcelSheetState = SheetState.Visible
            };

            if (_configuration.DynamicSheets == null || _configuration.DynamicSheets.Length <= 0)
                return info;

            var dynamicSheet = _configuration.DynamicSheets.SingleOrDefault(_ => _.Key == sheetName);
            if (dynamicSheet == null)
            {
                return info;
            }

            if (dynamicSheet.Name != null)
                info.ExcelSheetName = dynamicSheet.Name;
            info.ExcelSheetState = dynamicSheet.State;

            return info;
        }

        private static void WriteColumnsWidths(MiniExcelStreamWriter writer, IEnumerable<ExcelColumnInfo> props)
        {
            var ecwProps = props.Where(x => x?.ExcelColumnWidth != null).ToList();
            if (ecwProps.Count <= 0)
                return;
            writer.Write($@"<x:cols>");
            foreach (var p in ecwProps)
            {
                writer.Write(
                    $@"<x:col min=""{p.ExcelColumnIndex + 1}"" max=""{p.ExcelColumnIndex + 1}"" width=""{p.ExcelColumnWidth?.ToString(CultureInfo.InvariantCulture)}"" customWidth=""1"" />");
            }

            writer.Write($@"</x:cols>");
        }

        private static void WriteC(MiniExcelStreamWriter writer, string r, string columnName)
        {
            writer.Write($"<x:c r=\"{r}\" t=\"str\" s=\"1\">");
            writer.Write($"<x:v>{ExcelOpenXmlUtils.EncodeXML(columnName)}"); //issue I45TF5
            writer.Write($"</x:v>");
            writer.Write($"</x:c>");
        }

        private void GenerateEndXml()
        {
            AddFilesToZip();

            GenerateStylesXml();

            GenerateDrawinRelXml();

            GenerateDrawingXml();

            GenerateWorkbookXml();

            GenerateContentTypesXml();
        }

        private void AddFilesToZip()
        {
            foreach (var item in _files)
            {
                this.CreateZipEntry(item.Path, item.Byte);
            }
        }

        /// <summary>
        /// styles.xml
        /// </summary>
        private void GenerateStylesXml()
        {
            var styleXml = GetStylesXml();
            CreateZipEntry(@"xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", styleXml);
        }

        private void GenerateDrawinRelXml()
        {
            for (int sheetIndex = 0; sheetIndex < _sheets.Count; sheetIndex++)
            {
                var drawing = GetDrawingRelationshipXml(sheetIndex);
                CreateZipEntry(
                    $"xl/drawings/_rels/drawing{sheetIndex + 1}.xml.rels",
                    null,
                    _defaultDrawingXmlRels.Replace("{{format}}", drawing));
            }
        }

        private void GenerateDrawingXml()
        {
            for (int sheetIndex = 0; sheetIndex < _sheets.Count; sheetIndex++)
            {
                var drawing = GetDrawingXml(sheetIndex);

                CreateZipEntry(
                    $"xl/drawings/drawing{sheetIndex + 1}.xml",
                    "application/vnd.openxmlformats-officedocument.drawing+xml",
                    _defaultDrawing.Replace("{{format}}", drawing));
            }
        }

        /// <summary>
        /// workbook.xml 、 workbookRelsXml
        /// </summary>
        private void GenerateWorkbookXml()
        {
            GenerateWorkBookXmls(
                out StringBuilder workbookXml,
                out StringBuilder workbookRelsXml,
                out Dictionary<int, string> sheetsRelsXml);

            foreach (var sheetRelsXml in sheetsRelsXml)
            {
                CreateZipEntry(
                    $"xl/worksheets/_rels/sheet{sheetRelsXml.Key}.xml.rels",
                    null,
                    _defaultSheetRelXml.Replace("{{format}}", sheetRelsXml.Value));
            }

            CreateZipEntry(
                @"xl/workbook.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                _defaultWorkbookXml.Replace("{{sheets}}", workbookXml.ToString()));

            CreateZipEntry(
                @"xl/_rels/workbook.xml.rels",
                null,
                _defaultWorkbookXmlRels.Replace("{{sheets}}", workbookRelsXml.ToString()));
        }

        /// <summary>
        /// [Content_Types].xml
        /// </summary>
        private void GenerateContentTypesXml()
        {
            var contentTypes = GetContentTypesXml();

            CreateZipEntry(@"[Content_Types].xml", null, contentTypes);
        }

        private string GetDimensionRef(int maxRowIndex, int maxColumnIndex)
        {
            string dimensionRef;
            if (maxRowIndex == 0 && maxColumnIndex == 0)
                dimensionRef = "A1";
            else if (maxColumnIndex == 1)
                dimensionRef = $"A{maxRowIndex}";
            else if (maxRowIndex == 0)
                dimensionRef = $"A1:{ColumnHelper.GetAlphabetColumnName(maxColumnIndex - 1)}1";
            else
                dimensionRef = $"A1:{ColumnHelper.GetAlphabetColumnName(maxColumnIndex - 1)}{maxRowIndex}";
            return dimensionRef;
        }

        private void CreateZipEntry(string path, string contentType, string content)
        {
            ZipArchiveEntry entry = _archive.CreateEntry(path, CompressionLevel.Fastest);
            using (var zipStream = entry.Open())
            using (MiniExcelStreamWriter writer = new MiniExcelStreamWriter(zipStream, _utf8WithBom, _configuration.BufferSize))
                writer.Write(content);
            if (!string.IsNullOrEmpty(contentType))
                _zipDictionary.Add(path, new ZipPackageInfo(entry, contentType));
        }

        private void CreateZipEntry(string path, byte[] content)
        {
            ZipArchiveEntry entry = _archive.CreateEntry(path, CompressionLevel.Fastest);
            using (var zipStream = entry.Open())
                zipStream.Write(content, 0, content.Length);
        }

        public void Insert()
        {
            throw new NotImplementedException();
        }
    }
}