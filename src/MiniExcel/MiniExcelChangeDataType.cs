using MiniExcelLibs.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime;
using System.Xml;

namespace MiniExcelLibs
{
    internal class MiniExcelChangeDataType
    {
        public static void ChangeDataType(string fileName, bool useHeader, Dictionary<string, Dictionary<int, NewColumnInfo>> newSheetColumnDic, Func<string, NewColumnInfo, bool, string> reWriteValue = null)
        {
            //创建好临时文件夹
            var tempPath = Path.Combine(Path.GetTempPath(), AppDomain.CurrentDomain.FriendlyName);
            if (!Directory.Exists(tempPath))
                Directory.CreateDirectory(tempPath);

            XmlReaderSettings _xmlSettings = new XmlReaderSettings
            {
                IgnoreComments = true,
                IgnoreWhitespace = true,
                XmlResolver = null,
            };

            using (ZipArchive zip = ZipFile.Open(fileName, ZipArchiveMode.Update))
            {
                //读取xl/_rels/workbook.xml.rels
                var workbookRelsXml = new XmlDocument();
                {
                    var workbookRelsEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
                    using (Stream workbookRelStream = workbookRelsEntry.Open())
                    {
                        workbookRelsXml.Load(workbookRelStream);
                    }
                }
                XmlNamespaceManager nsMRels = new XmlNamespaceManager(workbookRelsXml.NameTable);
                nsMRels.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships");
                //读取xl/workbook.xml
                var workbookXml = new XmlDocument();
                {
                    var workbookEntry = zip.GetEntry("xl/workbook.xml");
                    using (Stream workbookStream = workbookEntry.Open())
                    {
                        workbookXml.Load(workbookStream);
                    }
                }
                XmlNamespaceManager nsMWorkbook = new XmlNamespaceManager(workbookXml.NameTable);
                nsMWorkbook.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");


                //获取样式文件路径
                var styleNode = workbookRelsXml.SelectSingleNode($"/r:Relationships/r:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles']", nsMRels);
                var stypesFile = styleNode.Attributes["Target"].Value;
                if (!stypesFile.StartsWith("/"))    //相对路径,则补全
                    stypesFile = $"/xl/{stypesFile}";
                //增加样式，并修改样式索引
                ChangeStyles(zip, stypesFile, tempPath, newSheetColumnDic);


                //获取所有的sheet信息
                var sheets = workbookXml.SelectNodes("/x:workbook/x:sheets/x:sheet", nsMWorkbook);
                foreach (XmlNode sheetNode in sheets)
                {
                    var sheetName = sheetNode.Attributes["name"].Value;
                    var sheetRelId = sheetNode.Attributes["id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"].Value;

                    var relationshipNode = workbookRelsXml.SelectSingleNode($"/r:Relationships/r:Relationship[@Id='{sheetRelId}']", nsMRels);
                    var target = relationshipNode.Attributes["Target"].Value;
                    if (!target.StartsWith("/"))    //相对路径,则补全
                        target = $"/xl/{target}";

                    var newColumns = newSheetColumnDic[sheetName];

                    ChangeSingleSheetFile(useHeader, target, newColumns, reWriteValue, zip, tempPath, _xmlSettings);
                }

            }

        }

        private static void ChangeSingleSheetFile(bool useHeader, string sheetFile, Dictionary<int, NewColumnInfo> newColumns, Func<string, NewColumnInfo, bool, string> reWriteValue, ZipArchive zip, string tempPath, XmlReaderSettings _xmlSettings)
        {
            var tmpSheetXml = Path.Combine(tempPath, Path.GetRandomFileName());
            var sheetEntry = zip.GetEntry(sheetFile.TrimStart('/'));
            using (var sheetStream = sheetEntry.Open())
            {
                using (XmlReader reader = XmlReader.Create(sheetStream, _xmlSettings))
                {
                    using (var writer = XmlWriter.Create(tmpSheetXml, new XmlWriterSettings() { Indent = false }))   //设置缩进，测试完就不设置了
                    {
                        bool canWriteData = false;
                        NewColumnInfo newColumnInfo = null;
                        bool writeNewValue = false;
                        bool writeValue = false;

                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.XmlDeclaration)
                            {
                                //XML声明
                                writer.WriteStartDocument();
                                continue;
                            }
                            else if (reader.NodeType == XmlNodeType.Element)
                            {
                                //XML元素
                                var isEmptyElement = reader.IsEmptyElement; //无子元素
                                var elementName = reader.LocalName;

                                if (isEmptyElement)
                                {
                                    //写入元素起始
                                    writer.WriteStartElement(reader.Prefix, reader.LocalName, reader.NamespaceURI);
                                    //写入Attribute
                                    while (reader.MoveToNextAttribute())
                                        writer.WriteAttributeString(reader.Prefix, reader.LocalName, reader.NamespaceURI, reader.Value);
                                    //空元素写入结束符
                                    writer.WriteEndElement();
                                    writer.Flush(); //缓冲区的数据写入流
                                }
                                else    //包含子元素
                                {
                                    writeValue = canWriteData && elementName == "v";
                                    writeNewValue = writeValue && newColumnInfo != null;    //是否需要写入新的值

                                    writer.WriteStartElement(reader.Prefix, elementName, reader.NamespaceURI);  //写入元素起始

                                    //写入Attribute
                                    while (reader.MoveToNextAttribute())
                                    {
                                        var attrName = reader.LocalName;
                                        var attrValue = reader.Value;
                                        var attrPrefix = reader.Prefix;
                                        var attrNamespaceURI = reader.NamespaceURI;

                                        if (!canWriteData && elementName == "row" && attrName == "r")   //未开始读取数据时，读取到row元素的r(行号)属性
                                        {
                                            if (!useHeader || useHeader && int.Parse(attrValue) > 1)   //没有表头时，直接开始读取数据；有表头时，读取到第二行开始读取数据
                                                canWriteData = true;
                                        }
                                        else if (canWriteData && elementName == "c")
                                        {
                                            if (attrName == "r")
                                            {
                                                //设置当前单元格索引
                                                ReferenceHelper.ParseReference(attrValue, out var columnIdx, out _);    //columnIdx从1开始计算索引
                                                                                                                        //获取转换的newColumnInfo
                                                newColumns.TryGetValue(columnIdx, out newColumnInfo);
                                            }
                                            else if (attrName == "t" && (newColumnInfo?.TargetDataType == CellDataType.DateTime || newColumnInfo?.TargetDataType == CellDataType.Number))
                                                attrValue = "n"; //newColumnInfo不为null(需要转换),写入新的数据类型
                                            else if (attrName == "s" && (newColumnInfo?.TargetDataType == CellDataType.DateTime || newColumnInfo?.TargetDataType == CellDataType.Number))
                                                attrValue = newColumnInfo.NewStyleIndex.ToString(); //newColumnInfo不为null(需要转换),写入新的样式索引
                                        }
                                        if (!(attrName == "t" && attrValue == "n")) //t的默认值是n，可以不写入，减小文件体积
                                        {
                                            writer.WriteAttributeString(attrPrefix, attrName, attrNamespaceURI, attrValue);
                                        }

                                    }
                                    //if(!writeTAttribute)
                                    //    writer.WriteAttributeString("t","");
                                }
                            }
                            else if (reader.NodeType == XmlNodeType.Text)
                            {
                                var newValue = reader.Value;
                                if (writeValue)
                                {
                                    if (reWriteValue != null)
                                    {
                                        newValue = reWriteValue(newValue, newColumnInfo, writeNewValue);
                                    }
                                    else if (writeNewValue)
                                    {
                                        if (newColumnInfo.TargetDataType == CellDataType.DateTime && DateTime.TryParse(newValue, out var newTime))
                                            newValue = newTime.ToOADate().ToString();
                                        else if (newColumnInfo.TargetDataType == CellDataType.Number && decimal.TryParse(newValue, out var newNumber))
                                            newValue = newNumber.ToString();
                                    }
                                }
                                writer.WriteString(newValue);
                            }
                            else if (reader.NodeType == XmlNodeType.EndElement)
                            {
                                writer.WriteEndElement();
                                writer.Flush(); //缓冲区的数据写入流
                                if (canWriteData)
                                {
                                    //c单元格结束时，重置newColumnInfo
                                    if (reader.LocalName == "c")
                                        newColumnInfo = null;
                                    //v结束时，重置isValueElement
                                    else if (reader.LocalName == "v")
                                        writeNewValue = writeValue = false;
                                }
                            }
                            else
                            {
                                Console.WriteLine("未处理的NodeType" + reader.NodeType);
                            }
                        }

                        writer.WriteEndDocument();
                    }
                }
            }
            sheetEntry.Delete();   //删除原始的sheet1.xml
            zip.CreateEntryFromFile(tmpSheetXml, sheetFile.TrimStart('/'));  //添加新的sheet1.xml
            File.Delete(tmpSheetXml);
        }

        /// <summary>
        /// 把新的样式追加到原来的style.xml文件中
        /// </summary>
        /// <param name="zip"></param>
        /// <param name="stypesFile"></param>
        /// <param name="tempPath"></param>
        /// <param name="newSheetColumnDic"></param>
        private static void ChangeStyles(ZipArchive zip, string stypesFile, string tempPath, Dictionary<string, Dictionary<int, NewColumnInfo>> newSheetColumnDic)
        {
            // styles.xml文件中的numFmtId表示内置的数字格式，其ID范围通常在0 - 164。

            var newColumns = newSheetColumnDic.SelectMany(x => x.Value);
            var formats = newColumns.Select(x => x.Value.FormatStr).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList();

            if (formats.Count == 0) //没有需要转换的数据类型
                return;
            var tmpStylesXml = Path.Combine(tempPath, Path.GetRandomFileName());
            var stylesEntry = zip.GetEntry(stypesFile.TrimStart('/'));
            XmlDocument doc = new XmlDocument();
            using (var stylesStream = stylesEntry.Open())
            {
                // 加载XML文档  
                doc.Load(stylesStream);
            }

            var xNamespaceURI = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("x", xNamespaceURI);

            var dicStyleIndex = new Dictionary<string, int>();

            var numFmtIdIdx = 180;  //从181开始
            XmlNode numFmtsNode = doc.DocumentElement.SelectSingleNode("/x:styleSheet/x:numFmts", nsmgr);
            foreach (var item in newColumns)
            {
                if (string.IsNullOrEmpty(item.Value.FormatStr))
                    continue;
                if (!dicStyleIndex.TryGetValue(item.Value.FormatStr, out int styleIndex))
                {
                    numFmtIdIdx++;
                    // 1. 增加numFmt
                    XmlElement newNumFmt = doc.CreateElement("x:numFmt", xNamespaceURI);
                    newNumFmt.SetAttribute("numFmtId", numFmtIdIdx.ToString());
                    newNumFmt.SetAttribute("formatCode", item.Value.FormatStr);
                    numFmtsNode.AppendChild(newNumFmt);
                    // 更新numFmts的count
                    numFmtsNode.Attributes["count"].Value = numFmtsNode.ChildNodes.Count.ToString();

                    //从原始的efaultStylesXml中复制一个xf ,数组索引为2，时间索引为3

                    // 增加cellXfs
                    XmlNode cellXfsNode = doc.DocumentElement.SelectSingleNode("x:cellXfs", nsmgr);
                    XmlNode sourceXfNode = cellXfsNode.ChildNodes[item.Value.TargetDataType == CellDataType.DateTime ? 3 : 2];
                    XmlElement newXf = sourceXfNode.CloneNode(true) as XmlElement;
                    styleIndex = cellXfsNode.ChildNodes.Count;
                    newXf.SetAttribute("numFmtId", numFmtIdIdx.ToString());
                    cellXfsNode.AppendChild(newXf);
                    // 更新cellXfs的count
                    cellXfsNode.Attributes["count"].Value = (styleIndex + 1).ToString();

                    dicStyleIndex.Add(item.Value.FormatStr, styleIndex);

                }

                item.Value.NewStyleIndex = styleIndex;
            }

            var settings = new XmlWriterSettings { Indent = false }; // 禁止缩进
            using (XmlWriter writer = XmlWriter.Create(tmpStylesXml, settings))
            {
                //这样保存没缩进
                doc.Save(writer);
            }

            stylesEntry.Delete();   //删除原始的styles.xml  ,这里报错Cannot delete an entry currently open for writing
            zip.CreateEntryFromFile(tmpStylesXml, stypesFile.TrimStart('/'));  //添加新的styles.xml
            File.Delete(tmpStylesXml);

        }

    }
}
