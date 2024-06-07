using MiniExcelLibs.Attributes;
using System.Drawing;

namespace MiniExcelLibs.OpenXml
{
    public class OpenXmlConfiguration : Configuration
    {
        internal static readonly OpenXmlConfiguration DefaultConfig = new OpenXmlConfiguration();
        public bool FillMergedCells { get; set; }
        public TableStyles TableStyles { get; set; } = TableStyles.Default;
        public bool AutoFilter { get; set; } = true;
        public bool EnableConvertByteArray { get; set; } = true;
        public bool IgnoreTemplateParameterMissing { get; set; } = true;
        public bool EnableWriteNullValueCell { get; set; } = true;
        public bool EnableSharedStringCache { get; set; } = true;
        public long SharedStringCacheSize { get; set; } = 5 * 1024 * 1024;
        public DynamicExcelSheet[] DynamicSheets { get; set; }

        /// <summary>
        /// 头部背景色，默认#4472c4
        /// </summary>
        public Color HeadBackgroundColor { get; set; } = Color.FromArgb(68, 114, 196);

        /// <summary>
        /// 头部字体颜色，默认#ffffff
        /// </summary>
        public Color HeadFontColor{ get; set; } = Color.FromArgb(255, 255, 255);
    }
}