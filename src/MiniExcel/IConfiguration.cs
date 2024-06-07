using MiniExcelLibs.Attributes;
using System.Globalization;

namespace MiniExcelLibs
{
    public interface IConfiguration { }
    public abstract class Configuration : IConfiguration
    {
        public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;
        public DynamicExcelColumn[] DynamicColumns { get; set; }
        public int BufferSize { get; set; } = 1024 * 512;
        public bool FastMode { get; set; } = false;

        /// <summary>
        /// 当列不在DynamicColumns时是否忽略列
        /// </summary>
        public bool IgnoreColumnIfNotInDynamicColumns { get; set; } = false;
    }
}
