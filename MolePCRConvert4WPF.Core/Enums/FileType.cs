namespace MolePCRConvert4WPF.Core.Enums
{
    /// <summary>
    /// 文件类型
    /// </summary>
    public enum FileType
    {
        /// <summary>
        /// 未知
        /// </summary>
        Unknown = 0,
        
        /// <summary>
        /// Excel文件(.xlsx)
        /// </summary>
        Excel = 1,
        
        /// <summary>
        /// CSV文件(.csv)
        /// </summary>
        Csv = 2,
        
        /// <summary>
        /// 旧Excel文件(.xls)
        /// </summary>
        ExcelLegacy = 3,
        
        /// <summary>
        /// 文本文件(.txt)
        /// </summary>
        Text = 4
    }
} 