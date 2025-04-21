using System.Threading.Tasks;
using MolePCRConvert4WPF.Core.Enums;
using MolePCRConvert4WPF.Core.Models;

namespace MolePCRConvert4WPF.Core.Interfaces
{
    /// <summary>
    /// 文件处理接口
    /// </summary>
    public interface IFileHandler
    {
        /// <summary>
        /// 检测仪器类型
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>仪器类型</returns>
        Task<InstrumentType> DetectInstrumentTypeAsync(string filePath);
        
        /// <summary>
        /// 读取PCR数据
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="instrumentType">仪器类型</param>
        /// <param name="plateId">板ID</param>
        /// <returns>板数据</returns>
        Task<Plate> ReadPcrDataAsync(string filePath, InstrumentType instrumentType, Guid plateId);
        
        /// <summary>
        /// 获取文件类型
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>文件类型</returns>
        FileType GetFileType(string filePath);
        
        /// <summary>
        /// 校验文件格式
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="instrumentType">仪器类型</param>
        /// <returns>是否有效</returns>
        Task<bool> ValidateFileFormatAsync(string filePath, InstrumentType instrumentType);
    }
} 