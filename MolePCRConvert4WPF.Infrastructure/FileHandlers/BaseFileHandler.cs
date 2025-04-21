using System;
using System.IO;
using System.Threading.Tasks;
using MolePCRConvert4WPF.Core.Enums;
using MolePCRConvert4WPF.Core.Interfaces;
using MolePCRConvert4WPF.Core.Models;

namespace MolePCRConvert4WPF.Infrastructure.FileHandlers
{
    /// <summary>
    /// 基础文件处理器
    /// </summary>
    public abstract class BaseFileHandler : IFileHandler
    {
        /// <summary>
        /// 检测仪器类型
        /// </summary>
        public abstract Task<InstrumentType> DetectInstrumentTypeAsync(string filePath);
        
        /// <summary>
        /// 读取PCR数据
        /// </summary>
        public abstract Task<Plate> ReadPcrDataAsync(string filePath, InstrumentType instrumentType, Guid plateId);
        
        /// <summary>
        /// 校验文件格式
        /// </summary>
        public abstract Task<bool> ValidateFileFormatAsync(string filePath, InstrumentType instrumentType);
        
        /// <summary>
        /// 获取文件类型
        /// </summary>
        public virtual FileType GetFileType(string filePath)
        {
            string extension = Path.GetExtension(filePath)?.ToLowerInvariant() ?? string.Empty;
            
            return extension switch
            {
                ".xlsx" => FileType.Excel,
                ".csv" => FileType.Csv,
                ".xls" => FileType.ExcelLegacy,
                ".txt" => FileType.Text,
                _ => FileType.Unknown
            };
        }

        /// <summary>
        /// 创建新板
        /// </summary>
        protected virtual Plate CreateNewPlate(string name, InstrumentType instrumentType, string importFilePath, Guid plateId)
        {
            return new Plate
            {
                Id = plateId,
                Name = name,
                InstrumentType = instrumentType,
                CreatedAt = DateTime.UtcNow,
                ImportFilePath = importFilePath,
                // 根据不同仪器类型设置行列数 (可以调整为更通用的方式)
                Rows = instrumentType switch
                {
                    //InstrumentType.ABI7500 or InstrumentType.SLAN96S or InstrumentType.SLAN96P or 
                    //InstrumentType.ABIQ5 or InstrumentType.CFX96 or InstrumentType.CFX96DeepWell => 8,
                    //InstrumentType.MA600 => 16, // Example if MA600 has different layout
                    _ => 8 // Default to 8x12
                },
                Columns = instrumentType switch
                {
                   // InstrumentType.ABI7500 or InstrumentType.SLAN96S or InstrumentType.SLAN96P or 
                   // InstrumentType.ABIQ5 or InstrumentType.CFX96 or InstrumentType.CFX96DeepWell => 12,
                   // InstrumentType.MA600 => 24,
                    _ => 12
                }
            };
        }

        /// <summary>
        /// 创建样本
        /// </summary>
        protected virtual Sample CreateSample(Guid plateId, string name)
        {
            return new Sample
            {
                PlateId = plateId,
                Name = name,
                CreatedAt = DateTime.UtcNow
            };
        }

        /// <summary>
        /// 创建孔位
        /// </summary>
        protected virtual WellLayout CreateWell(Guid plateId, string row, int column, WellType wellType = WellType.Unknown)
        {
            return new WellLayout
            {
                PlateId = plateId,
                Row = row,
                Column = column,
                WellType = wellType
            };
        }
    }
} 