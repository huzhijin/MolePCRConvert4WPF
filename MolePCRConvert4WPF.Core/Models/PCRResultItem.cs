using CommunityToolkit.Mvvm.ComponentModel;

namespace MolePCRConvert4WPF.Core.Models
{
    /// <summary>
    /// PCR分析结果项模型
    /// </summary>
    public partial class PCRResultItem : ObservableObject
    {
        /// <summary>
        /// 患者姓名
        /// </summary>
        [ObservableProperty]
        private string _patientName = "-";
        
        /// <summary>
        /// 病历号
        /// </summary>
        [ObservableProperty]
        private string _patientCaseNumber = "-";
        
        /// <summary>
        /// 孔位（如A1）
        /// </summary>
        [ObservableProperty]
        private string _wellPosition = string.Empty;
        
        /// <summary>
        /// 荧光通道（如FAM, VIC, ROX, CY5）
        /// </summary>
        [ObservableProperty]
        private string _channel = string.Empty;
        
        /// <summary>
        /// 靶标名称
        /// </summary>
        [ObservableProperty]
        private string _targetName = string.Empty;
        
        /// <summary>
        /// CT值
        /// </summary>
        [ObservableProperty]
        private double? _ctValue;
        
        /// <summary>
        /// CT值显示文本 (包括"-"等特殊值)
        /// </summary>
        [ObservableProperty]
        private string _ctValueDisplay = "-";
        
        /// <summary>
        /// 浓度值
        /// </summary>
        [ObservableProperty]
        private string _concentration = "-";
        
        /// <summary>
        /// 检测结果（阳性/阴性/-）
        /// </summary>
        [ObservableProperty]
        private string _result = "-";
        
        /// <summary>
        /// 是否为阳性结果（用于UI样式）
        /// </summary>
        [ObservableProperty]
        private bool _isPositive;
        
        /// <summary>
        /// 是否为患者的第一行（用于UI分组显示）
        /// </summary>
        [ObservableProperty]
        private bool _isFirstPatientRow;
    }
} 