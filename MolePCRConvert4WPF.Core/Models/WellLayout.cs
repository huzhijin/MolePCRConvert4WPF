using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using MolePCRConvert4WPF.Core.Enums;

namespace MolePCRConvert4WPF.Core.Models
{
    /// <summary>
    /// 孔位布局
    /// </summary>
    public class WellLayout : INotifyPropertyChanged
    {
        private int _row;
        private int _column;
        private WellType _wellType;
        private string? _sampleName = null;
        private bool _isSelected;
        private string? _patientName;
        private string? _patientCaseNumber;

        /// <summary>
        /// 孔位ID
        /// </summary>
        [Key]
        public Guid Id { get; set; } = Guid.NewGuid();
        
        /// <summary>
        /// 板ID
        /// </summary>
        [Required]
        public Guid PlateId { get; set; }
        
        /// <summary>
        /// 样本ID
        /// </summary>
        public Guid? SampleId { get; set; }
        
        /// <summary>
        /// 行索引（A, B, C...）
        /// </summary>
        [Required]
        [StringLength(2)]
        public string Row
        {
            get => ((char)('A' + _row)).ToString();
            set => _row = value[0] - 'A';
        }
        
        /// <summary>
        /// 列索引（1, 2, 3...）
        /// </summary>
        [Required]
        public int Column
        {
            get => _column;
            set => SetProperty(ref _column, value);
        }
        
        /// <summary>
        /// 孔位名称 (例如: A1, B2)
        /// </summary>
        [StringLength(10)]
        public string WellName => $"{(char)('A' + _row)}{_column}";
        
        /// <summary>
        /// 孔位类型
        /// </summary>
        public WellType WellType
        {
            get => _wellType;
            set => _wellType = value;
        }
        
        /// <summary>
        /// 样本名称
        /// </summary>
        public string? SampleName
        {
            get => _sampleName;
            set => SetProperty(ref _sampleName, value);
        }
        
        /// <summary>
        /// 患者姓名
        /// </summary>
        public string? PatientName
        {
            get => _patientName;
            set => SetProperty(ref _patientName, value);
        }
        
        /// <summary>
        /// 患者病历号
        /// </summary>
        public string? PatientCaseNumber
        {
            get => _patientCaseNumber;
            set => SetProperty(ref _patientCaseNumber, value);
        }
        
        /// <summary>
        /// 是否选中 (用于 UI 交互)
        /// </summary>
        [NotMapped] // Assuming IsSelected is UI state, not database field
        public bool IsSelected
        {
            get => _isSelected;
            // Ensure PropertyChanged is raised for IsSelected
            set => SetProperty(ref _isSelected, value); 
        }
        
        /// <summary>
        /// 位置标识（如A1, B2等）
        /// </summary>
        public string Position => $"{(char)('A' + _row)}{_column}";
        
        /// <summary>
        /// 孔位类型
        /// </summary>
        public WellType Type { get; set; } = WellType.Unknown;
        
        /// <summary>
        /// CT值
        /// </summary>
        public double? CtValue { get; set; }
        
        /// <summary>
        /// CT值特殊标记，存储"-"、"Undetermined"等特殊值
        /// </summary>
        public string? CtValueSpecialMark { get; set; }
        
        /// <summary>
        /// 荧光值
        /// </summary>
        public double? FluorescenceValue { get; set; }
        
        /// <summary>
        /// 标准曲线浓度（如果是标准品）
        /// </summary>
        public double? StandardConcentration { get; set; }
        
        /// <summary>
        /// 靶基因/探针名称
        /// </summary>
        [StringLength(50)]
        public string? TargetName { get; set; }
        
        /// <summary>
        /// 通道/颜色
        /// </summary>
        [StringLength(50)]
        public string? Channel { get; set; }
        
        /// <summary>
        /// 备注
        /// </summary>
        [StringLength(500)]
        public string? Notes { get; set; }
        
        /// <summary>
        /// 相关样本
        /// </summary>
        [ForeignKey("SampleId")]
        public virtual Sample? Sample { get; set; }

        /// <summary>
        /// 属性变更事件
        /// </summary>
        public event PropertyChangedEventHandler? PropertyChanged;

        /// <summary>
        /// 触发属性变更事件
        /// </summary>
        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        /// 设置属性值
        /// </summary>
        protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
        {
            if (Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }
    }
} 