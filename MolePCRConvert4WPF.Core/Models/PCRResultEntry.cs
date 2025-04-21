using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using MolePCRConvert4WPF.Core.Enums;
//using MolePCRConvert4WPF.Core.Services; // Assuming LogService might be moved or replaced

namespace MolePCRConvert4WPF.Core.Models
{
    /// <summary>
    /// PCR结果条目，用于存储PCR分析结果
    /// </summary>
    public class PCRResultEntry : INotifyPropertyChanged
    {
        private string? _patientName = null;
        private string? _sampleID = null;
        private string? _sampleName = null;
        private string? _wellPosition = null;
        private string? _channel = null;
        private double? _ctValue;
        private double _ctDisplayValue = double.NaN; // For binding
        private double? _concentration;
        private string? _targetName = null;
        private string? _judgementFormula = null;
        private string? _concentrationFormula = null;
        private string? _detectionResult = null;
        private bool _isFirstPatientRow;
        private string? _target = null;
        private string? _result = null;
        private string? _patientID = null;
        private string? _gender = null;
        private string? _department = null;
        private string? _doctor = null;
        private string? _clinicalInfo = null;
        private string? _sampleType = null;
        private string? _experimentID = null;
        private string? _fluorescence = null;
        private string? _well = null;
        private DateTime? _testTime;
        private DateTime? _collectionTime;
        private string? _sampleSource = null;
        private string? _remark = null;

        /// <summary>
        /// 患者姓名
        /// </summary>
        public string? PatientName
        {
            get => _patientName;
            set => SetProperty(ref _patientName, value);
        }

        /// <summary>
        /// 样本ID
        /// </summary>
        public string? SampleID
        {
            get => _sampleID;
            set => SetProperty(ref _sampleID, value);
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
        /// 孔位
        /// </summary>
        public string? WellPosition
        {
            get => _wellPosition;
            set => SetProperty(ref _wellPosition, value);
        }

        /// <summary>
        /// 通道
        /// </summary>
        public string? Channel
        {
            get => _channel;
            set => SetProperty(ref _channel, value);
        }

        /// <summary>
        /// 原始CT值
        /// </summary>
        public double? CtValue
        {
            get => _ctValue;
            set
            {
                if (SetProperty(ref _ctValue, value))
                {
                    // Update the display value whenever the raw CtValue changes
                    UpdateCtDisplayValue();
                }
            }
        }

        /// <summary>
        /// 用于显示的CT值 (NaN代表N/A或无效)
        /// </summary>
        public double CtDisplayValue
        {
            get => _ctDisplayValue;
            private set => SetProperty(ref _ctDisplayValue, value);
        }

        /// <summary>
        /// 更新用于显示的CT值
        /// </summary>
        private void UpdateCtDisplayValue()
        {
            if (!_ctValue.HasValue || _ctValue.Value > 40) // Threshold for display as N/A
            {
                CtDisplayValue = double.NaN;
            }
            else
            {
                CtDisplayValue = _ctValue.Value;
            }
        }
        
        /// <summary>
        /// 浓度值
        /// </summary>
        public double? Concentration
        {
            get => _concentration;
            set => SetProperty(ref _concentration, value);
        }

        /// <summary>
        /// 目标名称（微生物名称）
        /// </summary>
        public string? TargetName
        {
            get => _targetName;
            set
            {
                if (SetProperty(ref _targetName, value))
                {
                    Target = value; // Also update the simplified Target property
                }
            }
        }

        /// <summary>
        /// 检测目标（简化版本，用于报告）
        /// </summary>
        public string? Target
        {
            get => _target;
            set => SetProperty(ref _target, value);
        }

        /// <summary>
        /// 判定公式
        /// </summary>
        public string? JudgementFormula
        {
            get => _judgementFormula;
            set => SetProperty(ref _judgementFormula, value);
        }

        /// <summary>
        /// 浓度计算公式
        /// </summary>
        public string? ConcentrationFormula
        {
            get => _concentrationFormula;
            set => SetProperty(ref _concentrationFormula, value);
        }

        /// <summary>
        /// 检测结果（阳性/阴性）
        /// </summary>
        public string? DetectionResult
        {
            get => _detectionResult;
            set
            {
                if (SetProperty(ref _detectionResult, value))
                {
                    Result = value; // Also update the simplified Result property
                }
            }
        }

        /// <summary>
        /// 检测结果（简化版本，用于报告）
        /// </summary>
        public string? Result
        {
            get => _result;
            set => SetProperty(ref _result, value);
        }

        /// <summary>
        /// 是否为患者分组的第一行（用于报告合并单元格）
        /// </summary>
        public bool IsFirstPatientRow
        {
            get => _isFirstPatientRow;
            set => SetProperty(ref _isFirstPatientRow, value);
        }
        
        /// <summary>
        /// 患者ID
        /// </summary>
        public string? PatientID
        {
            get => _patientID;
            set => SetProperty(ref _patientID, value);
        }

        /// <summary>
        /// 性别
        /// </summary>
        public string? Gender
        {
            get => _gender;
            set => SetProperty(ref _gender, value);
        }

        /// <summary>
        /// 科室
        /// </summary>
        public string? Department
        {
            get => _department;
            set => SetProperty(ref _department, value);
        }

        /// <summary>
        /// 医生
        /// </summary>
        public string? Doctor
        {
            get => _doctor;
            set => SetProperty(ref _doctor, value);
        }

        /// <summary>
        /// 临床信息
        /// </summary>
        public string? ClinicalInfo
        {
            get => _clinicalInfo;
            set => SetProperty(ref _clinicalInfo, value);
        }

        /// <summary>
        /// 样本类型
        /// </summary>
        public string? SampleType
        {
            get => _sampleType;
            set => SetProperty(ref _sampleType, value);
        }

        /// <summary>
        /// 实验ID
        /// </summary>
        public string? ExperimentID
        {
            get => _experimentID;
            set => SetProperty(ref _experimentID, value);
        }

        /// <summary>
        /// 荧光值
        /// </summary>
        public string? Fluorescence
        {
            get => _fluorescence;
            set => SetProperty(ref _fluorescence, value);
        }

        /// <summary>
        /// 孔号（简化）
        /// </summary>
        public string? Well
        {
            get => _well;
            set => SetProperty(ref _well, value);
        }

        /// <summary>
        /// 检测时间
        /// </summary>
        public DateTime? TestTime
        {
            get => _testTime;
            set => SetProperty(ref _testTime, value);
        }

        /// <summary>
        /// 采集时间
        /// </summary>
        public DateTime? CollectionTime
        {
            get => _collectionTime;
            set => SetProperty(ref _collectionTime, value);
        }

        /// <summary>
        /// 样本来源
        /// </summary>
        public string? SampleSource
        {
            get => _sampleSource;
            set => SetProperty(ref _sampleSource, value);
        }

        /// <summary>
        /// 备注
        /// </summary>
        public string? Remark
        {
            get => _remark;
            set => SetProperty(ref _remark, value);
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }
    }
} 