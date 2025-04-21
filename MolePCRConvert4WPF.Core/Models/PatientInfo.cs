using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using CommunityToolkit.Mvvm.ComponentModel; // Base class for observable objects
using CommunityToolkit.Mvvm.ComponentModel; // Include this for ObservableValidator

namespace MolePCRConvert4WPF.Core.Models
{
    /// <summary>
    /// Represents basic patient information.
    /// </summary>
    // Inherit from ObservableValidator since we are using validation attributes ([StringLength]) with [ObservableProperty]
    public partial class PatientInfo : ObservableValidator
    {
        /// <summary>
        /// 患者ID
        /// </summary>
        [Key]
        public Guid Id { get; set; } = Guid.NewGuid();
        
        /// <summary>
        /// 患者姓名
        /// </summary>
        [StringLength(50)]
        [ObservableProperty]
        private string? _name;
        
        /// <summary>
        /// 病历号
        /// </summary>
        [StringLength(50)]
        [ObservableProperty]
        private string? _medicalRecordNumber;
        
        /// <summary>
        /// 所属区域
        /// </summary>
        [StringLength(50)]
        public string Region { get; set; } = string.Empty;
        
        /// <summary>
        /// 样本ID
        /// </summary>
        public Guid? SampleId { get; set; }
        
        /// <summary>
        /// 相关样本
        /// </summary>
        [ForeignKey("SampleId")]
        public virtual Sample? Sample { get; set; }
        
        /// <summary>
        /// 创建时间
        /// </summary>
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        
        /// <summary>
        /// 创建人ID
        /// </summary>
        public Guid? CreatedById { get; set; }
        
        /// <summary>
        /// 更新时间
        /// </summary>
        public DateTime? UpdatedAt { get; set; }
        
        /// <summary>
        /// 更新人ID
        /// </summary>
        public Guid? UpdatedById { get; set; }

        // Constructor
        public PatientInfo(string? name = null, string? medicalRecordNumber = null)
        {
            _name = name;
            _medicalRecordNumber = medicalRecordNumber;
            // Validate all properties on creation if needed
            // ValidateAllProperties(); 
        }

        // Override Equals and GetHashCode if needed for comparisons
        public override bool Equals(object? obj)
        {
            return obj is PatientInfo info &&
                   Name == info.Name &&
                   MedicalRecordNumber == info.MedicalRecordNumber;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Name, MedicalRecordNumber);
        }
    }
} 