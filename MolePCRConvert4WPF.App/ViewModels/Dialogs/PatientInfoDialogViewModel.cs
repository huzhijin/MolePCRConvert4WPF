using CommunityToolkit.Mvvm.ComponentModel;

namespace MolePCRConvert4WPF.App.ViewModels.Dialogs // Create a Dialogs subfolder if desired
{
    public partial class PatientInfoDialogViewModel : ObservableObject
    {
        [ObservableProperty]
        private string? _patientName;

        [ObservableProperty]
        private string? _medicalRecordNumber;

        // You can add validation logic here using attributes or IDataErrorInfo

        public PatientInfoDialogViewModel(string? initialName = null, string? initialMrn = null)
        {
            _patientName = initialName;
            _medicalRecordNumber = initialMrn;
        }
    }
} 