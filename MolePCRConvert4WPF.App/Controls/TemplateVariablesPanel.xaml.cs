using MolePCRConvert4WPF.Core.Models;
using MolePCRConvert4WPF.App.Commands;
using MolePCRConvert4WPF.App.Models;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;

namespace MolePCRConvert4WPF.App.Controls
{
    /// <summary>
    /// TemplateVariablesPanel.xaml 的交互逻辑
    /// </summary>
    public partial class TemplateVariablesPanel : UserControl
    {
        public static readonly DependencyProperty VariablesByCategoryProperty = 
            DependencyProperty.Register(
                nameof(VariablesByCategory), 
                typeof(ObservableCollection<VariableCategoryGroup>), 
                typeof(TemplateVariablesPanel), 
                new PropertyMetadata(null, OnVariablesByCategoryChanged));

        public static readonly DependencyProperty InsertVariableCommandProperty =
            DependencyProperty.Register(
                nameof(InsertVariableCommand),
                typeof(Commands.RelayCommand<MolePCRConvert4WPF.Core.Models.TemplateVariable>),
                typeof(TemplateVariablesPanel),
                new PropertyMetadata(null));

        public ObservableCollection<VariableCategoryGroup> VariablesByCategory
        {
            get => (ObservableCollection<VariableCategoryGroup>)GetValue(VariablesByCategoryProperty);
            set => SetValue(VariablesByCategoryProperty, value);
        }

        public Commands.RelayCommand<MolePCRConvert4WPF.Core.Models.TemplateVariable> InsertVariableCommand
        {
            get => (Commands.RelayCommand<MolePCRConvert4WPF.Core.Models.TemplateVariable>)GetValue(InsertVariableCommandProperty);
            set => SetValue(InsertVariableCommandProperty, value);
        }

        public TemplateVariablesPanel()
        {
            InitializeComponent();
            VariablesTreeView.DataContext = this;
        }

        private static void OnVariablesByCategoryChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is TemplateVariablesPanel panel)
            {
                panel.VariablesTreeView.ItemsSource = panel.VariablesByCategory;
            }
        }
    }
} 