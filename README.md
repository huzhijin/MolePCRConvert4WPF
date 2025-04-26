<!--
 * @Author: 胡志锦 13089116+huzhijin1213@user.noreply.gitee.com
 * @Date: 2025-04-21 11:30:39
 * @LastEditors: 胡志锦 13089116+huzhijin1213@user.noreply.gitee.com
 * @LastEditTime: 2025-04-21 15:39:32
 * @FilePath: /MolePCRConvert4WPF/README.md
 * @Description: 这是默认设置,请设置`customMade`, 打开koroFileHeader查看配置 进行设置: https://github.com/OBKoro1/koro1FileHeader/wiki/%E9%85%8D%E7%BD%AE
-->
# MolePCRConvert4WPF

PCR结果分析与报告生成系统

## 导航问题修复

如果您遇到类似"导航处理对发生错误：未注册视图模型 XXXViewModel 的导航操作"的错误，这是由于导航服务未正确注册导航操作造成的。问题已在以下代码中修复：

1. 在MainWindow.xaml.cs中添加了各ViewModel的导航操作注册：
```csharp
// 注册各个ViewModel的导航操作
concreteNavService.RegisterNavigationAction<DataInputViewModel>(() => {
    // 使用MainContent的依赖注入创建视图
    var vm = concreteNavService.GetViewModel<DataInputViewModel>();
    var view = new Views.DataInput.DataInputView { DataContext = vm };
    MainContent.Content = view;
    Debug.WriteLine("导航到DataInputView完成");
});

// 其他ViewModel的导航操作注册...
```

2. 修正了ReportTemplateConfigView和ReportTemplateDesignerView的命名空间路径。

这些修复确保了应用程序可以正确导航到各个页面，不再显示导航错误弹窗。

## 其他问题

如果您在MacOS上构建此项目，您可能需要在项目文件中添加以下属性（已包含在MolePCRConvert4WPF.App.csproj中）：

```xml
<PropertyGroup>
    <EnableWindowsTargeting>true</EnableWindowsTargeting>
</PropertyGroup>
```

这允许在非Windows平台上构建面向Windows的应用程序。请注意，虽然可以在MacOS上构建，但WPF应用程序只能在Windows上运行。

## 报告生成功能

应用程序支持生成PCR分析报告，包括患者报告和整板报告。

### 模板目录

报告模板存放在应用程序根目录下的`Templates`文件夹中。系统会自动扫描该目录中的所有Excel模板文件(.xlsx)，并在报告生成时显示模板选择对话框。

### 使用方法

1. 确保在`Templates`目录中存在有效的Excel模板文件(.xlsx格式)
2. 在结果分析界面，点击"生成患者报告"或"生成整板报告"按钮
3. 在弹出的模板选择对话框中，选择要使用的报告模板
4. 选择保存位置
5. 系统会根据当前分析结果和所选模板生成报告

### 报告类型

- **患者报告**：按患者分组展示结果，适合生成给患者的个人报告
- **整板报告**：显示整个板的结果，适合实验室内部使用和整体分析

同一个模板文件可以用于生成患者报告或整板报告，系统会根据您选择的报告类型以不同方式填充数据。

### 数据导出

点击"导出结果"按钮可以将当前分析结果导出为Excel文件，无需使用模板。

### 自定义模板

您可以根据需要自定义报告模板。详细说明请参考`Templates/README.md`文件。 