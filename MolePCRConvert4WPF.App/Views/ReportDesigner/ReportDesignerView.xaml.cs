using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using MolePCRConvert4WPF.App.ViewModels;
using MolePCRConvert4WPF.Core.Models;
using unvell.ReoGrid;
using unvell.ReoGrid.Events;
using unvell.ReoGrid.Graphics;
using System.IO;
using MolePCRConvert4WPF.App.Utils;
using System.Drawing;
using Microsoft.Win32;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using DrawingRectangle = System.Drawing.Rectangle;
using DrawingPoint = System.Drawing.Point;
using WpfRectangle = System.Windows.Shapes.Rectangle;
using WpfPoint = System.Windows.Point;

namespace MolePCRConvert4WPF.Views.ReportDesigner
{
    /// <summary>
    /// ReportDesignerView.xaml 的交互逻辑
    /// </summary>
    public partial class ReportDesignerView : Window
    {
        private ReportDesignerViewModel ViewModel => DataContext as ReportDesignerViewModel;
        private string _currentCellPos = "A1"; // 当前选中单元格位置
        private bool _isUpdatingSelection = false;

        public ReportDesignerView()
        {
            try
            {
                InitializeComponent();
                
                // 手动创建并设置ViewModel实例，防止XAML和代码中的循环引用
                this.DataContext = new MolePCRConvert4WPF.App.ViewModels.ReportDesignerViewModel();
                
                // 延迟初始化，等待数据上下文加载完成
                Loaded += ReportDesignerView_Loaded;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化设计器视图时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ReportDesignerView_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // 设置ViewModel的View引用
                if (ViewModel != null)
                {
                    ViewModel.View = this;
                }
                
                // 在组件完全加载后初始化
                InitializeReoGrid();
                HookEvents();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载设计器视图时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void HookEvents()
        {
            try
            {
                // 绑定ReoGrid事件 - 添加空检查和标志位防止递归
                if (ReportGrid?.CurrentWorksheet != null)
                {
                    ReportGrid.CurrentWorksheet.SelectionRangeChanged += Worksheet_SelectionRangeChanged;
                    
                    // 监听单元格变化事件
                    ReportGrid.CurrentWorksheet.CellDataChanged += Worksheet_CellDataChanged;
                    // 监听样式变化
                    ReportGrid.CurrentWorksheet.RangeStyleChanged += Worksheet_RangeStyleChanged;
                    
                    // 监听鼠标双击事件，用于编辑图片单元格（由于没有专门的双击事件，使用MouseDown事件模拟）
                    ReportGrid.CurrentWorksheet.CellMouseDown += Worksheet_CellMouseDown;
                }
                
                // 添加快捷键支持
                this.KeyDown += ReportDesignerView_KeyDown;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化事件时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void ReportDesignerView_KeyDown(object sender, KeyEventArgs e)
        {
            // Ctrl+S：保存
            if (e.Key == Key.S && Keyboard.Modifiers == ModifierKeys.Control)
            {
                if (ViewModel.SaveTemplateCommand.CanExecute(null))
                {
                    ViewModel.SaveTemplateCommand.Execute(null);
                    e.Handled = true;
                }
            }
        }

        private void Worksheet_SelectionRangeChanged(object sender, RangeEventArgs e)
        {
            if (_isUpdatingSelection) return; // 防止递归调用
            
            try
            {
                _isUpdatingSelection = true;
                
                if (e.Range != null && !e.Range.IsEmpty)
                {
                    // 更新当前选中的单元格位置
                    _currentCellPos = $"{(char)('A' + e.Range.Col)}{e.Range.Row + 1}";
                }
            }
            finally
            {
                _isUpdatingSelection = false;
            }
        }

        /// <summary>
        /// 监听单元格数据变化
        /// </summary>
        private void Worksheet_CellDataChanged(object sender, CellEventArgs e)
        {
            if (ViewModel != null)
            {
                ViewModel.MarkAsModified();
            }
        }

        /// <summary>
        /// 监听区域样式变化
        /// </summary>
        private void Worksheet_RangeStyleChanged(object sender, RangeEventArgs e)
        {
            if (ViewModel != null)
            {
                ViewModel.MarkAsModified();
            }
        }

        // 跟踪鼠标点击，用于检测双击事件
        private DateTime _lastClickTime = DateTime.MinValue;
        private CellPosition _lastClickPos = new CellPosition(-1, -1);
        private const double DoubleClickThreshold = 500; // 毫秒

        private void Worksheet_CellMouseDown(object sender, CellMouseEventArgs e)
        {
            try
            {
                var now = DateTime.Now;
                
                // 检查是否是双击（同一单元格的两次点击间隔小于阈值）
                if ((now - _lastClickTime).TotalMilliseconds < DoubleClickThreshold &&
                    e.CellPosition.Row == _lastClickPos.Row && 
                    e.CellPosition.Col == _lastClickPos.Col)
                {
                    // 模拟双击事件
                    HandleCellDoubleClick(e);
                }
                
                // 更新最后点击信息
                _lastClickTime = now;
                _lastClickPos = e.CellPosition;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"处理单元格鼠标按下事件时出错: {ex.Message}");
            }
        }

        private void HandleCellDoubleClick(CellMouseEventArgs e)
        {
            try
            {
                var worksheet = ReportGrid?.CurrentWorksheet;
                if (worksheet == null) return;
                
                // 检查是否是图片单元格
                var cell = worksheet.GetCell(e.CellPosition);
                if (cell?.Body is unvell.ReoGrid.CellTypes.ImageCell imageCell)
                {
                    // 双击图片单元格时，弹出图片选择对话框允许更换图片
                    OpenFileDialog openFileDialog = new OpenFileDialog
                    {
                        Filter = "图片文件(*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif|所有文件(*.*)|*.*",
                        Title = "更换图片"
                    };
                    
                    if (openFileDialog.ShowDialog() == true)
                    {
                        // 加载新图片
                        string imagePath = openFileDialog.FileName;
                        
                        // 使用内存流避免文件锁定
                        using (var stream = new MemoryStream(File.ReadAllBytes(imagePath)))
                        {
                            // 创建WPF的BitmapImage
                            var bitmapImage = new BitmapImage();
                            bitmapImage.BeginInit();
                            bitmapImage.StreamSource = stream;
                            bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                            bitmapImage.EndInit();
                            
                            // 更新图片，保持原有的显示模式
                            var viewMode = imageCell.ViewMode;
                            imageCell.Image = bitmapImage;
                            imageCell.ViewMode = viewMode;
                            
                            // 刷新显示
                            worksheet.RequestInvalidate();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"处理单元格双击事件时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void InitializeReoGrid()
        {
            try
            {
                var worksheet = ReportGrid?.CurrentWorksheet;
                if (worksheet == null) return;
                
                // 设置默认列数和行数
                const int defaultColumnCount = 10;
                const int defaultRowCount = 30;
                
                // 添加防护代码，避免重复设置触发事件
                bool originalSelectionChangedState = _isUpdatingSelection;
                _isUpdatingSelection = true;
                
                try
                {
                    // 避免直接使用ReoGridHelper中有问题的扩展方法
                    worksheet.SetRows(defaultRowCount);
                    worksheet.SetCols(defaultColumnCount);
                    
                    // 设置默认样式
                    worksheet.SetSettings(WorksheetSettings.View_ShowGridLine, true);
                    worksheet.SetSettings(WorksheetSettings.View_ShowHeaders, true);
                    
                    // 设置默认行高和列宽 - 使用ReoGrid原生方法
                    worksheet.SetRowsHeight(0, defaultRowCount, 22);
                    worksheet.SetColumnsWidth(0, defaultColumnCount, 80);
                    
                    // 设置标题 - 避免使用可能导致问题的索引器
                    try
                    {
                        var cell = worksheet.GetCell(0, 0); // A1
                        if (cell != null)
                        {
                            cell.Data = "报告模板";
                            
                            var style = new WorksheetRangeStyle
                            {
                                Flag = PlainStyleFlag.FontSize | PlainStyleFlag.FontName | PlainStyleFlag.FontStyleBold,
                                Bold = true,
                                FontSize = 14,
                                FontName = "Microsoft YaHei",
                            };
                            
                            // 创建范围对象时不使用字符串索引，改为直接使用坐标
                            worksheet.SetRangeStyles(new RangePosition(0, 0, 1, 2), style);
                        }
                    }
                    catch (Exception ex)
                    {
                        // 单独捕获设置标题时可能出现的异常
                        MessageBox.Show($"设置标题时出错：{ex.Message}", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    
                    // 初始化右键菜单
                    InitializeContextMenu();
                }
                finally
                {
                    // 恢复原始状态
                    _isUpdatingSelection = originalSelectionChangedState;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化ReoGrid时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 初始化单元格的右键菜单
        /// </summary>
        private void InitializeContextMenu()
        {
            // 创建右键菜单
            ContextMenu cellContextMenu = new ContextMenu();
            
            // 合并单元格
            MenuItem mergeMenuItem = new MenuItem();
            mergeMenuItem.Header = "合并单元格";
            mergeMenuItem.Click += MergeCells_Click;
            cellContextMenu.Items.Add(mergeMenuItem);
            
            // 取消合并单元格
            MenuItem unmergeMenuItem = new MenuItem();
            unmergeMenuItem.Header = "取消合并单元格";
            unmergeMenuItem.Click += UnmergeCells_Click;
            cellContextMenu.Items.Add(unmergeMenuItem);
            
            cellContextMenu.Items.Add(new Separator());
            
            // 设置边框
            MenuItem borderMenuItem = new MenuItem();
            borderMenuItem.Header = "设置边框";
            
            MenuItem allBordersItem = new MenuItem();
            allBordersItem.Header = "所有边框";
            allBordersItem.Click += AllBorders_Click;
            borderMenuItem.Items.Add(allBordersItem);
            
            MenuItem outerBordersItem = new MenuItem();
            outerBordersItem.Header = "外边框";
            outerBordersItem.Click += OuterBorders_Click;
            borderMenuItem.Items.Add(outerBordersItem);
            
            MenuItem noBordersItem = new MenuItem();
            noBordersItem.Header = "无边框";
            noBordersItem.Click += NoBorders_Click;
            borderMenuItem.Items.Add(noBordersItem);
            
            cellContextMenu.Items.Add(borderMenuItem);
            
            cellContextMenu.Items.Add(new Separator());
            
            // 对齐方式
            MenuItem alignmentMenuItem = new MenuItem();
            alignmentMenuItem.Header = "对齐方式";
            
            MenuItem alignLeftItem = new MenuItem();
            alignLeftItem.Header = "左对齐";
            alignLeftItem.Click += AlignLeft_Click;
            alignmentMenuItem.Items.Add(alignLeftItem);
            
            MenuItem alignCenterItem = new MenuItem();
            alignCenterItem.Header = "居中对齐";
            alignCenterItem.Click += AlignCenter_Click;
            alignmentMenuItem.Items.Add(alignCenterItem);
            
            MenuItem alignRightItem = new MenuItem();
            alignRightItem.Header = "右对齐";
            alignRightItem.Click += AlignRight_Click;
            alignmentMenuItem.Items.Add(alignRightItem);
            
            cellContextMenu.Items.Add(alignmentMenuItem);
            
            cellContextMenu.Items.Add(new Separator());
            
            // 插入图片
            MenuItem insertImageMenuItem = new MenuItem();
            insertImageMenuItem.Header = "插入图片";
            insertImageMenuItem.Click += InsertImage_Click;
            cellContextMenu.Items.Add(insertImageMenuItem);
            
            // 图片选项菜单（仅当选中图片单元格时显示）
            MenuItem imageOptionsMenuItem = new MenuItem();
            imageOptionsMenuItem.Header = "图片选项";
            
            MenuItem imageViewModeMenuItem = new MenuItem();
            imageViewModeMenuItem.Header = "显示模式";
            
            MenuItem stretchMenuItem = new MenuItem();
            stretchMenuItem.Header = "拉伸填充";
            stretchMenuItem.Click += (s, e) => SetImageViewMode(unvell.ReoGrid.CellTypes.ImageCellViewMode.Stretch);
            imageViewModeMenuItem.Items.Add(stretchMenuItem);
            
            MenuItem zoomMenuItem = new MenuItem();
            zoomMenuItem.Header = "等比缩放";
            zoomMenuItem.Click += (s, e) => SetImageViewMode(unvell.ReoGrid.CellTypes.ImageCellViewMode.Zoom);
            imageViewModeMenuItem.Items.Add(zoomMenuItem);
            
            MenuItem clipMenuItem = new MenuItem();
            clipMenuItem.Header = "原始尺寸";
            clipMenuItem.Click += (s, e) => SetImageViewMode(unvell.ReoGrid.CellTypes.ImageCellViewMode.Clip);
            imageViewModeMenuItem.Items.Add(clipMenuItem);
            
            imageOptionsMenuItem.Items.Add(imageViewModeMenuItem);
            
            // 替换图片选项
            MenuItem replaceImageMenuItem = new MenuItem();
            replaceImageMenuItem.Header = "替换图片";
            replaceImageMenuItem.Click += InsertImage_Click;  // 复用插入图片功能
            imageOptionsMenuItem.Items.Add(replaceImageMenuItem);
            
            cellContextMenu.Items.Add(imageOptionsMenuItem);
            
            // 设置单元格格式
            MenuItem formatCellMenuItem = new MenuItem();
            formatCellMenuItem.Header = "设置单元格格式";
            formatCellMenuItem.Click += FormatCell_Click;
            cellContextMenu.Items.Add(formatCellMenuItem);
            
            // 设置右键菜单打开前的事件处理
            cellContextMenu.Opened += (s, e) => 
            {
                var worksheet = ReportGrid?.CurrentWorksheet;
                if (worksheet == null) return;
                
                var pos = worksheet.SelectionRange;
                if (pos == null || pos.IsEmpty) return;
                
                // 检查当前选中单元格是否是图片单元格
                var cell = worksheet.GetCell(pos.Row, pos.Col);
                bool isImageCell = cell?.Body is unvell.ReoGrid.CellTypes.ImageCell;
                
                // 根据单元格类型控制菜单项的可见性
                imageOptionsMenuItem.Visibility = isImageCell ? Visibility.Visible : Visibility.Collapsed;
            };
            
            // 设置菜单到ReoGrid控件
            ReportGrid.ContextMenu = cellContextMenu;
        }

        /// <summary>
        /// 设置图片的显示模式
        /// </summary>
        private void SetImageViewMode(unvell.ReoGrid.CellTypes.ImageCellViewMode viewMode)
        {
            try
            {
                var worksheet = ReportGrid?.CurrentWorksheet;
                if (worksheet == null) return;
                
                var selection = worksheet.SelectionRange;
                if (selection == null || selection.IsEmpty) return;
                
                var cell = worksheet.GetCell(selection.Row, selection.Col);
                if (cell?.Body is unvell.ReoGrid.CellTypes.ImageCell imageCell)
                {
                    imageCell.ViewMode = viewMode;
                    worksheet.RequestInvalidate();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置图片显示模式时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #region 右键菜单事件处理
        private void MergeCells_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var worksheet = ReportGrid?.CurrentWorksheet;
                if (worksheet == null) return;
                
                var selection = worksheet.SelectionRange;
                if (selection != null && !selection.IsEmpty && (selection.Cols > 1 || selection.Rows > 1))
                {
                    // 使用ReoGrid的单元格合并操作
                    worksheet.MergeRange(selection);
                }
                else
                {
                    MessageBox.Show("请先选择多个单元格再进行合并操作。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"合并单元格时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UnmergeCells_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var worksheet = ReportGrid?.CurrentWorksheet;
                if (worksheet == null) return;
                
                var selection = worksheet.SelectionRange;
                if (selection != null && !selection.IsEmpty)
                {
                    // 使用ReoGrid的单元格拆分操作
                    worksheet.UnmergeRange(selection);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"取消合并单元格时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AllBorders_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var worksheet = ReportGrid?.CurrentWorksheet;
                if (worksheet == null) return;
                
                var selection = worksheet.SelectionRange;
                if (selection != null && !selection.IsEmpty)
                {
                    // 使用ReoGrid的边框设置方法
                    worksheet.SetRangeBorders(selection, BorderPositions.All, RangeBorderStyle.BlackSolid);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置边框时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OuterBorders_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var worksheet = ReportGrid?.CurrentWorksheet;
                if (worksheet == null) return;
                
                var selection = worksheet.SelectionRange;
                if (selection != null && !selection.IsEmpty)
                {
                    // 使用ReoGrid的边框设置方法
                    worksheet.SetRangeBorders(selection, BorderPositions.Outside, RangeBorderStyle.BlackSolid);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置边框时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void NoBorders_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var worksheet = ReportGrid?.CurrentWorksheet;
                if (worksheet == null) return;
                
                var selection = worksheet.SelectionRange;
                if (selection != null && !selection.IsEmpty)
                {
                    // 使用ReoGrid的边框设置方法
                    worksheet.SetRangeBorders(selection, BorderPositions.All, RangeBorderStyle.Empty);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置边框时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AlignLeft_Click(object sender, RoutedEventArgs e)
        {
            SetTextAlignment(ReoGridHorAlign.Left);
        }

        private void AlignCenter_Click(object sender, RoutedEventArgs e)
        {
            SetTextAlignment(ReoGridHorAlign.Center);
        }

        private void AlignRight_Click(object sender, RoutedEventArgs e)
        {
            SetTextAlignment(ReoGridHorAlign.Right);
        }

        private void SetTextAlignment(ReoGridHorAlign alignment)
        {
            try
            {
                var worksheet = ReportGrid?.CurrentWorksheet;
                if (worksheet == null) return;
                
                var selection = worksheet.SelectionRange;
                if (selection != null && !selection.IsEmpty)
                {
                    var style = new WorksheetRangeStyle
                    {
                        Flag = PlainStyleFlag.HorizontalAlign,
                        HAlign = alignment
                    };
                    
                    worksheet.SetRangeStyles(selection, style);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置对齐方式时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void InsertImage_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 打开文件选择对话框
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "图片文件(*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif|所有文件(*.*)|*.*",
                    Title = "选择要插入的图片"
                };
                
                if (openFileDialog.ShowDialog() == true)
                {
                    // 获取当前选中单元格
                    var worksheet = ReportGrid?.CurrentWorksheet;
                    if (worksheet == null) return;
                    
                    var selection = worksheet.SelectionRange;
                    if (selection != null && !selection.IsEmpty)
                    {
                        // 加载图片
                        string imagePath = openFileDialog.FileName;
                        
                        // 使用内存流避免文件锁定
                        using (var stream = new MemoryStream(File.ReadAllBytes(imagePath)))
                        {
                            // 创建WPF的BitmapImage
                            var bitmapImage = new BitmapImage();
                            bitmapImage.BeginInit();
                            bitmapImage.StreamSource = stream;
                            bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                            bitmapImage.EndInit();
                            
                            // 使用ReoGrid的图片功能
                            var imageCell = new unvell.ReoGrid.CellTypes.ImageCell(bitmapImage, unvell.ReoGrid.CellTypes.ImageCellViewMode.Zoom);
                            
                            // 合并几个单元格以容纳图片
                            if (selection.Rows == 1 && selection.Cols == 1)
                            {
                                // 如果只选择了一个单元格，则自动扩展合并区域
                                int rowCount = Math.Max(5, (int)(bitmapImage.Height / 50)); // 根据图片高度估算行数
                                int colCount = Math.Max(5, (int)(bitmapImage.Width / 80));  // 根据图片宽度估算列数
                                
                                // 检查合并范围是否超出网格边界
                                if (selection.Row + rowCount > worksheet.RowCount)
                                {
                                    worksheet.InsertRows(worksheet.RowCount, selection.Row + rowCount - worksheet.RowCount);
                                }
                                
                                if (selection.Col + colCount > worksheet.ColumnCount)
                                {
                                    worksheet.InsertColumns(worksheet.ColumnCount, selection.Col + colCount - worksheet.ColumnCount);
                                }
                                
                                // 创建合并范围
                                var mergeRange = new RangePosition(selection.Row, selection.Col, rowCount, colCount);
                                
                                // 先放置图片再合并，避免合并操作覆盖图片
                                worksheet.SetCellBody(selection.Row, selection.Col, imageCell);
                                
                                // 尝试合并单元格
                                try
                                {
                                    worksheet.MergeRange(mergeRange);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show($"合并单元格失败：{ex.Message}\n将仅在选定单元格插入图片。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                                }
                            }
                            else
                            {
                                // 如果已经选择了多个单元格，则使用选择的范围
                                worksheet.SetCellBody(selection.Row, selection.Col, imageCell);
                                
                                // 确保选择区域已合并
                                try
                                {
                                    if (!worksheet.IsMergedCell(selection.Row, selection.Col))
                                    {
                                        worksheet.MergeRange(selection);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show($"合并单元格失败：{ex.Message}", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                                }
                            }
                            
                            // 调整单元格的参数，允许内容填充
                            var style = new WorksheetRangeStyle
                            {
                                Flag = PlainStyleFlag.HorizontalAlign | PlainStyleFlag.VerticalAlign,
                                HAlign = ReoGridHorAlign.Center,
                                VAlign = ReoGridVerAlign.Middle
                            };
                            
                            worksheet.SetRangeStyles(selection.Row, selection.Col, 1, 1, style);
                        }
                    }
                    else
                    {
                        MessageBox.Show("请先选择至少一个单元格", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"插入图片时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void FormatCell_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var worksheet = ReportGrid?.CurrentWorksheet;
                if (worksheet == null) return;
                
                var selection = worksheet.SelectionRange;
                if (selection != null && !selection.IsEmpty)
                {
                    // TODO: 实现单元格格式设置对话框
                    MessageBox.Show("单元格格式设置功能正在开发中...", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置单元格格式时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion

        private void AddColumn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var worksheet = ReportGrid?.CurrentWorksheet;
                if (worksheet == null) return;
                
                int columnCount = worksheet.ColumnCount;
                
                // 在当前列数的基础上添加一列 - 使用原生方法
                worksheet.InsertColumns(columnCount, 1);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"添加列时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AddRow_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var worksheet = ReportGrid?.CurrentWorksheet;
                if (worksheet == null) return;
                
                int rowCount = worksheet.RowCount;
                
                // 在当前行数的基础上添加一行 - 使用原生方法
                worksheet.InsertRows(rowCount, 1);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"添加行时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 插入变量到当前选定的单元格
        /// </summary>
        /// <param name="variable">要插入的变量</param>
        public void InsertVariable(TemplateVariable variable)
        {
            try
            {
                if (variable == null) return;
                
                var worksheet = ReportGrid?.CurrentWorksheet;
                if (worksheet == null) return;
                
                // 获取当前选中的单元格位置
                var selection = worksheet.SelectionRange;
                if (selection != null && !selection.IsEmpty)
                {
                    // 插入变量占位符
                    string placeholderValue = $"${{{variable.Name}}}";
                    
                    // 将变量放入选中的单元格
                    var cell = worksheet.GetCell(selection.Row, selection.Col);
                    if (cell != null)
                    {
                        cell.Data = placeholderValue;
                        
                        // 设置字体颜色为蓝色以标识这是一个变量
                        var style = new WorksheetRangeStyle
                        {
                            Flag = PlainStyleFlag.TextColor,
                            TextColor = SolidColor.Blue
                        };
                        
                        worksheet.SetRangeStyles(selection, style);
                    }
                }
                else
                {
                    MessageBox.Show("请先选择一个单元格", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"插入变量时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 插入数据表格标记
        /// </summary>
        public void InsertDataTableMarker()
        {
            try
            {
                var worksheet = ReportGrid?.CurrentWorksheet;
                if (worksheet == null) return;
                
                // 获取当前选中的单元格位置
                var selection = worksheet.SelectionRange;
                if (selection != null && !selection.IsEmpty)
                {
                    // 插入数据表格标记
                    var cell = worksheet.GetCell(selection.Row, selection.Col);
                    if (cell != null)
                    {
                        cell.Data = "[[DataStart]]";
                        
                        // 设置背景色为浅绿色以标识这是数据表格起始位置
                        var style = new WorksheetRangeStyle
                        {
                            Flag = PlainStyleFlag.BackColor | PlainStyleFlag.TextColor | PlainStyleFlag.FontStyleBold,
                            BackColor = SolidColor.LightGreen,
                            TextColor = SolidColor.FromArgb((byte)0, (byte)100, (byte)0), // 使用深绿色的RGB值
                            Bold = true
                        };
                        
                        worksheet.SetRangeStyles(selection, style);
                    }
                }
                else
                {
                    MessageBox.Show("请先选择一个单元格", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"插入数据表格标记时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 收集模板数据
        /// </summary>
        /// <returns>单元格数据列表</returns>
        public List<TemplateCellData> CollectTemplateData()
        {
            var cells = new List<TemplateCellData>();
            
            try
            {
                var worksheet = ReportGrid?.CurrentWorksheet;
                if (worksheet == null) return cells;
                
                // 获取工作表的使用范围
                var usedRange = worksheet.UsedRange;
                if (usedRange == null || usedRange.IsEmpty) return cells;
                
                // 安全检查，确保范围值有效
                if (usedRange.Row < 0 || usedRange.Col < 0 || 
                    usedRange.Rows <= 0 || usedRange.Cols <= 0 ||
                    usedRange.Row + usedRange.Rows > worksheet.RowCount ||
                    usedRange.Col + usedRange.Cols > worksheet.ColumnCount)
                {
                    return cells;
                }
                
                // 遍历所有非空单元格
                for (int row = usedRange.Row; row < usedRange.Row + usedRange.Rows; row++)
                {
                    for (int col = usedRange.Col; col < usedRange.Col + usedRange.Cols; col++)
                    {
                        try
                        {
                            var cell = worksheet.GetCell(row, col);
                            if (cell != null && cell.Data != null)
                            {
                                var cellValue = cell.Data.ToString();
                                
                                if (!string.IsNullOrWhiteSpace(cellValue))
                                {
                                    // 获取单元格样式
                                    var style = cell.Style;
                                    string fontName = style?.FontName ?? "Arial";
                                    float fontSize = style?.FontSize ?? 10;
                                    string bgColor = style?.BackColor != null ? SolidColorToHex(style.BackColor) : "#FFFFFF";
                                    string textColor = style?.TextColor != null ? SolidColorToHex(style.TextColor) : "#000000";
                                    bool isBold = style?.Bold ?? false;
                                    bool isItalic = style?.Italic ?? false;
                                    
                                    // 创建单元格数据对象
                                    var cellData = new TemplateCellData
                                    {
                                        RowIndex = row,
                                        ColumnIndex = col,
                                        Content = cellValue,
                                        FontName = fontName,
                                        FontSize = fontSize,
                                        BackgroundColor = bgColor,
                                        ForegroundColor = textColor,
                                        IsBold = isBold,
                                        IsItalic = isItalic
                                    };
                                    
                                    cells.Add(cellData);
                                }
                            }
                        }
                        catch (Exception cellEx)
                        {
                            // 单个单元格处理错误不应该中断整个收集过程
                            System.Diagnostics.Debug.WriteLine($"处理单元格({row},{col})时出错: {cellEx.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"收集模板数据时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
            return cells;
        }

        /// <summary>
        /// 加载模板数据到设计器
        /// </summary>
        /// <param name="template">要加载的模板</param>
        public void LoadTemplateData(ReportCustomTemplate template)
        {
            try
            {
                if (template == null || template.Cells == null) return;
                
                var worksheet = ReportGrid?.CurrentWorksheet;
                if (worksheet == null) return;
                
                // 清空当前工作表
                worksheet.Reset();
                
                // 确保工作表有足够的行和列
                int maxRowIndex = 0;
                int maxColIndex = 0;
                
                foreach (var cell in template.Cells)
                {
                    maxRowIndex = Math.Max(maxRowIndex, cell.RowIndex + 1);
                    maxColIndex = Math.Max(maxColIndex, cell.ColumnIndex + 1);
                }
                
                // 确保至少有默认的行列数
                maxRowIndex = Math.Max(maxRowIndex, 30);
                maxColIndex = Math.Max(maxColIndex, 10);
                
                // 使用原生方法设置行列数
                worksheet.SetRows(maxRowIndex);
                worksheet.SetCols(maxColIndex);
                
                // 设置默认样式
                worksheet.SetSettings(WorksheetSettings.View_ShowGridLine, true);
                worksheet.SetSettings(WorksheetSettings.View_ShowHeaders, true);
                
                // 加载单元格数据
                foreach (var cellData in template.Cells)
                {
                    var cell = worksheet.GetCell(cellData.RowIndex, cellData.ColumnIndex);
                    if (cell != null)
                    {
                        cell.Data = cellData.Content;
                        
                        // 应用样式
                        var flag = PlainStyleFlag.FontName | PlainStyleFlag.FontSize | 
                                   PlainStyleFlag.TextColor | PlainStyleFlag.BackColor | 
                                   PlainStyleFlag.FontStyleBold | PlainStyleFlag.FontStyleItalic;
                        
                        var style = new WorksheetRangeStyle
                        {
                            Flag = flag,
                            FontName = cellData.FontName,
                            FontSize = cellData.FontSize,
                            TextColor = HexToSolidColor(cellData.ForegroundColor),
                            BackColor = HexToSolidColor(cellData.BackgroundColor),
                            Bold = cellData.IsBold,
                            Italic = cellData.IsItalic
                        };
                        
                        worksheet.SetRangeStyles(new RangePosition(cellData.RowIndex, cellData.ColumnIndex, 1, 1), style);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载模板数据时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 获取模板数据为ReoGrid格式
        /// </summary>
        public byte[] GetTemplateBytes()
        {
            try
            {
                using (var ms = new MemoryStream())
                {
                    if (ReportGrid != null)
                    {
                        ReportGrid.Save(ms, unvell.ReoGrid.IO.FileFormat.ReoGridFormat);
                        return ms.ToArray();
                    }
                    return new byte[0];
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"获取模板数据时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return new byte[0];
            }
        }

        #region 辅助方法

        /// <summary>
        /// 将SolidColor转换为十六进制字符串
        /// </summary>
        private string SolidColorToHex(SolidColor color)
        {
            return $"#{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}";
        }

        /// <summary>
        /// 将十六进制字符串转换为SolidColor
        /// </summary>
        private SolidColor HexToSolidColor(string hex)
        {
            if (string.IsNullOrEmpty(hex)) return SolidColor.Transparent;
            
            hex = hex.TrimStart('#');
            
            if (hex.Length == 8)
            {
                return SolidColor.FromArgb(
                    (byte)Convert.ToByte(hex.Substring(0, 2), 16),
                    (byte)Convert.ToByte(hex.Substring(2, 2), 16),
                    (byte)Convert.ToByte(hex.Substring(4, 2), 16),
                    (byte)Convert.ToByte(hex.Substring(6, 2), 16));
            }
            else if (hex.Length == 6)
            {
                return SolidColor.FromArgb(
                    (byte)255,
                    (byte)Convert.ToByte(hex.Substring(0, 2), 16),
                    (byte)Convert.ToByte(hex.Substring(2, 2), 16),
                    (byte)Convert.ToByte(hex.Substring(4, 2), 16));
            }
            
            return SolidColor.Black;
        }

        #endregion

        /// <summary>
        /// 工具栏按钮 - 添加边框
        /// </summary>
        private void AddBorders_Click(object sender, RoutedEventArgs e)
        {
            // 复用右键菜单中的添加边框功能
            AllBorders_Click(sender, e);
        }

        private void Variable_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                WpfRectangle rectangle = sender as WpfRectangle;
                if (rectangle != null && rectangle.Tag is TemplateVariable variable)
                {
                    // 开始拖放操作
                    DataObject data = new DataObject();
                    data.SetData("TemplateVariable", variable);
                    
                    // 启动拖放
                    DragDrop.DoDragDrop(rectangle, data, DragDropEffects.Copy);
                    
                    e.Handled = true;
                }
            }
        }

        private void ReportGrid_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent("TemplateVariable"))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            
            e.Handled = true;
        }

        private void ReportGrid_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent("TemplateVariable"))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            
            e.Handled = true;
        }

        private void ReportGrid_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent("TemplateVariable"))
            {
                var variable = e.Data.GetData("TemplateVariable") as TemplateVariable;
                if (variable != null)
                {
                    // 获取鼠标在ReoGrid中的位置
                    WpfPoint position = e.GetPosition(ReportGrid);
                    
                    try
                    {
                        var worksheet = ReportGrid?.CurrentWorksheet;
                        if (worksheet == null) return;
                        
                        // 尝试将变量插入到当前选中的单元格
                        var selection = worksheet.SelectionRange;
                        if (selection != null && !selection.IsEmpty)
                        {
                            // 使用当前选择范围的开始单元格
                            int row = selection.Row;
                            int col = selection.Col;
                            
                            // 将变量放入选中的单元格
                            var cell = worksheet.GetCell(row, col);
                            if (cell != null)
                            {
                                string placeholderValue = $"${{{variable.Name}}}";
                                cell.Data = placeholderValue;
                                
                                // 设置字体颜色为蓝色以标识这是一个变量
                                var style = new WorksheetRangeStyle
                                {
                                    Flag = PlainStyleFlag.TextColor,
                                    TextColor = SolidColor.Blue
                                };
                                
                                worksheet.SetRangeStyles(row, col, 1, 1, style);
                                
                                // 选中该单元格
                                worksheet.SelectRange(row, col, 1, 1);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"放置变量时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            
            e.Handled = true;
        }
    }
} 