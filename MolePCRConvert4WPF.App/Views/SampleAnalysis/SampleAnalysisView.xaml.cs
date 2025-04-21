using System.Globalization;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using MolePCRConvert4WPF.App.ViewModels;
using MolePCRConvert4WPF.Core.Models;
using System;
using System.Linq;
using System.Windows;
using System.Collections.Generic; // For List
using System.Windows.Controls.Primitives; // For UniformGrid, ListViewItem

namespace MolePCRConvert4WPF.App.Views.SampleAnalysis
{
    /// <summary>
    /// SampleAnalysisView.xaml 的交互逻辑
    /// Contains interaction logic adapted from SampleNamingView for consistent plate view behavior.
    /// </summary>
    public partial class SampleAnalysisView : System.Windows.Controls.UserControl
    {
        // Fields for drag selection logic - REINSTATED
        private Point _startPoint;
        private bool _isDragging;
        private ListViewItem? _lastSelectedItemOnShiftClick = null; // Still useful for Shift+Click

        public SampleAnalysisView()
        {
            InitializeComponent();
             // Optional: Keep DataContextChanged/Loaded if needed for other init
             // this.DataContextChanged += (s, e) => { }; 
             this.Loaded += SampleAnalysisView_Loaded;
        }

         private void SampleAnalysisView_Loaded(object sender, RoutedEventArgs e)
         {
            if (WellsListView != null && WellsListView.Items.Count > 0)
             {
                 WellsListView.Focus();
             }
         }

        // --- Interaction Logic with Drag Selection --- 

        /// <summary>
        /// Keyboard interaction for ListView - Esc, Delete, Ctrl+A.
        /// </summary>
        private void WellsListView_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            var listView = sender as ListView;
            if (listView == null || listView.Items.Count == 0 || !(this.DataContext is SampleAnalysisViewModel vm))
            {
                return;
            }

            if (e.Key == Key.Escape)
            {
                 if (vm.ClearSelectionCommand.CanExecute(null))
                     vm.ClearSelectionCommand.Execute(null);
                e.Handled = true;
            }
            else if (e.Key == Key.Delete)
            {   
                if (vm.ClearPatientInfoCommand.CanExecute(null))
                    vm.ClearPatientInfoCommand.Execute(null);
                e.Handled = true;
            }
            else if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
            {
                if (e.Key == Key.A)
                {
                    if (vm.SelectAllCommand.CanExecute(null))
                        vm.SelectAllCommand.Execute(null);
                    e.Handled = true;
                }
            }
        }

        /// <summary>
        /// Mouse Left Button Down - Records start point and captures mouse for potential drag.
        /// </summary>
        private void WellsListView_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _startPoint = e.GetPosition(this); // Position relative to the UserControl for coordinate consistency
            _isDragging = false;

            var listView = sender as ListView;
            if (listView == null) { return; }

            // Capture the mouse to reliably receive MouseMove and MouseButtonUp events
            if (e.ButtonState == MouseButtonState.Pressed)
            { 
                 listView.CaptureMouse();
            }

            // Determine anchor for potential Shift+Click range selection
            var clickedItemContainer = GetListViewItemFromPoint(listView, e.GetPosition(listView));
            if (clickedItemContainer != null && !(Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift)))
            {
                _lastSelectedItemOnShiftClick = clickedItemContainer;
            }
        }

        /// <summary>
        /// Mouse Move - Handles drag selection logic.
        /// </summary>
        private void WellsListView_MouseMove(object sender, MouseEventArgs e)
        {
             var listView = sender as ListView;
             if (listView == null || !(this.DataContext is SampleAnalysisViewModel vm) || e.LeftButton != MouseButtonState.Pressed || !listView.IsMouseCaptured)
             {
                 return; 
             }

            Point currentPos = e.GetPosition(this); 

            if (!_isDragging &&
                (Math.Abs(currentPos.X - _startPoint.X) > SystemParameters.MinimumHorizontalDragDistance ||
                 Math.Abs(currentPos.Y - _startPoint.Y) > SystemParameters.MinimumVerticalDragDistance))
            {
                _isDragging = true;
                 bool isCtrlPressed = Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl);
                 if (!isCtrlPressed) 
                 { 
                     if (vm.ClearSelectionCommand.CanExecute(null))
                         vm.ClearSelectionCommand.Execute(null);
                 }
            }

            if (_isDragging)
            {
                Rect selectionRect = new Rect(_startPoint, currentPos);
                bool isCtrlPressed = Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl);

                foreach (var item in listView.Items)
                {
                    if (item is WellLayout well)
                    {
                        var container = listView.ItemContainerGenerator.ContainerFromItem(item) as ListViewItem;
                        if (container != null && container.IsVisible) 
                        {
                             Point itemOriginRelativeToUserControl = container.TransformToAncestor(this).Transform(new Point(0, 0));
                             Rect itemRect = new Rect(itemOriginRelativeToUserControl, new Size(container.ActualWidth, container.ActualHeight));
                            bool intersects = selectionRect.IntersectsWith(itemRect);
                            
                            if (intersects)
                            {
                                if (!well.IsSelected) well.IsSelected = true; 
                            }
                            else if (!isCtrlPressed)
                            {
                                if (well.IsSelected) well.IsSelected = false;
                            }
                        }
                    }
                }
                 e.Handled = true; 
            }
        }

        /// <summary>
        /// Mouse Left Button Up - Finalizes selection (distinguishes click from drag).
        /// </summary>
        private void WellsListView_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            var listView = sender as ListView;
            if (listView == null || !(this.DataContext is SampleAnalysisViewModel vm))
            { return; }

            if (listView.IsMouseCaptured)
            {
                listView.ReleaseMouseCapture();
            }

            if (_isDragging)
            {
                _isDragging = false;
                e.Handled = true; 
            }
            else 
            {
                 var clickedItemContainer = GetListViewItemFromPoint(listView, e.GetPosition(listView));
                 if (clickedItemContainer != null && clickedItemContainer.Content is WellLayout clickedWell)
                 {
                     bool isCtrlPressed = Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl);
                     bool isShiftPressed = Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift);

                     if (!isCtrlPressed && !isShiftPressed)
                     {
                         if (vm.ClearSelectionCommand.CanExecute(null))
                             vm.ClearSelectionCommand.Execute(null);
                         clickedWell.IsSelected = true;
                     }
                     else if (isCtrlPressed)
                     {
                         clickedWell.IsSelected = !clickedWell.IsSelected;
                         _lastSelectedItemOnShiftClick = clickedItemContainer; 
                     }
                     else if (isShiftPressed)
                     {
                          // Shift + Click: Range selection (Rectangular)
                          if (_lastSelectedItemOnShiftClick != null && _lastSelectedItemOnShiftClick.Content is WellLayout startWell && clickedItemContainer.Content is WellLayout endWell)
                          {
                              int startRow = GetRowIndex(startWell.Row);
                              int startCol = startWell.Column;
                              int endRow = GetRowIndex(endWell.Row);
                              int endCol = endWell.Column;

                              if (startRow != -1 && endRow != -1) // Check valid row indices
                              {
                                   if (!isCtrlPressed) // If only Shift is pressed, clear previous selection first
                                   {
                                       if (vm.ClearSelectionCommand.CanExecute(null))
                                           vm.ClearSelectionCommand.Execute(null);
                                   }

                                  int minRow = Math.Min(startRow, endRow);
                                  int maxRow = Math.Max(startRow, endRow);
                                  int minCol = Math.Min(startCol, endCol);
                                  int maxCol = Math.Max(startCol, endCol);

                                  // Iterate through all wells and select if within the rectangle
                                  foreach (var well in vm.WellLayouts)
                                  {
                                      int wellRow = GetRowIndex(well.Row);
                                      int wellCol = well.Column;
                                      if (wellRow >= minRow && wellRow <= maxRow && wellCol >= minCol && wellCol <= maxCol)
                                      {
                                          well.IsSelected = true;
                                      }
                                  }
                              }
                          }
                          else // No valid anchor or clicking the anchor itself, treat as single select (respecting Ctrl)
                          {
                               if (!isCtrlPressed)
                               { 
                                   if (vm.ClearSelectionCommand.CanExecute(null))
                                        vm.ClearSelectionCommand.Execute(null);
                               } 
                               clickedWell.IsSelected = true;
                              _lastSelectedItemOnShiftClick = clickedItemContainer; 
                          }
                     }
                     e.Handled = true; 
                 }
            }
        }
        
        /// <summary>
        /// Mouse Right Button Up - Show context menu, invoking ViewModel commands.
        /// </summary>
        private void WellsListView_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
             var listView = sender as ListView;
             if (listView == null || !(this.DataContext is SampleAnalysisViewModel vm)) { return; }

            var clickedItemContainer = GetListViewItemFromPoint(listView, e.GetPosition(listView));
             if (clickedItemContainer != null && clickedItemContainer.Content is WellLayout clickedWell)
             {
                 if (!clickedItemContainer.IsSelected)
                 {
                     if (!(Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)))
                     { 
                         listView.SelectedItems.Clear(); 
                     }
                     listView.SelectedItems.Add(clickedWell); 
                     clickedItemContainer.IsSelected = true; 
                     listView.UpdateLayout(); 
                 }
             }
             else if (listView.SelectedItems.Count == 0) 
             {
                 return; 
             }

             if (listView.SelectedItems.Count == 0) return;

            ContextMenu contextMenu = new ContextMenu();
            MenuItem setPatientInfoItem = new MenuItem { Header = "设置患者信息", Command = vm.ShowSetPatientInfoDialogCommand };
            setPatientInfoItem.IsEnabled = vm.ShowSetPatientInfoDialogCommand.CanExecute(null);
            contextMenu.Items.Add(setPatientInfoItem);
            MenuItem clearPatientInfoItem = new MenuItem { Header = "清除患者信息", Command = vm.ClearPatientInfoCommand };
            clearPatientInfoItem.IsEnabled = vm.ClearPatientInfoCommand.CanExecute(null);
            contextMenu.Items.Add(clearPatientInfoItem);
            contextMenu.Items.Add(new Separator());
            MenuItem selectAllItem = new MenuItem { Header = "全选", Command = vm.SelectAllCommand };
            selectAllItem.IsEnabled = vm.SelectAllCommand.CanExecute(null);
            contextMenu.Items.Add(selectAllItem);
            MenuItem clearSelectionItem = new MenuItem { Header = "取消选择", Command = vm.ClearSelectionCommand };
            clearSelectionItem.IsEnabled = listView.SelectedItems.Count > 0;
            contextMenu.Items.Add(clearSelectionItem);
            contextMenu.PlacementTarget = listView;
            contextMenu.IsOpen = true;
            e.Handled = true;
        }

        /// <summary>
        /// Handles Mouse Double Click on a ListView item to trigger editing.
        /// </summary>
        private void WellsListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
             var listView = sender as ListView;
             if (listView == null || !(this.DataContext is SampleAnalysisViewModel vm)) { return; }
             var clickedItemContainer = GetListViewItemFromPoint(listView, e.GetPosition(listView));
             if (clickedItemContainer != null && clickedItemContainer.Content is WellLayout clickedWell)
             { 
                 if (!clickedItemContainer.IsSelected || listView.SelectedItems.Count > 1) 
                 {
                     listView.SelectedItems.Clear();
                     listView.SelectedItems.Add(clickedWell);
                     clickedItemContainer.IsSelected = true;
                 }

                 if (vm.ShowSetPatientInfoDialogCommand.CanExecute(null))
                 {
                     vm.ShowSetPatientInfoDialogCommand.Execute(null);
                     e.Handled = true;
                 }
             }
        }

        /// <summary>
        /// Helper to get the ListViewItem at a specific point relative to the ListView.
        /// </summary>
        private ListViewItem? GetListViewItemFromPoint(ListView listView, Point point)
        {
            HitTestResult result = VisualTreeHelper.HitTest(listView, point);
            if (result == null) { return null; }
            DependencyObject? current = result.VisualHit;
            while (current != null)
            {
                if (current is ListViewItem item) return item;
                current = VisualTreeHelper.GetParent(current);
            }
            return null;
        }
        
        // Helper to convert Row label (string 'A', 'B') to index (0, 1...)
        private int GetRowIndex(string? rowLabel)
        {
            if (string.IsNullOrEmpty(rowLabel) || rowLabel.Length != 1) return -1;
            char label = char.ToUpper(rowLabel[0]);
            // Assuming standard A-H for 8 rows, but make it more general A-Z
            return (label >= 'A' && label <= 'Z') ? label - 'A' : -1;
        }
    }
} 