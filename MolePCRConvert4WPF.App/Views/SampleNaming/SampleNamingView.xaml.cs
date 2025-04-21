using MolePCRConvert4WPF.App.ViewModels;
using MolePCRConvert4WPF.Core.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using WinPoint = System.Windows.Point; // Alias for System.Windows.Point

namespace MolePCRConvert4WPF.App.Views.SampleNaming
{
    /// <summary>
    /// Interaction logic for SampleNamingView.xaml
    /// </summary>
    public partial class SampleNamingView : UserControl
    {
        private SampleNamingViewModel? ViewModel => DataContext as SampleNamingViewModel;
        private WinPoint _startPoint;
        private bool _isDragging;
        private ListViewItem? _lastSelectedItem; // Keep track for Shift-select
        private object? _dragSelectionAnchor; // Anchor for Shift+Arrow key selection

        public SampleNamingView()
        {
            InitializeComponent();
            // Ensure ListView is focusable to receive keyboard events
            // Loaded += (s, e) => WellsListView.Focus(); 
            // Consider setting FocusManager.FocusedElement in XAML instead
        }

        // --- Keyboard Navigation and Selection --- 
        private void WellsListView_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            var listView = sender as ListView;
            if (listView == null || listView.Items.Count == 0 || ViewModel == null) return;

            // Handle Esc: Clear selection
            if (e.Key == Key.Escape)
            {
                listView.SelectedItems.Clear();
                _dragSelectionAnchor = null; // Reset anchor
                e.Handled = true;
                return;
            }

            // Handle Ctrl+A: Select all
            if (e.Key == Key.A && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                listView.SelectAll();
                _dragSelectionAnchor = null; // Reset anchor after full selection
                e.Handled = true;
                return;
            }

            // Handle Arrow Keys for Navigation and Selection
            int currentIndex = listView.SelectedIndex;
            if (currentIndex < 0 && listView.SelectedItems.Count > 0)
            {
                // If multiple selected, maybe start from the first one?
                currentIndex = listView.Items.IndexOf(listView.SelectedItems[0]);
            }
            if (currentIndex < 0) currentIndex = 0; // Default to first item if nothing selected

            int rowCount = 8; // Assuming 8 rows
            int colCount = 12; // Assuming 12 columns
            int totalItems = rowCount * colCount;
            if (totalItems != listView.Items.Count) 
            {
                 // Mismatch, recalculate based on actual items if necessary or log warning
                 totalItems = listView.Items.Count;
                 // Simple division might be wrong if layout isn't perfect grid
                 colCount = 12; // Re-assert known column count
                 rowCount = (int)Math.Ceiling((double)totalItems / colCount);
            }

            int currentRow = currentIndex / colCount;
            int currentCol = currentIndex % colCount;
            int newIndex = currentIndex;

            switch (e.Key)
            {
                case Key.Up:
                    if (currentRow > 0) newIndex = currentIndex - colCount;
                    break;
                case Key.Down:
                    if (currentRow < rowCount - 1) newIndex = currentIndex + colCount;
                    break;
                case Key.Left:
                    if (currentCol > 0) newIndex = currentIndex - 1;
                    break;
                case Key.Right:
                    if (currentCol < colCount - 1) newIndex = currentIndex + 1;
                    break;
                default:
                    return; // Not an arrow key we handle here
            }

            // Ensure new index is within bounds
            newIndex = Math.Clamp(newIndex, 0, totalItems - 1);

            if (newIndex != currentIndex)
            {
                var newItem = listView.Items[newIndex];
                if (Keyboard.Modifiers == ModifierKeys.Shift)
                {
                    // Shift + Arrow: Range selection
                    if (_dragSelectionAnchor == null)
                    {
                        _dragSelectionAnchor = listView.SelectedItem ?? listView.Items[currentIndex];
                    }
                    int anchorIndex = listView.Items.IndexOf(_dragSelectionAnchor);
                    if (anchorIndex < 0) anchorIndex = currentIndex;
                    
                    // Clear previous selection if Ctrl is not held
                    if (Keyboard.Modifiers != (ModifierKeys.Shift | ModifierKeys.Control))
                    {
                         listView.SelectedItems.Clear();
                    }
                    
                    // Select items in the range
                    int start = Math.Min(anchorIndex, newIndex);
                    int end = Math.Max(anchorIndex, newIndex);
                    for (int i = start; i <= end; i++)
                    {
                         if (!listView.SelectedItems.Contains(listView.Items[i]))
                             listView.SelectedItems.Add(listView.Items[i]);
                    }
                }
                else if (Keyboard.Modifiers == ModifierKeys.Control)
                {
                    // Ctrl + Arrow: Move focus without changing selection (standard behavior handles this?)
                    // Or potentially add/remove from selection if needed? Let standard behavior handle focus for now.
                    listView.SelectedItem = newItem; // Move focus
                }
                else
                {
                    // Arrow only: Clear selection and select the new item
                    listView.SelectedItems.Clear();
                    listView.SelectedItem = newItem;
                    _dragSelectionAnchor = newItem; // Set new anchor for future Shift selection
                }

                // Scroll into view and focus
                var itemContainer = listView.ItemContainerGenerator.ContainerFromIndex(newIndex) as ListViewItem;
                itemContainer?.Focus();
                itemContainer?.BringIntoView();
                
                e.Handled = true;
            }
        }

        // --- Mouse Drag Selection --- 

        private void WellsListView_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _startPoint = e.GetPosition(null); // Position relative to screen or root element
            _isDragging = false; 

            var listView = sender as ListView;
            if (listView == null) return;

            // Find the item under the mouse
            var clickedItemContainer = GetListViewItemFromPoint(listView, e.GetPosition(listView));
             _lastSelectedItem = clickedItemContainer; // Track the clicked item for shift selection start
             if (clickedItemContainer == null) 
             {
                 // Clicked outside an item, potentially clear selection if not Ctrl/Shift clicking
                 if (Keyboard.Modifiers == ModifierKeys.None)
                 {
                     listView.SelectedItems.Clear();
                     _dragSelectionAnchor = null;
                 }
                 return; // Don't start drag if not clicking on an item
             }

             var clickedItem = clickedItemContainer.Content;

             if (Keyboard.Modifiers == ModifierKeys.Shift)
             {
                 // Shift+Click: Extend selection from anchor
                  if (_dragSelectionAnchor == null) _dragSelectionAnchor = clickedItem;
                  int anchorIndex = listView.Items.IndexOf(_dragSelectionAnchor);
                  int clickedIndex = listView.Items.IndexOf(clickedItem);
                  if(anchorIndex < 0 || clickedIndex < 0) return;

                  listView.SelectedItems.Clear(); // Standard Shift+Click clears previous range
                  int start = Math.Min(anchorIndex, clickedIndex);
                  int end = Math.Max(anchorIndex, clickedIndex);
                  for (int i = start; i <= end; i++) listView.SelectedItems.Add(listView.Items[i]);
                  e.Handled = true; // Prevent drag start on shift-click
             }
             else if (Keyboard.Modifiers == ModifierKeys.Control)
             {
                 // Ctrl+Click: Toggle selection
                 if (listView.SelectedItems.Contains(clickedItem))
                     listView.SelectedItems.Remove(clickedItem);
                 else
                     listView.SelectedItems.Add(clickedItem);
                 _dragSelectionAnchor = clickedItem; // Set anchor to the toggled item
                 e.Handled = true; // Prevent drag start on ctrl-click
             }
             else
             {
                 // Simple Click: Select only this item (unless it's already the only selected one)
                 if (!listView.SelectedItems.Contains(clickedItem) || listView.SelectedItems.Count != 1)
                 {
                     listView.SelectedItems.Clear();
                     listView.SelectedItem = clickedItem;
                 }
                  _dragSelectionAnchor = clickedItem; // Set anchor
                  // Do not set e.Handled = true here, allow MouseMove to potentially start dragging
             }
        }

        private void WellsListView_MouseMove(object sender, MouseEventArgs e)
        {
            // Only proceed if left button is pressed and mouse has moved significantly
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                 WinPoint currentPoint = e.GetPosition(null);
                 Vector diff = _startPoint - currentPoint;

                // Start dragging only if mouse moved significantly
                if (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance ||
                    Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance)
                {
                     if(!_isDragging) 
                     {
                        // This is the start of the drag operation
                        _isDragging = true;
                        // If drag starts, ensure initial item is selected if modifiers weren't pressed
                        if (_lastSelectedItem != null && Keyboard.Modifiers == ModifierKeys.None)
                        {
                            var listView = sender as ListView;
                            if (listView != null && !listView.SelectedItems.Contains(_lastSelectedItem.Content))
                            {
                                listView.SelectedItems.Clear();
                                listView.SelectedItems.Add(_lastSelectedItem.Content);
                                _dragSelectionAnchor = _lastSelectedItem.Content;
                            }
                        }
                     } 
                     
                     // Perform drag selection logic
                     PerformDragSelection(sender as ListView, e.GetPosition(sender as IInputElement));
                }
            }
        }

         private void WellsListView_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
         {
             if (_isDragging)
             {
                // Optionally finalize selection or clear drag state
                 _isDragging = false;
                 e.Handled = true; // Prevent potential click event after drag
             }
             Mouse.Capture(null); // Release mouse capture
         }

        private void PerformDragSelection(ListView? listView, WinPoint currentPosition)
        {
            if (listView == null || _lastSelectedItem == null) return;

            WinPoint startItemPos = _lastSelectedItem.TransformToAncestor(listView).Transform(new WinPoint(0, 0));
            Rect selectionRect = new Rect(startItemPos, currentPosition);

            // If not holding Ctrl, clear previous selection before applying new drag selection
            if (Keyboard.Modifiers != ModifierKeys.Control) // Check if Ctrl is NOT held
            {
                 listView.SelectedItems.Clear();
                 // Ensure the item where drag started is selected
                 if (!listView.SelectedItems.Contains(_lastSelectedItem.Content))
                    listView.SelectedItems.Add(_lastSelectedItem.Content);
            }

            foreach (var item in listView.Items)
            {
                var container = listView.ItemContainerGenerator.ContainerFromItem(item) as ListViewItem;
                if (container != null)
                {
                    WinPoint itemPos = container.TransformToAncestor(listView).Transform(new WinPoint(0, 0));
                    Rect itemBounds = new Rect(itemPos, container.RenderSize);

                    if (selectionRect.IntersectsWith(itemBounds))
                    {
                         if (!listView.SelectedItems.Contains(item))
                            listView.SelectedItems.Add(item);
                    }
                    // Optional: Deselect if Ctrl is not held and item is outside rect?
                    // else if (Keyboard.Modifiers != ModifierKeys.Control && listView.SelectedItems.Contains(item))
                    // {
                    //     listView.SelectedItems.Remove(item);
                    // }
                }
            }
        }

        // --- Helper Methods --- 

        private ListViewItem? GetListViewItemFromPoint(ListView listView, WinPoint point)
        {
            if (VisualTreeHelper.HitTest(listView, point) is HitTestResult hitTestResult)
            {
                DependencyObject? obj = hitTestResult.VisualHit;
                while (obj != null && !(obj is ListViewItem) && obj != listView)
                {
                    obj = VisualTreeHelper.GetParent(obj);
                }
                return obj as ListViewItem;
            }
            return null;
        }

        // Helper to get 0-based row index from 'A', 'B', etc.
        private int GetRowIndex(string? rowLabel)
        {
            if (string.IsNullOrEmpty(rowLabel) || rowLabel.Length != 1)
                return -1; // Invalid label
            char c = char.ToUpper(rowLabel[0]);
            if (c >= 'A' && c <= 'Z')
                return c - 'A';
            return -1; // Not a valid row letter
        }

         // Optional: Update command CanExecute on selection change if needed
         private void WellsListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
         {
             if (ViewModel != null)
             {
                // Example: Manually update CanExecute if command depends on selection count/content
                //          and binding CommandParameter isn't sufficient.
                 var selectedItemsList = (sender as ListView)?.SelectedItems.Cast<object>().ToList() ?? new List<object>();
                 ViewModel.SetPatientInfoCommand.NotifyCanExecuteChanged(); 
             }
         }
    }
} 