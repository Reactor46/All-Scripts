   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089  �� MergedDictionaries �� Source    BottomNavigationViewTemplate �� 
TargetType$    Wpf.Ui.Controls.NavigationView � RowDefinitions    Header    HeaderVisibility4   .Wpf.Ui.Controls.NavigationViewContentPresenter  Name    
Transition   
Transition	    TransitionDuration
   TransitionDuration    ContentOverlay �Name � ColumnDefinitions   Wpf.Ui.Controls.Button  Name    IsBackEnabled  BasePaneButtonStyle - )BackButtonVisibilityToVisibilityConverter �� 	Converter   Icon    Wpf.Ui.Controls.SymbolIcon   Symbol    
PaneHeader ��Name-  'System.Windows.Input.KeyboardNavigation   DirectionalNavigation K� IsItemsHost �� Orientation    AutoSuggestBox    
PaneFooter + 'BottomCompactNavigationViewItemTemplate(   "Wpf.Ui.Controls.NavigationViewItem ��Name ��Name   Icon   NavigationViewItemForeground   	InfoBadge ��Name . *NavigationViewSelectionIndicatorForeground J� Opacity  �� RadiusX! �� RadiusY" �� Triggers# P� Property$ P� Value% �� 
TargetName& �� Property' �� Value�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls    �� $l fpack://application:,,,/Wpf.Ui;component/Controls/NavigationView/NavigationViewBasePaneButtonStyle.xaml?��� $i cpack://application:,,,/Wpf.Ui;component/Controls/NavigationView/NavigationViewItemDefaultStyle.xaml?��� $b \pack://application:,,,/Wpf.Ui;component/Controls/NavigationView/NavigationViewConstants.xaml?�%g  &	        7  7  &	 v    �� "   .R��  �� $��* ��� $
��Auto �.+�� $��0�� �� $
��Auto ��� $��* �.+��� $��0��$��0q�#��z #y�z   $) #PART_NavigationViewContentPresenter��$��1��#��z��# z #
 z	 �� $��0��$��2��$��0q�#��z � $ PaneGrid��$��1�� �� $
��Auto ��� $
��Auto ��� $��* ��� $
��Auto ��� $
��Auto ��� $
��Auto �.+�  $ PART_BackButton��$��0��#{�z 8��  y��� *IsBackButtonVisible+8     $ ArrowLeft48=��� $��1��$��0q�#��z �� $  PART_MenuItemsItemsControl��$��2��$|�False��$ 	Contained=�$��Disabled=�$��Disabled=����� .R��� $��0q�$
 True��$ 
Horizontal=��� $��3��#��z �� $��4��$��0q�#��z �� $&  PART_FooterMenuItemsItemsControl��$��5��$|�False��$ 	Contained=�$��Disabled=�$��Disabled=����� .R��� $��0q�$
 True��$ 
Horizontal=��� "  .R��� $ 
MainBorder��._�� .+�� $��	12,8,12,8q�$��Stretch=�$��Stretch=� �� $
��Auto ��� $��* ��� $
��Auto �.+��� $ IconContentPresenter��$��0��$��0,0,6,0q�$��Center=�$��Center=�#��z $��16(�#���  �� $ ElementContentPresenter��$��1��$
��Left=�$��Center=�#��z��$��14(�#��z���� $ PART_InfoBadge��$��2��$��2q�$��Right=�$��Center=�#��z �� $ ActiveRectangle��$��0��$��16��$��3��$��0,0,0,-4q�$��Center=�$��Bottom=�#���  $	 0.0Q�$  2��$! 2��" P� # �@ IsActive$
$ True��. ��� $% ActiveRectangle��& �@J�Opacity$	' 1.0Q�P� $ Z� # �@ Icon. ��� $% IconContentPresenter��& � y�$' 	Collapsed=�P� $ Z� # � ��. ��� $% IconContentPresenter��& � ��$' 0q��� $% ElementContentPresenter��& � y�$' 	Collapsed=�P� # � ��$$  ��. ��� $% IconContentPresenter��& � ��$' 0q��� $% ElementContentPresenter��& � y�$' 	Collapsed=�