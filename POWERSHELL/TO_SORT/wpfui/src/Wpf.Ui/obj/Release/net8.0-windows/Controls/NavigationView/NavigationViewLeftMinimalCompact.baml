   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089  �� MergedDictionaries �� Source 0  ,LeftMinimalCompactNavigationViewItemTemplate �� 
TargetType(    "Wpf.Ui.Controls.NavigationViewItem � RowDefinitions ��Name �� Padding �� CornerRadius � ColumnDefinitions ��Name	    Icon   NavigationViewItemForeground
 ��Name . *NavigationViewSelectionIndicatorForeground J� Opacity �� RadiusX �� RadiusY ��Name    	InfoBadge    Wpf.Ui.Controls.SymbolIcon  Name   Symbol    MenuItemsSource-  'System.Windows.Input.KeyboardNavigation   DirectionalNavigation K� IsItemsHost � 	Resources " NavigationViewItemDefaultStyle �� BasedOn �� 
TargetType �� Property �� Value �� Triggers P� Property P� Value �� 
TargetName ( $NavigationViewItemBackgroundSelected + 'NavigationViewItemBackgroundPointerOver ]� 
Conditions �� Property  �� Value�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls    �� $i cpack://application:,,,/Wpf.Ui;component/Controls/NavigationView/NavigationViewItemDefaultStyle.xaml?�%�  &	        7  �� "   .R��  �� $
��Auto ��� $
��Auto �.+��� $ 
MainBorder��$��0��$	��310��$��40��$��0q�$ 0q�
������ #��z��$��1q�# z ._��  �� $��40 ��� $��* ��� $
��Auto ��� $
��Auto �.+�� $��0��.+��� $ IconContentPresenter��$��-1,0,0,0q�$��Center=�$��Center=�#��z	 #���  �� $
 ActiveRectangle��$��0��$��3��$��16��$��0q�$
��Left=�$��Center=�#���  $	 0.0Q�$ 2��$ 2���� $ ElementContentPresenter��$��1��$
��Left=�$��Center=�#��z��$��14(�#��z���� $ PART_InfoBadge��$��2��$��0q�#��z   $ ChevronIcon��$��3��$ ChevronDown24=�$y�	Collapsed=��� $��1��$|�False��#��z $ 	Contained=�$��Disabled=�$��Disabled=����� .R��� $��0q�$
 True�� ���� 8   '         "   .���  � �� �� "   .R��� $ 
MainBorder��$��0��$��40��$��0q�$ 0q�
������ #��z��$��1q�# z ._��  �� $��40 ��� $��* �$	��180���� $
��Auto ��� $
��Auto �.+�� $��0��.+��� $ IconContentPresenter��$��-1,0,0,0q�$��Center=�$��Center=�#��z	 #���  �� $
 ActiveRectangle��$��0��$��3��$��16��$��0q�$
��Left=�$��Center=�#���  $	 0.0Q�$ 2��$ 2���� $ ElementContentPresenter��$��1��$
��Left=�$��Center=�#��z��$��14(�#��z���� $ PART_InfoBadge��$��2��$��0q�#��z   $ ChevronIcon��$��3��$ ChevronDown24=�$y�	Collapsed=� P�  �@  
IsPaneOpen$
 True��. ��� $ PART_InfoBadge�� � ��$ 2��P�  �@  
IsPaneOpen$ False��. ��� $ PART_InfoBadge�� � ��$ 0���� $ PART_InfoBadge�� � ��$	 Top=��� $ PART_InfoBadge�� � ��$ Right=�P�  �@  HasMenuItems$
 True��. ��� $ ChevronIcon�� � y�$ Visible=��� $ ChevronIcon�� � ��$ 8,0,8,0q�P�  �@  IsActive$
 True��. ��� $ ActiveRectangle�� �@J�Opacity$	 1.0Q��� $ 
MainBorder��# �   � ��P�  Z�  �@  Icon. ��� $ IconContentPresenter�� � y�$ 	Collapsed=�P�  �@J�IsMouseOver$
 True��. ��� $ 
MainBorder��# �   � ��]�  ��  �@J�IsMouseOver$
  True����  �@  IsActive$  False��/��� $ IconContentPresenter��# �   � ���� $ ElementContentPresenter��# �   � ��