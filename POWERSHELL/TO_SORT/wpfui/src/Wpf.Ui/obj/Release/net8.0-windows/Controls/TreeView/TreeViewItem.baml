   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  U QSystem.Runtime, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a7,clr-namespace:System;assembly=System.RuntimeSystem W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089   TreeViewItemChevronSize  TreeViewItemFontSize # ExpandCollapseToggleButtonStyle  �� 
TargetType �� Property " DefaultControlFocusVisualStyle �� Value �� 
TargetType �Name J� RenderTransformOrigin �� Angle     Wpf.Ui.Controls.SymbolIcon   Name   Wpf.Ui.Controls.FontIcon   FontSize	    Symbol
 �� Triggers P� Property P� Value N� EnterActions �� 
TargetName �� TargetProperty	 W� To g� Duration N� ExitActions  DefaultTreeViewItemStyle  TreeViewItemForeground  TreeViewItemBackground  ControlCornerRadius � RowDefinitions ��Name �� CornerRadius � ColumnDefinitions ��Name , (TreeViewItemSelectionIndicatorForeground �� RadiusX �� RadiusY d�Name �� 	ClickMode d� 	IsChecked �� RelativeSource ��Name ��Name  �� 
TargetName %	 !TreeViewItemBackgroundPointerOver "
 TreeViewItemBackgroundSelected! ]� 
Conditions" �� Property# �� Value$ �� Triggers  DefaultUiTreeViewItemStyle"   Wpf.Ui.Controls.TreeViewItem  TreeViewItemFocusVisual% �� Padding& �� Opacity  SystemAccentColor' �� Color(   Icon'  !System.Windows.VisualStateManager)   VisualStateGroups%  System.Windows.VisualStateGroup*   States   System.Windows.VisualState+   
Storyboard, W� From- R� KeyTime  System.Windows.Visibility.   Visible/ R� Value0 �� BasedOn�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls  7system,clr-namespace:System;assembly=System.Runtime  %]  &	        &	      &	      7   &	 �    7  7  &	 �    7  7  'Q� �    7  '  �    7  X�10Q�X�14Q��� "  d�.���  �@�FocusVisualStyle# �  ��  � |�$ False����  �@�OverridesDefaultStyle$
 True����  �@J�SnapsToDevicePixels$
 True����  � �� �� " d�.R�� $ ChevronContainer��$��15��$��15��
������ $ 0.5, 0.5&�z��� $ 0Q�.+�   $ ChevronIcon��$��Center=�8   $	 ChevronRight28=�
 P�  �@d�	IsChecked$
 TrueY� �� .���� .��W� $ ChevronContainer��$4 .(Grid.RenderTransform).(RotateTransform.Angle)� 90$ 0:0:0.16D� �� .���� .��W� $ ChevronContainer��$4 .(Grid.RenderTransform).(RotateTransform.Angle)� 0$ 0:0:0.16D��� "  Q�.���  �@�FocusVisualStyle# �  ��  � ��# �  ��  � ��# �  ��  � ��$ 4q���  � ��8   ��  �@��CornerRadius# �  ��  � ��$
 True����  �@�OverridesDefaultStyle$
 True����  �@J�SnapsToDevicePixels$
 True����  � �� �� " Q�.R�� #��z�� �� $
��Auto ��� $
��Auto �.+��� $ Border��$��0��# z ._��  �� $
��Auto �$��19���� $
��Auto ��� $��* �.+��� $ ActiveRectangle��$��0��$��3��$��16��$��0,0,0,0q�$
��Left=�$��Center=�#���  $ 2��$ 2��$y�	Collapsed=�d� $ Expander��$��0��$	��8,0q�$ Press=�8��  �� *
IsExpanded+ �� *TemplatedParent+�� $ PART_Header��$��1��#��z��#��z��$��Header��#��z��� $��1�� �� $
��Auto �$��19���� $
��Auto ��� $��* �.+��� $ 	ItemsHost��$��1��$��2��$y�	Collapsed=�
 P�  �@��HasItems$ False��. ��� $  Expander�� � y�$ Hidden=�P�  �@Q�
IsExpanded$
 True��. ��� $  	ItemsHost�� � y�$ Visible=�P�  �@J�IsMouseOver$
 True��. ��� $  Border��# � 	  � ��P�  �@Q�
IsSelected$
 True��. ��� $  Border��# � 
  � ���� $  ActiveRectangle�� � y�$ Visible=�]� ! �� " � ��$# False���� " � ��$
# Auto��./��� $  PART_Header�� � ��$ 75��]� ! �� " � ��$# False���� " � ��$
# Auto��./��� $  PART_Header�� � ��$ 19��$ P�  �@%�IsVirtualizing$
 True��. ���  � �� �� .R�$� �� "   .���  � ��# �  ��  � ��# �  ��  � ��$ 4q���  �@�FocusVisualStyle8   ��  �@�OverridesDefaultStyle$
 True����  �@J�SnapsToDevicePixels$
 True����  � �� �� "  .R�� #��z�� �� $
��Auto ��� $
��Auto ��� $��* � �� $
��Auto ��� .+�d� $ Expander��$��0��$��0��$ Press=�8��  �� *
IsExpanded+ �� *TemplatedParent+�� $ Bd��$��0��$��1��#% z��#��z��#��z������ $	& 0.0Q�#' �  ._��  �� $
��Auto ��� $��* �.+��� $ 
SymbolIcon��$��0��$��0,0,8,0q�#��z( #��z���� $ PART_Header��$��1��#��z��$��Header���� $ 	ItemsHost��$��1��$��1��$��2��$y�	Collapsed=�)   SelectionStates��.*   Selected��.+ �� .��W� $2 ,(Panel.Background).(SolidColorBrush.Opacity)�, 0.0 1.0$ 0:0:.16D�  
Unselected��  SelectedInactive��.+ �� .��W� $2 ,(Panel.Background).(SolidColorBrush.Opacity)�, 0.0 1.0$ 0:0:.16D�  ExpansionStates��.*   Expanded��.+ �� .��T� $ 	ItemsHost��$ (UIElement.Visibility)�..�i� $- 0��#/ Z.   	Collapsed��
 P�  Z�  �@ Icon. ��� $  
SymbolIcon�� � y�$ 	Collapsed=��� $  
SymbolIcon�� � ��$ 0q�P�  �@��HasItems$ False��. ��� $  Expander�� � y�$ Hidden=�]� ! �� " � ��$# False���� " � ��$
# Auto��./��� $  PART_Header�� � ��$ 75��]� ! �� " � ��$# False���� " � ��$
# Auto��./��� $  PART_Header�� � ��$ 19��$ P�  �@%�IsVirtualizing$
 True��. ���  � �� �� .R�$� �� 80   "  Q��� 80   "   