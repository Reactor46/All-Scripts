   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  U QSystem.Runtime, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a7,clr-namespace:System;assembly=System.RuntimeSystem W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089   ComboBoxPadding   ComboBoxBorderThemeThickness & "ComboBoxAccentBorderThemeThickness  ComboBoxChevronMargin  ComboBoxItemMargin  ComboBoxItemContentMargin  ComboBoxChevronSize  ComboBoxPopupMinHeight  DefaultComboBoxTextBoxStyle  �� 
TargetType �� Property �� Value 	 DefaultControlContextMenu 
 ComboBoxForeground �� 
TargetType {�Name $  DefaultComboBoxToggleButtonStyle ��Name �� CornerRadius ��Name  DefaultComboBoxItemStyle " DefaultControlFocusVisualStyle  ControlCornerRadius �� Padding	 J� SnapsToDevicePixels
 ��Name  ComboBoxItemPillFillBrush �� RadiusX �� RadiusY �� Triggers P� Property P� Value �� 
TargetName " ComboBoxItemBackgroundSelected " ComboBoxItemForegroundSelected  ComboBoxForegroundDisabled  DefaultComboBoxStyle  ComboBoxBackground  ControlElevationBorderBrush  ControlContentThemeFontSize  TextControlThemeMinHeight  TextControlThemeMinWidth � ColumnDefinitions �� SelectionBoxItem �� SelectionBoxItemTemplate J� IsHitTestVisible     Wpf.Ui.Controls.SymbolIcon   Name   Wpf.Ui.Controls.FontIcon   FontSize # ComboBoxDropDownGlyphForeground!   Wpf.Ui.Controls.IconElement   
Foreground J� RenderTransformOrigin    Symbol �� Angle �� 	ClickMode d� 	IsChecked �� Path �� Mode �� RelativeSource  ��Name! �� 
IsReadOnly" �� 
IsReadOnly# �Name$ � ActualWidth% � AllowsTransparency& �� IsDropDownOpen' � VerticalOffset  ComboBoxDropDownBackground  ComboBoxDropDownBorderBrush  PopupCornerRadius)   #Wpf.Ui.Controls.DynamicScrollViewer( �� MaxDropDownHeight) K� IsItemsHost-  'System.Windows.Input.KeyboardNavigation*   DirectionalNavigation  ComboBoxBorderBrushFocused+ N� EnterActions, �� 
TargetName- �� TargetProperty. W� From	/ W� To0 g� Duration1 W� EasingFunction/  )System.Windows.Media.Animation.CircleEase7  1System.Windows.Media.Animation.EasingFunctionBase2   
EasingMode3 N� ExitActions4 P� 
SourceName5 ]� 
Conditions6 �� Property7 �� Value  ComboBoxBackgroundFocused ! ComboBoxBackgroundPointerOver   ComboBoxBackgroundDisabled ! ComboBoxBorderBrushDisabled8 �� BasedOn�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls  7system,clr-namespace:System;assembly=System.Runtime  %z  &	        &	      &	 $     &	 5     &	 G     &	 X     &	 j     &	 x     &	 �     &	 G    &	 �    7  7  &	      7  7  7  7  7  7  7  '�� �    7  '�� �    7  u�	10,8,10,8q�u�1,1,1,1q�u�0,0,0,2q�u�8,0,10,0q�u�3,2,3,0q�u�10,8,8,8q�X�11.0Q�X�32.0Q��� "  ��.���  �@�FocusVisualStyle Z� ��  �@�ContextMenu# � 	 ��  � ��# � 
 ��  �@��
CaretBrush# � 
 ��  � ��
 ���� ��  � y�$ Hidden=���  �@�Cursor$ IBeam����  � ��$ 0q���  � ��$ 0q���  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� " ��.R�{� $ PART_ContentHost��#��z��$��Stretch=�$��Stretch=�#��z���� "  d�.���  �@�FocusVisualStyle Z� ��  � ��# � 
 ��  � ��
 ���� ��  � ��
 ���� ��  � ��$ 0q���  �@��CornerRadius$ 0����  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� " d�.R��� $ ContentBorder��#��z��#��z��#��z��# z ._��� $ PART_ContentHost��#��z��#��z��#��z���� "  ��.���  �@�FocusVisualStyle# �  ��  � ��# � 
 ��  � ��
 ���� ��  � ��8   ��  � ��8  ��  �@�Cursor$
 Hand����  �@��CornerRadius# �  ��  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� " ��.R�� .+��� $��ContentBorder��#��z��$ 0q�$��Stretch=�# z $
	 True��._�� .+��� $ PART_ContentPresenter��#��z��$��Center=��� $
 ActiveRectangle��$��3��$��16��$��0q�$
��Left=�$��Center=�#���  $ 2��$ 2��$y�	Collapsed=� P�  �@��IsHighlighted$
 True��. ��� $ ContentBorder��# �   � ��P�  �@��
IsSelected$
 True��. ��� $ ActiveRectangle�� � y�$ Visible=��� $ PART_ContentPresenter��# �   � ��P�  � {�$ False��. ��� # �   � ���� "  ��.���  �@�FocusVisualStyle# �  ��  �@�ContextMenu# � 	 ��  � ��# � 
 ��  � ��# �  ��  � ��# �  ��  � ��8   ��  � ��# �  ��  � ��$ False����  � ��$ Hidden=���  � ��$ Hidden=��� " �@��IsDeferredScrollingEnabled$ False����  � ��$ Stretch=���  � ��$ Center=���  � ��$ Stretch=���  � ��$	 Top=���  � ��# �  ��  � ��# �  ��  � ��# �   ��  �@��CornerRadius# �  ��  � ��$
 None=���  � ��$ Bottom=���  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� " ��.R�� #��z��#��z��.+��� $ ContentBorder��$��0��#��z��#��z��$ 0q�$��Stretch=�$��Stretch=�#��z��#��z��#��z��# z ._�� #��z��#��z��.+��  �� $��* ��� $
��Auto �.+�� $��0��#��z��.+��� $��PART_ContentPresenter��$��Stretch=�$��Stretch=�#��z #��z #��z��$ False��#��z��� $��1��8�� .+�   $ ChevronIcon��$��0q�$��Center=�8  # �  $ 0.5, 0.5&�$ ChevronDown24=�z��� $ 0Q�� $��0��$��2��$��0q�.+�d� $��ToggleButton��$��Stretch=�$��Stretch=�$ Press=�$|�False��#��z��8��  �� $ IsDropDownOpen�$ TwoWay=� �� *TemplatedParent+� $��0��#��z��.+��� $  PART_EditableTextBox��$��Stretch=�$��Stretch=�#��z��#��z��#" z! 8�� � $# Popup��#��z$ $��Center=�$
% True��$|�False��#��z& #��z��#��z��$' 1��.���� $ DropDownBorder��$��0q�$ 0,4,0,6q�#���  #���  $��1q�# �  $
	 True��z�T� ._�� .+�  #��z( $��0q�#��z��$
	 True��#��z��#��z��#��z��#��z��.���� $
) True��$* 	Contained=�#��z���� $ AccentBorder��$��Stretch=�$��Stretch=�#���  8�� # z $y�	Collapsed=� P�  �@��IsDropDownOpen$
 True��+ �� .���� .��W� $, ChevronIcon��$C- =(controls:SymbolIcon.RenderTransform).(RotateTransform.Angle)�. 0/ 180$0 00:00:00.167D�W� $, DropDownBorder��$5- /(Border.RenderTransform).(TranslateTransform.Y)�. -90/ 0$0 00:00:00.167D�1   $2 EaseOut=�3 �� .���� .��W� $, ChevronIcon��$C- =(controls:SymbolIcon.RenderTransform).(RotateTransform.Angle)�. 180/ 0$0 00:00:00.167D�P�  �@��HasItems$ False��. ��� $ DropDownBorder��8   � ��P� $4 Popup�� �@�AllowsTransparency$ False��. ��� $ DropDownBorder�� �@��CornerRadius$ 0��P�  �@��
IsGrouping$
 True��. ���  � ��$ False��P�  �@��
IsEditable$
 True��. ���  � ��$ False���� $ PART_EditableTextBox�� � y�$ Visible=��� $ PART_ContentPresenter�� � y�$ Hidden=�]� 5 �� 6 � {�$
7 True���� 6 �@J�IsKeyboardFocusWithin$
7 True���� 6 �@��
IsEditable$
7 True��./��� $ ContentBorder��# �   � ���� $ AccentBorder�� � y�$ Visible=�]� 5 �� 6 � {�$
7 True���� 6 �@J�IsMouseOver$
7 True���� 6 �@J�IsKeyboardFocusWithin$7 False��./��� $ ContentBorder��# �   � ��P�  � {�$ False��. ��� $ ContentBorder��# �    � ���� $ ContentBorder��# � !  � ���� # �   � ���� 88   "  ���� 88   "  ��