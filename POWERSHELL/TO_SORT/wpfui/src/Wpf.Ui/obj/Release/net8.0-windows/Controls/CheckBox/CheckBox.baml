   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  5clr-namespace:Wpf.Ui.ConvertersWpf.Ui.Converters  O Kmscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e0891&clr-namespace:System;assembly=mscorlibSystem W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089.   (Wpf.Ui.Converters.FallbackBrushConverter   FallbackBrushConverter  FallbackColor  CheckBoxPadding   CheckBoxBorderThemeThickness  CheckBoxContentMargin  CheckBoxIconSize  CheckBoxHeight  CheckBoxWidth  DefaultCheckBoxStyle  �� 
TargetType �� Property "	 DefaultControlFocusVisualStyle �� Value 
 CheckBoxBackground  CheckBoxForeground  CheckBoxBorderBrush  ControlCornerRadius  ControlContentThemeFontSize-  'System.Windows.Input.KeyboardNavigation �� 
TargetType �� 
Background �� Bullet ��Name �� CornerRadius    Wpf.Ui.Controls.SymbolIcon  Name   Wpf.Ui.Controls.FontIcon	   FontSize
   
FontWeight   CheckBoxCheckGlyphForeground!   Wpf.Ui.Controls.IconElement   
Foreground   Symbol ��Name �� Triggers P� Value P� Property �� 
TargetName & "CheckBoxCheckBackgroundFillChecked  CheckBoxCheckBorderBrush ]� 
Conditions �� Property �� Value 3 /CheckBoxCheckBackgroundFillUncheckedPointerOver 1 -CheckBoxCheckBackgroundFillCheckedPointerOver 0 ,CheckBoxCheckBackgroundFillUncheckedDisabled 2 .CheckBoxCheckBackgroundStrokeUncheckedDisabled ' #CheckBoxForegroundUncheckedDisabled  TextFillColorDisabledBrush �� BasedOn�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls  .
convertersclr-namespace:Wpf.Ui.Converters  1system&clr-namespace:System;assembly=mscorlib  %y	  &	        &	      &	      &	 +     &	 6     &	 G     &	 S     &	 _     &	 k     7  7  7  7  7  7  '�� �    7     ��	#FFFF0000��u�	11,5,11,6q�u�1q�u�8,0,0,0q�X�14Q�X�22Q�X�22Q��� "  ��.���  �@�FocusVisualStyle# � 	 ��  � ��# � 
 ��  � ��# �  ��  � ��# �  ��  � ��8   ��  � ��8  ��  �@��CornerRadius# �  ��  � ��$
 Left=���  � ��$ Center=���  � ��$
 Left=���  � ��$ Center=���  �@�Cursor$
 Hand����  � ��# �  ��  � ��$ Normal"���  �@ 	IsTabStop$
 True����  � |�$
 True����  � ��$	 120����  � ��$ 32����  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� " ��.R��� #��z��#��z��
 ����  �� $  ControlBorderIconPresenter��8�� 8�� $
��Left=�$��Center=�#��z��# z ._��� $ StrokeBorder��#��z��#��z��# z ._�� .+�  $ ControlIcon��$��0q�$��Center=�$��Center=�8	  $

 Bold"�# �  $ Checkmark48=�$y�	Collapsed=�._��� $ ContentPresenter��8�� #��z��#��z��$
��True�� P�  Z�  � ��. ��� $ ContentPresenter�� � ��$ 0q���  � ��$ 30��P�  � ��$  ��. ��� $ ContentPresenter�� � ��$ 0q���  � ��$ 30��P�  Z�  �@d�	IsChecked. ��� $ ControlIcon�� �@ Symbol$ 
Subtract16=��� $ ControlIcon�� � y�$ Visible=��� $  ControlBorderIconPresenter��# �   � ���� $ StrokeBorder��# �   � ��P�  �@d�	IsChecked$
 TrueY�. ��� $ ControlIcon�� � y�$ Visible=��� $  ControlBorderIconPresenter��# �   � ���� $ StrokeBorder��# �   � ��]�  ��  �@J�IsMouseOver$
 True����  �@d�	IsChecked$ FalseY�./��� $ StrokeBorder��# �   � ��]�  ��  �@J�IsMouseOver$
 True����  Z�  �@d�	IsChecked./��� $  ControlBorderIconPresenter��# �   � ���� $ StrokeBorder��# �   � ��]�  ��  �@J�IsMouseOver$
 True����  �@d�	IsChecked$
 TrueY�./��� $  ControlBorderIconPresenter��# �   � ���� $ StrokeBorder��# �   � ��P�  � {�$ False��. ��� $  ControlBorderIconPresenter��# �   � ���� $ StrokeBorder��# �   � ���� $ ControlIcon��# �   �@ 
Foreground�� # �   � ���� 8   "  ��