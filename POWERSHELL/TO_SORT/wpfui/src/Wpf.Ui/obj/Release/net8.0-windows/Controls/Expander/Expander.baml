   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  5clr-namespace:Wpf.Ui.ConvertersWpf.Ui.Converters  U QSystem.Runtime, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a7,clr-namespace:System;assembly=System.RuntimeSystem W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089   ExpanderPadding   ExpanderBorderThemeThickness  ExpanderChevronSize $  DefaultExpanderToggleButtonStyle  �� 
TargetType � ColumnDefinitions ��Name �Name J� RenderTransformOrigin �� Angle     Wpf.Ui.Controls.SymbolIcon   Name   Wpf.Ui.Controls.FontIcon   FontSize!   Wpf.Ui.Controls.IconElement   
Foreground	    Symbol
 �� Triggers P� Property P� Value N� EnterActions �� 
TargetName �� TargetProperty	 W� To g� Duration N� ExitActions  DefaultExpanderStyle �� 
TargetType �� Property " DefaultControlFocusVisualStyle �� Value  ExpanderHeaderBackground  ExpanderHeaderForeground  ExpanderHeaderBorderBrush 	 ControlContentThemeFontSize 
 ControlCornerRadius � 	Resources7  1Wpf.Ui.Converters.AnimationFactorToValueConverter  Key � RowDefinitions ��Name �� CornerRadius d�Name � OverridesDefaultStyle d� 	IsChecked �� Mode �� RelativeSource  ExpanderContentBackground
 � Tag  T� Y # AnimationFactorToValueConverter! `� 	Converter" `� ConverterParameter# �� ElementName$ �� Path% �� 
TargetName& R� KeyTime  System.Windows.Visibility'   Visible( R� Value) O� KeyTime* O� Value+ �� 	KeySpline,   	Collapsed $  ExpanderHeaderDisabledForeground % !ExpanderHeaderDisabledBorderBrush- P� 
SourceName ( $ExpanderHeaderBorderPointerOverBrush. �� BasedOn�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls  .
convertersclr-namespace:Wpf.Ui.Converters  7system,clr-namespace:System;assembly=System.Runtime  %�
  &	        &	      &	       &	 .     7  &	 p    7  7   7  7  '9� r
    7  u�11,11,11,11q�u�1q�X�16.0Q��� "  d�.R�� #��z��
������  �� $��* ��� $
��Auto �.+��� $ ContentPresenter��$��0��#��z��#��z��#��z��#��z��� $ ChevronGrid��$��1��$��0q�$��Center=�
������ $ 0.5, 0.5&�z��� $ 0Q�.+�   $ ControlChevronIcon��8   # z��$	 ChevronDown24=�
 P�  �@d�	IsChecked$
 TrueY� �� .���� .��W� $ ChevronGrid��$4 .(Grid.RenderTransform).(RotateTransform.Angle)� 180$ 00:00:00.167D� �� .���� .��W� $ ChevronGrid��$4 .(Grid.RenderTransform).(RotateTransform.Angle)� 0$ 00:00:00.167D��� " 9�.���  �@�FocusVisualStyle# �  ��  � ��# �  ��  � ��# �  ��  � ��# �  ��  � ��8   ��  � ��8  ��  � ��$ Stretch=���  � ��$ Center=���  � ��$ Stretch=���  � ��$ Center=���  � ��# � 	 ��  � ��$ Normal"���  �@��CornerRadius# � 
 ��  �@9�
IsExpanded$ False����  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� "  9� ��  #AnimationFactorToValueConverter .R��  �� $
��Auto ��� $
��Auto �.+��� $ ToggleButtonBorder��$��0��#��z��#��z��$��1q�# z ._�d� $ ExpanderToggleButton��$��0q�#��z��$��Stretch=�#��z��$��Stretch=�$��Center=�#��z��#��z��#��z��#{�z{�$
 True��8��  �� *
IsExpanded+$ TwoWay=� �� *TemplatedParent+� $��1��$
}�True��.+��� $ ContentPresenterBorder��#���  #��z��$��1,0,1,1q�$ 0,0,4,4��$y�	Collapsed=�._��� $ ContentPresenter��#��z��#��z��#��z��#��z�� X�0.0Q�z�T�   `� 8!  $" negative��.1��� $# ContentPresenterBorder��$$ ActualHeight��� $# ContentPresenterBorder��$	$ Tag�
 P�  �@9�
IsExpanded$
 True��. ��� $% ToggleButtonBorder�� �@��CornerRadius$ 4,4,0,0�� �� .���� .��T� $ ContentPresenterBorder��$ (Border.Visibility)�..�i� $& 0��#( Z' U� $ ContentPresenterBorder��$	 Tag�.\�n� $) 0��* 1.0�� $+ 0.0, 0.0, 0.0, 1.0��$) 	0:0:0.333��* 0.0 �� .���� .��T� $ ContentPresenterBorder��$ (Border.Visibility)�..�i� $& 0��#( Z' i� $& 0:0:0.2��#( Z, U� $ ContentPresenterBorder��$	 Tag�.\�n� $) 0��* 0.0�� $+ 1.0, 1.0, 0.0, 1.0��$) 	0:0:0.167��* 1.0P�  � {�$ False��. ��� $% ContentPresenter��# �   � ���� $% ExpanderToggleButton��# �   � ���� $% ExpanderToggleButton��# �   � ��P� $- ExpanderToggleButton�� �@J�IsMouseOver$
 True��. ��� $% ExpanderToggleButton��# �   � ���� 8.   " 9�