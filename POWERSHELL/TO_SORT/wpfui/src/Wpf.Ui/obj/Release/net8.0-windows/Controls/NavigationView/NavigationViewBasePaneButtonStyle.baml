   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089  �� MergedDictionaries �� Source   BasePaneButtonStyle �� 
TargetType    Wpf.Ui.Controls.Button �� Property �� Value  ButtonForeground  PaneToggleButtonThickness  PaneToggleButtonHeight  PaneToggleButtonWidth  ControlCornerRadius �� 
TargetType ��Name �� Padding �� CornerRadius	 � ColumnDefinitions
 ��Name    Icon J� RenderTransformOrigin �� ScaleX �� ScaleY �� Triggers P� Property P� Value �� 
TargetName  ButtonBackgroundPointerOver  ButtonBackgroundDisabled  ButtonForegroundDisabled :� RoutedEvent �� 
TargetName �� TargetProperty W� From	 W� To g� Duration�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls    �� $b \pack://application:,,,/Wpf.Ui;component/Controls/NavigationView/NavigationViewConstants.xaml?�%  &	        7  7  7  �� "   .���  � ��
 ���� ��  � ��# �  ��  � ��
 ���� ��  � ��8   ��  � ��$ 16(���  � ��8  ��  � ��8  ��  � ��$ 10q���  �@��CornerRadius# �  ��  � ��$ Stretch=���  � ��$	 Top=���  � ��$ Center=���  � ��$ Center=���  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� "   .R��� $ 
LayoutRoot��#��z��#��z��$��0q�# z��#��z��#��z��
������ #��z��#��z��# z ._�� 	 �� $
��Auto ��� $��* �.+��� $

 Icon��$��0��$��0,0,20,0q�$��Center=�#��z $|�False��$ 0.5, 0.5&�#��z��z��� IconScaleTransform��$	 1.0Q�$	 1.0Q��� $��1��#��z�� P�  �@J�IsMouseOver$
 True��. ��� $ 
LayoutRoot��# �   � ��P�  � {�$ False��. ��� # �   � ���� $
 Icon��# �   � ��P�  Z�  � ��. ��� $
 Icon�� � ��$ 0q�:� $' !Button.PreviewMouseLeftButtonDown��.[��� .���� .��W� $ IconScaleTransform��$ (ScaleTransform.ScaleX)� 1.0 0.85$ 0:0:0.08D�:� $% Button.PreviewMouseLeftButtonUp��.[��� .���� .��W� $ IconScaleTransform��$ (ScaleTransform.ScaleX)� 0.85 1.0$ 0:0:0.08D�