   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  5clr-namespace:Wpf.Ui.ConvertersWpf.Ui.Converters  W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e0892   ,Wpf.Ui.Converters.ProgressThicknessConverter   ProgressThicknessConverter"   Wpf.Ui.Controls.ProgressRing  �� 
TargetType �� Property �� Value $  ProgressRingForegroundThemeBrush $  ProgressRingBackgroundThemeBrush �� 
TargetType   Wpf.Ui.Controls.Arc   EndAngle   
StartAngle   CoverRingStroke   CoverRingVisibility �� Path	 �� RelativeSource
 �� Mode �� 	Converter   EngAngle �� Triggers P� Property P� Value  Name   IndeterminateAngle J� RenderTransformOrigin �� Triggers N� EnterActions g� RepeatBehavior �� 
TargetName �� TargetProperty	 W� To g� Duration g� AutoReverse W� From N� ExitActions�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls  .
convertersclr-namespace:Wpf.Ui.Converters  %V  &	        '       7   7   7   7      �� "   .���  � ��$ 60����  � ��$ 60����  � |�$ False����  � ��# �  ��  �@ CoverRingStroke# �  ��  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� "  .R�� #��z��#��z��.+�  $	 359Q�$ 0Q�#��z #y�z ���� $ Height�	 �� *TemplatedParent+$
 OneWay=�8     # z $ 0Q�#��z������ $ Height�	 �� *TemplatedParent+$
 OneWay=�8   P�  �@ IsIndeterminate$
 True��. ���  � �� �� "  .R�� #��z��#��z��.+�  $	 359Q�$ 0Q�#��z #y�z ���� $ Height�	 �� *TemplatedParent+$
 OneWay=�8    $	 Arc��# z $ 0.5, 0.5&�$ 0Q�#��z������ $ Height�	 �� *TemplatedParent+$
 OneWay=�8  z���  P�  � {�$
 True�� �� .���� .��W� $ Forever��$	 Arc��$6 0(Canvas.RenderTransform).(RotateTransform.Angle)� 360$ 0:0:2D�W� $
 True��$ Forever��$	 Arc��$ EndAngle� 100 320$ 0:0:5D� �� .���� 