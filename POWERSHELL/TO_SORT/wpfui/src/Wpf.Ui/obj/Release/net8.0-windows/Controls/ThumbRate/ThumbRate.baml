   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089   DefaultUiThumbRateStyle  �� 
TargetType    Wpf.Ui.Controls.ThumbRate �� Property �� Value  ThumbRateForeground-  'System.Windows.Input.KeyboardNavigation �� 
TargetType � ColumnDefinitions   Wpf.Ui.Controls.Button$   Wpf.Ui.Controls.ThumbRateState   Liked �� Path �� Mode �� RelativeSource    Wpf.Ui.Controls.SymbolIcon	  Name   Wpf.Ui.Controls.FontIcon
   FontSize   Symbol   Disliked �� Triggers P� Property P� Value �� 
TargetName �� BasedOn�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls  %�  &	        '   e    7   �� "    .���  � ��
 ���� ��  � ��# �  ��  � ��$ 24(���  � |�$ False����  �@ 	IsTabStop$ False����  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� "   .R��  �� $
��Auto ��� $
��Auto �.+�  $��0��
������ 
������ #��Z #��z������ $ TemplateButtonCommand�$ OneWay=� �� *TemplatedParent+.��  $	 ThumbsUpButtonIcon��#
 z��$ ThumbLike24=�  $��1��
������ 
������ #��Z #��z������ $ TemplateButtonCommand�$ OneWay=� �� *TemplatedParent+.��  $	 ThumbsDownButtonIcon��#
 z��$ ThumbDislike24=� P�  �@  State$ Liked=�. ��� $ ThumbsUpButtonIcon�� �@ Filled$
 True��P�  �@  State$ Disliked=�. ��� $ ThumbsDownButtonIcon�� �@ Filled$
 True���� 8   "    