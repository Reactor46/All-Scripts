   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089   InfoBarPadding  InfoBarBorderThemeThickness    Wpf.Ui.Controls.InfoBar  �� 
TargetType �� Property �� Value  ControlFillColorDefault �� Color  InfoBarTitleForeground  InfoBarBorderBrush  ControlContentThemeFontSize  ControlCornerRadius-  'System.Windows.Input.KeyboardNavigation �� 
TargetType �Name ��Name �� Padding �� CornerRadius	 � ColumnDefinitions    Wpf.Ui.Controls.SymbolIcon
  Name   Filled   Wpf.Ui.Controls.FontIcon   FontSize    Title    Message   Wpf.Ui.Controls.Button  Name �� Path �� Mode �� RelativeSource   Icon   Symbol �� Triggers P� Property P� Value �� 
TargetName!   Wpf.Ui.Controls.IconElement  SystemAccentColorSecondary / +InfoBarInformationalSeverityBackgroundBrush (	 $InfoBarSuccessSeverityIconBackground )
 %InfoBarSuccessSeverityBackgroundBrush ( $InfoBarWarningSeverityIconBackground ) %InfoBarWarningSeverityBackgroundBrush & "InfoBarErrorSeverityIconBackground ' #InfoBarErrorSeverityBackgroundBrush�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls  %q  &	        &	      '        7   7  u�	14,8,14,8q�u�1q��� "    .���  � �� �� # �  ��  � ��# �  ��  � ��8   ��  � ��$ Stretch=���  � ��$ Center=���  � ��$ Stretch=���  � ��$ Center=���  � ��# �  ��  � ��8  ��  � ��# �  ��  � ��$ Normal"���  �@��CornerRadius# �  ��  � |�$ False����  �@ 	IsTabStop$ False����  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� "   .R�� $ InfoBarRoot��.+��� $ ContentBorder��$ 16,14,16,10q�#��z��#��z��#��z��#��z��#��z��# z ._�� 	 �� $
��Auto ��� $��* ��� $
��Auto �.+��� $��0,2,14,0q�$	��Top=�._�  $
 
SymbolIcon��$	��Top=�$
 True��$ 16(�� $��1��$	��Top=�.+��� $��0,0,14,0q�$��False��#��z #��z��$��SemiBold"�$
��Wrap=��� $��0q�$��False��#��z #��z��$
��Wrap=��� $��2��$��	12,-2,0,0q�$	��Top=�._�  $ CloseButton��$��4q�
������ $��0q�$
{�True������ $ TemplateButtonCommand�$ OneWay=� �� *TemplatedParent+   $ 	Dismiss24=� P�  �@  
IsClosable$ False��. ��� $ CloseButton�� � y�$ 	Collapsed=�P�  �@  IsOpen$
 True��. ��� $ InfoBarRoot�� � y�$ Visible=�P�  �@  IsOpen$ False��. ��� $ InfoBarRoot�� � y�$ 	Collapsed=�P�  �@  Severity$ Informational=�. ��� $ 
SymbolIcon�� �@ 
Foreground �� # �  �� $ 
SymbolIcon�� �@ Symbol$ Info24=��� $ ContentBorder��# �   � ��P�  �@  Severity$ Success=�. ��� $ 
SymbolIcon��# � 	  �@ 
Foreground�� $ 
SymbolIcon�� �@ Symbol$ CheckmarkCircle24=��� $ ContentBorder��# � 
  � ��P�  �@  Severity$ Warning=�. ��� $ 
SymbolIcon��# �   �@ 
Foreground�� $ 
SymbolIcon�� �@ Symbol$ ErrorCircle24=��� $ ContentBorder��# �   � ��P�  �@  Severity$ Error=�. ��� $ 
SymbolIcon��# �   �@ 
Foreground�� $ 
SymbolIcon�� �@ Symbol$ DismissCircle24=��� $ ContentBorder��# �   � ��