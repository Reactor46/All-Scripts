   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089   HyperlinkButtonPadding ' #HyperlinkButtonBorderThemeThickness  HyperlinkButtonIconMargin ! DefaultUiHyperlinkButtonStyle  �� 
TargetType%    Wpf.Ui.Controls.HyperlinkButton �� Property " DefaultControlFocusVisualStyle �� Value  HyperlinkButtonBackground  HyperlinkButtonForeground  ControlContentThemeFontSize   Wpf.Ui.Controls.Button  ControlCornerRadius �� 
TargetType ��Name �� Padding   CornerRadius �� CornerRadius � ColumnDefinitions	 ��Name
   Icon ��Name �� Triggers ]� 
Conditions �� Property �� Value �� 
TargetName (	 $HyperlinkButtonBackgroundPointerOver (
 $HyperlinkButtonForegroundPointerOver $  HyperlinkButtonBackgroundPressed $  HyperlinkButtonForegroundPressed P� Value P� Property % !HyperlinkButtonBackgroundDisabled % !HyperlinkButtonForegroundDisabled �� BasedOn�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls  %�  &	        &	      &	      &	 /     7  7   7  '   y    7  u�	11,5,11,6q�u�0q�u�0,0,8,0q��� "    .���  �@�FocusVisualStyle# �  ��  � ��# �  ��  � ��# �  ��  � ��8   ��  � ��
 ���� ��  � ��8  ��  � ��$
 Left=���  � ��$ Center=���  � ��$ Center=���  � ��$ Center=���  � ��# �  ��  � ��$ Normal"���  �@ CornerRadius# �  ��  �@ 
Appearance$ 	Secondary=���  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� "   .R��� $ ContentBorder��#��z��#��z��#��z��#��z��# z��#��z��#��z��#��z��#��z��#��z��# z ._�� #��z��#��z�� �� $
��Auto ��� $��* �.+��� $	 ControlIcon��$��0��8�� $��Center=�#��z
 $|�False��#��z���� $ ContentPresenter��$��1��#��z��#��z�� ]�  ��  �@J�IsMouseOver$
 True����  � ��$ False��./��� $ ContentBorder��# � 	  � ���� $ ContentPresenter��# � 
  � ��]�  ��  �@J�IsMouseOver$
 True����  � ��$
 True��./��� # �   � ���� $ ControlIcon��# �   � ���� $ ContentPresenter��# �   � ��P�  Z�  �@ Icon. ��� $ ControlIcon�� � ��$ 0q��� $ ControlIcon�� � y�$ 	Collapsed=�P�  � {�$ False��. ��� # �   � ���� $ ControlIcon��# �   � ���� $ ContentPresenter��# �   � ���� 8   "    