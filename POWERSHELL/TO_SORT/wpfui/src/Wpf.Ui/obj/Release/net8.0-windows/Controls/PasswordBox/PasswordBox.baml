   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  O Kmscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e0891&clr-namespace:System;assembly=mscorlibSystem W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089 #  PasswordBoxBorderThemeThickness ) %PasswordBoxAccentBorderThemeThickness  PasswordBoxLeftIconMargin  PasswordBoxRightIconMargin  PasswordBoxButtonMargin  PasswordBoxButtonPadding  PasswordBoxButtonHeight  PasswordBoxButtonIconSize ! DefaultPasswordBoxContextMenu  m� Command 	 DefaultPasswordBoxStyle �� 
TargetType �� Property "
 DefaultControlFocusVisualStyle �� Value  TextControlForeground  TextControlBackground # TextControlElevationBorderBrush  ControlContentThemeFontSize  TextControlThemeMinHeight  TextControlThemeMinWidth  TextControlThemePadding  ControlCornerRadius �� 
TargetType ��Name �� Padding �� CornerRadius {�Name " ControlStrokeColorDefaultBrush	 �� Triggers
 P� Property P� Value �� 
TargetName ! TextControlFocusedBorderBrush   TextControlBackgroundFocused ]� 
Conditions �� Property �� Value $  TextControlBackgroundPointerOver ! TextControlBackgroundDisabled " TextControlBorderBrushDisabled ! TextControlForegroundDisabled  DefaultUiPasswordBoxStyle!    Wpf.Ui.Controls.PasswordBox   Wpf.Ui.Controls.TextBox � ColumnDefinitions ��Name   Icon ��Name �� Padding $  TextControlPlaceholderForeground   PlaceholderText   Wpf.Ui.Controls.Button  Name   
Appearance � Cursor  TextControlButtonForeground �� Path �� RelativeSource   Icon    Wpf.Ui.Controls.SymbolIcon   Wpf.Ui.Controls.FontIcon   FontSize   Symbol �� BasedOn�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls  1system&clr-namespace:System;assembly=mscorlib  %  &	        &	      &	 "     &	 4     &	 F     &	 W     &	 h     &	 t     &	 �     &		 �     7   7  &	 �    7   7  7  7  7  7  7  7  7  7  7  7  7  7  'G� )    7 	 '   8    7  u�1,1,1,0q�u�0,0,0,1q�u�10,8,0,0q�u�0,8,10,0q�u�0,5,4,0q�u�0,0,0,0q�X�24Q�X�14Q��� .@�m� $  ApplicationCommands.Paste���� " G�.���  �@�FocusVisualStyle# � 
 ��  �@�ContextMenu# �  ��  � ��# �  ��  �@G�
CaretBrush# �  ��  � ��# �  ��  � ��# �  ��  � ��8   ��  � ��# �  ��  � ��$ Stretch=���  � ��$ Center=���  � ��$ Stretch=���  � ��$	 Top=���  � ��# �  ��  � ��# �  ��  � ��# �  ��  �@��CornerRadius# �  ��  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� " G�.R�� #��z��#��z��.+��� $ ContentBorder��#��z��#��z��$ 0q�$��Stretch=�$��Stretch=�#��z��#��z��#��z��# z ._�� #��z��#��z��#��z��.+�{� $ PART_ContentHost��$��0q�$��Center=�#��z���� $ AccentBorder��$��Stretch=�$��Stretch=�#���  8�� # z 	 P� 
 �@J�	IsFocused$
 True��. ��� $ AccentBorder�� � ��$ 0,0,0,2q��� $ AccentBorder��# �   � ���� $ ContentBorder��# �   � ��]�  ��  � {�$
 True����  �@J�IsMouseOver$
 True����  �@J�	IsFocused$ False��./��� $ ContentBorder��# �   � ��P� 
 � {�$
 True��. ���  �@�Cursor$ IBeam��P� 
 � {�$ False��. ��� $ ContentBorder��# �   � ���� $ ContentBorder��# �   � ���� $ AccentBorder��# �   � ���� # �   � ���� "   .���  �@�FocusVisualStyle# � 
 ��  �@�ContextMenu# �  ��  � ��# �  ��  �@��
CaretBrush# �  ��  � ��# �  ��  � ��# �  ��  � ��8   ��  � ��# �  ��  � ��$ Stretch=���  � ��$ Center=���  � ��$ Stretch=���  � ��$	 Top=���  � ��# �  ��  � ��# �  ��  � ��# �  ��  �@��CornerRadius# �  ��  �@ ClearButtonEnabled$
 True����  �@  RevealButtonEnabled$
 True����  �@ IconPlacement$
 Left=���  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� "   .R�� #��z��#��z��.+��� $ ContentBorder��#��z��#��z��$ 0q�$��Stretch=�$��Stretch=�#��z��#��z��#��z��# z ._�� #��z��#��z�� �� $
��Auto ��� $��* ��� $
��Auto ��� $
��Auto ��� $
��Auto �.+��� $ ControlIconLeft��$��0��8�� $	��Top=�#��z $��16(�#��z��� $��1��#��z��.+�{� $ PART_ContentHost��$��0q�$��Center=�#��z���� $ PlaceholderTextBox��$��0q�$	 1,0q�$	��Top=�#���  #��z   $ ClearButton��$��2��8�� 8�� 8�� 8�� $��Center=�$	��Top=�$��Center=�$��Center=�$ 	Secondary=�
������ 
������ $��clear��$ Arrow��#���  ���� $ TemplateButtonCommand� �� *TemplatedParent+   8  $ 	Dismiss24=�  $ RevealButton��$��3��8�� 8�� 8��	 8��
 $��Center=�$	��Top=�$��Center=�$��Center=�$ 	Secondary=�
������ 
������ $��reveal��$ Arrow��#���  ���� $ TemplateButtonCommand� �� *TemplatedParent+   8  $ Eye24=��� $ ControlIconRight��$��4��8�� $	��Top=�#��z $��16(�#��z���� $ AccentBorder��$��Stretch=�$��Stretch=�#���  8�� # z 	 P� 
 �@ PlaceholderEnabled$ False��. ��� $ PlaceholderTextBox�� � y�$ 	Collapsed=�P� 
 �@ ShowClearButton$ False��. ��� $ ClearButton�� � y�$ 	Collapsed=��� $ ClearButton�� � ��$ 0q�P� 
 �@ ClearButtonEnabled$ False��. ��� $ ClearButton�� � y�$ 	Collapsed=��� $ ClearButton�� � ��$ 0q�P� 
 �@  RevealButtonEnabled$ False��. ��� $ RevealButton�� � y�$ 	Collapsed=��� $ RevealButton�� � ��$ 0q�P� 
 �@  IsPasswordRevealed$
 True��. ��� $ RevealButton�� �@ Icon   $ EyeOff24=�P� 
 �@ IconPlacement$
 Left=�. ��� $ ControlIconRight�� � y�$ 	Collapsed=��� $ ControlIconRight�� � ��$ 0q�P� 
 �@ IconPlacement$ Right=�. ��� $ ControlIconLeft�� � y�$ 	Collapsed=��� $ ControlIconLeft�� � ��$ 0q�P�  Z� 
 �@ Icon. ��� $ ControlIconRight�� � y�$ 	Collapsed=��� $ ControlIconRight�� � ��$ 0q��� $ ControlIconLeft�� � y�$ 	Collapsed=��� $ ControlIconLeft�� � ��$ 0q�P� 
 �@J�	IsFocused$
 True��. ��� $ AccentBorder�� � ��$ 0,0,0,2q��� $ AccentBorder��# �   � ���� $ ContentBorder��# �   � ��]�  ��  � {�$
 True����  �@J�IsMouseOver$
 True����  �@J�	IsFocused$ False��./��� $ ContentBorder��# �   � ��P� 
 � {�$
 True��. ���  �@�Cursor$ IBeam��P� 
 � {�$ False��. ��� $ ContentBorder��# �   � ���� $ ContentBorder��# �   � ���� $ AccentBorder��# �   � ���� $ ControlIconLeft��# �   � ���� $ ControlIconRight��# �   � ���� $ PlaceholderTextBox��# �   � ���� # �   � ����  �@ ClearButtonEnabled$ False����  �@  RevealButtonEnabled$ False���� 8   " G��� 8   "   