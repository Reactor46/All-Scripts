   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089   ButtonPadding  ButtonBorderThemeThickness  ButtonIconMargin  DefaultButtonStyle  �� 
TargetType �� Property " DefaultControlFocusVisualStyle �� Value  ButtonBackground  ButtonForeground  ControlElevationBorderBrush  ControlContentThemeFontSize 	 ControlCornerRadius �� 
TargetType ��Name �� Padding �� CornerRadius ��Name �� Triggers	 P� Property
 P� Value 
 ButtonBackgroundPointerOver �� 
TargetName  ButtonForegroundPointerOver  ButtonBackgroundDisabled  ButtonBorderBrushDisabled  ButtonForegroundDisabled  ButtonBackgroundPressed  ButtonBorderBrushPressed  ButtonForegroundPressed  DefaultUiButtonStyle    Wpf.Ui.Controls.Button    CornerRadius � ColumnDefinitions    Icon ]� 
Conditions �� Property �� Value �� RelativeSource �� Triggers  AccentButtonBackground % !AccentButtonBackgroundPointerOver ! AccentButtonBackgroundPressed  AccentButtonForeground % !AccentButtonForegroundPointerOver % !AccentControlElevationBorderBrush " AccentButtonBorderBrushPressed  ControlStrongFillColorDark �� Color  TextFillColorLightPrimary  TextFillColorLightSecondary  ControlStrongFillColorLight  TextFillColorDarkPrimary  TextFillColorDarkSecondary   PaletteLightBlueColor �� Opacity ! PaletteRedColor " PaletteGreenColor # PaletteOrangeColor �� BasedOn�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls  %)  &	        &	      &	      &	 /     7  7   &	 �    7  7   7  '�� �    7  '   �    7  u�	11,5,11,6q�u�1q�u�0,0,8,0q��� "  ��.���  �@�FocusVisualStyle# �  ��  � ��# �  ��  � ��# �  ��  � ��# �  ��  � ��8   ��  � ��8  ��  � ��$
 Left=���  � ��$ Center=���  � ��$ Center=���  � ��$ Center=���  � ��# �  ��  � ��$ Normal"���  �@��CornerRadius# � 	 ��  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� " ��.R��� $ ContentBorder��#��z��#��z��#��z��#��z��# z��#��z��#��z��#��z��#��z��#��z��# z ._��� $ ContentPresenter��#��z��#��z��#��z��#��z��#��z�� P� 	 �@J�IsMouseOver$

 True��. ��� # � 
  � ���� # �   � ���� $ ContentPresenter��# �   � ��P� 	 � {�$
 False��. ��� # �   � ���� # �   � ���� $ ContentPresenter��# �   � ��P� 	 � ��$

 True��. ��� # �   � ���� # �   � ���� $ ContentPresenter��# �   � ���� "    .���  �@�FocusVisualStyle# �  ��  � ��# �  ��  �@  MouseOverBackground# � 
 ��  �@  PressedBackground# �  ��  � ��# �  ��  �@  PressedForeground# �  ��  � ��# �  ��  �@  PressedBorderBrush# �  ��  �@  MouseOverBorderBrush# �  ��  � ��8   ��  � ��8  ��  � ��$
 Left=���  � ��$ Center=���  � ��$ Center=���  � ��$ Center=���  � ��# �  ��  � ��$ Normal"���  �@  CornerRadius# � 	 ��  �@  
Appearance$ 	Secondary=���  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� "   .R��� $ ContentBorder��#��z��#��z��#��z��#��z��#��z��#��z��#��z��# z ._��� # z��#��z��#��z��# z ._�� #��z��#��z�� �� $
��Auto ��� $��* �.+��� $ ControlIcon��$��0��8�� $��Center=�#��z #��z��$|�False��#��z��#��z���� $ ContentPresenter��$��1��$��Center=�#��z��#��z�� ]�  ��  �@J�IsMouseOver$
 True����  � ��$ False��./��� $ ContentBorder�� �� *MouseOverBackground+ �� *TemplatedParent+ � ���� $ ContentBorder�� �� *MouseOverBorderBrush+ �� *TemplatedParent+ � ��]�  ��  �@J�IsMouseOver$
 True����  � ��$
 True��./��� $ ContentBorder�� �� *PressedBackground+ �� *TemplatedParent+ � ���� $ ContentBorder�� �� *PressedBorderBrush+ �� *TemplatedParent+ � ���� $ ContentPresenter�� �� *PressedForeground+ �� *TemplatedParent+ � ���� $ ControlIcon�� �� *PressedForeground+ �� *TemplatedParent+ � ��P� 	 � {�$
 False��. ��� $ ContentBorder��# �   � ���� $ ContentBorder��# �   � ���� $ ControlIcon��# �   � ���� $ ContentPresenter��# �   � ��P� 
 Z� 	 � ��. ��� $ ControlIcon�� � ��$ 0q�P� 	 � ��$
  ��. ��� $ ControlIcon�� � ��$ 0q�P� 
 Z� 	 �@  Icon. ��� $ ControlIcon�� � ��$ 0q��� $ ControlIcon�� � y�$ 	Collapsed=� P� 	 �@  
Appearance$
 Transparent=�. ���  � ��
 ���� P� 	 �@  
Appearance$
 Primary=�. ���  � ��# �  ��  �@  MouseOverBackground# �  ��  �@  PressedBackground# �  ��  � ��# �  ��  �@  PressedForeground# �  ��  � ��# �  ��  �@  PressedBorderBrush# �  P� 	 �@  
Appearance$

 Dark=�. ���  � �� �� # �  ��  �@  MouseOverBackground �� $ 	#62000000����  �@  PressedBackground �� $ 	#52000000����  � �� �� # �  ��  �@  PressedForeground �� # �  P� 	 �@  
Appearance$
 Light=�. ���  � �� �� # �  ��  �@  MouseOverBackground �� $ 	#D3FFFFFF����  �@  PressedBackground �� $ 	#F3FFFFFF����  � �� �� # �  ��  �@  PressedForeground �� # �  P� 	 �@  
Appearance$

 Info=�. ���  � �� �� # �   ��  �@  MouseOverBackground �� $	 0.9Q�# �   ��  �@  PressedBackground �� $	 0.7Q�# �   ��  � �� �� # �  P� 	 �@  
Appearance$
 Danger=�. ���  � �� �� # � ! ��  �@  MouseOverBackground �� $	 0.9Q�# � ! ��  �@  PressedBackground �� $	 0.7Q�# � ! P� 	 �@  
Appearance$
 Success=�. ���  � �� �� # � " ��  �@  MouseOverBackground �� $	 0.9Q�# � " ��  �@  PressedBackground �� $	 0.7Q�# � " P� 	 �@  
Appearance$
 Caution=�. ���  � �� �� # � # ��  �@  MouseOverBackground �� $	 0.9Q�# � # ��  �@  PressedBackground �� $	 0.7Q�# � # �� 8   "  ���� 8   "    