   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089   ToolBarButtonBaseStyle  �� 
TargetType �� Property �� Value �� Opacity  ControlFillColorDefault �� Color �� 
TargetType ��Name �� Padding �� Triggers	 P� Property
 P� Value N� EnterActions �� 
TargetName �� TargetProperty W� From	 W� To g� Duration N� ExitActions �� BasedOn-   'System.Windows.Input.KeyboardNavigation)   #Wpf.Ui.Controls.PassiveScrollViewer  Name  ToolBarThumbStyle J� SnapsToDevicePixels j� TileMode j� Viewbox j� ViewboxUnits j� Viewport j� ViewportUnits L� Drawing � Brush � Geometry  ToolBarOverflowButtonStyle  TextFillColorTertiary �� CornerRadius    Wpf.Ui.Controls.SymbolIcon!   Wpf.Ui.Controls.IconElement   
Foreground   Symbol  �� 	ClickMode! c� HasOverflowItems" d� 	IsChecked# �� Mode$ �� RelativeSource% �Name& � AllowsTransparency' � 	StaysOpen  MenuBorderColorDefaultBrush ) %SystemFillColorSolidNeutralBackground( b�Name)    DirectionalNavigation*    TabNavigation+ b� 	WrapWidth, � FocusVisualStyle- n�Name. a�Name/ K� IsItemsHost0 �� 
TargetName�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls  %:  &	        (�� y    *ToolBar.ButtonStyleKey+)7   (�� �    *ToolBar.ToggleButtonStyleKey+)7   (�� �    *ToolBar.CheckBoxStyleKey+)7   (�� �    *ToolBar.RadioButtonStyleKey+)7   (�� �    *ToolBar.TextBoxStyleKey+)&	     &	 s    'c� �    7  7  '`� �
    �� "  ��.���  � �� �� $	 0.0Q�# �  ��  � ��
 ���� ��  � ��$ 0q���  �@�Cursor$
 Hand����  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� " ��.R��� $ Border��$
 12,6q�#��z��#��z��#��z��._��� $��2q�$��Center=�$��Center=�$
��True�� P� 	 �@J�IsMouseOver$

 True�� �� .���� .��W� $ Border��$3 -(Border.Background).(SolidColorBrush.Opacity)� 0.0 1.0$ 0:0:0.16D� �� .���� .��W� $ Border��$3 -(Border.Background).(SolidColorBrush.Opacity)� 1.0 0.0$ 0:0:0.16D��� 8   "  ���� 8   "  d��� 8   "  ���� 8   "  	��� "  ��.���  �@  TabNavigation$
 None=���  �@�FocusVisualStyle Z� ��  �@J�	AllowDrop$
 True����  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� " ��.R��� $ Border��$ 2q�
������ 
������ $��0q�._�  $ PART_ContentHost��$��0q��� "  n�.���  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  �@�Cursor$ SizeAll����  � �� �� " n�.R��� 
������ $
 True��._��� $	��0,2q���L� $
 Tile=�$ 0,0,8,8�$ Absolute=�$ 0,0,4,4�$ Absolute=� I� .���  �#AAA+ �� 	=  	= � 	=  z � z  z � z  	= 	�� "  d�.���  � �� �� # �  ��  � ��
 ���� ��  �@�Cursor$
 Hand����  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� " d�.R��� $ Border��#��z��$ 0,3,3,0��$
 true��._�� .+�  $��0q�$��Bottom=�# z��$ ChevronDown20=���  P� 	 � {�$
 False��. ���  � ��$ 0���� "  c�.���  � ��
 ���� ��  � ��
 ���� ��  � ��$ 0q���  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� " c�.R��� $ Border��#��z��#��z��#��z��$ 4��._�]� .+�d� $  Press=�$��Right=�#{�z! 8��  " �� *IsOverflowOpen+$# TwoWay=�$ �� *TemplatedParent+.��� $% OverflowPopup��$
& True��$|�False��$��Bottom=�$��Slide=�$' False������ *IsOverflowOpen+$ �� *TemplatedParent+.���� $ DropDownBorder��$��
12,0,12,18q�$ 0,3,0,3q�#���  $��1q�$ 8��$
 True������ # �  ._�b� $( PART_ToolBarOverflowPanel��$��0q�$
|�True��$) Cycle=�$* Cycle=�$	+ 200Q�, Z� n� $- ToolBarThumb��$��10��8�� a� $. PART_ToolBarPanel��$��0,1,2,2q�$
/ True�� P� 	 �@c�IsOverflowOpen$

 true��. ��� $0 ToolBarThumb�� � {�$ false��P� 	 �@`�IsLocked$

 true��. ��� $0 ToolBarThumb�� � y�$ 	Collapsed=��� "  `�.���  �@`�
Background �� # �  ��  � ��$ 0q�