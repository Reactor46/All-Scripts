   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  O Kmscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e0891&clr-namespace:System;assembly=mscorlibSystem W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089 &  "AutoSuggestBoxBorderThemeThickness , (AutoSuggestBoxAccentBorderThemeThickness   AutoSuggestBoxLeftIconMargin ! AutoSuggestBoxRightIconMargin # AutoSuggestBoxClearButtonMargin $  AutoSuggestBoxClearButtonPadding # AutoSuggestBoxClearButtonHeight % !AutoSuggestBoxClearButtonIconSize + 'DefaultAutoSuggestBoxItemContainerStyle  �� 
TargetType �� Property 	 ListViewItemForeground �� Value 
 ControlCornerRadius �� 
TargetType ��Name �� Padding �� CornerRadius ��Name  ListViewItemPillFillBrush �� RadiusX	 �� RadiusY
 �� Triggers P� Property P� Value �� 
TargetName % !ListViewItemBackgroundPointerOver ]� 
Conditions �� Property �� Value N� EnterActions �� 
TargetName �� TargetProperty W� From	 W� To g� Duration N� ExitActions   DefaultUiAutoSuggestBoxStyle$    Wpf.Ui.Controls.AutoSuggestBox  TextControlThemePadding  ControlContentThemeFontSize'   !Wpf.Ui.Controls.IconSourceElement   
IconSource&    Wpf.Ui.Controls.SymbolIconSource   Symbol   Wpf.Ui.Controls.TextBox  Name    Icon   Icon   IconPlacement    PlaceholderText   PlaceholderText  �� RelativeSource! �� Path" �� Mode# �� UpdateSourceTrigger$ �Name% � ActualWidth& � AllowsTransparency'    IsSuggestionListOpen  FlyoutBackground  FlyoutBorderBrush( J� SnapsToDevicePixels) ��Name*    MaxSuggestionListHeight+ �� DisplayMemberPath-  'System.Windows.Input.KeyboardNavigation,   DirectionalNavigation- �� SelectionMode. K� IsItemsHost/ %� IsVirtualizing0 %� VirtualizationMode1 �� BasedOn�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls  1system&clr-namespace:System;assembly=mscorlib  %�  &	        &	      &	 "     &	 4     &	 F     &	 W     &	 h     &	 t     &	 �     &	 =    7   7  '   �    7  u�1,1,1,0q�u�0,0,0,1q�u�10,8,0,0q�u�0,8,10,0q�u�0,5,4,0q�u�0,0,0,0q�X�24Q�X�14Q��� "  ��.���  � ��# � 	 ��  � ��
 ���� ��  �@��CornerRadius# � 
 ��  � ��$ 0,0,0,2q���  � ��$ 4q���  �@�OverridesDefaultStyle$
 True����  � �� �� " ��.R��� $ ContentBorder��$��4,0,4,0q�$
 6,10q�
������ $��0q�$ 6��._�� .+��� $��12,0,0,0q��� $ ActiveRectangle��$��3��$��18��$��0q�$
��Left=�$��Center=�#���  $ 2��$	 2��$y�	Collapsed=�
 P�  �@��
IsSelected$
 True��. ��� $ ContentBorder��# �   � ���� $ ActiveRectangle�� � y�$ Visible=�]�  ��  �@��
IsSelected$ False����  �@J�IsMouseOver$
 True��/��� $ ContentBorder��# �   � �� �� .���� .��W� $ ContentBorder��$3 -(Border.Background).(SolidColorBrush.Opacity)� 0.0 1.0$ 00:00:00.167D� �� .���� .��W� $ ContentBorder��$3 -(Border.Background).(SolidColorBrush.Opacity)� 1.0 0.0$ 00:00:00.167D��� "    .���  �@  MaxSuggestionListHeight$	 240Q���  � ��$ Stretch=���  � ��$ Center=���  � ��$ Stretch=���  � ��$	 Top=���  � ��# �  ��  �@��CornerRadius# � 
 ��  � ��8   ��  � ��# �  ��  �@J�SnapsToDevicePixels$
 True����  � ��8  ��  �@�OverridesDefaultStyle$
 True����  �@  Icon      $ Search24=���  � �� �� "   .R�� .+�  $ PART_TextBox��$��0��# z $ Right=�# z ����   �� *TemplatedParent+$
! Text�$" OneWayToSource=�$# PropertyChanged=�� $$ PART_SuggestionsPopup��#��z% $��0q�$��Stretch=�$
& True��$|�False��#��z' $��Bottom=�$��Slide=�.���� $��0q�$ 0,6,0,6q�$��Stretch=�#���  #���  $��1q�$ 8��$
( True��._��� $) PART_SuggestionsList��#��z* #+ z+ #��z��#��z��#��z��#��z��$, Cycle=�$- Single=����� .R�$� $
. True��$
/ True��$0 	Recycling=��� 81   "    