   M S B A M L   `   `   `  ���� W  SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089  �� 
TargetType �� Property   TabViewForeground �� Value  TabViewBackground  TabViewBorderBrush �� 
TargetType-   'System.Windows.Input.KeyboardNavigation    TabNavigation � RowDefinitions ��Name K� ZIndex K� IsItemsHost	    TabIndex
 ��Name �� CornerRadius    DirectionalNavigation ��Name'  !System.Windows.VisualStateManager   VisualStateGroups%  System.Windows.VisualStateGroup   States   System.Windows.VisualState  TabViewItemHeaderBackground " DefaultControlFocusVisualStyle �Name �� Padding   
Storyboard �� 
TargetName �� TargetProperty W� From	 W� To g� Duration �� Triggers P� Property P� Value �� 
TargetName ' #TabViewItemHeaderBackgroundSelected " TabViewSelectedItemBorderBrush ! TabViewItemForegroundSelected�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation      :x,http://schemas.microsoft.com/winfx/2006/xaml      %6  '��       '�� B    �� "  ��.���  � ��# �   ��  � ��# �  ��  � ��# �  ��  �@�OverridesDefaultStyle$
 True����  �@J�SnapsToDevicePixels$
 True����  � �� �� " ��.R�� $ Local=� �� $
��Auto ��� $��* �.+��� $ HeaderPanel��$��0��$��0q�$ 1��
������ $
 True��$	 1���� $
 Border��$��1��#��z��#��z��$��0,1,0,0q�$ 0,4,4,4��$ 	Contained=�$	 2��$ Local=�._��� $ PART_SelectedContentHost��$��0q�$��Stretch=�$��Stretch=�$��SelectedContent��#��z��   CommonStates��.   Disabled���� "  ��.���  � ��# �  ��  � ��
 ���� ��  � ��$ Stretch=���  � ��$ Stretch=���  � ��$ Stretch=���  � ��$ Stretch=���  �@�FocusVisualStyle# �  ��  �@  	IsTabStop$
 True����  � |�$
 True����  �@�OverridesDefaultStyle$
 True����  �@J�SnapsToDevicePixels$
 True����  � �� �� " ��.R�� $
 Root��.+��� $
 Border��$	��180��$��36��$��0q�$ 6q�#��z��#��z��$��1,1,1,0q�$ 8,8,0,0��._��� $ ContentSite��$��0q�$
��Left=�$��Center=�$��Header��$
��True��   SelectionStates��.   
Unselected��  Selected��  CommonStates��.   
Normal��  	MouseOver��  Disabled��. �� .��W� $ Border��$2 ,(Panel.Background).(SolidColorBrush.Opacity)� 0.0 0.5$ 0:0:.16D� P�  �@��
IsSelected$
 True��. ���  �@K�ZIndex$	 100���� $ Border��# �   � ���� $ Border��# �   � ���� # �   � ��