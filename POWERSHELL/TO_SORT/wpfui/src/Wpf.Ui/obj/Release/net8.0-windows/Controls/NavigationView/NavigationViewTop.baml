   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  5clr-namespace:Wpf.Ui.ConvertersWpf.Ui.Converters  W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089  �� MergedDictionaries �� Source   TopNavigationViewTemplate �� 
TargetType$    Wpf.Ui.Controls.NavigationView � RowDefinitions �Name � ColumnDefinitions   Wpf.Ui.Controls.Button  Name    IsBackEnabled  BasePaneButtonStyle - )BackButtonVisibilityToVisibilityConverter �� 	Converter	   Icon    Wpf.Ui.Controls.SymbolIcon
   Symbol    
PaneHeader ��Name-  'System.Windows.Input.KeyboardNavigation   DirectionalNavigation K� IsItemsHost �� Orientation    AutoSuggestBox    
PaneFooter    Header    HeaderVisibility4   .Wpf.Ui.Controls.NavigationViewContentPresenter  Name    
Transition   
Transition    TransitionDuration   TransitionDuration    ContentOverlay ( $TopCompactNavigationViewItemTemplate(   "Wpf.Ui.Controls.NavigationViewItem ��Name ��Name   Icon   NavigationViewItemForeground   	InfoBadge ��Name . *NavigationViewSelectionIndicatorForeground J� Opacity  �� RadiusX! �� RadiusY" �Name# � AllowsTransparency$ � HorizontalOffset% � 	StaysOpen& � VerticalOffset' �� RelativeSource( �� Mode) � PlacementTarget* �� ElementName+ �� Padding  FlyoutBorderBrush  FlyoutBackground, �� CornerRadius- J� SnapsToDevicePixels. J� Effect3  -System.Windows.Media.Effects.DropShadowEffect/   
BlurRadius0   	Direction1   Opacity2   ShadowDepth3   	MenuItems4 � 	Resources " NavigationViewItemDefaultStyle5 �� BasedOn6 �� 
TargetType7 �� Property8 �� Value9 �� Triggers: P� Property; P� Value< �� 
TargetName (	 $NavigationViewItemBackgroundSelected +
 'NavigationViewItemBackgroundPointerOver ' #NavigationViewItemBackgroundPressed= N� EnterActions> �� 
TargetName? �� TargetProperty@ W� From	A W� ToB g� DurationC N� ExitActions�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls  .
convertersclr-namespace:Wpf.Ui.Converters    �� $l fpack://application:,,,/Wpf.Ui;component/Controls/NavigationView/NavigationViewBasePaneButtonStyle.xaml?��� $i cpack://application:,,,/Wpf.Ui;component/Controls/NavigationView/NavigationViewItemDefaultStyle.xaml?��� $b \pack://application:,,,/Wpf.Ui;component/Controls/NavigationView/NavigationViewConstants.xaml?�%h  &	        7  7  &	 v    7  �� "   .R��  �� $
��Auto ��� $��* �.+�� $ PaneGrid��$��0�� �� $
��Auto ��� $
��Auto ��� $��* ��� $
��Auto ��� $
��Auto ��� $
��Auto �.+�  $ PART_BackButton��$��0��#{�z 8��  y��� *IsBackButtonVisible+8  	   $
 ArrowLeft48=��� $��1��$��0q�#��z �� $  PART_MenuItemsItemsControl��$��2��$|�False��$ 	Contained=�$��Disabled=�$��Disabled=����� .R��� $��0q�$
 True��$ 
Horizontal=��� $��3��#��z �� $��4��$��0q�#��z �� $&  PART_FooterMenuItemsItemsControl��$��5��$|�False��$ 	Contained=�$��Disabled=�$��Disabled=����� .R��� $��0q�$
 True��$ 
Horizontal=�� $��1�� �� $
��Auto ��� $��* �.+��� $��0��$��0q�#��z #y�z   $) #PART_NavigationViewContentPresenter��$��1��#��z��# z # z �� $��0��$��2��$��0q�#��z �� "  .R��  �� $
��Auto ��� $
��Auto �.+��� $ 
MainBorder��._�� .+�� $��8,8,8,8q�$��Stretch=�$��Stretch=� �� $
��Auto ��� $��* ��� $
��Auto �.+��� $ IconContentPresenter��$��0��$��0,0,6,0q�$��Center=�$��Center=�#��z $��16(�#���  �� $ ElementContentPresenter��$��1��$
��Left=�$��Center=�#��z��$��14(�#��z��� $ PART_ChevronGrid��$��2��$��0��$y�	Collapsed=��� $ PART_InfoBadge��$��2��$��2q�$��Right=�$��Center=�#��z �� $ ActiveRectangle��$��0��$��16��$��3��$��0,0,0,-4q�$��Center=�$��Bottom=�#���  $	 0.0Q�$  2��$! 2��� $" SubMenuPopup��$��1��$
# True��$|�False��$	$ -12��$��Bottom=�$
��None=�$% False��$& 1������ *
IsExpanded+' �� *TemplatedParent+$( TwoWay=�) �� $* Border��.���� $ SubMenuBorder��$��
12,0,12,18q�$+ 0,3,0,3q�#���  #���  $��1q�$, 8��$
- True��z�T� .   $/ 20Q�$	0 270Q�$
1 0.25Q�$2 6Q�._��� $ SubMenuItemsPresenter��$|�False��#��z3 $ 	Contained=�$	 0.0Q�$��Disabled=�$��Disabled=����� .R��� $��0q�$
 True��4 ���� 85   '        "6  .��� 7 � ��8 �� "  .R��� $ 
MainBorder��$��25��$��0q�$+ 0q�$��Stretch=�$��1q�$, 4��._�� $��0q�$��Stretch=� �� $��3 ��� $��15 ��� $��* �$	��120��.+��� $ ActiveRectangle��$��0��$��3��$��16��$��0q�$
��Left=�$��Center=�#���  $	 0.0Q�$  2��$! 2���� $ IconContentPresenter��$��1��$��0,0,0,0q�$��Center=�$��Center=�#��z $|�False��$��16(�#���  �� $ ElementContentPresenter��$��2��$��5,0,0,0q�$��Stretch=�$��Center=�#��z��$��14(�#���  9 P� : �@ IsActive$
; True��. ��� $< ActiveRectangle��7 �@J�Opacity$	8 1.0Q��� $< 
MainBorder��#8 � 	 7 � ��P� : �@J�IsMouseOver$
; True��. ��� $< 
MainBorder��#8 � 
 7 � ��P� : � ��$
; True��. ��� $< 
MainBorder��#8 �  7 � ��9 P� : �@ HasMenuItems$
; True��. ��� $< PART_ChevronGrid��7 � y�$8 Visible=��� $< PART_ChevronGrid��7 � ��$8 20��P� : �@ 
IsExpanded$
; True��= �� .���� .��W� $> SubMenuItemsPresenter��$? (UIElement.Opacity)�@ 0.0A 1.0$B 	0:0:0.167D�W� $> ChevronIcon��$7? 1(Control.RenderTransform).(RotateTransform.Angle)�A -180$B 00:00:00.167D�C �� .���� .��W� $> SubMenuItemsPresenter��$? (UIElement.Opacity)�@ 1.0A 0.0$B 	0:0:0.167D�W� $> ChevronIcon��$7? 1(Control.RenderTransform).(RotateTransform.Angle)�A 0$B 00:00:00.167D�P� : �@ IsActive$
; True��. ��� $< ActiveRectangle��7 �@J�Opacity$	8 1.0Q�P� ; Z� : �@ Icon. ��� $< IconContentPresenter��7 � y�$8 	Collapsed=�P� ; Z� : � ��. ��� $< IconContentPresenter��7 � ��$8 0q��� $< ElementContentPresenter��7 � y�$8 	Collapsed=�P� : � ��$;  ��. ��� $< IconContentPresenter��7 � ��$8 0q��� $< ElementContentPresenter��7 � y�$8 	Collapsed=�