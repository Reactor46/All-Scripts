   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089  �� MergedDictionaries �� Source )  %LeftCompactNavigationViewItemTemplate �� 
TargetType(    "Wpf.Ui.Controls.NavigationViewItem � RowDefinitions ��Name �� Padding �� CornerRadius �� 
TargetType �� Property	 �� Value
 �� Triggers P� Property P� Value + 'NavigationViewItemBackgroundPointerOver � ColumnDefinitions ��Name    Icon   NavigationViewItemForeground ��Name . *NavigationViewSelectionIndicatorForeground J� Opacity �� RadiusX �� RadiusY    	InfoBadge �Name    Wpf.Ui.Controls.SymbolIcon  Name ! NavigationViewItemChevronSize   Wpf.Ui.Controls.FontIcon   FontSize!   Wpf.Ui.Controls.IconElement   
Foreground J� RenderTransformOrigin   Symbol �� Angle ��Name    	MenuItems-  'System.Windows.Input.KeyboardNavigation   DirectionalNavigation K� IsItemsHost  � 	Resources " NavigationViewItemDefaultStyle! �� BasedOn  NullToVisibilityConverter" �� 	Converter# �� Triggers$ �� 
TargetName ( $NavigationViewItemBackgroundSelected ' #NavigationViewItemBackgroundPressed% N� EnterActions& �� 
TargetName' �� TargetProperty( W� From	) W� To* g� Duration+ N� ExitActions, ]� 
Conditions- �� Property. �� Value +	 'NavigationViewItemForegroundPointerOver 
 LeftNavigationViewTemplate$   Wpf.Ui.Controls.NavigationView/   OpenPaneLength   Wpf.Ui.Controls.Button0  Name1   IsBackEnabled  BasePaneButtonStyle - )BackButtonVisibilityToVisibilityConverter2   Icon  BoolToVisibilityConverter3   	PaneTitle4   
PaneHeader5   AutoSuggestBox6 �� RelativeSource7 �� Path)   #Wpf.Ui.Controls.DynamicScrollViewer8   
PaneFooter9   FrameMargin # NavigationViewContentBackground ( $NavigationViewContentGridBorderBrush:   Header;   HeaderVisibility4   .Wpf.Ui.Controls.NavigationViewContentPresenter<  Name=   
Transition>   
Transition?   TransitionDuration@   TransitionDurationA   ContentOverlay'	  !System.Windows.VisualStateManagerB 	  VisualStateGroups%
  System.Windows.VisualStateGroupC 
  NameD 
  States   System.Windows.VisualStateE   NameF   
StoryboardG g� AccelerationRatio�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls    �� $l fpack://application:,,,/Wpf.Ui;component/Controls/NavigationView/NavigationViewBasePaneButtonStyle.xaml?��� $i cpack://application:,,,/Wpf.Ui;component/Controls/NavigationView/NavigationViewItemDefaultStyle.xaml?��� $b \pack://application:,,,/Wpf.Ui;component/Controls/NavigationView/NavigationViewConstants.xaml?�%  &	        7  7  7  &	
 D    7  7  7  7  7  �� "   .R��  �� $
��Auto ��� $
��Auto �.+��� $ 
MainBorder��$��0��$	��310��$��40��$��0q�$ 0q�#��z��$��1q�# z ���� " ��.���  � ��
	 ���� 
 P�  �@J�IsMouseOver$
 True��. ���  � ��#	 �  ._�� 
������  �� $��40 ��� $��* ��� $
��Auto ��� $
��Auto �.+�� $��0��.+��� $ IconContentPresenter��$��-1,0,0,0q�$��Center=�$��Center=�#��z $|�False��$��16(�#���  �� $ ActiveRectangle��$��0��$��3��$��16��$��0q�$
��Left=�$��Center=�#���  $	 0.0Q�$ 2��$ 2���� $ ElementContentPresenter��$��1��$
��Left=�$��Center=�#��z��$��14(�#��z���� $ PART_InfoBadge��$��2��$��0q�#��z � $ PART_ChevronGrid��$��3��$��0��$y�	Collapsed=�.+�  $ ChevronIcon��$��0q�$��Center=�$��Center=�8   # �  $ 0.5, 0.5&�$ ChevronDown24=�z��� $ 0Q��� $ MenuItemsPresenter��$��1��$|�False��#��z $ 	Contained=�$	 0.0Q�$��Disabled=�$��Disabled=�$y�	Collapsed=����� .R��� $��0q�$
 True��  ���� 8!  '         "   .���  � ��	 �� "   .R��� $ 
MainBorder��$��40��$��Stretch=�$��1q�$ 4��._�� $��30,0,0,0q�$��Stretch=� �� $
��Auto ��� $
��Auto ��� $��* �$	��180���� $
��Auto �.+��� $ ActiveRectangle��$��0��$��3��$��16��$��0q�$
��Left=�$��Center=�#���  $	 0.0Q�$ 2��$ 2���� $ IconContentPresenter��$��1��$��6,0,3,0q�$��Center=�$��Center=�#��z $|�False��$��18(�#���  y��� *Icon+8"  �� $ ElementContentPresenter��$��2��$��10,0,0,0q�$��Stretch=�$��Center=�#��z��$��14(�#��z���� $��3��#��z # P�  �@  IsActive$
 True��. ��� $$ ActiveRectangle�� �@J�Opacity$		 1.0Q��� $$ 
MainBorder��#	 �   � ��P�  �@J�IsMouseOver$
 True��. ��� $$ 
MainBorder��#	 �   � ��P�  � ��$
 True��. ��� $$ 
MainBorder��#	 �   � ��# P�  �@  
IsPaneOpen$
 True��. ��� $$ PART_InfoBadge�� � ��$	 2��P�  �@  
IsPaneOpen$ False��. ��� $$ PART_InfoBadge�� � ��$	 0���� $$ PART_InfoBadge�� � ��$		 Top=��� $$ PART_InfoBadge�� � ��$	 Right=�P�  Z�  �@  	InfoBadge. ��� $$ PART_InfoBadge�� � ��$	 0q��� $$ PART_InfoBadge�� � y�$	 	Collapsed=�P�  �@  HasMenuItems$
 True��. ��� $$ PART_ChevronGrid�� � y�$	 Visible=��� $$ PART_ChevronGrid�� � ��$	 40��P�  �@  
IsExpanded$
 True��. ��� $$ MenuItemsPresenter�� � y�$	 Visible=�% �� .���� .��W� $& MenuItemsPresenter��$' (ItemsControl.Opacity)�( 0.0) 1.0$* 00:00:00.167D�W� $& ChevronIcon��$7' 1(Control.RenderTransform).(RotateTransform.Angle)�) -180$* 00:00:00.167D�+ �� .���� .��W� $& MenuItemsPresenter��$' (ItemsControl.Opacity)�( 1.0) 0.0$* 00:00:00.167D�W� $& ChevronIcon��$7' 1(Control.RenderTransform).(RotateTransform.Angle)�) 0$* 00:00:00.167D�P�  �@  IsActive$
 True��. ��� $$ ActiveRectangle�� �@J�Opacity$		 1.0Q��� $$ 
MainBorder��#	 �   � ��P�  Z�  �@  Icon. ��� $$ IconContentPresenter�� � y�$	 	Collapsed=�]� , �� - �@J�IsMouseOver$
. True���� - �@  IsActive$. False��/��� $$ IconContentPresenter��#	 � 	  � ���� $$ ElementContentPresenter��#	 � 	  � ���� "  .R��  �� $
��Auto ��� $��* �.+�� $ PaneGrid��$��0��#��z/ $	��4,0q�.+��� $ 
PaneBorder��._��  �� $
��Auto ��� $��* ��� $
��Auto �.+�� $��0�� �� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto �.+�  $0 PART_BackButton��$��0��$��0,5,0,5q�$
��Left=�#{�z1 8��  y��� *IsBackButtonVisible+8"  2   $ ArrowLeft48=�  $0 PART_ToggleButton��$��1��$��0,0,0,5q�8�� y��� *IsPaneToggleVisible+8"  2   $ LineHorizontal320=����� $��Medium"�#��z3 �� $��2��$��0q�#��z4 �� $$ AutoSuggestBoxContentPresenter��$��3��$��0q�#��z5   $%0 PART_AutoSuggestBoxSymbolButton��$��3��$��0,0,0,6q�8�� $y�	Collapsed=�2 �� 6 �� *TemplatedParent+$7 AutoSuggestBox.Icon�  $��1��$��0,0,-4,0q�$��0,0,4,0q�$
��True��$��Disabled=�$
��Auto=�.���  �� $��* ��� $
��Auto �.+��� $  PART_MenuItemsItemsControl��$��0��$|�False��$ 	Contained=�$��Disabled=�$��Disabled=����� .R��� $��0q�$
 True���� $&  PART_FooterMenuItemsItemsControl��$��1��$��0,0,0,4q�$|�False��$ 	Contained=�$��Disabled=�$��Disabled=����� .R��� $��0q�$
 True���� $��2��$��0q�#��z8 �� $��1��#��z9 #���  #���  $��1,1,0,0q�$ 8,0,0,0��._��  �� $
��Auto ��� $��* �.+��� $��0��$��0q�#��z: #y�z;   $)< #PART_NavigationViewContentPresenter��$��1��#��z��#> z= #@ z? �� $��0��$��2��$��0q�#��zA B 
  $C 
PaneStates��.D   $E PaneOpen��.F �� .��W� $	G 0.4Q�$& PaneGrid��$' Width�( 40#) z/ $* 0:0:.16D�  $E PaneCompact��.F �� .��W� $	G 0.4Q�$& PaneGrid��$' Width�#( z/ ) 40$* 0:0:.16D�# P�  �@ 
IsPaneOpen$ False��. ��� $%$ PART_AutoSuggestBoxSymbolButton�� � y�$	 Visible=��� $$$ AutoSuggestBoxContentPresenter�� � y�$	 	Collapsed=��� $$ PART_ToggleButton��	 Z�  � ��P�  Z�  �@ AutoSuggestBox. ��� $%$ PART_AutoSuggestBoxSymbolButton�� � y�$	 	Collapsed=�P�  Z�  �@ 	PaneTitle. ��� $$ PART_ToggleButton��	 Z�  � ���� $$ PART_ToggleButton�� � ��$
	 Left=�