   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  5clr-namespace:Wpf.Ui.ConvertersWpf.Ui.Converters  U QSystem.Runtime, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a7,clr-namespace:System;assembly=System.RuntimeSystem W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089    BooleanToVisibilityConverter  DefaultDataGridFontSize  DisabledControlLightColor  DisabledControlDarkColor  DisabledForegroundColor  ControlLightColor  ControlMediumColor  ControlDarkColor  ControlMouseOverColor 	 ControlPressedColor 
 
GlyphColor  MenuPopupBrush  �� 
StartPoint �� EndPoint � Offset � Color $  ProgressBarIndicatorAnimatedFill * &DefaultDataGridCheckBoxColumnCellStyle �� 
TargetType,   &System.Windows.Controls.DataGridColumn �� 
ResourceId �� TypeInTargetAssembly&   System.Windows.Controls.DataGrid �� Property �� Value	 �� 
TargetType
 ��Name J� SnapsToDevicePixels �Name J� Opacity � Points �� Stretch �� Color'  !System.Windows.VisualStateManager   VisualStateGroups%  System.Windows.VisualStateGroup   States   System.Windows.VisualState   
Storyboard �� 
TargetName �� TargetProperty8  2System.Windows.Media.Animation.EasingColorKeyFrame �� KeyTime �� Value  DefaultDataGridCellStyle*  $System.Windows.Controls.DataGridCell ��Name  DefaultDataGridRowStyle)  #System.Windows.Controls.DataGridRow   SystemFillColorCriticalBrush?  9System.Windows.Controls.Primitives.SelectiveScrollingGrid � ColumnDefinitions � RowDefinitions?	  9System.Windows.Controls.Primitives.DataGridCellsPresenter   
ItemsPanelA
  ;System.Windows.Controls.Primitives.DataGridDetailsPresenter   DetailsVisibility$   SelectiveScrollingOrientation;  5System.Windows.Controls.SelectiveScrollingOrientation   Vertical �� ConverterParameter#    RowDetailsScrollingConverter! �� 	Converter" �� RelativeSource# �� AncestorType:  4System.Windows.Controls.Primitives.DataGridRowHeader7  1System.Windows.Controls.DataGridHeadersVisibility
$   Row!%   HeadersVisibilityConverter  RowHeaderGripperStyle& �� Padding ! DefaultDataGridRowHeaderStyle  ControlElevationBorderBrush' �� Orientation( n�Name  ColumnHeaderGripperStyle $  DefaultDataGridColumnHeaderStyle=  7System.Windows.Controls.Primitives.DataGridColumnHeader  ControlFillColorDefault . *DefaultDataGridColumnHeadersPresenterStyleG  ASystem.Windows.Controls.Primitives.DataGridColumnHeadersPresenter)  Name* J� IsHitTestVisible+   SeparatorVisibility & "DefaultDataGridCellsPresenterStyle " DefaultDragIndicatorStyleStyle  TextFillColorPrimaryBrush    Wpf.Ui.Controls.SymbolIcon,  Name   Wpf.Ui.Controls.FontIcon-   FontSize.   Symbol	/ W� To0 g� Duration  TextBoxStyle1 �� Triggers2 P� Property3 P� Value0  *System.Windows.Controls.DataGridTextColumn1  +System.Windows.Controls.DataGridBoundColumn  DefaultDataGridStyle  TextFillColorPrimary  ControlCornerRadius4 �� CornerRadius)   #Wpf.Ui.Controls.PassiveScrollViewer5  Name6   SelectAllCommand
7   All8  Name9   Column: ��Name; �� CanContentScroll< ��Name= �� ScrollableHeight> � Maximum? �� Orientation@ �� ViewportHeightA �� ViewportSize*B �� #ComputedVerticalScrollBarVisibilityC � ValueD �� ModeE �� ScrollableWidthF �� ViewportWidth,G �� %ComputedHorizontalScrollBarVisibility.  (Wpf.Ui.Converters.FallbackBrushConverter  FallbackBrushConverter   FallbackColor ! CheckBoxPadding  " CheckBoxBorderThemeThickness # CheckBoxContentMargin $ CheckBoxIconSize % CheckBoxHeight & CheckBoxWidth '' #DataGridCheckBoxElementDefaultStyle "( DefaultControlFocusVisualStyle ) ControlContentThemeFontSize-  'System.Windows.Input.KeyboardNavigationH �� 
BackgroundI �� BulletJ   
FontWeight!   Wpf.Ui.Controls.IconElementK   
Foreground  * TextOnAccentFillColorPrimaryL ��NameM �� TriggersN �� 
TargetNameO �� Opacity + SystemAccentColorSecondaryP ]� 
ConditionsQ �� PropertyR �� Value ", ControlFillColorSecondaryBrush !- ControlFillColorDisabledBrush ". ControlStrokeColorDefaultBrush / TextFillColorDisabledBrush .0 *DataGridCheckBoxEditingElementDefaultStyleS �� BasedOn 1 DefaultUiDataGridStyle   Wpf.Ui.Controls.DataGrid�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls  .
convertersclr-namespace:Wpf.Ui.Converters  7system,clr-namespace:System;assembly=System.Runtime  %�(  &	        &	      &	      &	 $     &	 7     &	 J     &	 Y     &	 l     &	      &		 �     &	
 �     &	 �     &	     &	 �    (�� �    $" DataGridSelectAllButtonStyle��"  )7  &	 �    &	 !    &	 	    &	 �	    7   7  7  &	 �    &	 �    7  7  &	 �    &	 M    &	 �    &	      '  �    7  &	     7  7  7  7  7  7  &	 �    &	      &	!     &	" '    &	# 2    &	$ C    &	% O    &	& [    &	' g    7 " 7 ! 7 % 7 & 7 $ 7 # &	0 |%    7 ' &	1 �%    7  7 0 7 ' '  -&    7  '  <&    7  '  K&    7 1 �� X�14Q���	#FFE8EDF9����	#FFC5CBF9����	#FF888888����	White����	#FF7381F9����	#FF211AA9����	#FF3843C4����	#FF211AA9����	#FF444444���� $  0.5,0&�$ 0.5,1&�.��� $ 0Q�# �  � $	 0.5Q�# �  � $ 1Q�# �  �� $	  0,0&�$	 1,0&���� � $ 0Q�$ 	#000000FF��� $	 0.4Q�$ 	#600000FF��� $	 0.6Q�$ 	#600000FF��� $ 1Q�$ 	#000000FF���� "   �� " ��.���  �@�OverridesDefaultStyle$
 True����  � �� �� "	 ��.R�� .+��� $
 Border��
������ $
 True��
������ � $ Arrow��$��8,8,3,3q�$��Right=�$��Bottom=�$
 0.15Q� �   ��� ��� ��� ��� $ Uniform=����� # � 
    CommonStates��.   
Normal��  	MouseOver��. �� .���� $ Border��$* $(Shape.Fill).(SolidColorBrush.Color)�.g�  $ 0��8     Pressed��  Disabled���� "  .���  � ��$ Stretch=���  � ��$ Stretch=���  � ��$ 32����  � ��$ False����  �@�OverridesDefaultStyle$
 True����  � �� �� "	  .R��� $ Border��#��z��#��z��
������ 
������ $��0q�$
 True��._��� # z    FocusStates��.   	Unfocused��  Focused��  CurrentStates��.   Regular��  Current���� "  .���  � ��
 ���� ��  � ��
 ���� ��  � ��$ False����  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  �@=�ErrorTemplate Z� ��  �@ ValidationErrorTemplate �� .R��� $��2,0,0,0q�$��Center=�#���  $��!����  � �� �� "	  .R��� $ 
DGR_Border��
������ #��z��#��z��$
 True��._�   �� $
��Auto ��� $��* � �� $��* ��� $
��Auto �.+�	  $��1��#��z # z 
  $��1��$��1��#y�z  �� *AreRowDetailsFrozen+# Z #! Z  " �� "#    $��0��$��2��$��0��$ Vertical=�y��� *HeadersVisibility+# Z$ #! Z% " �� "#     CommonStates��.   
Normal��  Normal_AlternatingRow��  Normal_Selected��  Unfocused_Selected��  Normal_Editing��  MouseOver_Editing��  MouseOver_Unfocused_Editing��  Unfocused_Editing��  	MouseOver��  MouseOver_Selected��   MouseOver_Unfocused_Selected���� " n�.���  � ��$ 8����  � ��
 ���� ��  �@�Cursor$ SizeNS����  �@�OverridesDefaultStyle$
 True����  � �� �� "	 n�.R��� #& z��#��z���� "  .���  � ��
 ���� ��  � ��# �  ��  �@�OverridesDefaultStyle$
 True����  � �� �� "	  .R�� .+��� $ rowHeaderBorder��$��10��#��z��#��z��$��0,0,1,0q�._��� $' 
Horizontal=�.+��� $��Center=�# z �� $ False������ *ValidationErrorTemplate+" �� "#  y��� *(Validation.HasError)+8!   " �� "#  n� $( PART_TopHeaderGripper��$	��Top=�8�� n� $( PART_BottomHeaderGripper��$��Bottom=�8��    CommonStates��.   
Normal��  Normal_CurrentRow��  Unfocused_EditingRow��  Normal_EditingRow��  	MouseOver��  MouseOver_CurrentRow��  "MouseOver_Unfocused_EditingRow��  MouseOver_EditingRow��   MouseOver_Unfocused_Selected��  MouseOver_Selected��  +'MouseOver_Unfocused_CurrentRow_Selected��  !MouseOver_CurrentRow_Selected��  Unfocused_Selected��  !Unfocused_CurrentRow_Selected��  Normal_CurrentRow_Selected��  Normal_Selected���� " n�.���  � ��$ 8����  � ��
 ���� ��  �@�Cursor$ SizeWE����  �@�OverridesDefaultStyle$
 True����  � �� �� "	 n�.R��� #& z��#��z���� "  .���  � ��
 ���� ��  � ��# �  ��  � ��$ 12,0,0,0q���  � ��$ 32����  �@�FocusVisualStyle Z� ��  � ��$ Stretch=���  � ��$ Center=���  �@ SeparatorBrush# �  ��  � ��$ False����  �@�OverridesDefaultStyle$
 True����  � �� �� "	  .R�� .+��� $ columnHeaderBorder��$& 3,0,3,0q�#��z��#��z��$��0,0,1,1q�._��� #��z��#��z��# z n� $( PART_LeftHeaderGripper��$
��Left=�8��  n� $( PART_RightHeaderGripper��$��Right=�8�� �� "  .���  �@�OverridesDefaultStyle$
 True����  � �� �� "	  .R�� .+�  $) PART_FillerColumnHeader��$* False��$+ 	Collapsed=��� �� " 	 .���  �@�OverridesDefaultStyle$
 True����  � �� �� "	  .R�� .+�  $) PART_FillerColumnHeader��$* False��$+ 	Collapsed=��� �� " ��.���  � ��# �  ��  � ��
 ���� ��  �@�OverridesDefaultStyle$
 True����  � �� �� "	 ��.R�� #��z��.+�� #��z��#��z��#��z�� �� $��* ��� $
��Auto �$��32��.+�  $, SortIcon��$��1��$��Center=�$��Center=�$- 12(�$ 0Q�$. TextSortAscending24=�   
SortStates��.   Unsorted��  SortAscending��. �� .��W� $ SortIcon��$ Opacity�/ 1$0 0D�  SortDescending��. �� .��W� $ SortIcon��$ Opacity�/ 1$0 0D��� " ��.���  � ��$ 3q���  � ��$ 30����  � ��$ 20����  � ��$ Center=���  � ��$ Center=���  � ��
 � ���1 P� 2 �@��
IsReadOnly$
3 True��. ���  � ��
 ������� "  .���  �@ EditingElementStyle8   �� "  .���  �@�FocusVisualStyle Z� ��  � �� �� # �  ��  �@ RowBackground
 ���� ��  � �� �� # �  ��  � ��# �  ��   �@ HorizontalGridLinesBrush# �  ��  �@ VerticalGridLinesBrush# �  ��  � ��$ 1q���  � ��8   ��   �@ RowDetailsVisibilityMode$ VisibleWhenSelected=���  � ��$
 True����  �@��PanningMode$
 Both=���  �@��IsFlicksEnabled$ False����  �@��CornerRadius# �  ��  �@ RowStyle8  ��  �@ RowHeaderStyle8  ��  �@ 	CellStyle8  ��  �@ ColumnHeaderStyle8  ��  �@ DragIndicatorStyle8  ��  �@�OverridesDefaultStyle$
 True����  � �� �� "	  .R��� $ border��#& z��#��z��#��z��#��z��#4 z4 $
 True��._�  $5 DG_ScrollViewer��$|�False������ "	  .R��  �� $
��Auto ��� $��* ��� $
��Auto � �� $
��Auto ��� $��* ��� $
��Auto �.+��� #��Z6 $|�false������ *CellsPanelHorizontalOffset+" �� "#  ��C� *�� $" DataGridSelectAllButtonStyle��"  +y��� *HeadersVisibility+# Z7 #! Z% " �� "#    $!8 PART_ColumnHeadersPresenter��$��0��$��1��y��� *HeadersVisibility+# Z9 #! Z% " �� "#  �� $!: PART_ScrollContentPresenter��$��1��$��0��$��2��#; z���� $< PART_VerticalScrollBar��$��1��$��2��#> z= $? Vertical=�#A z@ #y�zB C �� *VerticalOffset+$D OneWay=�" �� *TemplatedParent+� $��2��$��1�� �� ���� **(NonFrozenColumnsViewportHorizontalOffset+" �� "#  �� $��* �.+��� $< PART_HorizontalScrollBar��$��1��#> zE $? 
Horizontal=�#A zF #y�zG C �� *HorizontalOffset+$D OneWay=�" �� *TemplatedParent+.���� # z    CommonStates��.   Disabled��. �� .���� $ border��$G A(Panel.Background).                       (SolidColorBrush.Color)�.g�  $ 0��# �    
Normal��1 P� 2 �@��
IsGrouping$
3 True��. ���  � ��$ False��  ��	#FFFF0000��u�	11,5,11,6q�u�1q�u�8,0,0,0q�X�14Q�X�22Q�X�22Q��� " ��.���  �@�FocusVisualStyle# � ( ��  � ��
 ���� ��  � �� �� # �  ��  � ��# �  ��  � ��8   ��  � ��8  ��  �@��CornerRadius# �  ��  � ��$ Center=���  � ��$ Center=���  � ��$
 Left=���  � ��$ Center=���  �@�Cursor$
 Hand����  � ��# � ) ��  � ��$ Normal"���  �@ 	IsTabStop$
 True����  � |�$
 True����  � ��$	 120����  � ��$ 32����  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� "	 ��.R��� #��z��#��z��
H ���� I �� $  ControlBorderIconPresenter��8�� 8�� $
��Left=�$��Center=�#��z��#��z��#��z��#4 z4 ._�� .+�  $, ControlIcon��$��0q�$��Center=�$��Center=�8-  $
J Bold"�$. Checkmark48=�$y�	Collapsed=�K �� # � * ._��� $L ContentPresenter��8�� #��z��#��z��$
��True��M P� 3 Z� 2 � ��. ��� $N ContentPresenter�� � ��$ 0q�P� 2 � ��$3  ��. ��� $N ContentPresenter�� � ��$ 0q�P� 3 Z� 2 �@d�	IsChecked. ��� $N ControlIcon�� �@ Symbol$ 
Subtract16=��� $N ControlIcon�� � y�$ Visible=��� $ N ControlBorderIconPresenter�� � �� �� $	O 1.0Q�# � + P� 2 �@d�	IsChecked$
3 TrueY�. ��� $N ControlIcon�� � y�$ Visible=��� $ N ControlBorderIconPresenter�� � �� �� $	O 1.0Q�# � + ]� P �� Q �@J�IsMouseOver$
R True���� Q �@d�	IsChecked$R FalseY�./��� $ N ControlBorderIconPresenter��# � ,  � ��]� P �� Q �@J�IsMouseOver$
R True���� R Z� Q �@d�	IsChecked./��� $ N ControlBorderIconPresenter�� � �� �� $	O 0.8Q�# � + ]� P �� Q �@J�IsMouseOver$
R True���� Q �@d�	IsChecked$
R TrueY�./��� $ N ControlBorderIconPresenter�� � �� �� $	O 0.8Q�# � + P� 2 � {�$3 False��. ��� $ N ControlBorderIconPresenter��# � -  � ���� $ N ControlBorderIconPresenter��# � .  � ���� $N ControlIcon��# � /  �@ 
Foreground�� # � /  � ���� 8S   " ��.���  �@J�SnapsToDevicePixels$
 True���� 8S   "  .��� ) �@ !CheckBoxColumnEditingElementStyle8  �� " �@ CheckBoxColumnElementStyle8  �� 8S   "  �� 8S   "  �� 8S   "  