   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089 !  DefaultCalendarDayButtonStyle  �� 
TargetType:   4System.Windows.Controls.Primitives.CalendarDayButton �� Property �� Value �� 
TargetType ��Name  CalendarViewTodayBackground J� Opacity �� RadiusX �� RadiusY # CalendarViewSelectedBorderBrush ) %CalendarViewItemBackgroundPointerOver ��Name  CalendarViewForeground	 J� IsHitTestVisible
 �� Color'  !System.Windows.VisualStateManager   VisualStateGroups%  System.Windows.VisualStateGroup   Name   Transitions%  System.Windows.VisualTransition   GeneratedDuration   States   System.Windows.VisualState   Name   
Storyboard �� 
TargetName �� TargetProperty	 W� To g� Duration R� KeyTime R� Value  System.Windows.Visibility  CalendarViewTodayForeground  DefaultCalendarButtonStyle7  1System.Windows.Controls.Primitives.CalendarButton " CalendarViewSelectedBackground  DefaultCalendarItemStyle5  /System.Windows.Controls.Primitives.CalendarItem �Name � RowDefinitions � ColumnDefinitions   Wpf.Ui.Controls.Button  Name   
Appearance  	 CalendarViewButtonForeground   Icon 	   Wpf.Ui.Controls.SymbolIcon 	  Filled
   Wpf.Ui.Controls.FontIcon 
  FontSize  	  Symbol! �� Stretch 
 ControlFillColorDefault" �� Triggers# P� Property$ P� Value% �� 
TargetName& �� Value' �� Binding( �� RelativeSource) �� AncestorType&   System.Windows.Controls.Calendar* � 	Resources  DefaultCalendarStyle+ ��Name, �� Padding  CalendarViewBackground  CalendarViewBorderBrush- �� CornerRadius.  Name/   CalendarItemStyle0 �� BasedOn�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls  %  &	        &	 �    &	 �    &	     7  7   7  '  �    7  �� "    .���  � ��$ 30����  � ��$ 30����  � ��$ 14(���  � ��$ Center=���  � ��$ Center=���  � �� �� "   .R�� .+��� $ TodayBackground��#���  $ 0Q�$ 99��$ 99���� $ SelectedBackground��$ 0Q�$ 99��$ 99��#���  $��1���� #��z��#��z��#��z���� $ HighlightBackground��#���  $ 0Q�$ 99��$ 99���� $ 
NormalText��$��5,1,5,1q�#��z��#��z��#���  �� $ DayButtonFocusVisual��$	 false��$ 1��$ 1��$y�	Collapsed=����� $
 Transparent��   $ CommonStates��   $ 0:0:0.1D�.   $ Normal��  $ 	MouseOver��. �� .��W� $ HighlightBackground��$ Opacity� 1$ 0D�  $ Pressed��. �� .��W� $ HighlightBackground��$ Opacity� 1$ 0D�  $ Disabled��. �� .��W� $ HighlightBackground��$ Opacity� 0$ 0D�W� $ 
NormalText��$ Opacity� 1$ 0D�  $ SelectionStates��   $ 0D�.   $ 
Unselected��  $ Selected��. �� .��W� $ SelectedBackground��$ Opacity� 1$ 0D�  $ CalendarButtonFocusStates��   $ 0D�.   $ CalendarButtonFocused��. �� .��T� $ DayButtonFocusVisual��$ 
Visibility�$ 0D�..�i� $ 0��  Visible=�  $ CalendarButtonUnfocused��. �� .��T� $ DayButtonFocusVisual��$ 
Visibility�$ 0D�..�i� $ 0��  	Collapsed=�  $ ActiveStates��   $ 0D�.   $ Active��  $ 	DayStates��   $ 0D�.   $ 
RegularDay��  $ Today��. �� .��W� $ TodayBackground��$ Opacity� 1$ 0D�T� $ 
NormalText��$ (TextBlock.Foreground)�.�i� $ 0:0:0��# �    $ BlackoutDayStates���� "   .���  � ��$ 40����  � ��$ 40����  � ��$ 12(���  � ��$ Center=���  � ��$ Center=���  � �� �� "  .R�� .+��� $ SelectedBackground��#���  $ 0Q�$ 99��$ 99���� $ 
Background��#���  $ 0Q�$ 99��$ 99���� $ 
NormalText��$��1,0,1,1q�#��z��#��z��#���  �� $ CalendarButtonFocusVisual��$	 false��$ 99��$ 99��#���  $y�	Collapsed=�   $ CommonStates��   $ 0:0:0.1D�.   $ Normal��  $ 	MouseOver��. �� .��W� $ 
Background��$ Opacity� 1$ 0D�  $ Pressed��. �� .��W� $ 
Background��$ Opacity� 1$ 0D�  $ SelectionStates��   $ 0D�.   $ 
Unselected��  $ Selected��. �� .��W� $ SelectedBackground��$ Opacity� 1$ 0D�T� $ 
NormalText��$ (TextBlock.Foreground)�.�i� $ 0:0:0��# �    $ ActiveStates��   $ 0D�.   $ Active��  $ CalendarButtonFocusStates��   $ 0D�.   $ CalendarButtonFocused��. �� .��T� $ CalendarButtonFocusVisual��$ 
Visibility�$ 0D�..�i� $ 0��  Visible=�  $ CalendarButtonUnfocused���� "   .���  � ��$ 0q���  � ��$ 0q���  � ��
 ���� ��  � ��
 ���� ��  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� "  .R�� $ 	PART_Root��$��12q�.+��  �� $
��Auto ��� $��* �.+�� $��0��$��8,0,8,14q� �� $��* ��� $
��Auto ��� $
��Auto �.+�  $ PART_HeaderButton��$��0��$��-6,0,0,0q�$	��6,2q�$
��Left=�$��Center=�$ 	Secondary=�
������ 
������ $|�False��$��14(�$
��Bold"�#���    $ PART_PreviousButton��$��1��$��26��$��26��$��0,0,8,0q�$��0q�$��Right=�$��Center=�$ 	Secondary=�
������ 
������ $�� ��$|�False��#��� 	  	  $
 True��$ 26(�$  	CaretUp20=�  $ PART_NextButton��$��2��$��26��$��26��$��0q�$��0q�$��Right=�$��Center=�$ 	Secondary=�
������ 
������ $�� ��$|�False��#��� 	  	  $
 True��$ 26(�$  CaretDown20=�� $ PART_MonthView��$��1��$��Center=�$��Center=�$y�Visible=� �� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto � �� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $��1��$��Center=�$��Center=�._�� $ PART_YearView��$y�Hidden=� �� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto � �� $
��Auto ��� $
��Auto ��� $
��Auto ��� $ PART_DisabledVisual��$��0��$��2��$ 0Q�$ 2��$ 2��$
! Fill=�
������ $��0��$y�	Collapsed=����� #
 � 
 " P� # � {�$$ False��. ��� $% PART_DisabledVisual�� � y�$ Visible=��� $
& Year��' �� *DisplayMode+( �� *FindAncestor+")  .a��� $% PART_MonthView�� � y�$ Hidden=��� $% PART_YearView�� � y�$ Visible=��� $& Decade��' �� *DisplayMode+( �� *FindAncestor+")  .a��� $% PART_MonthView�� � y�$ Hidden=��� $% PART_YearView�� � y�$ Visible=�* ���� (��       **(CalendarItem.DayTitleTemplateResourceKey+).R��� $��0,0,0,6q�$��Center=�$��Center=�$��12(�$��SemiBold"�#���  ���� �� "   .���  �@ CalendarButtonStyle8   ��  �@ CalendarDayButtonStyle8  ��  �@ CalendarItemStyle8  ��  � ��# �  ��  � ��
 ���� ��  � ��
 ���� ��  � ��$ 0q���  � �� �� "  .R��� $+ 	PART_Root��$��0q�$, 0q�#���  #���  $��1q�$- 8��._�  $. PART_CalendarItem��$��0q�$��0q�$��Stretch=�$��Stretch=�#��z��#��z��#��z��#��z/ �� 80   "   