   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  O Kmscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e0891&clr-namespace:System;assembly=mscorlibSystem W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089   ScrollAnimationDuration   ButtonHoverAnimationDuration  LineButtonHeight  LineButtonWidth  UiScrollBarLineButton  �� 
TargetType �� Property " ScrollBarButtonArrowForeground �� Value �� 
TargetType ��Name  ScrollBarButtonBackground �� CornerRadius     Wpf.Ui.Controls.SymbolIcon    Filled   Wpf.Ui.Controls.FontIcon   FontSize!   Wpf.Ui.Controls.IconElement   
Foreground	    Symbol
 �� Path �� RelativeSource �� Mode �� Triggers P� Property P� Value N� EnterActions �� 
TargetName �� TargetProperty W� From	 W� To g� Duration N� ExitActions  UiScrollBarPageButton  UiScrollBarThumb 	 ScrollBarThumbFill 
 UiVerticalScrollBar � RowDefinitions ! ScrollBarTrackFillPointerOver J� Opacity ��Name#   Wpf.Ui.Controls.SymbolRegular   	CaretUp24 ]�Name ]� DecreaseRepeatButton ]� Thumb ]� IncreaseRepeatButton   CaretDown24  UiHorizontalScrollBar  � ColumnDefinitions!   CaretLeft24"   CaretRight24  UiScrollBar# �� Triggers$ �� BasedOn�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls  .sys&clr-namespace:System;assembly=mscorlib  %  &	        &	      &	 $     &	 0     &	 <     7  7  7  7  &	     &	 �    &	
 �    7  7  7  7  7  7   7   7   7   7   7   7   7   &	 	    7  7  7  7  7  7   7   7   7   7   7   7   7   &	 |    7  7 
 '��     7  E�0:0:0.16D�E�0:0:0.16D�X�12Q�X�12Q��� "  ��.���  � ��# �  ��  � ��8   ��  � ��8  ��  � ��$ 11(���  � ��$ 0q���  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � |�$ False����  � �� �� " ��.R��� $ Border��#��z��#��z��#��z��#���  $ 6��._�   $��0,0,0,0q�$��Center=�$��Center=�$
 True��# z��# z��	 �� $
 Content� �� *TemplatedParent+$ OneWay=� P�  �@J�IsMouseOver$
 True�� �� .���� .��W� $ Border��$3 -(Border.Background).(SolidColorBrush.Opacity)� 0.0 1.08   �� .���� .��W� $ Border��$3 -(Border.Background).(SolidColorBrush.Opacity)� 1.0 0.08  �� "  ��.���  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 true����  � ��$ False����  � |�$ False����  � �� �� " ��.R��� 
������ �� "  n�.���  � ��# � 	 ��  �@��CornerRadius$ 4����  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � ��$ False����  � |�$ False����  � �� �� " n�.R��� #��z��#��z��$��1q�# z �� " ��.R��  �� $��14���� $��0.00001* ��� $��14��.+��� $ PART_Border��$��3��$��12��$��Center=�#���  $ 6��$ 0Q��� $ PART_ButtonScrollUp��$��0��$
��Left=�$��ScrollBar.LineUpCommand��#��Z $ 0Q�8��  ]� $ 
PART_Track��$��1��$��6��$
��True�� �� $��ScrollBar.PageUpCommand��8��  n� $��0q�$��0q�8��  �� $��ScrollBar.PageDownCommand��8�� �� $ PART_ButtonScrollDown��$��2��$
��Left=�$��ScrollBar.LineDownCommand��#��Z $ 0Q�8��  P�  �@J�IsMouseOver$
 True�� �� .���� .��W� $ 
PART_Track��$ Width� 6 88  W� $ PART_Border��$ Opacity� 0.0 1.08  W� $ PART_ButtonScrollUp��$ Opacity� 0.0 1.08  W� $ PART_ButtonScrollDown��$ Opacity� 0.0 1.08   �� .���� .��W� $ 
PART_Track��$ Width� 8 68 	 W� $ PART_Border��$ Opacity� 1.0 0.08 
 W� $ PART_ButtonScrollUp��$ Opacity� 1.0 0.08  W� $ PART_ButtonScrollDown��$ Opacity� 1.0 0.08  �� " ��.R��   �� $��18���� $��0.00001* ��� $��18��.+��� $ PART_Border��$��3��$��12��$��Center=�#���  $ 6��$ 0Q��� $ PART_ButtonScrollLeft��$��0��$��Center=�$��ScrollBar.LineLeftCommand��#��Z! $ 0Q�8��  ]� $ 
PART_Track��$��1��$��6��$��Center=�$��False�� �� $��ScrollBar.PageLeftCommand��8��  n� $��0q�$��0q�8��  �� $ ��ScrollBar.PageRightCommand��8�� �� $ PART_ButtonScrollRight��$��2��$��Center=�$ ��ScrollBar.LineRightCommand��#��Z" $ 0Q�8��  P�  �@J�IsMouseOver$
 True�� �� .���� .��W� $ 
PART_Track��$ Height� 6 88  W� $ PART_Border��$ Opacity� 0.0 1.08  W� $ PART_ButtonScrollLeft��$ Opacity� 0.0 1.08  W� $ PART_ButtonScrollRight��$ Opacity� 0.0 1.08   �� .���� .��W� $ 
PART_Track��$ Height� 8 68 	 W� $ PART_Border��$ Opacity� 1.0 0.08 
 W� $ PART_ButtonScrollLeft��$ Opacity� 1.0 0.08  W� $ PART_ButtonScrollRight��$ Opacity� 1.0 0.08  �� "  ��.���  � ��
 ���� ��  � ��$ 0q���  � ��$ 0q���  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True��# P�  �@��Orientation$ 
Horizontal=�. ���  � ��$
 Auto����  � ��$ 14����  � ��8   P�  �@��Orientation$ Vertical=�. ���  � ��$ 14����  � ��$
 Auto����  � ��8  �� 8$   "  ��