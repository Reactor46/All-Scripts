   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089    Wpf.Ui.Controls.Snackbar  �� 
TargetType �� Property   SnackBarForeground �� Value  SnackBarBackground  SnackBarBorderBrush-  'System.Windows.Input.KeyboardNavigation �� 
TargetType �Name J� Opacity ��Name �� Padding �� CornerRadius	    SlideTransform
 J� Effect3  -System.Windows.Media.Effects.DropShadowEffect   
BlurRadius   	Direction   Opacity   ShadowDepth   Color � ColumnDefinitions ��Name ��Name    Icon � RowDefinitions    Title    TitleTemplate � 	Resources �� BasedOn �� ContentStringFormat �� ContentStringFormat    ContentForeground   Wpf.Ui.Controls.Button  Name    
Appearance   
Appearance �� Path  �� Mode! �� RelativeSource"   Icon    Wpf.Ui.Controls.SymbolIcon#   Symbol$ �� Triggers% P� Value& P� Property' �� 
TargetName( N� EnterActions) g� 	BeginTime* �� 
TargetName+ �� TargetProperty, p� KeyTime- p� Value. W� From	/ W� To0 g� Duration1 N� ExitActions2 �� Triggers   TextOnAccentFillColorPrimary3 �� Color " TextOnAccentFillColorSecondary  SystemAccentColorSecondary  PaletteLightBlueColor  PaletteRedColor  PaletteGreenColor 	 PaletteOrangeColor�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls  %�  '         7��7���� "    .���  � ��# �   ��  �@  ContentForeground# �   ��  � ��# �  ��  � ��# �  ��  � ��$ 24q���  � ��$ 68����  � |�$ False����  �@ 	IsTabStop$ False����  � ��$ Bottom=���  � ��$ Stretch=���  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� "   .R�� $ CardBorderContainer��$��
0,0,0,-200q�$ 0Q�.+��� $ 
CardBorder��#��z��$ 12q�$��Bottom=�#��z��#��z��$��1q�$ 4��#z�z	 #y�zy�
   $ 30Q�$ 0Q�$	 0.4Q�$ 0Q�$ #202020��._��  �� $
��Auto ��� $��* ��� $ ChevronColumn��$
��Auto �.+��� $ CardIcon��$��0��$��0,0,12,0q�$��Center=�#��z $��28(�#��z��� $��1�� �� $
��Auto ��� $��* �.+��� $ TitleContentPresenter��$��0��$��0q�#��z #��z $��16(�$��SemiBold"�#��z�� ���� 8   '��       "  ��.���  � ��$ WrapWithOverflow=��� $��1��#��z��# z #��z��#��z��$��14(�$��Regular"�#��z  ���� 8  '��       "  ��.���  � ��$ WrapWithOverflow=�  $ CloseButton��$��2��$��12,0,0,0q�$��6q�# z #��z������ $ TemplateButtonCommand�$  OneWay=�! �� *TemplatedParent+"   $# 	Dismiss48=�$ P� % Z� & �@  Icon. ��� $' CardIcon�� � ��$ 0q�P� & �@  IsCloseButtonEnabled$% False��. ��� $' CloseButton�� � y�$ 	Collapsed=��� $' CloseButton�� � ��$ 0q�P� & �@  IsShown$
% True��. ���  � |�$
 True����  �@ 	IsTabStop$
 True��( �� .���� .��r� ) 00:00:00$* CardBorderContainer��$+ Margin�.	��� $, 00:00:00��- 
0,0,0,-200�� $, 	00:00:.26��- 0,0,0,0W� $* CardBorderContainer��$+ (Grid.Opacity)�. 0/ 1$0 0:0:.26D�1 �� .���� .��r� ) 00:00:00$* CardBorderContainer��$+ Margin�.	��� $, 00:00:00��- 0,0,0,0�� $, 	00:00:.32��- 
0,0,0,-200W� $* CardBorderContainer��$+ (Grid.Opacity)�. 1/ 0$0 0:0:.32D�2 P� & �@  
Appearance$% Transparent=�. ���  � ��
 ���� P� & �@  
Appearance$% Primary=�. ���  � �� �� #3 �  ��  �@  ContentForeground �� #3 �  ��  � �� �� #3 �  P� & �@  
Appearance$
% Dark=�. ���  � �� �� $3 #FFFFFF����  �@  ContentForeground �� $3 	#87FFFFFF����  � �� �� $3 #2e2e2e��P� & �@  
Appearance$% Light=�. ���  � �� �� $3 	#E4000000����  �@  ContentForeground �� $3 	#9E000000����  � �� �� $3 #F3F3F3��P� & �@  
Appearance$
% Info=�. ���  � �� �� $3 #FFFFFF����  �@  ContentForeground �� $3 	#87FFFFFF����  � �� �� #3 �  P� & �@  
Appearance$% Danger=�. ���  � �� �� $3 #FFFFFF����  �@  ContentForeground �� $3 	#87FFFFFF����  � �� �� #3 �  P� & �@  
Appearance$% Success=�. ���  � �� �� $3 #FFFFFF����  �@  ContentForeground �� $3 	#87FFFFFF����  � �� �� #3 �  P� & �@  
Appearance$% Caution=�. ���  � �� �� $3 #FFFFFF����  �@  ContentForeground �� $3 	#87FFFFFF����  � �� �� #3 � 	 