   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089   UiMenuItemScrollViewer  �� BasedOn �� 
TargetType �� Property �� Value  MenuBarItemBorderBrush �� 
TargetType ��Name �� CornerRadius � RowDefinitions � ColumnDefinitions	 ��Name
 m� Icon �Name � AllowsTransparency � HorizontalOffset m� IsSubmenuOpen � VerticalOffset � PlacementTarget �� ElementName �� Padding  FlyoutBackground  FlyoutBorderBrush J� SnapsToDevicePixels)    #Wpf.Ui.Controls.PassiveScrollViewer K� IsItemsHost-  'System.Windows.Input.KeyboardNavigation   DirectionalNavigation J� Effect3  -System.Windows.Media.Effects.DropShadowEffect   
BlurRadius   	Direction   Opacity   ShadowDepth �� Triggers P� Value P� Property �� 
TargetName ! MenuBarItemBackgroundSelected  TextFillColorDisabled �� Color  N� EnterActions! �� 
TargetName" �� TargetProperty# W� From	$ W� To% g� Duration& W� EasingFunction/  )System.Windows.Media.Animation.CircleEase7  1System.Windows.Media.Animation.EasingFunctionBase'   
EasingMode( z� SharedSizeGroup  CheckBoxBackground  CheckBoxBorderBrush) ��Name  FluentSystemIcons* �� TextAlignment 	 TextFillColorDisabledBrush+ m� InputGestureText, �Name    Wpf.Ui.Controls.SymbolIcon-  Name   Wpf.Ui.Controls.FontIcon.   FontSize/   Symbol!   Wpf.Ui.Controls.IconElement 
 
UiMenuItem " DefaultControlFocusVisualStyle0 �� Triggers 1 m� TopLevelHeaderTemplateKey2 m� TopLevelItemTemplateKey3 m� SubmenuHeaderTemplateKey4 m� SubmenuItemTemplateKey * &WpfUiMenuItemTopLevelHeaderTemplateKey   Wpf.Ui.Controls.MenuItem5 ��Name ( $WpfUiMenuItemTopLevelItemTemplateKey ' #WpfUiMenuItemSubmenuItemTemplateKey ) %WpfUiMenuItemSubmenuHeaderTemplateKey  WpfUiMenuItem�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls  %�&  &	        7��(�� F     *MenuItem.SeparatorStyleKey+)(�� U    *$"MenuItem.TopLevelHeaderTemplateKey+)7   (�� �    *" MenuItem.TopLevelItemTemplateKey+)(�� �    *!MenuItem.SubmenuItemTemplateKey+)(��     *#!MenuItem.SubmenuHeaderTemplateKey+)7   &	
 �    71 72 73 74 'm� �    7 
 &	 �    7   &	 y    &	 �    &	 �    7   &	 �#    7  7  7  7  '  H%    7  �� 8    " ��.���  � ��$ Disabled=���  � ��$
 Auto=��� " ��.���  � ��# �  ��  � ��
 ���� ��  � ��$ 0,1,0,1q���  � ��$ 1,1,0,0q���  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� " ��.R��� #��z��#��z��#��z��#��z��#��z���� " m�.R��� $ Border��$��4q�
������ $ 6��._��  �� $��* ��� $
��Auto �.+�� $��10q� �� $
��Auto ��� $��* �.+��� $
	 Icon��$��0��$��0,0,6,0q�$��Center=�#��z
 �� $	 HeaderPresenter��$��1��$��Center=�$��Header��$
��True��#��z��� $ Popup��$��1��$��0��$
 True��$|�False��$	 -12��#��z $��Bottom=�$
��None=�$ 1�� �� $ Border��.��� .+��� $ SubmenuBorder��$��
12,0,12,18q�$ 0,3,0,3q�#���  #���  $��1q�$ 8��$
 True��z�T� ._�   $
��True��8��  .���� $
 True��$ Cycle=�   $ 20Q�$	 270Q�$
 0.25Q�$ 6Q� P�  Z�  �@m�Icon. ��� $
 Icon�� � y�$ 	Collapsed=�P�  Z�  � ��. ��� $
 Icon�� � ��$ 0q��� $ HeaderPresenter�� � y�$ 	Collapsed=�P�  �@m�IsHighlighted$
 True��. ��� $ Border��# �   � ��P�  � {�$ False��. ���  � �� �� # �  P�  �@m�IsSubmenuOpen$
 True��  �� .���� .��W� $! SubmenuBorder��$5" /(Border.RenderTransform).(TranslateTransform.Y)�# -90$ 0$% 00:00:00.167D�&   $' EaseOut=��� " m�.R��� $ Border��$��4q�
������ $ 6��._�� $��10q� �� $
��Auto ��� $��* �.+��� $
	 Icon��$��0��$��0,0,6,0q�$��Center=�#��z
 �� $	 HeaderPresenter��$��1��$��Center=�$��Header��$
��True��#��z�� P�  �@m�IsHighlighted$
 True��. ��� $ Border��# �   � ��P�  Z�  �@m�Icon. ��� $
 Icon�� � y�$ 	Collapsed=�P�  Z�  � ��. ��� $
 Icon�� � ��$ 0q��� $ HeaderPresenter�� � y�$ 	Collapsed=�P�  � {�$ False��. ���  � �� �� # �  �� " m�.R��� $ Border��$��4,1,4,1q�
������ $ 4��._�� $��8,6,8,6q� �� $
��Auto ��� $
��Auto ��� $��* ��� $
��Auto �$( Shortcut��.+��� $ CheckBoxIconBorder��$��0��$��20��$��20��$��0,0,6,0q�$��Center=�#���  #���  $��1q�$ 4��$y�	Collapsed=�._��� $) CheckBoxIcon��$��Center=�$��Center=�#���  $��16(�$�� ��$* Center=��� $
	 Icon��$��1��$��0,0,6,0q�$��Center=�#��z
 �� $��2��$��Header��$
��True��#��z���� $) InputGestureText��$��3��$��25,0,0,0q�$��Bottom=�$��Right=�$��11(�#��� 	 #��z+  P�  �@m�IsHighlighted$
 True��. ��� $ Border��# �   � ��P�  Z�  �@m�Icon. ��� $
 Icon�� � y�$ 	Collapsed=�P�  �@m�IsCheckable$
 True��. ��� $ CheckBoxIconBorder�� � y�$ Visible=�P�  �@m�	IsChecked$
 True��. ��� $ CheckBoxIcon�� � ��$	 ��P�  � {�$ False��. ���  � �� �� # �  P�  �@m�InputGestureText$  ��. ��� $ InputGestureText�� � y�$ 	Collapsed=��� " m�.R��  �� $��* ��� $
��Auto �.+��� $ Border��$��1��$��4,1,4,1q�$��1q�$ 4��._�� $, MenuItemContent��$��8,6,8,6q� �� $
��Auto ��� $��* ��� $
��Auto �.+��� $
	 Icon��$��0��$��0,0,6,0q�$��Center=�#��z
 �� $	 
HeaderHost��$��1��$��Header��$
��True��� $��2��.+�  $- Chevron��$��0,3,0,0q�$��Center=�#. z��$/ ChevronRight20=�� $ Popup��$��1��$
 True��$|�False��#��z $��Right=�$
��None=�$	 -20�� �� $ MenuItemContent��.��� .+��� $ SubmenuBorder��$��12,10,12,18q�$ 0,3,0,3q�#���  #���  $��1q�$ 8��$
 True��z�T� ._�   $
��True��8��  .���� $
 True��$ Cycle=�   $ 20Q�$	 270Q�$	 0.5Q�$ 6Q� P�  Z�  �@m�Icon. ��� $
 Icon�� � y�$ 	Collapsed=��� $
 Icon�� � ��$ 0q�P�  �@m�IsHighlighted$
 true��. ��� $ Border��# �   � ��P�  � {�$ false��. ���  � �� �� # �  �� $ Chevron�� �@ 
Foreground �� # �  P�  �@m�IsSubmenuOpen$
 True��  �� .���� .��W� $! SubmenuBorder��$5" /(Border.RenderTransform).(TranslateTransform.Y)�# -90$ 0$% 00:00:00.167D�&   $' EaseOut=��� " m�.���  �@�FocusVisualStyle# �  ��  �@ 	IsTabStop$
 True����  � |�$
 True����  �@�OverridesDefaultStyle$
 True��0 P�  �@m�Role$ TopLevelHeader=�. ���  � ��8   ��  �@�IsSharedSizeScope$
 True��P�  �@m�Role$ TopLevelItem=�. ���  � ��8  P�  �@m�Role$ SubmenuHeader=�. ���  � ��8  P�  �@m�Role$ SubmenuItem=�. ���  � ��8  �� 8    " m��� "  .R��� $ Border��$��4q�
������ $ 6��._��  �� $��* ��� $
��Auto �.+�� $��10q� �� $
��Auto ��� $
��Auto ��� $��* �.+��� $5 IconElement��$��0��$��0,0,6,0q�$��Center=�#��z
 #��z���� $	 HeaderPresenter��$��2��$��Center=�$��Header��$
��True��#��z��� $ Popup��$��1��$��0��$��2��$
 True��$|�False��$	 -12��#��z $��Bottom=�$
��None=�$ 1�� �� $ Border��.��� .+��� $ SubmenuBorder��$��
12,0,12,18q�$ 0,3,0,3q�#���  #���  $��1q�$ 8��$
 True��z�T� ._�   $
��True��8��  .���� $
 True��$ Cycle=�   $ 20Q�$	 270Q�$
 0.25Q�$ 6Q� P�  Z�  �@ Icon. ��� $ IconElement�� � y�$ 	Collapsed=��� $ IconElement�� � ��$ 0q�P�  Z�  � ��. ��� $ IconElement�� � ��$ 0q��� $ HeaderPresenter�� � y�$ 	Collapsed=�P�  �@m�IsHighlighted$
 True��. ��� $ Border��# �   � ��P�  � {�$ False��. ���  � �� �� # �  P�  �@m�IsSubmenuOpen$
 True��  �� .���� .��W� $! SubmenuBorder��$5" /(Border.RenderTransform).(TranslateTransform.Y)�# -90$ 0$% 00:00:00.167D�&   $' EaseOut=��� "  .R��� $ Border��$��4q�
������ $ 6��._�� $��10q� �� $
��Auto ��� $
��Auto ��� $��* �.+��� $5 IconElement��$��0��$��0,0,6,0q�$��Center=�#��z
 #��z���� $	 HeaderPresenter��$��2��$��Center=�$��Header��$
��True��#��z�� P�  �@m�IsHighlighted$
 True��. ��� $ Border��# �   � ��P�  Z�  �@ Icon. ��� $ IconElement�� � y�$ 	Collapsed=��� $ IconElement�� � ��$ 0q�P�  Z�  � ��. ��� $ IconElement�� � ��$ 0q��� $ HeaderPresenter�� � y�$ 	Collapsed=�P�  � {�$ False��. ���  � �� �� # �  �� "  .R��� $ Border��$��4,1,4,1q�
������ $ 4��._�� $��8,6,8,6q� �� $
��Auto ��� $
��Auto ��� $
��Auto ��� $��* ��� $
��Auto �$( Shortcut��.+��� $ CheckBoxIconBorder��$��0��$��20��$��20��$��0q�$��Center=�#���  #���  $��1q�$ 4��$y�	Collapsed=�._��� $) CheckBoxIcon��$��Center=�$��Center=�#���  $��16(�$�� ��$* Center=��� $5 IconElement��$��1��$��0,0,6,0q�$��Center=�#��z
 $��16(��� $��3��$��Header��$
��True��#��z���� $) InputGestureText��$��4��$��25,0,0,0q�$��Bottom=�$��Right=�$��11(�#��� 	 #��z+  P�  �@m�IsHighlighted$
 True��. ��� $ Border��# �   � ��P�  Z�  �@ Icon. ��� $ IconElement�� � y�$ 	Collapsed=��� $ IconElement�� � ��$ 0q�P�  �@m�IsCheckable$
 True��. ��� $ CheckBoxIconBorder�� � y�$ Visible=��� $ CheckBoxIconBorder�� � ��$ 0,0,6,0q�P�  �@m�	IsChecked$
 True��. ��� $ CheckBoxIcon�� � ��$	 ��P�  � {�$ False��. ���  � �� �� # �  P�  �@m�InputGestureText$  ��. ��� $ InputGestureText�� � y�$ 	Collapsed=��� "  .R��  �� $��* ��� $
��Auto �.+��� $ Border��$��1��$��4,1,4,1q�$��1q�$ 4��._�� $, MenuItemContent��$��8,6,8,6q� �� $
��Auto ��� $
��Auto ��� $��* ��� $
��Auto �.+��� $5 IconElement��$��0��$��0,0,6,0q�$��Center=�#��z
 $��16(��� $	 
HeaderHost��$��2��$��Header��$
��True��� $��3��.+�  $- Chevron��$��0,3,0,0q�$��Center=�$. 14(�$/ ChevronRight20=�� $ Popup��$��1��$
 True��$|�False��#��z $��Right=�$
��None=�$	 -20�� �� $ MenuItemContent��.��� .+��� $ SubmenuBorder��$��12,10,12,18q�$ 0,3,0,3q�#���  #���  $��1q�$ 8��$
 True��z�T� ._�   $
��True��8��  .���� $
 True��$ Cycle=�   $ 20Q�$	 270Q�$	 0.5Q�$ 6Q� P�  Z�  �@ Icon. ��� $ IconElement�� � y�$ 	Collapsed=��� $ IconElement�� � ��$ 0q�P�  �@m�IsHighlighted$
 true��. ��� $ Border��# �   � ��P�  � {�$ false��. ���  � �� �� # �  �� $ Chevron�� �@ 
Foreground �� # �  P�  �@m�IsSubmenuOpen$
 True��  �� .���� .��W� $! SubmenuBorder��$5" /(Border.RenderTransform).(TranslateTransform.Y)�# -90$ 0$% 00:00:00.167D�&   $' EaseOut=��� "  .���  �@�OverridesDefaultStyle$
 True����  �@�FocusVisualStyle# �  ��  �@ 	IsTabStop$
 True����  � |�$
 True��0 P�  �@m�Role$ TopLevelHeader=�. ���  � ��8   ��  �@�IsSharedSizeScope$
 True��P�  �@m�Role$ TopLevelItem=�. ���  � ��8  P�  �@m�Role$ SubmenuHeader=�. ���  � ��8  P�  �@m�Role$ SubmenuItem=�. ���  � ��8  �� 8    "  