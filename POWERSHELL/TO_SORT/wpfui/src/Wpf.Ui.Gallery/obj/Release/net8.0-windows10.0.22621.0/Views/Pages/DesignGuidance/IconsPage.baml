   M S B A M L   `   `   `  ���� I  EWpf.Ui.Gallery, Version=3.0.0.0, Culture=neutral, PublicKeyToken=nullA%clr-namespace:Wpf.Ui.Gallery.ControlsWpf.Ui.Gallery.Controls  e7clr-namespace:Wpf.Ui.Gallery.Views.Pages.DesignGuidance)Wpf.Ui.Gallery.Views.Pages.DesignGuidance  =#clr-namespace:Wpf.Ui.Gallery.ModelsWpf.Ui.Gallery.Models  9    3Wpf.Ui.Gallery.Views.Pages.DesignGuidance.IconsPage   W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     :controls%clr-namespace:Wpf.Ui.Gallery.Controls     @d2http://schemas.microsoft.com/expression/blend/2008     Ilocal7clr-namespace:Wpf.Ui.Gallery.Views.Pages.DesignGuidance     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     6models#clr-namespace:Wpf.Ui.Gallery.Models     A =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null/ui&http://schemas.lepo.co/wpfui/2022/xaml    N� Title$  	IconsPage��6   0Wpf.Ui.Gallery.Controls.PageControlDocumentation   DocumentationType   Wpf.Ui.Controls.SymbolIcon"     ApplicationBackgroundBrush  Wpf.Ui.Markup.Design   
Background# �     TextFillColorPrimaryBrush   
Foreground# �   N� 
Foreground# �  $��False��.���  � RowDefinitions �� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $��* �.+�  Wpf.Ui.Controls.TextBlock  $��0��$����WPF UI uses Fluent System Icons. Although this font was also created by Microsoft, it does not contain all the icons for Windows 11. If you need the missing icons, add Segoe Fluent Icons to your application.��$��WrapWithOverflow=�%  Wpf.Ui.Controls.HyperlinkButton  $��1��$��0,4,0,0q�$��0q�$-��'Find out more about Fluent System Icons��   NavigateUri$8 2https://github.com/microsoft/fluentui-system-icons��  $��2��$��0,24,0,0q�   FontTypography$ 
BodyStrong=�$!��Fluent System Icons Library��� $��3�� � ColumnDefinitions �� $
��Auto ��� $
��Auto �.+�$  Wpf.Ui.Controls.AutoSuggestBox  $��0��$	��320��$��0,4,0,0q�$
��Left=�	   PlaceholderText$	 Search icons��
   Text
 �� *ViewModel.AutoSuggestBoxText+ �� Mode$ TwoWay=� �� UpdateSourceTrigger$ PropertyChanged=��� $��1��$��0��$��12,0,0,0q�$��	Is filled������ *" ViewModel.CheckboxCheckedCommand+ �� RelativeSource ��  �� Mode$ FindAncestor=� �� AncestorType local:IconsPage$ OneWay=�����  �� *Self+$ OneWay=�� $��4��$��0,12,0,0q� �� $��* ��� $
��Auto �.+��� $��0��$��1,1,0,0q� �� CornerRadius$ 8,0,0,0����*  $Wpf.Ui.Markup.ThemeResourceExtension  *%#CardBackgroundFillColorDefaultBrush+._�.  (Wpf.Ui.Controls.VirtualizingItemsControl  $��0,24,4,0q�$��0q� %� CacheLengthUnit$ Pixel=����� *#!ViewModel.FilteredIconsCollection+$ OneWay=�����  �� DataType+	   %Wpf.Ui.Gallery.Models.DisplayableIcon" 	 .R��� $��80��$��80��$��2q�$��0q�$��Stretch=�$��Stretch=� ^� InitialShowDelay$	 240������ *ViewModel.IconSelectedCommand+ �� $ FindAncestor=� local:IconsPage����  �� Path$ Id�$ OneTime=� � ToolTip �� *Name+$ OneTime=�.��  $��0��
  Wpf.Ui.Controls.FontIcon 
  FontSize$ 28(�   Symbol �� *Icon+$ OneTime=��� $��1��$	��300��$��0,1,1,0q�$ 0,8,0,0����  *'%CardBackgroundFillColorSecondaryBrush+��  *%#CardBackgroundFillColorDefaultBrush+._�� $��24q� �� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto �.+�  $��0��$ 
BodyStrong=����� *ViewModel.SelectedSymbolName+$ OneWay=�  $��1��$��	0,30,0,24q�$
��Left=�$ 62(�   Filled �� *ViewModel.IsIconFilled+$ OneWay=� �� *ViewModel.SelectedSymbol+$ OneWay=�  $��2��$ 
BodyStrong=�$��	Icon name��  $��3����  *TextFillColorSecondaryBrush+���� *ViewModel.SelectedSymbolName+$ OneWay=�  $��4��$��0,8,0,0q�$ 
BodyStrong=�$��Unicode point��  $��5����  *TextFillColorSecondaryBrush+���� *&$ViewModel.SelectedSymbolUnicodePoint+$ OneWay=�  $��6��$��0,8,0,0q�$ 
BodyStrong=�$��
Text glyph��  $��7����  *TextFillColorSecondaryBrush+���� *#!ViewModel.SelectedSymbolTextGlyph+$ OneWay=�  $��8��$��0,8,0,0q�$ 
BodyStrong=�$
��XAML��  $��9��$	��250��$��WrapWithOverflow=���  *TextFillColorSecondaryBrush+���� *ViewModel.SelectedSymbolXaml+$ OneWay=�