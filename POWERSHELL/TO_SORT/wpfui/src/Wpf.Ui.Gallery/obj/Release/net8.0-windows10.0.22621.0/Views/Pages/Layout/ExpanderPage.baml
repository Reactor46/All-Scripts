   M S B A M L   `   `   `  ���� I  EWpf.Ui.Gallery, Version=3.0.0.0, Culture=neutral, PublicKeyToken=nullA%clr-namespace:Wpf.Ui.Gallery.ControlsWpf.Ui.Gallery.Controls  U/clr-namespace:Wpf.Ui.Gallery.Views.Pages.Layout!Wpf.Ui.Gallery.Views.Pages.Layout  U QSystem.Runtime, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a7,clr-namespace:System;assembly=System.RuntimeSystem 4    .Wpf.Ui.Gallery.Views.Pages.Layout.ExpanderPage   W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     :controls%clr-namespace:Wpf.Ui.Gallery.Controls     @d2http://schemas.microsoft.com/expression/blend/2008     Alocal/clr-namespace:Wpf.Ui.Gallery.Views.Pages.Layout     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     ?system,clr-namespace:System;assembly=System.Runtime     A =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null/ui&http://schemas.lepo.co/wpfui/2022/xaml    N� Title$  ExpanderPage��   ApplicationBackgroundBrush  Wpf.Ui.Markup.Design   
Background# �     TextFillColorPrimaryBrush   
Foreground# �   N� 
Foreground# �  .���� $��0,0,0,24q�.+�,   &Wpf.Ui.Gallery.Controls.ControlExample  $��0q�   
HeaderText$; 5An Expander with text in the header and content areas��   XamlCode$W Q<Expander Header="This text is in the header" Content="This is in the content" />��   ExampleContent. �  � ColumnDefinitions �� $��* ��� $
��Auto �.+�9� $��0��$��This is in the content��$ ��This text is in the header��  $��0,32,0,0q�$Z TAn Expander with an Icon, a header text, a description, a control, and content areas��$7 1<Expander Icon="{ui:SymbolIcon PlaySettings20}"/>��. �  �� $��* ��� $
��Auto �.+�"  Wpf.Ui.Controls.CardExpander  $��0��   Icon '  !Wpf.Ui.Markup.SymbolIconExtension  *PlaySettings20+��� 	 � RowDefinitions	 �� $
��Auto ��� $
��Auto � �� $��* ��� $
��Auto �.+�  Wpf.Ui.Controls.TextBlock  $��0��$��0��$��16(�
   FontTypography$

 Body=�$��This is a header text��  $��1��$��0��$��12(�$ ��This is a description text����*  $Wpf.Ui.Markup.ThemeResourceExtension  *TextFillColorSecondaryBrush+"  Wpf.Ui.Controls.ToggleSwitch  $��0��$��2��$��1��$��0,0,16,0q�   
OffContent$	 Off��   	OnContent$ On��.���� $��24,0.5,24,0q�.+�!  Wpf.Ui.Controls.CardControl  $��20,10,20,10q�   Header$ This is an item��.��  $
��Left=�$
{�True��$	 Off��$ On��  $��	0,0.5,0,0q�$��20,10,20,10q�   Icon   *FlashSettings24+ �� .+�  $��0q�$
 
BodyStrong=�#���  $��This is a header text��$��WrapWithOverflow=�     
Appearance$ 	Secondary=�  TextFillColorSecondaryBrush#���  $!��This is a description text.��$��WrapWithOverflow=�.��$	  Wpf.Ui.Controls.DropDownButton	  $��Hello��
  Wpf.Ui.Controls.Button 
  Icon   *
Fluent24+ 	  Flyout �� .@�m� $	��Add��m� $��Remove��m� $
��Send��m� $��Hello��