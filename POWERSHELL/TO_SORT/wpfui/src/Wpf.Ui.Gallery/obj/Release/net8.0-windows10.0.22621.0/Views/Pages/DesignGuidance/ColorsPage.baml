   M S B A M L   `   `   `  ���� I  EWpf.Ui.Gallery, Version=3.0.0.0, Culture=neutral, PublicKeyToken=nullA%clr-namespace:Wpf.Ui.Gallery.ControlsWpf.Ui.Gallery.Controls  e7clr-namespace:Wpf.Ui.Gallery.Views.Pages.DesignGuidance)Wpf.Ui.Gallery.Views.Pages.DesignGuidance  :    4Wpf.Ui.Gallery.Views.Pages.DesignGuidance.ColorsPage   W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     :controls%clr-namespace:Wpf.Ui.Gallery.Controls     @d2http://schemas.microsoft.com/expression/blend/2008     Ilocal7clr-namespace:Wpf.Ui.Gallery.Views.Pages.DesignGuidance     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     A =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null/ui&http://schemas.lepo.co/wpfui/2022/xaml    N� Title$  
ColorsPage��   ApplicationBackgroundBrush  Wpf.Ui.Markup.Design   
Background# �     TextFillColorPrimaryBrush   
Foreground# �   N� 
Foreground# �  .���� $��0,0,0,24q�.+�  Wpf.Ui.Controls.TextBlock  $����Color provides an intuitive way of communicating information to users in your app: it can be used to indicate interactivity, give feedback to user actions, and give your interface a sense of visual continuity.��$��WrapWithOverflow=�  $��0,12,0,0q�   FontTypography$ 
BodyStrong=�$��Using colors��$��WrapWithOverflow=�  $y��sThe colors below are provided as part of WPF UI. You can reference them in your app using DynamicResource bindings.��$��WrapWithOverflow=�  $��0,48,0,0q�$ Subtitle=�$
��Text��$��WrapWithOverflow=��� $��0,12,0,0q� �� Padding$ 12q� �� CornerRadius$ 8����*  $Wpf.Ui.Markup.ThemeResourceExtension  *ControlFillColorDefaultBrush+._��� .+�  $#��For UI labels and static text��  $��	0,12,0,24q�$��Center=�$��42(�$��SemiBold"�$��Aa��� $	��120��$��0,12,0,0q� � ColumnDefinitions �� $��* ��� $��* ��� $��* ��� $��* �.+��� $��0��$ 8,0,0,8����  *TextFillColorPrimaryBrush+��  *TextFillColorInverseBrush+._�� $��12q� � RowDefinitions �� $
��Auto ��� $��* ��� $
��Auto �.+�  $��0��$ 
BodyStrong=�$��Text / Primary��  $��1��$��Rest or hover��  $��2��$��TextFillColorPrimaryBrush���� $��1��$ 0����  *TextFillColorSecondaryBrush+��  *TextFillColorInverseBrush+._�� $��12q� �� $
��Auto ��� $��* ��� $
��Auto �.+�  $��0��$ 
BodyStrong=�$��Text / Secondary��  $��1��$��Rest or hover��  $��2��$!��TextFillColorSecondaryBrush���� $��2��$ 0����  *TextFillColorTertiaryBrush+��  *TextFillColorInverseBrush+._�� $��12q� �� $
��Auto ��� $��* ��� $
��Auto �.+�  $��0��$ 
BodyStrong=�$��Text / Tertiary��  $��1��$#��Pressed only (not accessible)��  $��2��$ ��TextFillColorTertiaryBrush���� $��3��$ 0,8,8,0����  *TextFillColorDisabledBrush+��  *TextFillColorInverseBrush+._�� $��12q� �� $
��Auto ��� $��* ��� $
��Auto �.+�  $��0��$ 
BodyStrong=�$��Text / Disabled��  $��1��$$��Disabled only (not accessible)��  $��2��$ ��TextFillColorDisabledBrush��  $��0,48,0,0q�$ Subtitle=�$��Accent Text��$��WrapWithOverflow=��� $��0,12,0,0q�$ 12q�$ 8����  *ControlFillColorDefaultBrush+._��� .+�  $��Recommended for links��  $��	0,12,0,24q�$��Center=�$��42(�$��SemiBold"�$��Aa����  *!AccentTextFillColorPrimaryBrush+� $	��120��$��0,12,0,0q� �� $��* ��� $��* ��� $��* ��� $��* �.+��� $��0��$ 8,0,0,8����  *!AccentTextFillColorPrimaryBrush+��  *TextFillColorInverseBrush+._�� $��12q� �� $
��Auto ��� $��* ��� $
��Auto �.+�  $��0��$ 
BodyStrong=�$��Accent Text / Primary��  $��1��$��Rest or hover��  $��2��$%��AccentTextFillColorPrimaryBrush���� $��1��$ 0����  *#!AccentTextFillColorSecondaryBrush+��  *TextFillColorInverseBrush+._�� $��12q� �� $
��Auto ��� $��* ��� $
��Auto �.+�  $��0��$ 
BodyStrong=�$��Accent Text / Secondary��  $��1��$��Rest or hover��  $��2��$'��!AccentTextFillColorSecondaryBrush���� $��2��$ 0����  *" AccentTextFillColorTertiaryBrush+��  *TextFillColorInverseBrush+._�� $��12q� �� $
��Auto ��� $��* ��� $
��Auto �.+�  $��0��$ 
BodyStrong=�$��Accent Text / Tertiary��  $��1��$#��Pressed only (not accessible)��  $��2��$&�� AccentTextFillColorTertiaryBrush���� $��3��$ 0,8,8,0����  *" AccentTextFillColorDisabledBrush+��  *TextFillColorInverseBrush+._�� $��12q� �� $
��Auto ��� $��* ��� $
��Auto �.+�  $��0��$ 
BodyStrong=�$��Accent Text / Disabled��  $��1��$$��Disabled only (not accessible)��  $��2��$&�� AccentTextFillColorDisabledBrush��