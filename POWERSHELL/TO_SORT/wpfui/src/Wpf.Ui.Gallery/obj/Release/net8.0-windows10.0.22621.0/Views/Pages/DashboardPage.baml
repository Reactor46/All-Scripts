   M S B A M L   `   `   `  ���� I  EWpf.Ui.Gallery, Version=3.0.0.0, Culture=neutral, PublicKeyToken=nullG(clr-namespace:Wpf.Ui.Gallery.Views.PagesWpf.Ui.Gallery.Views.Pages  .    (Wpf.Ui.Gallery.Views.Pages.DashboardPage   W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     :local(clr-namespace:Wpf.Ui.Gallery.Views.Pages     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     A =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null/ui&http://schemas.lepo.co/wpfui/2022/xaml    N� Title$  Dashboard Page��$��0,32,0,0q�   ApplicationBackgroundBrush  Wpf.Ui.Markup.Design   
Background# �     TextFillColorPrimaryBrush   
Foreground# �   N� 
Foreground# �  .���  � RowDefinitions �� $
��Auto ��� $
��Auto ��� $
��Auto �.+��� $��0��$	��200�� �� CornerRadius$ 8������  �� ImageSource$F @pack://application:,,,/Assets/pexels-johannes-plenio-1103970.jpg�� �� BitmapScalingMode$ HighQuality=� j� Stretch$ UniformToFill=�._��� $ 8����
� .��� 	 � Offset$	 0Q�
 � Color$
 	#1F000000��� $	 1Q�$
 	#4F000000��._�� .+��� $
��48,0q�$
��Left=�$��Center=�
������ .+�  Wpf.Ui.Controls.TextBlock     FontTypography$ Title=����#FFFFFF$��	Fluent UI��  $ Subtitle=����	#B7FFFFFF$%��Windows Presentation Foundation��  $ 
BodyStrong=����	#B7FFFFFF$7��1Build Fluent experiences on Windows using WPF UI.��  $��12q�$��Right=�$��Bottom=�$ Caption=����	#57FFFFFF$��Created by lepo.co��� $��1��$��0,24,0,0q� � ColumnDefinitions �� $��* ��� $��* ��� $��* � �� $
��Auto �.+�   Wpf.Ui.Controls.CardAction  $��0��$��0��$��0,0,4,0q�$��0q�$��Stretch=�$��
BasicInput��   IsChevronVisible$ False������ *ViewModel.CardClickCommand+ �� Mode$ OneWay=�.���  �� $
��Auto ��� $��* �.+��� $��60��$��24,0,0,0q�$;��5pack://application:,,,/Assets/WinUiGallery/Button.png���� $��1��$��24q�$��Center=�.+�  $��0q�$ 
BodyStrong=�$��Basic input��$��WrapWithOverflow=�     
Appearance$ 	Secondary=�  TextFillColorSecondaryBrush#���  $%��Buttons, CheckBoxes, Sliders...��$��WrapWithOverflow=�  $��0��$��1��$��4,0,4,0q�$��0q�$��Stretch=�$��DialogsAndFlyouts��$ False������ *ViewModel.CardClickCommand+$ OneWay=�.���  �� $
��Auto ��� $��* �.+��� $��60��$��24,0,0,0q�$;��5pack://application:,,,/Assets/WinUiGallery/Flyout.png���� $��1��$��24q�$��Center=�.+�  $��0q�$ 
BodyStrong=�$��Dialogs & Flyouts��$��WrapWithOverflow=�  $ 	Secondary=�#���  $��Contextual notifications.��$��WrapWithOverflow=�  $��0��$��2��$��4,0,4,0q�$��0q�$��Stretch=�$��
Navigation��$ False������ *ViewModel.CardClickCommand+$ OneWay=�.���  �� $
��Auto ��� $��* �.+��� $��60��$��24,0,0,0q�$<��6pack://application:,,,/Assets/WinUiGallery/MenuBar.png���� $��1��$��24q�$��Center=�.+�  $��0q�$ 
BodyStrong=�$��
Navigation��$��WrapWithOverflow=�  $ 	Secondary=�#���  $#��Managing the displayed pages.��$��WrapWithOverflow=��� $��2��$��0,24,0,0q�.+��� $��18(�$��DemiBold"�$��
Learn more��%  Wpf.Ui.Controls.HyperlinkButton  $��8,8,0,0q�$��4q�$��Documentation��   NavigateUri$ https://wpfui.lepo.co/�� � ToolTip$ https://wpfui.lepo.co/��  $��8,0,0,0q�$��4q�$��Microsoft Store��$A ;https://apps.microsoft.com/store/detail/wpf-ui/9N9LKV8R9VGM��$A ;https://apps.microsoft.com/store/detail/wpf-ui/9N9LKV8R9VGM��  $��8,0,0,0q�$��4q�$��GitHub��$% https://github.com/lepoco/wpfui��$% https://github.com/lepoco/wpfui��  $��8,0,0,0q�$��4q�$��NuGet��$, &https://www.nuget.org/packages/wpf-ui/��$, &https://www.nuget.org/packages/wpf-ui/��  $��8,0,0,0q�$��4q�$��Visual Studio Marketplace��$E ?https://marketplace.visualstudio.com/items?itemName=lepo.wpf-ui��$E ?https://marketplace.visualstudio.com/items?itemName=lepo.wpf-ui��