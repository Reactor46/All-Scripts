   M S B A M L   `   `   `  ���� I  EWpf.Ui.Gallery, Version=3.0.0.0, Culture=neutral, PublicKeyToken=nullA%clr-namespace:Wpf.Ui.Gallery.ControlsWpf.Ui.Gallery.Controls  K*clr-namespace:Wpf.Ui.Gallery.Views.WindowsWpf.Ui.Gallery.Views.Windows  -    'Wpf.Ui.Gallery.Views.Windows.MainWindow   W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     :controls%clr-namespace:Wpf.Ui.Gallery.Controls     @d2http://schemas.microsoft.com/expression/blend/2008     <local*clr-namespace:Wpf.Ui.Gallery.Views.Windows     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     F BWpf.Ui.Tray, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null6tray+http://schemas.lepo.co/wpfui/2022/xaml/tray  A =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null/ui&http://schemas.lepo.co/wpfui/2022/xaml  -   $
��1450��$	��802��$	��900��   ApplicationBackgroundBrush  Wpf.Ui.Markup.Design    
Background#  �     TextFillColorPrimaryBrush   
Foreground# �  "  Wpf.Ui.Controls.FluentWindow!   ExtendsContentIntoTitleBar$
 True��   WindowBackdropType$
 Mica=�   WindowCornerPreference$ Default=� � WindowStartupLocation$ CenterScreen=� � Title �� *ViewModel.ApplicationTitle+ �� Mode$ OneWay=� J� InputBindings �� 
	 �� Key$	 F��
 �� 	Modifiers$
 Controlf� �� Command ��  �� ElementName$ AutoSuggestBox�� �� Path$ FocusCommand�.��� .+�$  Wpf.Ui.Controls.NavigationView  -     Name$ NavigationView��$��	42,0,42,0q�   EnableDebugMessages$
 True��   FrameMargin$ 0q�   IsBackButtonVisible$ Visible=�   IsPaneToggleVisible$
 True��   OpenPaneLength$	 310Q�   PaneDisplayMode$
 Left=�   
Transition$ FadeInWithSlide=�   BreadcrumbBar �� $ BreadcrumbBar��   FooterMenuItemsSource �� *ViewModel.FooterMenuItems+$ OneWay=�   MenuItemsSource �� *ViewModel.MenuItems+$ OneWay=�   TitleBar �� $ TitleBar��$ OneWay=�   Header �� $��42,32,42,20q�.+�#  Wpf.Ui.Controls.BreadcrumbBar  -     Name$ BreadcrumbBar��6   0Wpf.Ui.Gallery.Controls.PageControlDocumentation  $��0,10,0,0q�   NavigationView �� $ NavigationView��   AutoSuggestBox $  Wpf.Ui.Controls.AutoSuggestBox  -     Name$ AutoSuggestBox��   PlaceholderText$ Search��    Icon  '  !Wpf.Ui.Controls.IconSourceElement  !   
IconSource.! &   Wpf.Ui.Controls.SymbolIconSource  "   Symbol$" Search24=�#   ContentOverlay# � .+�'	  !Wpf.Ui.Controls.SnackbarPresenter	  -   $ 	 Name$$ SnackbarPresenter���� -   % ��Name$% RootContentDialog��$��0��
  Wpf.Ui.Controls.TitleBar
  -   & 
 Name$& TitleBar��$��0��%' 
  CloseWindowByDoubleClickOnIcon$
' True��( 
  Title( �� *ViewModel.ApplicationTitle+) 
  Icon)   Wpf.Ui.Controls.ImageIcon  *   Source$-* 'pack://application:,,,/Assets/wpfui.png��%  Wpf.Ui.Tray.Controls.NotifyIcon  $��0��+   FocusOnLeftClick$
+ True��,   Icon$-, 'pack://application:,,,/Assets/wpfui.png��-   MenuOnRightClick$
- True��.   TooltipText$. WPF UI Gallery��/   Menu/ �� ���� *ViewModel.TrayMenuItems+$ OneWay=�