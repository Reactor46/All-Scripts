   M S B A M L   `   `   `  ���� I  EWpf.Ui.Gallery, Version=3.0.0.0, Culture=neutral, PublicKeyToken=nullK*clr-namespace:Wpf.Ui.Gallery.Views.WindowsWpf.Ui.Gallery.Views.Windows  e aMicrosoft.Web.WebView2.Wpf, Version=1.0.2151.40, Culture=neutral, PublicKeyToken=2a8ab48044d2601ekLclr-namespace:Microsoft.Web.WebView2.Wpf;assembly=Microsoft.Web.WebView2.WpfMicrosoft.Web.WebView2.Wpf /    )Wpf.Ui.Gallery.Views.Windows.MonacoWindow   W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     <local*clr-namespace:Wpf.Ui.Gallery.Views.Windows     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     A =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null/ui&http://schemas.lepo.co/wpfui/2022/xaml  Vww2Lclr-namespace:Microsoft.Web.WebView2.Wpf;assembly=Microsoft.Web.WebView2.Wpf    � Title$  WPF UI - Monaco Editor��$
��1250��$	��652��   ApplicationBackgroundBrush  Wpf.Ui.Markup.Design   
Background# �     TextFillColorPrimaryBrush   
Foreground# �  "  Wpf.Ui.Controls.FluentWindow!   ExtendsContentIntoTitleBar$
 True��   WindowBackdropType$
 Mica=�   WindowCornerPreference$ Default=� � WindowStartupLocation$ CenterOwner=�.���  � RowDefinitions �� $
��Auto ��� $
��Auto ��� $��* ��� $
��Auto �.+�  Wpf.Ui.Controls.TitleBar     Title$1 +WPF UI - Monaco Editor (Visual Studio Code)��$��0��	   Icon*  $Wpf.Ui.Controls.IconElementConverter$-	 'pack://application:,,,/Assets/wpfui.png o� $��1��$��0q�
������ $��14(�.@�  Wpf.Ui.Controls.MenuItem  $
��File��
   Icon
 '  !Wpf.Ui.Markup.SymbolIconExtension  *DocumentSplitHint20+.@�m�  m� CommandParameter$ newFile��$	��New�� m� Command �� *ViewModel.MenuActionCommand+ �� Mode$ OneWay=�m� $ 	newWindow��$��
New window�� �� *ViewModel.MenuActionCommand+$ OneWay=�m� $ openFile��$��Open...�� �� *ViewModel.MenuActionCommand+$ OneWay=�m� $ saveFile��$
��Save�� �� *ViewModel.MenuActionCommand+$ OneWay=�m� $ 
saveFileAs��$��
Save As...�� �� *ViewModel.MenuActionCommand+$ OneWay=��� m� $
 exit��$
��Exit�� �� *ViewModel.MenuActionCommand+$ OneWay=�  $��Debug��
   *DeveloperBoard24+.@�m� $ editUndo��$
��Undo�� �� *ViewModel.MenuActionCommand+$ OneWay=��� m� $ editCut��$	��Cut�� �� *ViewModel.MenuActionCommand+$ OneWay=�m� $ editCopy��$
��Copy�� �� *ViewModel.MenuActionCommand+$ OneWay=�m� $ 	editPaste��$��Paste�� �� *ViewModel.MenuActionCommand+$ OneWay=�m� $ 
editDelete��$��Delete��${�False�� �� *ViewModel.MenuActionCommand+$ OneWay=��� m� $ browserSearch��$��Search with browser�� �� *ViewModel.MenuActionCommand+$ OneWay=�m� $
 find��$��Find...�� �� *ViewModel.MenuActionCommand+$ OneWay=�m� $ findNext��$��	Find next�� �� *ViewModel.MenuActionCommand+$ OneWay=��� m� $ 	selectAll��$��
Select All�� �� *ViewModel.MenuActionCommand+$ OneWay=���   $ 	hotReload��  PaletteDeepOrangeBrush#���   �� *ViewModel.MenuActionCommand+$ OneWay=�
   *Fire24True+  $ build��  PaletteGreenBrush#���   �� *ViewModel.MenuActionCommand+$ OneWay=�
   *Play24+  $ build��  PaletteRedBrush#���   �� *ViewModel.MenuActionCommand+$ OneWay=�
   *Stop24+  $ build��  PaletteLightBlueBrush#���   �� *ViewModel.MenuActionCommand+$ OneWay=�
   *ArrowClockwise24+��   $
��Help��
   *
ChatHelp20+.@�m� $ viewHelp��$��	View help�� �� *ViewModel.MenuActionCommand+$ OneWay=�m� $ viewFeedback��$��Send feedback�� �� *ViewModel.MenuActionCommand+$ OneWay=��� m� $ 	viewAbout��$��About WPF UI�� �� *ViewModel.MenuActionCommand+$ OneWay=�)  #Microsoft.Web.WebView2.Wpf.WebView2  -     Name$ WebView��$��2��$��0q�$��Stretch=�$��Stretch=��� -    ��Name$ RootStatusBar��$��3��$��0��$��0q�$	��8,2q����� .R��  � ColumnDefinitions �� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $��* ��� $
��Auto �.@��� .��   Wpf.Ui.Controls.SymbolIcon    TextFillColorSecondaryBrush!	  Wpf.Ui.Controls.IconElement 	  
Foreground# �     Symbol$ BranchFork24=��� $��1��$��0,0,4,0q�.���� #���  $��Development *���� $��2��$��4,0,4,0q�.��  # �  $ Home24=��� $��3��$��4,0,4,0q�.��  # �  $ Fire24=��� $��4��$��4,0,4,0q�.���� #���  $��
Wpf.Ui.sln���� $��6��.���� #���  $/��)Ln 45, Col 30  Spaces: 2 UTF8 with BOM C#��