   M S B A M L   `   `   `  ���� I  EWpf.Ui.Gallery, Version=3.0.0.0, Culture=neutral, PublicKeyToken=nullK*clr-namespace:Wpf.Ui.Gallery.Views.WindowsWpf.Ui.Gallery.Views.Windows  /    )Wpf.Ui.Gallery.Views.Windows.EditorWindow   W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     <local*clr-namespace:Wpf.Ui.Gallery.Views.Windows     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     A =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null/ui&http://schemas.lepo.co/wpfui/2022/xaml    � Title$  WPF UI - Editor��$
��1250��$	��652��   ApplicationBackgroundBrush  Wpf.Ui.Markup.Design   
Background# �     TextFillColorPrimaryBrush   
Foreground# �  "  Wpf.Ui.Controls.FluentWindow!   ExtendsContentIntoTitleBar$
 True��   WindowBackdropType$ Tabbed=�   WindowCornerPreference$ Default=� � WindowStartupLocation$ CenterOwner=�.��� .+��  � RowDefinitions �� $
��Auto ��� $
��Auto ��� $��* ��� $
��Auto �.+�  Wpf.Ui.Controls.TitleBar     Title$ WPF UI - Editor��$��0��	   Icon*  $Wpf.Ui.Controls.IconElementConverter$-	 'pack://application:,,,/Assets/wpfui.png o� $��1��
������ $��14(�.@�  Wpf.Ui.Controls.MenuItem  $
��File��
   Icon
 '  !Wpf.Ui.Markup.SymbolIconExtension  *DocumentSplitHint20+.@�m�  m� CommandParameter$ newFile��$	��New�� m� Command �� *" ViewModel.StatusBarActionCommand+ �� Mode$ OneWay=�m� $ 	newWindow��$��
New window�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=�m� $ openFile��$��Open...�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=�m� $ saveFile��$
��Save�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=�m� $ 
saveFileAs��$��
Save As...�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=��� m� $
 exit��$
��Exit�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=�  $
��Edit��
   *DocumentEdit20+.@�m� $ editUndo��$
��Undo�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=��� m� $ editCut��$	��Cut�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=�m� $ editCopy��$
��Copy�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=�m� $ 	editPaste��$��Paste�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=�m� $ 
editDelete��$��Delete��${�False�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=��� m� $ browserSearch��$��Search with browser�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=�m� $
 find��$��Find...�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=�m� $ findNext��$��	Find next�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=��� m� $ 	selectAll��$��
Select All�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=���   $ textBold�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=�
   *
TextBold20+  $ 
textItalic�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=�
   *TextItalic20+  $ textUnderline�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=�
   *TextUnderline20+��   $ textFont�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=�
   *
TextFont20+��   $��Format��
   *
ScanText24+.@�m� $ wordWrap��$��	Word wrap�� m� IsCheckable$
 True�� m� 	IsChecked$ False�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=�  $
��View��
   *CalendarWeekStart24+.@�m� $
��Zoom��.@�m� $ zoomIn��$��Zoom in�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=�m� $ zoomOut��$��Zoom out�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=�m� $ zoomRestore��$��Restore default zoom�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=�m� $ 	statusBar��$��
Status bar��$
 True�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=� �� *ViewModel.IsWordWrapEnbaled+$ TwoWay=���   $
��Help��
   *
ChatHelp20+.@�m� $ viewHelp��$��	View help�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=�m� $ viewFeedback��$��Send feedback�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=��� m� $ 	viewAbout��$��About WPF UI�� �� *" ViewModel.StatusBarActionCommand+$ OneWay=��� -    ��Name$ RootTextBox��$��2��$	��0,6q�
������ $��0q�" �� VerticalScrollBarVisibility$ Visible=� �/�  ! ContentControlThemeFontFamily /� 
FontFamily# [ T�J� .*�.*�qoEvil Is Evil. Lesser, Greater, Middling, Makes No Difference. The Degree Is Arbitrary, The Definitions Blurred.�� 0.If I'm To Choose Between One Evil And Another,�� .�I'd Rather Not Choose At All.�� -    ��Name$ RootStatusBar��$��3������ .R��  � ColumnDefinitions �� $
��Auto ��� $
��Auto ��� $��* ��� $
��Auto ��� $
��Auto �.@��� .���� ��`�  `� Mode$ OneWay=� �� StringFormat$ Line: {0}, Char: {1}��.1���  �� Path$
 Line��� $ 	Character��� $��1���� $��2��.���� ���� *ViewModel.CurrentlyOpenedFile+$ OneWay=��� $��3���� $��4��.��� $��90�� � Value �� *ViewModel.Progress+$ OneWay=�