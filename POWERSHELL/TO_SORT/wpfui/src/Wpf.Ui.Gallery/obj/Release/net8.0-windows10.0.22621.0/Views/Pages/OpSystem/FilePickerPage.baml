   M S B A M L   `   `   `  ���� I  EWpf.Ui.Gallery, Version=3.0.0.0, Culture=neutral, PublicKeyToken=nullA%clr-namespace:Wpf.Ui.Gallery.ControlsWpf.Ui.Gallery.Controls  Y1clr-namespace:Wpf.Ui.Gallery.Views.Pages.OpSystem#Wpf.Ui.Gallery.Views.Pages.OpSystem  8    2Wpf.Ui.Gallery.Views.Pages.OpSystem.FilePickerPage   W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     :controls%clr-namespace:Wpf.Ui.Gallery.Controls     @d2http://schemas.microsoft.com/expression/blend/2008     Clocal1clr-namespace:Wpf.Ui.Gallery.Views.Pages.OpSystem     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     A =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null/ui&http://schemas.lepo.co/wpfui/2022/xaml  6   0Wpf.Ui.Gallery.Controls.PageControlDocumentation    Show$  False��   ApplicationBackgroundBrush  Wpf.Ui.Markup.Design   
Background# �     TextFillColorPrimaryBrush   
Foreground# �   N� 
Foreground# �  .���� $��0,0,0,24q�.+�,   &Wpf.Ui.Gallery.Controls.ControlExample  $��0q�   
CsharpCode$, &OpenFileDialog openFileDialog = new();��   
HeaderText$ Pick a single file��   ExampleContent. �  � ColumnDefinitions �� $
��Auto ��� $
��Auto �.+�  Wpf.Ui.Controls.Button  $��0��$��Open a file������ *ViewModel.OpenFileCommand+�� $��1��$��16,0,0,0q�$��Center=� �� Orientation$ 
Horizontal=�y��� *$"ViewModel.OpenedFilePathVisibility+.+�  Wpf.Ui.Controls.TextBlock  $��Picked file:��  $��4,0,0,0q�	   FontTypography$	 
BodyStrong=����� *ViewModel.OpenedFilePath+  $��0,32,0,0q�$, &OpenFileDialog openFileDialog = new();��$ Pick a specific file type��. �  �� $
��Auto ��� $
��Auto �.+�  $��0��$��Open a picture������ *ViewModel.OpenPictureCommand+�� $��1��$��16,0,0,0q�$��Center=�$ 
Horizontal=�y��� *'%ViewModel.OpenedPicturePathVisibility+.+�  $��Picked photo:��  $��4,0,0,0q�$	 
BodyStrong=����� *ViewModel.OpenedPicturePath+  $��0,32,0,0q�$, &OpenFileDialog openFileDialog = new();��$ Pick multiple files��. �  �� $
��Auto ��� $
��Auto �.+�  $��0��$��Open multiple files������ *ViewModel.OpenMultipleCommand+�� $��1��$��16,0,0,0q�$��Center=�y��� *(&ViewModel.OpenedMultiplePathVisibility+.+�  $��Picked files:��  $	 
BodyStrong=����� *ViewModel.OpenedMultiplePath+  $��0,32,0,0q�$, &OpenFileDialog openFileDialog = new();��$ Pick a folder��. �  �� $
��Auto ��� $
��Auto �.+�  $��0��$��Open a folder������ *ViewModel.OpenFolderCommand+�� $��1��$��16,0,0,0q�$��Center=�y��� *&$ViewModel.OpenedFolderPathVisibility+.+�  $��Picked folder:��  $	 
BodyStrong=����� *ViewModel.OpenedFolderPath+  $��0,32,0,0q�$O Iawait File.WriteAllTextAsync(filePath, FileToSaveContents, cancellation);��$ Save a file��. � 
 � RowDefinitions
 �� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto �.+�  $��0��$%��Prompt the user to save a file.��� $��1��$��0,16,0,0q� �� $
��Auto ��� $
��Auto ��� $
��Auto �.+�  $��0��$��80��$��Center=�$��
File name:��  Wpf.Ui.Controls.TextBox  $��1��$	��180��   PlaceholderText$! Type your file name here...������ *ViewModel.FileToSaveName+  $��2��$��4,0,0,0q�$��Center=�$
��.txt��� $��2��$��0,8,0,0q� �� $
��Auto ��� $
��Auto �.+�  $��0��$��80��$��Center=�$��File content:��  $��1��$	��210��$% Type your file contents here...������ *ViewModel.FileToSaveContents+� $��3��$��0,16,0,0q� �� $
��Auto ��� $
��Auto �.+�  $��0��$��Save a file������ *ViewModel.SaveFileCommand+  $��1��$��16,0,0,0q�$��Center=����� *ViewModel.SavedFileNotice+y��� *%#ViewModel.SavedFileNoticeVisibility+