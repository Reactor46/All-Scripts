   M S B A M L   `   `   `  ���� I  EWpf.Ui.Gallery, Version=3.0.0.0, Culture=neutral, PublicKeyToken=nullA%clr-namespace:Wpf.Ui.Gallery.ControlsWpf.Ui.Gallery.Controls  ]3clr-namespace:Wpf.Ui.Gallery.Views.Pages.BasicInput%Wpf.Ui.Gallery.Views.Pages.BasicInput  U QSystem.Runtime, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a7,clr-namespace:System;assembly=System.RuntimeSystem 6    0Wpf.Ui.Gallery.Views.Pages.BasicInput.SliderPage   W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     :controls%clr-namespace:Wpf.Ui.Gallery.Controls     @d2http://schemas.microsoft.com/expression/blend/2008     Elocal3clr-namespace:Wpf.Ui.Gallery.Views.Pages.BasicInput     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     ?system,clr-namespace:System;assembly=System.Runtime     A =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null/ui&http://schemas.lepo.co/wpfui/2022/xaml    N� Title$  
SliderPage��   ApplicationBackgroundBrush  Wpf.Ui.Markup.Design   
Background# �     TextFillColorPrimaryBrush   
Foreground# �   N� 
Foreground# �  .���� $��0,0,0,24q�.+�,   &Wpf.Ui.Gallery.Controls.ControlExample  $��0q�   
HeaderText$ A simple slider.��   XamlCode$ <Slider Width="200" />��   ExampleContent. �  � ColumnDefinitions �� $��* ��� $
��Auto �.+��� $	��200��$��0q�$
��Left=�$��Center=� � Maximum$	 100Q�	 � Minimum$	 0Q�
 � Value
 �� *ViewModel.SimpleSliderValue+ �� RelativeSource ��  �� Mode$ FindAncestor=� �� AncestorType local:SliderPage �� Mode$ TwoWay=�� $��1��.+��� $��Center=�.+��� $��Output:���� ���� *ViewModel.SimpleSliderValue+ �� $ FindAncestor=� local:SliderPage$ OneWay=�  $��0,32,0,0q�$. (A slider with steps and range specified.��$> 8<Slider TickFrequency="20" IsSnapToTickEnabled="True" />��. �  �� $��* ��� $
��Auto �.+��� $	��200��$��0q�$
��Left=�$��Center=� �� IsSnapToTickEnabled$
 True��$
 1000Q�$		 500Q� �� TickFrequency$ 20Q�
 �� *ViewModel.RangeSliderValue+ �� $ FindAncestor=� local:SliderPage$ TwoWay=�� $��1��.+��� $��Center=�.+��� $��Output:���� ���� *ViewModel.RangeSliderValue+ �� $ FindAncestor=� local:SliderPage$ OneWay=�  $��0,32,0,0q�$ A slider with tick marks.��$% <Slider TickPlacement="Both" />��. �  �� $��* ��� $
��Auto �.+��� $	��200��$��0q�$
��Left=�$��Center=�$
 True��$	 100Q�$	 0Q�$ 20Q� �� TickPlacement$
 Both=�
 �� *ViewModel.MarksSliderValue+ �� $ FindAncestor=� local:SliderPage$ TwoWay=�� $��1��.+��� $��Center=�.+��� $��Output:���� ���� *ViewModel.MarksSliderValue+ �� $ FindAncestor=� local:SliderPage$ OneWay=�  $��0,32,0,0q�$< 6A vertical slider with range and tick marks specified.��$' !<Slider Orientation="Vertical" />��. �  �� $��* ��� $
��Auto �.+��� $	��200��$��0q�$
��Left=�$��Center=�$
 True��$	 100Q�$	 0Q� �� Orientation$ Vertical=�$ 20Q�$
 Both=�
 �� *ViewModel.VerticalSliderValue+ �� $ FindAncestor=� local:SliderPage$ TwoWay=�� $��1��.+��� $��Center=�.+��� $��Output:���� ���� *ViewModel.VerticalSliderValue+ �� $ FindAncestor=� local:SliderPage$ OneWay=�