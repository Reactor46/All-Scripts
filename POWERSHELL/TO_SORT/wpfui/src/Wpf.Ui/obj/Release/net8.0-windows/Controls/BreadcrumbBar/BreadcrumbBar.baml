   M S B A M L   `   `   `  ���� A  =Wpf.Ui, Version=3.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Wpf.Ui.ControlsWpf.Ui.Controls  W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089 "  DefaultUiBreadcrumbButtonStyle  �� 
TargetType �� Property & "BreadcrumbBarNormalForegroundBrush �� Value �� 
TargetType ��Name �� Triggers P� Property P� Value �� 
TargetName % !BreadcrumbBarHoverForegroundBrush # DefaultUiBreadcrumbBarItemStyle'    !Wpf.Ui.Controls.BreadcrumbBarItem'   !Wpf.Ui.Controls.IconSourceElement	   
IconSource&    Wpf.Ui.Controls.SymbolIconSource
   FontSize   
FontWeight   Symbol � ColumnDefinitions ��Name �� Path �� Mode �� RelativeSource �� AncestorType#   Wpf.Ui.Controls.BreadcrumbBar ��Name    
IconMargin    Icon - )BreadcrumbBarCurrentNormalForegroundBrush  DefaultUiBreadcrumbBarStyle K� IsItemsHost �� Orientation �� BasedOn�� F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     :x,http://schemas.microsoft.com/winfx/2006/xaml     *controlsclr-namespace:Wpf.Ui.Controls  %a  &	        &	 J    7   &	 �    '       7  '      7  �� "  ��.���  � ��# �  ��  � ��
 ���� ��  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � �� �� " ��.R��� $ 	Presenter��#��z��#��z��#��z��#��z�� P�  �@J�IsMouseOver$
 True��. ��� $ 	Presenter��# �   � ��P�  � {�$ False���� "    .���  � ��$ 14(���  � ��$ DemiBold"���  �@J�SnapsToDevicePixels$
 True����  �@�OverridesDefaultStyle$
 True����  � ��$ Center=���  � ��$ Center=���  �@  
IconMargin$ 10, 0, 10, 0q���  � |�$ False����  �@  Icon   	   $

 18.0Q�$ DemiBold"�$ ChevronRight24=���  � �� �� "   .R��  �� $��* ��� $
��Auto �.+��� $ Button��$��0��#��z��#��z��#��z��#��z��#��z��#��z��8��  ���� $ TemplateButtonCommand�$ OneTime=� �� "  ���� $ ItemTemplate�$ OneTime=� �� "  ���� $ ItemTemplateSelector�$ OneTime=� �� "  �� $
 Icon��$��1��#��z #��z $|�False�� P�  �@  IsLast$
 True��. ��� $ Button��# �   � ���� $ Button�� � {�$ False���� $
 Icon�� � y�$ 	Collapsed=��� "   .���  � �� �� .R��� $
 True��$ 
Horizontal=��� 8   "    �� 8   "   