   M S B A M L   `   `   `  ����D  @inventory, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null5clr-namespace:inventory.Helpersinventory.Helpers  5   	       inventory.View.Login   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   3helpersclr-namespace:inventory.Helpers     5   	   -   5          �Name$  window��5   h    � Title$ Login��5   	   $	��220��6   $	��400��6$   $��-8,0,0,0q�60   $��Georgia)�6B    � WindowStartupLocation$ CenterScreen=�5   9    � ShowInTaskbar .     ModalWindowStyle#���   5   =   $��0��6x   ���	#FFFFFFFF
�������$   inventory.Helpers.DialogCloser   DialogResult 6	   �� *Close+ � 	Resources 5        StoryboardEntrance g� 	BeginTime �� TargetProperty �� 
TargetName	 O� KeyTime
 O� Value �� 	KeySpline��%�  &	       �� 5   
   .��5	      U� $ 00:00:00X�6,   $U O(UIElement.RenderTransform).(TransformGroup.Children)[3].(TranslateTransform.Y)�6b    6A   Z� .\�5
      �� $	 00:00:00��6'   $	
 230Q�6:   $ 1,0,0,1��6F   �� 5      $	 00:00:00.5000000��6'   $	
 -30Q�6B   $ 1,0,0,1��6N   �� 5      $	 00:00:00.7000000��6'   $
 0Q�6B   $ 0,0,0,1��6L   5      5      5      z�5      V� 5   
   .~�5      ��  �� ScaleX$ 1Q�6    �� ScaleY$ 1Q�6(   �� 5       �� AngleX$ 0Q�6    �� AngleY$ 0Q�6'   �� 5       �� Angle$ 0Q�6   T� 5       T� X$ 0Q�6!    T� Y$ 0Q�6'   5       � Triggers 5      :� 5   
    :� RoutedEvent$ FrameworkElement.Loaded��6   .[�5      �� #��[ 6   5       J� CommandBindings 5      )  #System.Windows.Input.CommandBinding  5   
   -   5           Command$ Close��5      5   	   .��5!      �  � RowDefinitions 5"   
   �� 5#      $��30 �6   �� 5$      $��30 �6   �� 5%      $��30 �6   �� 5&      $��30 �6    � ColumnDefinitions 5(   
   �� 5)      $	��115 �6   �� 5*      $	��285 �6   .+�5,   
   �� $��	User Name��6   $
��Left=�6$   $	��Top=�6?   $��36,5,0,0q�6W    J� RenderTransformOrigin$ 0.872,1.417&�6i   $��1��6�   �� 5-   
   $��Password��6   $
��Left=�6#   $	��Top=�6>   $��44,3,0,0q�6V   $ 0.872,1.417&�6h   $��2��6�   �� 5.   
   -   5        $��txt_UserName��5.      $
��Left=�6&   $��20��6j    �� TextWrapping$
 Wrap=�6v   $	��Top=�6�   $	��155��6�   $��30,5,0,0q�6�   $��1��6�   $��1��6�   ��6A   �� *
UserName+ �� Mode$ TwoWay=�6�   G� 5/   
   -   5        $��txt_Password��5/   1   $
��Left=�6   $	��Top=�6E   $��20��6]   $��30,3,0,0q�6i   $	��155��6{   $��1��6�   $��2��6�   �� 50   
   $��Login��6   $
��Left=�6!   $	��Top=�6�   $��75��6�   $��	30,11,0,0q�6�   $ 
0.267,1.05&�6�   $��1��6�   $��3��6�   ��6>   �� *ClickCommand+��6_   ��  �� ElementName$ txt_Password��6�   �� 51   
   $��Cancel��6   $
��Left=�6"   $	��Top=�6`   $��75��6x   $��
136,11,0,0q�6�   $ 
0.267,1.05&�6�   $��1��6�   $��3��6�   ��6?   �� *ClickCancel+6�   �� 52   
   $
��Left=�6   $
��Wrap=�6/   $	��Top=�6D   $��1��6\   $��30,7,0,0q�6l   �� 53   
   
����ح� �� Command 6,   �� *ClickConfig+.�6M   Click for Configuration Setup54      55      56      57      