   M S B A M L   `   `   `  ����D  @inventory, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null5clr-namespace:inventory.Helpersinventory.Helpers  5   	   (    "inventory.View.ConfigurationWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   3helpersclr-namespace:inventory.Helpers     5   	   -   5          �Name$  window��5   �    � Title$ Configuration Window��5   
   $	��300��6'   $	��400��64   $��-8,0,0,0q�6@   $��Georgia)�6R    � WindowStartupLocation$ CenterScreen=�5   9    � ShowInTaskbar .     ModalWindowStyle#���   5   `   $��0��6�   ���	#FFFFFFFF
�������$   inventory.Helpers.DialogCloser   DialogResult 6	   �� *Close+ J� CommandBindings 5      )  #System.Windows.Input.CommandBinding  5   
   -   5           Command$ Close��5      6)   .��5      �  � RowDefinitions 5   
   �� 5      $��30 �6   �� 5      $��30 �6   �� 5      $��30 �6   �� 5      $��30 �6   �� 5      $��30 �6   �� 5      $��30 �6   �� 5      $��* �6    � ColumnDefinitions 5   
   �� 5      $	��115 �6   �� 5      $��* �6   .+�5   
   �� $
��Left=�6   $
��Bold"�6/   $
��Wrap=�6A   $��Configure Settings��6U   $	��Top=�6o   	 J� RenderTransformOrigin$	 2.806,1.188&�6�   $��1��6�   $��	37,10,0,0q�6�   �� 5   
   $
��Left=�6   $
��Wrap=�60   $��Server Name��6D   $��Bottom=�6W   $	 2.806,1.188&�6r   $��20,0,0,6q�6�   $��2��6�   �� 5   
   $
��Left=�6   $��18��6/   $	��185��6;   
 �� TextWrapping$

 Wrap=�6G   $	��Top=�6w   $��1��6�   $��	23,10,0,0q�6�   $��2��6�   ��6[   �� *
ServerName+6�   �� 5   
   $
��Left=�6   $��18��6.   $	��185��6:   $

 Wrap=�6F   $	��Top=�6x   $��1��6�   $��	23,10,0,0q�6�   $��3��6�   ��6Z   �� *DatabaseName+6�   �� 5    
   $
��Left=�6   $
��Wrap=�60   $��	User Name��6D   $��Bottom=�6U   $	 2.806,1.188&�6p   $��29,0,0,6q�6�   $��4��6�   �� 5!   
   $
��Left=�6   $��18��6.   $	��185��6:   $

 Wrap=�6F   $	��Top=�6t   $��1��6�   $��	23,10,0,0q�6�   $��4��6�   ��6Z   �� *
UserName+6�   �� 5"   
   $
��Left=�6   $
��Wrap=�60   $��Password��6D   $��Bottom=�6T   $	 2.806,1.188&�6o   $��	38,0,0,10q�6�   $��5��6�   �� 5#   
   $
��Left=�6   $��18��6.   $	��185��6:   $

 Wrap=�6F   $	��Top=�6t   $��1��6�   $��23,6,0,0q�6�   $��5��6�   ��6Z   �� *
Password+6�   �� 5$   
   $��Test Connection��6   $
��Left=�6L   $	��Top=�6g   $��85��6   $��1��6�   $��	23,10,0,0q�6�   $��6��6�   ��6+   �� *ClickCommand+6�   �� 5%   
   -   5        $��Cancel��5%      $
��Left=�6"   $��75��6=   $��1��6H   $��
123,10,0,0q�6X   $��6��6l   $��20��6y   $	��Top=�6�   6�   �� 5&   
   $
��Left=�6   $
��Wrap=�60   $��Database Name��6D   $��Bottom=�6Y   $	 2.806,1.188&�6t   $��	19,0,0,10q�6�   $��3��6�   5(      5)      