   M S B A M L   `   `   `  ����  	inventory    inventory.View.AddDealers   5      R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5      @d2http://schemas.microsoft.com/expression/blend/2008     5      .��5      �   � RowDefinitions  5	   
   �� 5
      $��30 �6   �� 5      $��30 �6   �� 5      $��30 �6   �� 5      $��30 �6   �� 5      $��30 �6   �� 5      $��* �6    � ColumnDefinitions 5   
   �� 5      $	��115 �6   �� 5      $��* �6   .+�5   
   �� $
��Left=�6   $
��Wrap=�6/   $��Add Dealers��6C   $
��Bold"�6V   $	��Top=�6h    J� RenderTransformOrigin$ 3.7,1&�6�   $��1��6�   $��
110,10,0,0q�6�   �� 5   
   $��Dealer Name��6   $
��Left=�6&   $	��Top=�6A   $��23,3,0,0q�6Y   $��2��6k   �� 5   
   $��Address��6   $
��Left=�6"   $	��Top=�6=   $��23,4,0,0q�6U   $��3��6g   $ 0.528,0.808&�6t   �� 5   
   -   5        $��txt_DealerID��5   -   $
��Left=�6   $��18��6A    �� TextWrapping$
 Wrap=�6M   $	��Top=�6�   $	��150��6�   $��1��6�   $��20,5,0,0q�6�   $��1��6�   $y�Hidden=�6�   ��6a   ��  �� ElementName$ Delarlst�� �� Path$ SelectedItem.id�6�   �� 5   
   $
��Left=�6   $��18��6-   $
 Wrap=�69    �� MaxLines$ 1��6M   $	��150��6w   $��1��6�   $��20,5,0,2q�6�   $��2��6�   ��6[   �� *
DealerName+6�   �� 5   
   $
��Left=�6   $��18��6-   $ 1��6:   $	��Top=�6g   $	��150��6   $��1��6�   $��20,5,0,0q�6�   $��3��6�   ��6H   �� *DealerAddress+6�   �� 5   
   $��
Add Dealer��6   $
��Left=�6&   $	��Top=�6�   $��75��6�   $��1��6�   $��20,5,0,0q�6�   $��4��6�   ��6A   �� *
AddCommand+��6`   �� $ txt_DealerID��6�   5      5      