   M S B A M L   `   `   `  ����D  @inventory, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null9!clr-namespace:inventory.ViewModelinventory.ViewModel  5      "    inventory.View.ModifyDealers   5      R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5      @d2http://schemas.microsoft.com/expression/blend/2008     5      7	ViewModel!clr-namespace:inventory.ViewModel     5      
������ .��5	      �   � RowDefinitions  5
   
   �� 5      $��30 �6   �� 5      $��30 �6   �� 5      $��30 �6   �� 5      $��30 �6   �� 5      $��30 �6   �� 5      $��* �6    � ColumnDefinitions 5   
   �� 5      $	��115 �6   �� 5      $��* �6   .+�5   
   �� $
��Left=�6   $
��Wrap=�6/   $��Modify Dealers��6C   $
��Bold"�6Y   $	��Top=�6k    J� RenderTransformOrigin$ 3.7,1&�6�   $��1��6�   $��
110,10,0,0q�6�   �� 5   
   $��Dealer Name��6   $
��Left=�6&   $	��Top=�6A   $��23,3,0,0q�6Y   $��2��6k   �� 5   
   $��Address��6   $
��Left=�6"   $	��Top=�6=   $��23,4,0,0q�6U   $��3��6g   $ 0.528,0.808&�6t   �� 5   
   -   5        $��txt_DealerID��5   -   $
��Left=�6   $��23��6A    �� TextWrapping$
 Wrap=�6M   $	��Top=�6�   $	��150��6�   $��1��6�   $��20,5,0,0q�6�   $��1��6�   $y�Hidden=�6�   ��6a   ��  �� ElementName$ Delarlst�� �� Path$ SelectedItem.id�6�   �� 5   
   $
��Left=�6   $��18��6-   $
 Wrap=�69    �� MaxLines$ 1��6M   $	��Top=�6�   $	��150��6�   $��1��6�   $��20,5,0,0q�6�   $��2��6�   ��6[   �� $ Delarlst��$ SelectedItem.dealer_name�6�   �� 5   
   $
��Left=�6   $��18��6-   $
 Wrap=�69   $ 1��6M   $	��Top=�6�   $	��150��6�   $��1��6�   $��20,5,0,0q�6�   $��3��6�   ��6[   �� $ Delarlst��$! SelectedItem.dealer_address�6�   �� 5   
   $��Update��6   $
��Left=�6"   $	��Top=�6�   $��75��6�   $��1��6�   $��20,5,0,0q�6�   $��4��6�   ��6=   �� *UpdateCommand+��6_   �� $ txt_DealerID��6�   �� 5   
   $��Delete��6   $
��Left=�6"   $	��Top=�6�   $��75��6�   $��1��6�   $��	110,5,0,0q�6�   $��4��6�   $ 0.507,1&�6�   ��6=   �� *DeleteCommand+��6_   �� $ txt_DealerID��6�   �� 5   
   -   5        $��Delarlst��5      $��5��6#   $��1��6t   $��20,19,10,10q�6�   ��61   �� *Delars+ �� UpdateSourceTrigger$ PropertyChanged=� �� View 5       �� 5!      -   5        grdDelar��.P�5"      �� $��Delar ID��6%      GridHeaderLeft	 �� HeaderContainerStyle#	 �   67   
 �� Width$
 75��6�    �� DisplayMemberBinding 6q   �� *id+5#      �� 5$      $��
Delar Name��6%   #	 �   69   $	
 130��6�    6q   �� *dealer_name+5%      �� 5&      $��Delar Address��6%   #	 �   6<   $	
 190��6�    6t   �� *dealer_address+6�   5'      5)      5+      5,      