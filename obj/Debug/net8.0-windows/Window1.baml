   M S B A M L   `   `   `  ����E  AWordLekcia, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'clr-namespace:WordLekcia
WordLekcia  5   	       WordLekcia.Window1   5   	   W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   *localclr-namespace:WordLekcia     5   	     � Title$  Window1��5   	   $	��500��6   $	��800��6&   .��5	      �  � RowDefinitions 5
   
   �� 5      $
��Auto �6   �� 5      �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   .+�5   
   ��  �� Orientation$ 
Horizontal=�6   $��0��6.   $��5q�6;   .+�5      �� -   5         ��Name$ ColumnNameTextBox��5      $	��200��61   $��5q�6=   �� 5      -   5        $%��Добавить столбец��5      $��5q�6N   5      &   System.Windows.Controls.DataGrid  5   
   -   5          Name$ 	griiiiiid��5      $��1��6&      AutoGenerateColumns$
 True��63   �� 5   
   -   5        $��2��5      $"��Экспортировать��6   67   �� 5   
   -   5        $��3��5      $ ��Импортировать��6   66   �� 5   
   -   5        $��4��5      $��Отправить��6   62   5      5      