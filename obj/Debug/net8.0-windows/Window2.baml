   M S B A M L   `   `   `  ����E  AWordLekcia, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'clr-namespace:WordLekcia
WordLekcia  5   	       WordLekcia.Window2   5   	   W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   *localclr-namespace:WordLekcia     5   	     � Title$  OtpravkaPocgtExcel��5   	   $	��450��6$   $	��800��61   .��5	      � $��10q�6    � RowDefinitions 5
   
   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $��* �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   .+�5   
   �� $��ОТ КОГО:��6   $��0��6$   $��Center=�61   �� 5   
   -   5         ��Name$
 From��5      $��0��6    $��	100,0,0,5q�6-   �� 5   
   $*��$Пароль отправителя:��6   $��1��6/   $��Center=�6<   G� 5   
   -   5         G�Name$
 Pass��5      $��1��6$   $��	150,0,0,5q�61   �� 5   
   $��	Кому:��6   $��2��6!   $��Center=�6.   �� 5   
   -   5        $ To��5      $��2��6   $��50,0,0,5q�6+   �� 5   
   $"��Тема сообщения:��6   $��3��6+   $��Center=�68   �� 5   
   -   5        $ Subject��5      $��3��6#   $��	150,0,0,5q�60   �� 5    
   $��Сообщение:��6   $��4��6&   $	��Top=�63   $��0,10,0,0q�6K   �� 5!   
   -   5         ��Name$ 
MessegeRtb��5!      $��5��6*   $��0,5,0,0q�67   �� 5#   
   -   5        $5��/Выбрать файл для отправки��5#      $��6��65   $��0,5,0,0q�6a   �� 5$   
   -   5        $��Отправить��5$      $��6��6%   $��Right=�6K   $��0,5,0,0q�6g   5%      5&      