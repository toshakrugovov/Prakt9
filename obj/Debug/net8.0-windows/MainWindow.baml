   M S B A M L   `   `   `  ����  
WordLekcia    WordLekcia.MainWindow   5   	   _ [MaterialDesignThemes.Wpf, Version=5.0.0.0, Culture=neutral, PublicKeyToken=df2a72020bd7962aFmaterialDesign1http://materialdesigninxaml.net/winfx/xaml/themes  5   	   _ [Microsoft.Xaml.Behaviors, Version=1.1.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a3i+http://schemas.microsoft.com/xaml/behaviors  5   	   W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	     � Title$  
MainWindow��5	   	   $	��437��6   $	��816��6)   .��5
      �  � RowDefinitions 5   
   �� 5      $
��Auto �6   �� 5      �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   .+�5   
   `� .�5      c� .@�5      �� $��ApplicationCommands.Cut��6   ��6;   0  *MaterialDesignThemes.Wpf.PackIconExtension     Kind$ 
ContentCut=��� 5      $��ApplicationCommands.Copy��6   ��6<     $ ContentCopy=��� 5      $��ApplicationCommands.Paste��6   ��6=     $ ContentPaste=��� 5      �� 5      $��ApplicationCommands.Undo��6   ��6<     $
 Undo=��� 5      $��ApplicationCommands.Redo��6   ��6<     $
 Redo=��� 5      �� 5      $ ��EditingCommands.ToggleBold��6   ��6>     $ 
FormatBold=��� 5      $"��EditingCommands.ToggleItalic��6   ��6@     $ FormatItalic=��� 5      $%��EditingCommands.ToggleUnderline��6   ��6C     $ FormatUnderline=��� 5      �� 5       $&�� EditingCommands.IncreaseFontSize��6   $��А🠕��6D   $��15(�6R   �� 5!      $&�� EditingCommands.DecreaseFontSize��6   $��А🠗��6D   $��15(�6R   �� 5"      �� 5#      $��EditingCommands.AlignLeft��6   ��6=     $ FormatAlignLeft=��� 5$      $!��EditingCommands.AlignCenter��6   ��6?     $ FormatAlignCenter=��� 5%      $ ��EditingCommands.AlignRight��6   ��6>     $ FormatAlignRight=��� 5&      $"��EditingCommands.AlignJustify��6   ��6@     $ FormatAlignJustify=��� 5'      �� 5(      $)��#EditingCommands.IncreaseIndentation��6   ��6G     $ FormatIndentIncrease=��� 5)      $)��#EditingCommands.DecreaseIndentation��6   ��6G     $ FormatIndentDecrease=�5*      5+      �� 5,   
   -   5         ��Name$	 rtb��5,      $��1��6$    � 	Resources 5-       �� 
TargetType �� Property �� Value��%X   'J�       �� 5.      " J�6   .�5/      ��  �@��Margin$ 0q�6/   50      51      . �52      /� -   5         /�Name$ MyDoc��52      .T�53      J� .*�6   .*�54      �� $��50��6   $C��=https://pngicon.ru/file/uploads/1303507165_Cherry-256x256.png��6(   55      �� 56      .�57      �� .�58      �� .�59      �� .�5:   "   J� .*�6,   Ряд 0 Столбец 06=   5;      �� 5<      .�5=   "   J� .*�6,   Ряд 0 Столбец 16=   5>      �� 5?      .�5@   "   J� .*�6,   Ряд 0 Столбец 26=   5A      5B      �� 5C      .�5D      �� .�5E   "   J� .*�6,   Ряд 1 Столбец 06=   5F      �� 5G      .�5H   "   J� .*�6,   Ряд 1 Столбец 16=   5I      �� 5J      .�5K   "   J� .*�6,   Ряд 1 Столбец 26=   5L      5M      5N      5O      �� 5P      	 �� MarkerOffset$	 25��6   
 �� MarkerStyle$
 
UpperRoman=�6)    �� 
StartIndex$ 5��6B   .<�5Q      �� .9�5R      J� .*�6$   Красный6-   5S      �� 5T      .9�5U      J� .*�6$   Оранжевый6/   5V      �� 5W      .9�5X      J� .*�6$   Желтый6,   5Y      �� 5Z      .9�5[      J� .*�6$   Зеленый6-   5\      �� 5]      .9�5^      J� .*�6$   Голубой6-   5_      �� 5`      .9�5a      J� .*�6$   
Синий6+   5b      �� 5c      .9�5d      J� .*�6$   Фиолетовый60   5e      5f      J� 5g      .*�6   .*�5h      �� .�6   
sdfsdfsd6$   5i      5j      5k      �� 5l   
   -   5        $!�� Импортировать��5l      $��2��6*   67   �� 5m   
   -   5        $#�� Экспортировать��5m      $��3��6+   68   �� 5n   
   -   5        $��Эксель��5n      $��4��6"   6/   �� 5o   
   -   5        $��Отправить��5o      $��5��6%   62   5p      5q      