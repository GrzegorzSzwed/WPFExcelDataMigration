   M S B A M L   `   `   `  ����P  LWPFUIExcelDataCapture, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null=#clr-namespace:WPFUIExcelDataCaptureWPFUIExcelDataCapture  5   	   &     WPFUIExcelDataCapture.MainWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   V RMaterialDesignThemes.Wpf, Version=3.1.0.1244, Culture=neutral, PublicKeyToken=nullFmaterialDesign1http://materialdesigninxaml.net/winfx/xaml/themes  5   	   /local#clr-namespace:WPFUIExcelDataCapture  5   	     � Title$  
MainWindow��5	   	   $	��800��6   $
��1000��6)    � 
ResizeMode$ NoResize=�5
   	    � WindowStyle$
 None=�6    � WindowStartupLocation$ CenterScreen=�5   	   $��Segoe UI Symbol)�6.   .��5      � -   5        ���#FFFFFF � ColumnDefinitions 5   
   �� 5      $��2 �6   �� 5      $��* �6   �� 5      $��2 �6    � RowDefinitions 5   
   �� 5      $��2 �6   �� 5      $��50 �6   �� 5      $��* �6   �� 5      $��2 �6   .+�5   
   � $��0��6   $��0��6   $��4��6,   ���#9514E66@   � 5   
   $��3��6   $��3��6   $��4��6,   ���#9514E66@   � 5   
   $��0��6   $��0��6   $��3��6,   ���#9514E66=   � 5   
   $��0��6   $��3��6   $��3��6,   ���#9514E66=   �� 5   
    �� Orientation$ 
Horizontal=�6   $��1��6.   $��1��6>   $��Right=�6K   $	��Top=�6g   .+�5      �� -   5         ��Name$ BtnExit��5       0  ,MaterialDesignFloatingActionMiniAccentButton#��[  5      $��5 5 5 5q�6r   ���#406309$	��Top=�5   &   ��5   [   Z� .��5      '  !MaterialDesignThemes.Wpf.PackIcon     Kind$ 	ExitToApp=�6*   $��30��6;   $��30��6F   
�������6R   5       5!      �� 5"   
   $ 
Horizontal=�6   $��1��6.   $��1��6>   $
��Left=�6K   $
��auto��6f   $��5 5 5 5q�6s   .+�5#      �� -   5        $ btnHome��5#      #��[  6&   ���#406309$��5 0 5 0q�5$   [   ��6.   Z� .��5%        $
 Home=�6.   $��30��6:   $��30��6E   
�������6Q   5&      �� 5'      -   5        $ btnData��5'      #��[  6&   ���#406309$��5 0 5 0q�5(   [   ��6.   Z� .��5)        $ DatabaseImport=�6.   $��30��6D   $��30��6O   
�������6[   5*      �� 5+      -   5        $ btnSettings��5+      #��[  6*   ���#406309$��5 0 5 0q�5,   _   ��6.   Z� .��5-        $ Settings=�6*   $��30��6:   $��30��6E   
�������6Q   5.      5/      �� 50   
   -   5        	 ��Name$	 DashboardContent��50      $��1��63   $��2��6C   51      52      