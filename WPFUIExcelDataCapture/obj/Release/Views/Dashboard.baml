   M S B A M L   `   `   `  ���� P  LWPFUIExcelDataCapture, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null=#clr-namespace:WPFUIExcelDataCaptureWPFUIExcelDataCapture  &     WPFUIExcelDataCapture.MainWindow   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     V RMaterialDesignThemes.Wpf, Version=3.1.0.1244, Culture=neutral, PublicKeyToken=nullFmaterialDesign1http://materialdesigninxaml.net/winfx/xaml/themes  /local#clr-namespace:WPFUIExcelDataCapture    � Title$  
MainWindow��$	��800��$
��1000�� � 
ResizeMode$ NoResize=� � WindowStyle$
 None=� � WindowStartupLocation$ CenterScreen=�$��Segoe UI Symbol)�.��� -   ���#FFFFFF � ColumnDefinitions �� $��10 ��� $��* ��� $��10 � � RowDefinitions �� $��10 ��� $��50 ��� $��* ��� $��10 �.+�� $��0��$��0��$��4�����#3A5321� $��1��$��0��$��3�����#9BC371��  �� Orientation$ 
Horizontal=�$��1��$��1��$��Right=�$��40��$	��Top=�.+��� -    ��Name$ BtnExit�� 0  ,MaterialDesignFloatingActionMiniAccentButton#��[  $��25��$��25��$��5,8,5,0q����#9BC371$	��Top=���Z� .��'  !MaterialDesignThemes.Wpf.PackIcon     Kind$ 	ExitToApp=�$��15��$��15��
��������� $ Vertical=�$��1��$��1��$��3��$	��200�����#9BC371$
��Left=�.+��� $ 
Horizontal=�$��50��$��5 0 0 25q�$
��Left=�.+��� -   $ btnHome��#��[  ��Z� ��Z� .��  $
 Home=�$��30��$��30��
��������� -   $ btnData��#��[  ��Z� ��Z� .��  $ DatabaseImport=�$��30��$��30��
��������� -   $ btnSettings��#��[  ��Z� ��Z� .��  $ Settings=�$��30��$��30��
��������� -   	 ��Name$	 MenuListView��
�������$��5 0 5 0q�.@��� -   
 ��Name$
 listViewCompare��$��40��.���� $ 
Horizontal=�.+�  $ Compare=�$��30��$��30��$��Center=�$
��Left=��� $��Compare Worksheets��$��Right=�$��Center=�$��20 0 0 0q��� -   $
 listViewMerge��$��40��.���� $ 
Horizontal=�.+�  $ Merge=�$��30��$��30��$��Center=�$
��Left=��� $��Merge��$��Right=�$��Center=�$��20 0 0 0q��� -	    ��Name$ DashboardContent��$��1��$��2��$��Right=�$	��780��