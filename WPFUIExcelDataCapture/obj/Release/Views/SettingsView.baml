   M S B A M L   `   `   `  ���� P  LWPFUIExcelDataCapture, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullI)clr-namespace:WPFUIExcelDataCapture.ViewsWPFUIExcelDataCapture.Views  .    (WPFUIExcelDataCapture.Views.SettingsView   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     @d2http://schemas.microsoft.com/expression/blend/2008     ;local)clr-namespace:WPFUIExcelDataCapture.Views     V RMaterialDesignThemes.Wpf, Version=3.1.0.1244, Culture=neutral, PublicKeyToken=nullFmaterialDesign1http://materialdesigninxaml.net/winfx/xaml/themes  $��Segoe UI Symbol)�
�������$��12(����#272E22.���   � RowDefinitions  �� $��5 ��� $��* ��� $��* ��� $��5 � � ColumnDefinitions �� $��5 ��� $��* ��� $��5 �.+��� $��1��$��1�� �� Orientation$ 
Horizontal=�$	��600��.+��� $	��750��.@��� .���� $ 
Horizontal=�.+��� $��Source Workbook��$��Center=�$	��150���� -    ��Name$ btnLoadExcelSource�� 0  ,MaterialDesignFloatingActionMiniAccentButton#��[  ��Z� ��Z� .��'  !MaterialDesignThemes.Wpf.PackIcon     Kind$ 	FolderAdd=�$��30��$��30���� -    ��Name$ txtSourceFileName��$��Load Source Workbook��$��Center=�$��20 0 0 0q�$	��180���� -    ��Name$ cmbExcelSource��$	��250��$��20 0 0 0q��� $
��auto��$ Vertical=�.+��� $��Columns��$��Center=��� -   $ txtColumnSourceCount��$��Center=��� $
��auto��$ Vertical=�$��5 0 5 0q�.+��� $
��Rows��$��Center=��� -   $ txtRowSourceCount��$��Center=��� .���� $ 
Horizontal=�.+��� $��Destination Workbook��$��Center=�$	��150���� -   $ btnExcelDestination��#��[  ��Z� ��Z� .��  $ 	FolderAdd=�$��30��$��30���� -   $ txtDestinationFileName��$	��180��$��Load Destination Workbook��$��Center=�$��20 0 0 0q��� -   $ cmbExcelDestination��$	��250��$��20 0 0 0q��� $
��auto��$ Vertical=�.+��� $��Columns:��$��Center=��� -	   $ txtColumnDestinationCount��$��Center=��� $
��auto��$ Vertical=�$��5 0 5 0q�.+��� $
��Rows��$��Center=��� -
   $ txtRowDestinationCount��$��Center=��� .���� $ 
Horizontal=�$
��auto��.+��� $��Worksheet Name Similarity��$��Center=�$��Center=��� -   $ txtSimililarityPercent��$��Center=�$
��auto��$��5 0 5 0q�$��
UltraBlack"�