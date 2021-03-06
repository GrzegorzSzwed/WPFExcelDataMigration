﻿#pragma checksum "..\..\..\Views\Dashboard.xaml" "{8829d00f-11b8-4213-878b-770e8597ac16}" "806D0F616A047DDD7A98BBEC5AAF0BBE16FEF72550D3CA0588DEE2FA5AF17969"
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

using MaterialDesignThemes.Wpf;
using MaterialDesignThemes.Wpf.Converters;
using MaterialDesignThemes.Wpf.Transitions;
using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Media.TextFormatting;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Shell;
using WPFUIExcelDataCapture;


namespace WPFUIExcelDataCapture {
    
    
    /// <summary>
    /// MainWindow
    /// </summary>
    public partial class MainWindow : System.Windows.Window, System.Windows.Markup.IComponentConnector {
        
        
        #line 28 "..\..\..\Views\Dashboard.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button BtnExit;
        
        #line default
        #line hidden
        
        
        #line 34 "..\..\..\Views\Dashboard.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button btnHome;
        
        #line default
        #line hidden
        
        
        #line 38 "..\..\..\Views\Dashboard.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button btnData;
        
        #line default
        #line hidden
        
        
        #line 42 "..\..\..\Views\Dashboard.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button btnSettings;
        
        #line default
        #line hidden
        
        
        #line 47 "..\..\..\Views\Dashboard.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ListView MenuListView;
        
        #line default
        #line hidden
        
        
        #line 48 "..\..\..\Views\Dashboard.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ListViewItem listViewCompare;
        
        #line default
        #line hidden
        
        
        #line 54 "..\..\..\Views\Dashboard.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ListViewItem listViewMerge;
        
        #line default
        #line hidden
        
        
        #line 62 "..\..\..\Views\Dashboard.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ContentControl DashboardContent;
        
        #line default
        #line hidden
        
        private bool _contentLoaded;
        
        /// <summary>
        /// InitializeComponent
        /// </summary>
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        public void InitializeComponent() {
            if (_contentLoaded) {
                return;
            }
            _contentLoaded = true;
            System.Uri resourceLocater = new System.Uri("/WPFUIExcelDataCapture;component/views/dashboard.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\..\Views\Dashboard.xaml"
            System.Windows.Application.LoadComponent(this, resourceLocater);
            
            #line default
            #line hidden
        }
        
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily")]
        void System.Windows.Markup.IComponentConnector.Connect(int connectionId, object target) {
            switch (connectionId)
            {
            case 1:
            
            #line 12 "..\..\..\Views\Dashboard.xaml"
            ((System.Windows.Controls.Grid)(target)).MouseDown += new System.Windows.Input.MouseButtonEventHandler(this.Grid_MouseDown);
            
            #line default
            #line hidden
            return;
            case 2:
            this.BtnExit = ((System.Windows.Controls.Button)(target));
            
            #line 28 "..\..\..\Views\Dashboard.xaml"
            this.BtnExit.Click += new System.Windows.RoutedEventHandler(this.BtnExit_Click);
            
            #line default
            #line hidden
            return;
            case 3:
            this.btnHome = ((System.Windows.Controls.Button)(target));
            
            #line 35 "..\..\..\Views\Dashboard.xaml"
            this.btnHome.Click += new System.Windows.RoutedEventHandler(this.BtnHome_Click);
            
            #line default
            #line hidden
            return;
            case 4:
            this.btnData = ((System.Windows.Controls.Button)(target));
            
            #line 39 "..\..\..\Views\Dashboard.xaml"
            this.btnData.Click += new System.Windows.RoutedEventHandler(this.BtnData_Click);
            
            #line default
            #line hidden
            return;
            case 5:
            this.btnSettings = ((System.Windows.Controls.Button)(target));
            
            #line 43 "..\..\..\Views\Dashboard.xaml"
            this.btnSettings.Click += new System.Windows.RoutedEventHandler(this.BtnSettings_Click);
            
            #line default
            #line hidden
            return;
            case 6:
            this.MenuListView = ((System.Windows.Controls.ListView)(target));
            return;
            case 7:
            this.listViewCompare = ((System.Windows.Controls.ListViewItem)(target));
            return;
            case 8:
            this.listViewMerge = ((System.Windows.Controls.ListViewItem)(target));
            return;
            case 9:
            this.DashboardContent = ((System.Windows.Controls.ContentControl)(target));
            return;
            }
            this._contentLoaded = true;
        }
    }
}

