# App-specific names for Window.Tests.ps1.
$script:AppConfig = @{
    Title         = 'Maintenance Window Manager'
    Shell         = 'start-maintenancewindowmgr.ps1'
    Xaml          = 'MainWindow.xaml'
    OptionsButton = 'btnOptions'
    LogLabel      = 'lblLogOutput'
    Views         = @(
        @{ Name = 'Windows';   Container = 'viewWindows';   Button = 'btnViewWindows' }
        @{ Name = 'Coverage';  Container = 'viewCoverage';  Button = 'btnViewCoverage' }
        @{ Name = 'Templates'; Container = 'viewTemplates'; Button = 'btnViewTemplates' }
    )
}
