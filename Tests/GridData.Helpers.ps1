# Helpers for the grid data tests (Grid.Rows.Tests.ps1, Grid.Cells.Window.Tests.ps1).
# A grid view can show empty or stale cells when a column binds to a property name that the rows do not have, or when
# the rows never reach the control. These helpers read the bound columns of every DataGrid in a XAML file.

# One object per bound column of every DataGrid in the XAML files: Grid (x:Name), Header, Path (the bound property),
# Xaml (the file that holds the grid).
function Get-XamlGridColumn {
    param([Parameter(Mandatory)][string[]]$XamlPath)
    foreach ($file in $XamlPath) {
        [xml]$doc = Get-Content -LiteralPath $file -Raw
        $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
        $ns.AddNamespace('w', 'http://schemas.microsoft.com/winfx/2006/xaml/presentation')
        foreach ($grid in $doc.SelectNodes('//w:DataGrid', $ns)) {
            $gridName = $grid.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml')
            foreach ($column in $grid.SelectNodes('w:DataGrid.Columns/*', $ns)) {
                if ($column.GetAttribute('Binding') -match '^\{Binding\s+(?:Path=)?([A-Za-z_][\w\.]*)') {
                    [pscustomobject]@{ Grid = $gridName; Header = $column.GetAttribute('Header'); Path = $Matches[1]; Xaml = $file }
                }
            }
        }
    }
}
