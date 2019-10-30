# Printer Mapper
# Written by Joshua Woleben
# Written 10/3/19

# GUI Code
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="Print Mapper" Height="300" Width="200" MinHeight="500" MinWidth="400" ResizeMode="CanResizeWithGrip">
    <StackPanel>
        <Label x:Name="HostLabel" Content="Workstation ID:"/>
        <TextBox x:Name="WorkstationTextBox" Height="20"/>
        <Label x:Name="PrintServerLabel" Content="Print Server:"/>
        <TextBox x:Name="PrintServerTextBox" Height="20"/>
        <Label x:Name="PrinterNameLabel" Content="Printer Name:"/>
        <TextBox x:Name="PrinterNameTextBox" Height="20"/>
        <Button x:Name="MapButton" Content="Map this Printer!" Margin="10,10,10,0" VerticalAlignment="Top" Height="25"/> 
        <Button x:Name="ClearFormButton" Content="Clear Form" Margin="10,10,10,0" VerticalAlignment="Top" Height="25"/>
    </StackPanel>
</Window>
'@
 
 # <ListBox x:Name="ResultsSelect" Height = "300" SelectionMode="Extended" ScrollViewer.VerticalScrollBarVisibility="Visible"/>
$global:Form = ""
# XAML Launcher
$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
try{$global:Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Some possible causes for this problem include: .NET Framework is missing PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."; break}
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name ($_.Name) -Value $global:Form.FindName($_.Name)}

# Set up controls
$WorkstationTextBox = $global:Form.FindName('WorkstationTextBox')
$PrintServerTextBox = $global:Form.FindName('PrintServerTextBox')
$PrinterNameTextBox = $global:Form.FindName('PrinterNameTextBox')
$MapButton = $global:Form.FindName('MapButton')
$ClearFormButton = $global:Form.FindName('ClearFormButton')

$MapButton.Add_Click({
    $workstation = $WorkstationTextBox.Text
    $print_server = $PrintServerTextBox.Text
    $printer_name = $PrinterNameTextBox.Text

    if ([string]::IsNullOrEmpty($workstation)) {
        [System.Windows.MessageBox]::Show("No Workstation ID set!")
        return
    }
    if ([string]::IsNullOrEmpty($print_server)) {
        [System.Windows.MessageBox]::Show("No print server set!")
        return
    }    
    if ([string]::IsNullOrEmpty($printer_name)) {
        [System.Windows.MessageBox]::Show("No printer set!")
        return
    }
    

    & \\networkshare\Powershell\PsExec.exe -s -accepteula -accepteula \\$workstation C:\Windows\System32\rundll32.exe PRINTUI.DLL,PrintUIEntry /ga /c\\$workstation /n\\$print_server\$printer_name /q 2>$null

       
    [System.Windows.MessageBox]::Show("Done!`nPlease have the user wait 30 seconds before checking their printer list.")

})
$ClearFormButton.Add_Click({
    $PrintServerTextBox.Text = ""
    $WorkstationTextBox.Text = ""
    $PrinterNameTextBox.Text = ""
    $global:Form.invalidateVisual()
})

# Show GUI
$global:Form.ShowDialog() | out-null