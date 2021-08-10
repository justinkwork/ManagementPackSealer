[System.Reflection.Assembly]::LoadWithPartialName("System.collections.generic") | Out-Null
##Form definition
$form = @"
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        Title="Seal Management Pack" MinHeight="480" MaxHeight="480" Width="416" MinWidth="416" MaxWidth="560">
    <Grid Background="#EEEEEE">
        <Label Content="MP Path:" Margin="56,34,0,0" VerticalAlignment="Top"/>
        <TextBox Name="txtFolder"  Margin="60,60,60,360" Text="" IsEnabled="False" TextWrapping="Wrap" Height="24" />
        <Button Name="btnFolderBrowse" Content="Browse" Margin="226,88,0,332" Height="19" Width="56"/>
        <Label Content="MP File:" Margin="56,120,0,0" VerticalAlignment="Top"/>
        <ComboBox Name="cmbFile" Margin="60,146,60,270"  Height="24" />
        <Label Content="KeyFile:" Margin="56,206,0,0" VerticalAlignment="Top"/>
        <TextBox Name="txtKeyFile" Margin="60,230,60,190" Text="" IsEnabled="False" TextWrapping="Wrap"  Height="24"/>
        <Button Name="btnKeyBrowse" Content="Browse" Margin="226,258,0,162" Height="19" Width="56"/>
        <Label Content="Company:" Margin="56,290,0,0" VerticalAlignment="Top"/>
        <TextBox Name="txtCompany" Margin="60,314,60,100" Text="" TextWrapping="Wrap"  Height="24"/>
        <Button Name="btnSeal" Content="Seal" HorizontalAlignment="Center" IsEnabled="False" Margin="0,364,0,0" VerticalAlignment="Top" Height="44" Width="92"/>
    </Grid>
</Window>
"@

##Function Definitions
Function Get-Dialog {
    Param(
        [Parameter(Mandatory = $True, Position = 1)]
        [string]$XamlPath
    )
    [xml]$xmlWPF = $XamlPath
    try {
        Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase, system.windows.forms
    } 
    catch {
        Throw "Failed to load Windows Presentation Framework assemblies."
    }
    $xamGUI = [Windows.Markup.XamlReader]::Load((new-object System.Xml.XmlNodeReader $xmlWPF))
    $xmlWPF.SelectNodes("//*[@Name]") | ForEach-Object {
        Set-Variable -Name ($_.Name) -Value $xamGUI.FindName($_.Name) -Scope Global
    }
    return $xamGUI
}

Function Remove-Dialog {
    Param(
        [Parameter(Mandatory = $True, Position = 1)]
        [string]$XamlPath
    )
    [xml]$xmlWPF = $XamlPath
    try {
        Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase, system.windows.forms
    } 
    catch {
        Throw "Failed to load Windows Presentation Framework assemblies."
    }
    
    $xmlWPF.SelectNodes("//*[@Name]") | ForEach-Object {
        Remove-Variable -Name ($_.Name) -Scope Global
    }
}

Function Get-Folder {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select a folder"
    $foldername.rootfolder = "MyComputer"

    if ($foldername.ShowDialog() -eq "OK") {
        $folder += $foldername.SelectedPath
    }
    return $folder
}

Function Get-FileName {
    param(
        [string]$initialDirectory
    )

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowDialog() | Out-Null
    return $OpenFileDialog.filename
} 


Function Get-MessageBox {
    param([string]$message)
    [System.Windows.Forms.MessageBox]::Show($message)
}

Function Invoke-SealMP {
    param(
        [string]$mpDir,
        [string]$fileName,
        [string]$keyFile,
        [string]$company
    )

    Protect-SCManagementPack -ManagementPackFile "$mpdir\$fileName" -OutputDirectory $mpdir -KeyFilePath $keyFile -CompanyName $company 
}

Function Get-Validation {
    $eap = $ErrorActionPreference
    $ErrorActionPreference = "Ignore"
    $folder = Test-Path $txtFolder.Text 
    $mpFile = (Test-Path "$($txtFolder.Text)\$($cmbFile.SelectedItem)") -and ($cmbFile.SelectedItem.ToLower().IndexOf(".xml") -gt -1)
    $keyFile = (Test-Path $txtKeyFile.Text) -and ($txtKeyFile.Text.ToLower().IndexOf(".snk") -gt -1)
    $company = !([string]::IsNullOrEmpty($txtCompany.Text))
    $btnSeal.IsEnabled = $folder -and $mpFile -and $keyFile -and $company
    $ErrorActionPreference = $eap
}



#end function definitions

## Load form definition and create variables
$win = Get-Dialog $Form

##event handlers
$btnFolderBrowse.add_click( {
        $folder = Get-Folder
        $files = Get-ChildItem $folder | Where-Object { $_.Extension -eq ".xml" } | Select-Object -Expand Name
        $fileList = New-Object System.Collections.ArrayList
        switch ($files.count) {
            0 { Get-MessageBox -message "No XML Files Found!" }
            1 { $fileList.Add($files) }
            default { $fileList.AddRange($files) }
        }
        $txtFolder.Text = $folder
        $cmbFile.ItemsSource = $fileList
        Get-Validation
    })

$cmbFile.add_selectionChanged( {
        Get-Validation
    })

$btnKeyBrowse.add_click( {
        $keyFile = Get-FileName -initialDirectory $env:USERPROFILE
        $txtKeyFile.Text = $keyFile
        Get-Validation
    })

$txtCompany.add_keyup( {
        Get-Validation
    })

$btnSeal.add_click( {
        try {
            Invoke-SealMP -mpDir $txtFolder.Text -fileName $cmbFile.SelectedItem -keyFile $txtKeyFile.Text -company $txtCompany.Text
            Get-MessageBox -message "MP Sealed!"
        }
        catch {
            Get-MessageBox -message $_.Exception.Message
        }

    })

$win.add_closing( {
        #remove all the variables
        Remove-Dialog -XamlPath $form
    })


#launch window 
$win.ShowDialog() | Out-Null