<# TODO
    Add notification icon - Done
    Add menu items for the image manipulator
    Add keyboard listener
    Add GUI for assigning the shortcut for the keyboard listener
    Implement Imgur API - Uploading the image is done. Should look into more features
    Add web proxy support
    fix annoying red border around captured image
    Style the context menu
    Add multi-monitor support
#>
# http://stackoverflow.com/questions/1838163/click-and-drag-selection-box-in-wpf

#$ErrorActionPreference = 'Stop'
#Build the GUI

[xml]$Capture_XAML = Get-Content (Join-Path $PSScriptRoot \Xaml\Capture.xaml)

function Upload-ImgurImage {
    param (
        [string]$Image      
    )

    $BaseURL = 'https://api.imgur.com/3/upload'
    $ClientID = '6bdcf8472408440'
    $Headers = @{Authorization = "Client-ID $ClientID"}

    $ImageBase64 = [System.Convert]::ToBase64String((Get-Content $Image -Encoding Byte))

    $Result = Invoke-WebRequest -Uri $BaseURL -Method POST -Body $ImageBase64 -Headers $Headers
    $Json = ConvertFrom-Json $Result
    Write-Output $Json.data.link
 
}

function New-Capture {
    
}


function Show-BalloonTip {
    $NotifyIcon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
    $NotifyIcon.BalloonTipText = 'Captured!'
    $NotifyIcon.BalloonTipTitle = 'Captured!'
    $NotifyIcon.ShowBalloonTip(30000)
}

# Load requires assemblies
[reflection.assembly]::LoadWithPartialName('System.Drawing') | Out-Null
[reflection.assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null
# Enable Windows themes
[System.Windows.Forms.Application]::EnableVisualStyles()

$Reader = New-Object System.Xml.XmlNodeReader $Capture_XAML
$Window = [Windows.Markup.XamlReader]::Load($Reader)
$Canvas = $Window.FindName('Canvas')
$script:Rectangle = $Window.FindName('Rectangle')
$Visible = [System.Windows.Visibility]::Visible
$Hidden = [System.Windows.Visibility]::Hidden
$Collapsed = [System.Windows.Visibility]::Collapsed

# Create tray icon

$global:NotifyIcon = New-Object System.Windows.Forms.NotifyIcon
$NotifyIcon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Join-Path $PSScriptRoot '\Assets\screenshot.ico'))
$NotifyIcon.Visible = $true

# Create context menu for tray icon
$NotifyIconContextMenu = New-Object System.Windows.Forms.ContextMenu

# Create context menu items
$NotifyIconContextMenu_MenuItem_Exit = New-Object System.Windows.Forms.MenuItem
$NotifyIconContextMenu_MenuItem_Exit.Index = 0
$NotifyIconContextMenu_MenuItem_Exit.Text = 'Exit'

$NotifyIconContextMenu_MenuItem_Options = New-Object System.Windows.Forms.MenuItem
$NotifyIconContextMenu_MenuItem_Options.Index = 1
$NotifyIconContextMenu_MenuItem_Options.Text = 'Options'

# Add context menu to tray icon
$NotifyIconContextMenu.MenuItems.Add($NotifyIconContextMenu_MenuItem_Options)
$NotifyIconContextMenu.MenuItems.Add($NotifyIconContextMenu_MenuItem_Exit)
$NotifyIcon.ContextMenu = $NotifyIconContextMenu

$Window.Add_MouseDown({
    $script:MouseDown = $true
    $Canvas.CaptureMouse()
    $script:MouseDownPos = [System.Windows.Forms.Cursor]::Position
    [System.Windows.Controls.Canvas]::SetLeft($Rectangle, $MouseDownPos.X)
    [System.Windows.Controls.Canvas]::SetTop($Rectangle, $MouseDownPos.Y)
    $Rectangle.Width = 0
    $Rectangle.Height = 0
    $Rectangle.Visibility = $Visible
})

$Window.Add_MouseUp({
    if ($Rectangle.Width -lt 1 -or $Rectangle.Height -lt 1) {
        $MouseDown = $false
        $Canvas.ReleaseMouseCapture()
        $Rectangle.Visibility = $Collapsed
        return
    }

    $MouseDown = $false
    $Canvas.ReleaseMouseCapture()
    $Rectangle.Visibility = $Collapsed
    $script:MouseUpPos = [System.Windows.Forms.Cursor]::Position

    <#
    Write-Host "`$MouseUpPos.X =" $MouseUpPos.X
    Write-Host "`$MouseDownPos.X =" $MouseDownPos.X
    Write-Host "`$MouseUpPos.Y =" $MouseUpPos.Y
    Write-Host "`$MouseDownPos.Y =" $MouseDownPos.Y
    Write-Host "`$Rectangle.Height =" $Rectangle.Height
    Write-Host "`$Rectangle.Width =" $Rectangle.Width
    Write-Host "`$Bounds.Size =" $Bounds.Size
    Write-Host "=============================================="
    #>
    $Bounds = New-Object Drawing.Rectangle
    $Bounds.Width = $Rectangle.Width
    $Bounds.Height = $Rectangle.Height

    $Screenshot = New-Object Drawing.Bitmap $Bounds.Width, $Bounds.Height
    $Graphics = [Drawing.Graphics]::FromImage($Screenshot)

    if ($MouseUpPos.X -gt $MouseDownPos.X) {
        $Graphics.CopyFromScreen($MouseDownPos, [System.Drawing.Point]::Empty, $Bounds.Size)
    } else {
        $Graphics.CopyFromScreen($MouseUpPos, [System.Drawing.Point]::Empty, $Bounds.Size)
    }
    
    $ImagePath = $(Join-Path -Path $env:TEMP -ChildPath $([guid]::NewGuid())) + '.bmp'
    #$Screenshot.Save($ImagePath)
    $Screenshot.Save('C:\temp\image.bmp')
    $Graphics.Dispose()
    $Screenshot.Dispose()
    $Rectangle.Visibility = $Collapsed
    #$ImgurURL = Upload-ImgurImage $ImagePath
    #start $ImgurURL
})

$Window.Add_MouseMove({
    if ($MouseDown) {
        #Write-Host $([System.Windows.Forms.Cursor]::Position)
        [System.Drawing.Point]$MousePos = [System.Windows.Forms.Cursor]::Position
        if ($MouseDownPos.X -lt $MousePos.X) {
            [System.Windows.Controls.Canvas]::SetLeft($Rectangle, $MouseDownPos.X)
            $Rectangle.Width = $MousePos.X - $MouseDownPos.X
        } else {
            [System.Windows.Controls.Canvas]::SetLeft($Rectangle, $MousePos.X)
            $Rectangle.Width = $MouseDownPos.X - $MousePos.X
        }

        if ($MouseDownPos.Y -lt $MousePos.Y) {
            [System.Windows.Controls.Canvas]::SetTop($Rectangle, $MouseDownPos.Y)
            $Rectangle.Height = $MousePos.Y - $MouseDownPos.Y           
        } else {
            [System.Windows.Controls.Canvas]::SetTop($Rectangle, $MousePos.Y)
            $Rectangle.Height = $MouseDownPos.Y - $MousePos.Y
        }
    }
})

$Window.Add_MouseRightButtonDown({
    $Window.Close()
})

$Window.Add_Closing({
    $NotifyIcon.Dispose()
})

$Window.Add_Loaded({
    #$Window.Width = $VirtualScreenWidth
    #$Window.Height = $VirtualScreenHeight
    $Canvas.Width = $Window.Width
    $Canvas.Height = $Window.Height
    $MouseDown = $false
    $NotifyIcon.Visible = $true
    #$Window.IsEnabled = $false


})

$NotifyIconContextMenu_MenuItem_Exit.Add_Click({
    $Window.Close()
})

$NotifyIconContextMenu_MenuItem_Options.Add_Click({
    $Window.showdialog()
})

try {
    #$Window.ShowDialog() | Out-Null
}
catch {
    $Error[0].Exception.Message
}
finally {
    $NotifyIconContextMenu.Dispose()
    $NotifyIconContextMenu_MenuItem_Exit.Dispose()
    $NotifyIcon.Dispose()
}
