﻿Add-Type -Assembly System.Drawing
Add-Type -AssemblyName PresentationFramework

$PhotosFolder = [Environment]::GetFolderPath("MyPictures")
$DefaultWallpaperFolder = "$PhotosFolder\Wallpaper"

$DefaultSearch = "turtles"

if ($args[0]) {
  # Invoked with arguments => Change wallpaper

  $SearchFor = $args[0]

  if ($args[1]) {
    $WallpaperFolder = $args[1]
  }
  else {
    $WallpaperFolder = $DefaultWallpaperFolder
  }

  try {
    [int]$IntervalHours = [Math]::clamp([convert]::ToInt32($args[2], 10), 1, 24)
  } catch {
    $IntervalHours = 1
  }

  if ($args[3] -eq "portrait") {
    $Orientation = $args[3]
  }
  else {
    $Orientation = "landscape" 
  }

  # Hide WebRequest progress output
  $ProgressPreference = 'SilentlyContinue'

  # Stop going if we run into an error
  $ErrorActionPreference = "Stop"

  # Get a random image from Unsplash
  $BaseUri = "https://unsplash.com/napi"
  $ImageQuery = [uri]::EscapeUriString("$BaseUri/photos/random?query=$SearchFor&orientation=$Orientation")

  $MaxRetries = 5
  # Retry n times, spread out over half the interval duration
  $RetryBase = [Math]::Pow($IntervalHours * 60 * .5, (1 / $MaxRetries))

  for ($retry = 1; $retry -le $MaxRetries; $retry++) {
    try {
      $ImageResponse = (Invoke-WebRequest -UseBasicParsing -URI $ImageQuery).Content | ConvertFrom-Json
      break
    } catch {
      [float]$BackoffFudge = (Get-Random -Max 1000) / 1000
      [float]$BackoffMinutes = [Math]::Pow($RetryBase, $retry) + $BackoffFudge
      $BackoffDuration = $BackoffMinutes * 60

      if ($retry -ge $MaxRetries) {
        exit
      } else {
        Start-Sleep -Seconds $BackoffDuration
      }
    }
  }

  $ImageUri = $ImageResponse.links.download
  $ImageDescription = $ImageResponse.alt_description
  $ImageLink = $ImageResponse.links.html
  $ImageAuthor = $ImageResponse.user.name

  $ImageUri = if ($ImageUri) { $ImageUri } else { "" }
  $ImageDescription = if ($ImageDescription) { $ImageDescription } else { "" }
  $ImageLink = if ($ImageLink) { $ImageLink } else { "" }
  $ImageAuthor = if ($ImageAuthor) { $ImageAuthor } else { "" }

  # Write the image to a temp file
  $TempPath = [System.IO.Path]::GetTempFileName()
  Invoke-WebRequest -UseBasicParsing $ImageUri -OutFile $TempPath

  # Save the downloaded file as a JPEG image
  $WallpaperDirCreateRes = New-Item -ItemType Directory $WallpaperFolder -ea 0

  $Image = [System.Drawing.Image]::FromFile($TempPath);
  $TempPathJPG = [IO.Path]::ChangeExtension($TempPath, '.jpg');
  $ImageName = Split-Path -Leaf -Path $TempPathJPG
  $ImagePath = "$WallpaperFolder\$ImageName"

  # This approach is super weird but MS have disabled the constructor for PropertyItem
  # and their official documentation says to do it this way, so whatever I guess...
  $ReusablePropertyItem = $Image.PropertyItems | Select-Object -First 1

  if ($ReusablePropertyItem) {
    # Set meta info on image

    $ASCII = [System.Text.ASCIIEncoding]::new()
    $UTF16 = [System.Text.UnicodeEncoding]::new()

    $ReusablePropertyItem.id = 270 # Description/Title
    $ReusablePropertyItem.Type = 2 # String
    $ReusablePropertyItem.Value = $ASCII.GetBytes("$ImageDescription`0")
    $ReusablePropertyItem.len = $ReusablePropertyItem.Value.Length
    $Image.SetPropertyItem($ReusablePropertyItem)

    $ReusablePropertyItem.id = 315 # Author name
    $ReusablePropertyItem.Type = 2 # String
    $ReusablePropertyItem.Value = $ASCII.GetBytes("$ImageAuthor`0")
    $ReusablePropertyItem.len = $ReusablePropertyItem.Value.Length
    $Image.SetPropertyItem($ReusablePropertyItem)

    $ReusablePropertyItem.id = 40092 # Comment
    $ReusablePropertyItem.Type = 1 # Byte array
    $ReusablePropertyItem.Value = $UTF16.GetBytes("$ImageLink`0")
    $ReusablePropertyItem.len = $ReusablePropertyItem.Value.Length
    $Image.SetPropertyItem($ReusablePropertyItem)
  }

  $Image.Save($ImagePath, [System.Drawing.Imaging.ImageFormat]::Jpeg);
  $Image.Dispose();

  # Write to the image log
  $LogFile = "$WallpaperFolder\rotation-history.txt"
  $Date = Get-Date

  $(
    "$Date :: $ImageName :: $ImageDescription :: $ImageAuthor :: $ImageLink"
    Get-Content -Path $LogFile -Tail 168 -ea 0
  ) | Set-Content -Path $LogFile

  # Set wallpaper using Win32
  $SetWallpaperWin32 = @'
  using System.Runtime.InteropServices; 
  namespace Win32 {
    public class WinUser {
      [DllImport("user32.dll", CharSet=CharSet.Auto)]

      // https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-systemparametersinfoa
      static extern int SystemParametersInfo (int uiAction, int uiParam, string pvParam, int fWinIni); 

      public static void SetWallpaper (string imagePath) { 
        int SPI_SETDESKWALLPAPER = 0x0014;
        int SPIF_SENDCHANGE = 0x02;

        SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, imagePath, SPIF_SENDCHANGE); 
      }
    }
  }
'@
  add-type $SetWallpaperWin32

  [Win32.WinUser]::SetWallpaper($ImagePath)
}
else {
  # Invoked without arguments => Setup

  [XML]$XAML = @'
  <Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"     
    Title="Rotating Wallpaper" Height="230" Width="400" ResizeMode="CanMinimize">
    <Grid>
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="50*"/>
        <ColumnDefinition Width="120*"/>
        <ColumnDefinition Width="230*"/>
      </Grid.ColumnDefinitions>

      <Rectangle HorizontalAlignment="Center" Height="194" VerticalAlignment="Center" Width="50" Fill="#FF9DD0FF"/>

      <Label Grid.Column="1" Content="Search Photos:" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top"/>
      <TextBox x:Name="SearchInput" Grid.Column="2" HorizontalAlignment="Left" Margin="0,14,0,0" Text="" VerticalAlignment="Top" Width="188"/>

      <Label Grid.Column="1" Content="Rotate every:" HorizontalAlignment="Left" Margin="10,94,0,0" VerticalAlignment="Top"/>
      <Slider x:Name="RotationInput" Grid.Column="2" HorizontalAlignment="Left" Margin="0,92,0,0" VerticalAlignment="Top" Width="120" SmallChange="1" Maximum="12" Minimum="1" Value="1" Cursor="Hand" TickPlacement="Both"/>
      <Label x:Name="RotationLabel" Grid.Column="2" HorizontalAlignment="Left" Margin="125,96,0,0" VerticalAlignment="Top" Content="Hour"/>

      <Label Grid.Column="1" Content="Save to:" HorizontalAlignment="Left" Margin="11,41,0,0" VerticalAlignment="Top"/>
      <TextBox x:Name="WallpaperFolderInput" Grid.Column="2" HorizontalAlignment="Left" Margin="0,45,0,0" Text="" VerticalAlignment="Top" Width="188" Height="41" TextWrapping="Wrap"/>

      <Button x:Name="SaveButton" Grid.Column="2" Content="Save" HorizontalAlignment="Left" Margin="0,134,0,0" VerticalAlignment="Top" Width="120" IsDefault="True" Cursor="Hand"/>

      <Label Grid.Column="2" Content="Photos by Unsplash" HorizontalAlignment="Left" Margin="0,159,0,0" VerticalAlignment="Top" Width="120"/>
    </Grid>
  </Window>
'@

  $reader = (New-Object System.Xml.XmlNodeReader $XAML)
  $window = [Windows.Markup.XamlReader]::Load( $reader )

  Set-Variable -Name SearchInput -Value $window.FindName("SearchInput")
  Set-Variable -Name RotationInput -Value $window.FindName("RotationInput")
  Set-Variable -Name RotationLabel -Value $window.FindName("RotationLabel")
  Set-Variable -Name WallpaperFolderInput -Value $window.FindName("WallpaperFolderInput")
  Set-Variable -Name SaveButton -Value $window.FindName("SaveButton")

  $SearchInput.Text = $DefaultSearch
  $WallpaperFolderInput.Text = $DefaultWallpaperFolder

  $SaveButton.Add_Click({
    if ($SearchInput.Text -eq "") {
      $SearchFor = $DefaultSearch
    } else {
      $SearchFor = $SearchInput.Text
    }

    $IntervalHours = [Math]::floor($RotationInput.Value)

    $AppDir = "$env:APPDATA\rotating-wallpaper"
    $WallpaperFolder = $WallpaperFolderInput.Text

    if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript") {
      $InstancePath = $MyInvocation.MyCommand.Definition
      $InstanceFileName = Split-Path -Leaf -Path $InstancePath

      $TaskArgs = "-ExecutionPolicy Bypass -WindowStyle Hidden -NoProfile -NonInteractive -File $InstanceFileName ""$SearchFor"" ""$WallpaperFolder"" $IntervalHours"
      $Action = New-ScheduledTaskAction -Execute "powershell.exe" `
        -Argument $TaskArgs `
        -WorkingDirectory $AppDir
    } else {
      $InstancePath = [Environment]::GetCommandLineArgs()[0]
      $InstanceFileName = Split-Path -Leaf -Path $InstancePath

      $Action = New-ScheduledTaskAction -Execute $InstanceFileName `
        -Argument """$SearchFor"" ""$WallpaperFolder"" $IntervalHours" `
        -WorkingDirectory $AppDir
    }

    $AppDirCreateRes = New-Item -ItemType Directory $AppDir -ea 0
    Copy-Item $InstancePath -Destination $AppDir

    $TaskName = "Rotating Wallpaper"
    $Trigger = (New-ScheduledTaskTrigger -Daily -At 00:00)
    $Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries

    $Interval = New-TimeSpan -Hours $IntervalHours
    $Duration = New-TimeSpan -Days 1
    $RepeatTrigger = New-ScheduledTaskTrigger -Once -At 00:00 -RepetitionDuration $Duration -RepetitionInterval $Interval
    $Trigger.Repetition = $RepeatTrigger.Repetition

    if (Get-ScheduledTask | Where-Object {$_.TaskName -like $TaskName }) {
      $Task = Set-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Settings
    } else {
      $Task = New-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings
      $TaskRegistrationResult = Register-ScheduledTask $TaskName -InputObject $Task
    }

    $StartRes = Start-ScheduledTask -TaskName $TaskName
    exit
  })

  $RotationInput.Add_ValueChanged({
    $RotationValue = [Math]::floor($RotationInput.Value)
    
    if ($RotationValue -eq 1) {
      $RotationLabel.Content = "Hour"
    } else {
      $RotationLabel.Content = "$RotationValue Hours"
    }
  })

  $Dialog = $window.ShowDialog()
}
