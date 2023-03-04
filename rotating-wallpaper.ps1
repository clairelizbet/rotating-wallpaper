Add-Type -Assembly System.Drawing
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

  if ($args[2] -eq "portrait") {
    $Orientation = $args[2]
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
  $ImageResponse = (Invoke-WebRequest -UseBasicParsing -URI $ImageQuery).Content | ConvertFrom-Json
  $ImageUri = $ImageResponse.links.download

  # Write the image to a temp file
  $TempPath = [System.IO.Path]::GetTempFileName()
  Invoke-WebRequest -UseBasicParsing $ImageUri -OutFile $TempPath

  # Save the downloaded file as a JPEG image
  $WallpaperDirCreateRes = New-Item -ItemType Directory $WallpaperFolder -ea 0

  $Image = [System.Drawing.Image]::FromFile($TempPath);
  $TempPathJPG = [IO.Path]::ChangeExtension($TempPath, '.jpg');
  $ImageName = Split-Path -Leaf -Path $TempPathJPG
  $ImagePath = "$WallpaperFolder\$ImageName"
  $Image.Save($ImagePath, [System.Drawing.Imaging.ImageFormat]::Jpeg);
  $Image.Dispose();

  # Write to the image log
  $LogFile = "$WallpaperFolder\rotation-history.txt"
  $Date = Get-Date
  $ImageDescription = $ImageResponse.alt_description
  $ImageLink = $ImageResponse.links.html
  $ImageAuthor = $ImageResponse.user.name

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

      $TaskArgs = "-ExecutionPolicy Bypass -WindowStyle Hidden -NoProfile -NonInteractive -File $InstanceFileName ""$SearchFor"" ""$WallpaperFolder"""
      $Action = New-ScheduledTaskAction -Execute "powershell.exe" `
        -Argument $TaskArgs `
        -WorkingDirectory $AppDir
    } else {
      $InstancePath = [Environment]::GetCommandLineArgs()[0]
      $InstanceFileName = Split-Path -Leaf -Path $InstancePath

      $Action = New-ScheduledTaskAction -Execute $InstanceFileName `
        -Argument """$SearchFor"" ""$WallpaperFolder""" `
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
