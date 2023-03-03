Add-Type -Assembly System.Drawing

# Set the search parameters
if ($args[0]) {
  $SearchFor = $args[0]
}
else { 
  $SearchFor = "turtles" 
}

if ($args[1] -eq "portrait") {
  $Orientation = $args[1]
}
else { 
  $Orientation = "landscape" 
}

# Hide WebRequest progress output
$ProgressPreference = 'SilentlyContinue'

# Stop going if we run into an error
# $ErrorActionPreference = "Stop"

# Set base URI to Unsplash search API
$BaseUri = "https://unsplash.com/napi/search/photos?query=$SearchFor&page={0}&per_page={1}&orientation=$Orientation"

# Check how many images there are to choose from
$MetaQuery = [uri]::EscapeUriString(($BaseUri -f 1, 1))
$TotalImages = ((Invoke-WebRequest -UseBasicParsing -URI $MetaQuery).Content | ConvertFrom-Json).total
$ImagesPerPage = [Math]::Min($TotalImages, 20)
$AvailablePages = [Math]::Floor($TotalImages / $ImagesPerPage)

# Select a random page between 1 and the lesser of 10 or the number of pages available
$MaxPage = [Math]::Min($AvailablePages, 10)
$Page = [Math]::Floor((Get-Random -Minimum 1 -Maximum $MaxPage))

# Fetch images
$ImageQuery = [uri]::EscapeUriString(($BaseUri -f $Page, $ImagesPerPage))
$Images = ((Invoke-WebRequest -UseBasicParsing -URI $ImageQuery).Content | ConvertFrom-Json).results | Where-Object { $_.premium -eq $false }

# Select a random images from the results
$ImageIndex = [Math]::Floor((Get-Random -Minimum 0 -Maximum ($Images.count - 1)))
$ImageUri = $Images[$ImageIndex].urls.raw

# Write the image to a temp file
$TempPath = [System.IO.Path]::GetTempFileName()
Invoke-WebRequest -UseBasicParsing $ImageUri -OutFile $TempPath

# Save the downloaded file as JPEG
$image = [System.Drawing.Image]::FromFile($TempPath);
$ImagePath = [IO.Path]::ChangeExtension($TempPath, '.jpg');
$image.Save($ImagePath, [System.Drawing.Imaging.ImageFormat]::Jpeg);
$image.Dispose();

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
