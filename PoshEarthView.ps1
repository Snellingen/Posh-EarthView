Add-Type @"
using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;
namespace Wallpaper
{
   public enum Style : int
   {
       Tile, Center, Stretch, NoChange
   }
   public class Setter {
      public const int SetDesktopWallpaper = 20;
      public const int UpdateIniFile = 0x01;
      public const int SendWinIniChange = 0x02;
      [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
      private static extern int SystemParametersInfo (int uAction, int uParam, string lpvParam, int fuWinIni);
      public static void SetWallpaper ( string path, Wallpaper.Style style ) {
         SystemParametersInfo( SetDesktopWallpaper, 0, path, UpdateIniFile | SendWinIniChange );
         RegistryKey key = Registry.CurrentUser.OpenSubKey("Control Panel\\Desktop", true);
         switch( style )
         {
            case Style.Stretch :
               key.SetValue(@"WallpaperStyle", "2") ; 
               key.SetValue(@"TileWallpaper", "0") ;
               break;
            case Style.Center :
               key.SetValue(@"WallpaperStyle", "1") ; 
               key.SetValue(@"TileWallpaper", "0") ; 
               break;
            case Style.Tile :
               key.SetValue(@"WallpaperStyle", "1") ; 
               key.SetValue(@"TileWallpaper", "1") ;
               break;
            case Style.NoChange :
               break;
         }
         key.Close();
      }
   }
}
"@

$base_url = "http://www.gstatic.com/prettyearth/assets/full/{0}.jpg"
$threshold = 50
$image_out_path = "$Env:USERPROFILE\Pictures\wallpaper\"
if ((Test-Path $image_out_path) -eq $false)
{
    New-Item -ItemType Directory $image_out_path
}

$old_images = Get-ChildItem $image_out_path | Sort CreationTime
if ($old_images.count -gt 5)
{
    $old_images | Select-Object -First ($old_images.count - 5) | Remove-Item
}

for ($i = 0; $i -lt $threshold; $i++)
{
    $random_number = Get-Random -Minimum 1000 -Maximum 7023
    $image = $image_out_path + "EarthView{0}.jpg" -f $random_number
    $url = $base_url -f $random_number
    try
    {
        $result = Invoke-WebRequest $url -OutFile $image
        [Wallpaper.Setter]::SetWallpaper( $image, 3 )
        break;
    }
    catch
    {
        continue;
    }    
}

