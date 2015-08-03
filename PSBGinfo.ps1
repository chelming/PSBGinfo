# Script to set the wallpaper on a machine similar to BGInfo but without VBScripts
# Written by Chris Helming, Oscar Santiano, and Chase Caynoski

###################################################
##         TEXT TO DISPLAY ON THE DESKTOP        ##
###################################################
$line1     = "Computer name: $env:COMPUTERNAME"
$line2     = "To report a problem"
$line3     = "visit labsupport.rit.edu"   
$line4     = " "
$typeFace  = "Gill Sans MT Condensed"
$bgColor   = (255, 70, 70, 70) #aRGB
$txtColor  = (255, 255, 255, 255) #aRGB
$txtShadow = (200, 0, 0, 0) #aRGB

###################################################
## Filepath for logo for bottom right of screen  ##
##            (can use transparancy)             ##
###################################################
$logoPath = "\\fileshare\login.png"


# Pull wallpaper from a folder for a specific lab
# If it's not a normal formatted lab, fall back to
# a default folder on the fileshare
if($env:COMPUTERNAME -imatch "^[a-z]{3}-[0-9]{4}-.+"){
    $labname = $env:COMPUTERNAME.substring(0,8)
} else { $labname = "default" } 

#######################################################
## Sources and Destinations for theme and wallpaper  ##
#######################################################
$sourcewallpaperpath = "\\fileshare\$labname\wallpaper\"
$sourcetheme         = "\\fileshare\ITS.theme"

$wallpaperpath = "C:\Windows\Web\Wallpaper\ITS"
$themepath     = "C:\Windows\Resources\Themes"


Function AddTextToImage {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$false)][String] $logoPath,
        [Parameter(Mandatory=$false)][String] $sourcePath,
        [Parameter(Mandatory=$true)][String]  $destPath,
        [Parameter(Mandatory=$true)][String]  $line1,
        [Parameter(Mandatory=$true)][String]  $line2,
        [Parameter(Mandatory=$true)][String]  $line3,
        [Parameter(Mandatory=$true)][String]  $line4,
        [Parameter(Mandatory=$true)][String]  $typeFace,
        [Parameter(Mandatory=$true)][Array]  $bgColor,
        [Parameter(Mandatory=$true)][Array]  $txtColor,
        [Parameter(Mandatory=$true)][Array]  $txtShadow
    )
 
    Write-Verbose "Load System.Drawing"
    [Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null


    Write-Verbose "Set Screen Resolution"
    if($win8 -or $win10) {
        $screen = (wmic path win32_videocontroller get VideoModeDescription)
        $screenWidth = $screen[2].Split("x").trim()[0]
        $screenHeight = $screen[2].Split("x").trim()[1]
    }
    else {
        $screen = gwmi win32_desktopmonitor
        try{
            $bound = $screen.GetUpperBound(0)
            $i = 0
            while(($i -le $bound) -and ($screenWidth -isnot [int32])) {
                $screenWidth  = $screen.screenWidth[$i]
                $screenHeight = $screen.screenHeight[$i]
                $i++
            }
        }
        catch {
            $screenWidth  = $screen.screenWidth
            $screenHeight = $screen.screenHeight
        }
    }

    # Move the text in 20 pixels from the right side of the screen
    $textPos    = $screenwidth - 20
    $shadowPos = $textPos + 2
    
    Write-Verbose "Create a bitmap as $destPath"
    $bmpFile = new-object System.Drawing.Bitmap([int]($screenWidth)),([int]($screenHeight))
 
    Write-Verbose "Intialize Graphics"
    $Image = [System.Drawing.Graphics]::FromImage($bmpFile)
    $Image.clear([System.Drawing.Color]::FromArgb($bgcolor[0], $bgcolor[1], $bgcolor[2], $bgcolor[3]))
    $Image.SmoothingMode = "AntiAlias"

    Write-Verbose "Set text alignment"
    $stringFormat = New-Object System.Drawing.StringFormat
    $stringFormat.Alignment = [System.Drawing.StringAlignment]::Far

    if($sourcePath){
        Write-Verbose "Get the image from $sourcePath"
        $srcImg = [System.Drawing.Image]::FromFile($sourcePath)
        $Rectangle = New-Object Drawing.Rectangle ($screenWidth/2 - $srcImg.width/2), ($screenHeight/2 - $srcImg.height/2), $srcImg.Width, $srcImg.Height
        $Image.DrawImage($srcImg, $Rectangle, 0, 0, $srcImg.Width, $srcImg.Height, ([Drawing.GraphicsUnit]::Pixel)) 
        $srcImg.Dispose()
    }
    $srcLogo = [System.Drawing.Image]::FromFile($logoPath)
    $Rectangle = New-Object Drawing.Rectangle ($screenWidth - $srcLogo.width - 20), ($screenHeight - $srcLogo.height - 50), $srcLogo.Width, $srcLogo.Height
    $Image.DrawImage($srcLogo, $Rectangle, 0, 0, $srcLogo.Width, $srcLogo.Height, ([Drawing.GraphicsUnit]::Pixel))    
    
    ##########################
    ## DROP SHADOW FOR TEXT ##
    ##########################
    Write-Verbose "Draw line1 shadow: $line1"
    $Font = new-object System.Drawing.Font($typeFace, 35)
    $Brush = New-Object Drawing.SolidBrush ([System.Drawing.Color]::FromArgb($txtShadow[0], $txtShadow[1], $txtShadow[2], $txtShadow[3]))
	$Image.DrawString($line1, $Font, $Brush, $shadowPos, 12, $stringFormat)
	
    Write-Verbose "Draw line2 shadow: $line2"
    $Font = New-object System.Drawing.Font($typeFace, 35)
    $Brush = New-Object Drawing.SolidBrush ([System.Drawing.Color]::FromArgb($txtShadow[0], $txtShadow[1], $txtShadow[2], $txtShadow[3]))
    $Image.DrawString($line2, $Font, $Brush, $shadowPos, 82, $stringFormat)
    
    Write-Verbose "Draw line3 shadow: $line3"
    $Font = New-object System.Drawing.Font($typeFace, 35)
    $Brush = New-Object Drawing.SolidBrush ([System.Drawing.Color]::FromArgb($txtShadow[0], $txtShadow[1], $txtShadow[2], $txtShadow[3]))
    $Image.DrawString($line3, $Font, $Brush, $shadowPos, 133, $stringFormat)

    Write-Verbose "Draw line4 shadow: $line4"
    $Font = New-object System.Drawing.Font($typeFace, 35)
    $Brush = New-Object Drawing.SolidBrush ([System.Drawing.Color]::FromArgb($txtShadow[0], $txtShadow[1], $txtShadow[2], $txtShadow[3]))
    $Image.DrawString($line4, $Font, $Brush, $shadowPos, 184, $stringFormat)
    
    ##################
    ## REGULAR TEXT ##
    ##################
    Write-Verbose "Draw line1: $line1"
    $Font = new-object System.Drawing.Font($typeFace, 35)
    $Brush = New-Object Drawing.SolidBrush ([System.Drawing.Color]::FromArgb($txtColor[0], $txtColor[1], $txtColor[2], $txtColor[3]))
	$Image.DrawString($line1, $Font, $Brush, $textpos, 10, $stringFormat)
	

    Write-Verbose "Draw line2: $line2"
    $Font = New-object System.Drawing.Font($typeFace, 35)
    $Brush = New-Object Drawing.SolidBrush ([System.Drawing.Color]::FromArgb($txtColor[0], $txtColor[1], $txtColor[2], $txtColor[3]))
    $Image.DrawString($line2, $Font, $Brush, $textpos, 80, $stringFormat)
    
    Write-Verbose "Draw line3: $line3"
    $Font = New-object System.Drawing.Font($typeFace, 35)
    $Brush = New-Object Drawing.SolidBrush ([System.Drawing.Color]::FromArgb($txtColor[0], $txtColor[1], $txtColor[2], $txtColor[3]))
    $Image.DrawString($line3, $Font, $Brush, $textpos, 131, $stringFormat)

    Write-Verbose "Draw line3: $line3"
    $Font = New-object System.Drawing.Font($typeFace, 35)
    $Brush = New-Object Drawing.SolidBrush ([System.Drawing.Color]::FromArgb($txtColor[0], $txtColor[1], $txtColor[2], $txtColor[3]))
    $Image.DrawString($line4, $Font, $Brush, $textpos, 182, $stringFormat)
    
    
    ####################
    ## SAVE THE IMAGE ##
    ####################
    Write-Verbose "Save and close the files"
    $bmpFile.save($destPath, [System.Drawing.Imaging.ImageFormat]::Jpeg)
    $bmpFile.Dispose()
    $srcLogo.Dispose()
    
}


# Get operating system
$win10 = (gwmi win32_operatingsystem).caption -like "*10*"
$win8  = (gwmi win32_operatingsystem).caption -like "*8*"
$win7  = (gwmi win32_operatingsystem).caption -like "*7*"

# Copy down the theme file
Copy-Item $sourcetheme -Destination $themepath
$themeName = Get-ItemProperty $sourcetheme | select -Expand Name

# Check if the wallpaper directory is on the client machine
if (!(Test-Path -path $wallpaperpath)) {
    New-Item $wallpaperpath -type directory
}

# Get all the wallpaper images fom the wallpaper directory
$sourceimages = Get-ChildItem -Path "$sourcewallpaperpath" -include *jpg

## Create a slideshow if there are multiple images in the wallpaper folder
if ($sourceimages.count -gt 1){
    $n=0

    # We need to set up the footer information for slideshows
    $themefooter=@(
        "[Slideshow]"
        "Interval=1000"
        "Shuffle=0"
        "ImagesRootPath=%SystemRoot%\Web\Wallpaper"
    )

    # BGinfo the all images and save them, updating the theme each time
    foreach ($sourceimage in $sourceimages) {
        # Add image to slideshowheader array
        $themefooter+="Item$($n)Path=$wallpaperpath\img$($n).jpg"
        
        # BGinfo the image
        AddTextToImage -logopath $logoPath -sourcePath "$sourceimage" -destPath "$wallpaperpath\img$($n).jpg" -line1 $line1 -line2 $line2 -line3 $line3 -line4 $line4 -typeFace $typeFace -bgColor $bgColor
        
        # On to the next one
        $n++
    }

    # Update the ITS.theme file to include the slideshow information
    foreach ($element in $themefooter) { Add-Content $themepath\$themeName "$element" }

}

else {
    # Use theme as-is but copy it and save the image
    if(!$sourceimages){
        # No image, use plain colored background
        AddTextToImage -logopath $logoPath -destPath "$wallpaperpath\img0.jpg" -line1 $line1 -line2 $line2 -line3 $line3 -line4 $line4 -typeFace $typeFace -bgColor $bgColor
    } else {
        $sourceimage = $sourceimages
        AddTextToImage -logoPath $logoPath -sourcePath "$sourceimage" -destPath "$wallpaperpath\img0.jpg" -line1 $line1 -line2 $line2 -line3 $line3 -line4 $line4 -typeFace $typeFace -bgColor $bgColor
    
    }
}

if ($win7) {
    New-Item -ItemType directory -Path "C:\Windows\System32\oobe\info\backgrounds" -force
    AddTextToImage -logoPath $logoPath -destPath "C:\Windows\System32\oobe\info\backgrounds\backgroundDefault.jpg" -line1 $line1 -line2 $line2 -line3 $line3 -line4 $line4 -typeFace $typeFace -bgColor $bgColor
}
if ($win8 -or $win10) {
    Copy-Item -path "$wallpaperpath\img0.jpg" -Destination "C:\Windows\Web\Wallpaper\Windows\" -force
}

# Update the registry and we're done!
$RegKey ="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\"
Set-ItemProperty -path $RegKey -name InstallTheme -value $themepath\$themename

$LogonBGRegKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\Background"
Set-ItemProperty -path $LogonBGRegKey -name OEMBackground -value 1
Set-ItemProperty -path "HKCU:\Control Panel\Desktop" -name wallpaper -value "$wallpaperpath\img0.jpg"