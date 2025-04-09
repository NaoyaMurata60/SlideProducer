# SECTION: Constants
$TYPE1 = "試験向け"
$TYPE2 = "運用向け"
$BOTH = "試験/運用向け"

# Add the required .NET assembly
Add-Type -AssemblyName System.Windows.Forms

# SECTION: Function to Close PowerPoint and Release Resources
function Close-PowerPointAndReleaseResources {
    param (
        [Parameter(Mandatory = $true)]
        $powerpoint,
        [Parameter(ValueFromRemainingArguments = $true)]
        $presentations
    )

    foreach ($presentation in $presentations) {
        if ($presentation) {
            try {
                $presentation.Close()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
                Remove-Variable -Name ($presentation.Name) -ErrorAction SilentlyContinue
            } catch {
                Write-Warning "Failed to close and release presentation: $($_.Exception.Message)"
            }
        }
    }

    if ($powerpoint) {
        try {
            $powerpoint.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint) | Out-Null
            Remove-Variable powerpoint -ErrorAction SilentlyContinue
        } catch {
            Write-Warning "Failed to quit PowerPoint application: $($_.Exception.Message)"
        }
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
}

# SECTION: Function to Update Slide
function Update-Slide {
    param (
        [Parameter(Mandatory = $true)]
        $slide,
        [Parameter(Mandatory = $true)]
        $presentation,
        [Parameter(Mandatory = $true)]
        $textBoxShape,
        [Parameter(Mandatory = $true)]
        [string]$textBoxText,
        [Parameter(Mandatory = $true)]
        [string]$presentationType
    )

    $condition1 = $textBoxText -eq $TYPE2 -and $presentationType -eq $TYPE1 # delete the slide from type 1 presetation
    $condition2 = $textBoxText -eq $TYPE1 -and $presentationType -eq $TYPE2 # delete the slide from type 2 presetation
    $condition3 = $null -eq $textBoxShape # delete the slide from both presentation

    if ($condition1 -or $condition2 -or $condition3) {
        # スライドが存在する場合のみ削除
        if ($slide) {
            try {
                $slide.Delete()
            } catch {
                Write-Error "Failed to delete slide: $($_.Exception.Message)"
            }
        }
    }

    # テキストボックスを削除
    if ($textBoxShape) {
        for ($i = $slide.Shapes.Count; $i -gt 0; $i--) {
            $shape = $slide.Shapes.Item($i)
            if ($shape.Name -eq $textBoxShape.Name) {
                try {
                    $shape.Delete()
                    break
                } catch {
                    Write-Error "Failed to delete shape: $($_.Exception.Message)"
                }
            }
        }
    }
}

# Function to bring a window to the foreground using its title
function Set-WindowToForeground {
    param (
        [Parameter(Mandatory = $true)]
        [string]$windowTitle
    )

    Add-Type @"
        using System;
        using System.Runtime.InteropServices;
        public class Win32 {
            [DllImport("user32.dll")]
            public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool SetForegroundWindow(IntPtr hWnd);
        }
"@

    $hWnd = [Win32]::FindWindow([NullString]::Value, $windowTitle)

    if ($hWnd -ne [IntPtr]::Zero) {
        [Win32]::SetForegroundWindow($hWnd) | Out-Null
        Start-Sleep -Milliseconds 500 # Add delay to ensure window focus
        return $true
    } else {
        Write-Host "Window '$windowTitle' not found."
        return $false
    }
}

# Function to send key sequence
# Function to send key sequence with individual key presses
function Send-KeySequence {
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$sequence
    )

    foreach ($key in $sequence) {
        [System.Windows.Forms.SendKeys]::SendWait($key)
        Start-Sleep -Milliseconds 100 # Delay between key presses
    }
}

# SECTION: Initialize PowerPoint Application
try {
    $powerpoint = New-Object -ComObject PowerPoint.Application
    $powerpoint.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
} catch {
    Write-Error "Failed to initialize PowerPoint application: $($_.Exception.Message)"
    exit
}

# SECTION: Select Master Presentation from Current Directory
$pptxFiles = Get-ChildItem -Path (Get-Location) -Filter "*.pptx"
if ($pptxFiles.Count -eq 0) {
    Write-Error "No pptx files found in the current directory."
    Close-PowerPointAndReleaseResources -powerpoint $powerpoint
    exit
}

Write-Host "Available pptx files in the current directory:"
for ($i = 0; $i -lt $pptxFiles.Count; $i++) {
    Write-Host "$($i + 1). $($pptxFiles[$i].Name)"
}

if ($pptxFiles.Count -eq 1) {
    $masterPath = $pptxFiles[0].FullName
} else {
    do {
        $userInput = Read-Host "Enter the number of the pptx file to use"
        if ($userInput -match "^\d+$") {
            $selectedIndex = [int]$userInput - 1
            if ($selectedIndex -ge 0 -and $selectedIndex -lt $pptxFiles.Count) {
                $masterPath = $pptxFiles[$selectedIndex].FullName
                break
            }
        }
        Write-Host "Invalid input. Please enter again."
    } while ($true)
}


# SECTION: Copy and Create New Presentations from Master
try {
    $type1Path = Join-Path -Path (Get-Location) -ChildPath "$TYPE1.pptx"
    $type2Path = Join-Path -Path (Get-Location) -ChildPath "$TYPE2.pptx"

    Copy-Item -Path $masterPath -Destination $type1Path -Force
    Copy-Item -Path $masterPath -Destination $type2Path -Force

    $type1Presentation = $powerpoint.Presentations.Open($type1Path)
    $type2Presentation = $powerpoint.Presentations.Open($type2Path)

    $presentationInfoArray = @()
    $presentationInfoArray += @{ Type = $TYPE1; Presentation = $type1Presentation }
    $presentationInfoArray += @{ Type = $TYPE2; Presentation = $type2Presentation }
} catch {
    Write-Error "Failed to create type1.pptx or type2.pptx: $($_.Exception.Message)"
    Close-PowerPointAndReleaseResources -powerpoint $powerpoint -presentations $type1Presentation, $type2Presentation
    exit
}

# SECTION: Process Each Slide
$multipleTagSlideCount = 0
foreach ($presentationInfo in $presentationInfoArray) {
    $presentation = $presentationInfo.Presentation
    $presentationType = $presentationInfo.Type
    for ($i = $presentation.Slides.Count; $i -gt 0; $i--) {
        $slide = $presentation.Slides.Item($i)

        # SECTION: Find Target Textbox
        $textBoxText = ""
        $textBoxShape = $null
        $foundTextBox = $false
        $matchingShapes = @()
        foreach ($shape in $slide.Shapes) {
            if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                $textBoxText = $shape.TextFrame.TextRange.Text
                if ($textBoxText -eq $TYPE1 -or $textBoxText -eq $TYPE2 -or $textBoxText -eq $BOTH) {
                    $matchingShapes += $shape
                }
            }
        }

        if ($matchingShapes.Count -gt 1) {
            Write-Host "Multiple textboxes for file generation flag found on page $i of master.pptx.\nThis slide is deleted from new presentation files."
            $textBoxShape = $null
            $textBoxText = ""
            $foundTextBox = $ture
            $multipleTagSlideCount += 1
        } elseif ($matchingShapes.Count -eq 1) {
            $textBoxShape = $matchingShapes[0]
            $textBoxText = $textBoxShape.TextFrame.TextRange.Text
            $foundTextBox = $true
        }

        if (-not $foundTextBox) {
            Write-Error "Textbox for file generation flag not found on page $i of master.pptx."
            Write-Host "Press any key to continue..."
            $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
            Close-PowerPointAndReleaseResources -powerpoint $powerpoint -presentations $type1Presentation, $type2Presentation
            exit
        }

        # Function call to process slides
        Update-Slide -slide $slide -presentation $presentation -textBoxShape $textBoxShape -textBoxText $textBoxText -presentationType $presentationType
    }
}
if ($multipleTagSlideCount -gt 1) {
    Write-Host "Caution: Multiple textboxes for file generation flag on 2 or more slides."
}

# SECTION: Save and Close Type Presentations

# Activate the PowerPoint window and send key sequence
$array = @($TYPE1, $TYPE2)
foreach ($_ in $array) {
    $windowTitle = "$($_).pptx - PowerPoint"
    if (Set-WindowToForeground -windowTitle $windowTitle) {
        # Send Alt key first, then individual keys
        $keySequence = @("%", "f", "i", "p", "{ENTER}", "{ESC}")
        Send-KeySequence -sequence $keySequence
    } else {
        Write-Error "Failed to activate window: $windowTitle"
    }
}

# Proceed with saving presentations
try {
    $type1Presentation.SaveAs($type1Path)
    $type2Presentation.SaveAs($type2Path)
} catch {
    Write-Error "Failed to save type1.pptx or type2.pptx: $($_.Exception.Message)"
} finally {
    Close-PowerPointAndReleaseResources -powerpoint $powerpoint -presentations $type1Presentation, $type2Presentation
}