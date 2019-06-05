# https://devblogs.microsoft.com/powershell/colorized-capture-of-console-screen-in-html-and-rtf/

# Create RTF block from text using named console colors.
#
function Append-RtfBlock ($block, $tokenColor)
{
    $colorIndex = $rtfColorMap.$tokenColor
    $block = $block.Replace('\','\\').Replace("`r`n","\cf1\par`r`n").Replace("`t",'\tab').Replace('{','\{').Replace('}','\}')

    $wordBuilder = new-object system.text.stringbuilder

    foreach ($ch in $block.ToCharArray())
    {
        $n = [int]$ch
        if ($n -ge 255)
        {
            $nText = $n.ToString()
            $null = $wordBuilder.Append("\u$($nText)?")
        }
        else
        {
            $null = $wordBuilder.Append($ch)
        }
    }

    $null = $rtfBuilder.Append("\cf$colorIndex $wordBuilder")
}

function Copy-Script
{
    if (-not $psise.CurrentFile)
    {
        Write-Error 'No script is available for copying.'
        return
    }
    
    $text = $psise.CurrentFile.Editor.Text

    trap { break }

    # Do syntax parsing.
    $errors = $null
    $tokens = [system.management.automation.psparser]::Tokenize($Text, [ref] $errors)

    # Set the desired font and font size
    $fontName = 'Lucida Console'
    $fontSize = 10

    # Initialize RTF builder.
    $rtfBuilder = new-object system.text.stringbuilder
    # Append RTF header
    $null = $rtfBuilder.Append("{\rtf1\fbidis\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil\fcharset0 $fontName;}}")

    $null = $rtfBuilder.Append("`r`n")
    # Append RTF color table which will contain all Powershell console colors.
    $null = $rtfBuilder.Append("{\colortbl ;")
    # Generate RTF color definitions for each token type.
    $rtfColorIndex = 1
    $rtfColors = @{}
    $rtfColorMap = @{}
    [Enum]::GetNames([System.Management.Automation.PSTokenType]) | % {
        $tokenColor = $psise.Options.TokenColors[$_];
        $rtfColor = "\red$($tokenColor.R)\green$($tokenColor.G)\blue$($tokenColor.B);"
        if ($rtfColors.Keys -notcontains $rtfColor)
        {
            $rtfColors.$rtfColor = $rtfColorIndex
            $null = $rtfBuilder.Append($rtfColor)
            $rtfColorMap.$_ = $rtfColorIndex
            $rtfColorIndex ++
        }
        else
        {
            $rtfColorMap.$_ = $rtfColors.$rtfColor
        }
    }
    $null = $rtfBuilder.Append('}')
    $null = $rtfBuilder.Append("`r`n")
    # Append RTF document settings.
    $null = $rtfBuilder.Append('\viewkind4\uc1\f0\fs20 ')

    $position = 0
    # Iterate over the tokens and set the colors appropriately.
    foreach ($token in $tokens)
    {
        if ($position -lt $token.Start)
        {
            $block = $text.Substring($position, ($token.Start - $position))

            $tokenColor = 'Unknown'
            Append-RtfBlock $block $tokenColor
        }
        
        $block = $text.Substring($token.Start, $token.Length)
        $tokenColor = $token.Type.ToString()
        Append-RtfBlock $block $tokenColor
        
        $position = $token.Start + $token.Length
    }

    # Append RTF ending brace.
    $null = $rtfBuilder.Append('}')

    # Copy console screen buffer contents to clipboard in RTF formats.
    #
    $dataObject = New-Object Windows.DataObject

    $rtf = $rtfBuilder.ToString()
    $dataObject.SetText([string]$rtf, [Windows.TextDataFormat]"Rtf")

    [Windows.Clipboard]::SetDataObject($dataObject, $true)
}

Copy-Script
