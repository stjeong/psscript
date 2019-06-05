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
        if ($n -ge 44032 -and $n -le 55215 )
        {
            $nText = $n.ToString()
            $null = $wordBuilder.Append("\u$($nText)?")
        }
        else
        {
            $nText = $n.ToString()
            if ($nText -ne '\pard')
            {
                $null = $wordBuilder.Append($ch)
            }
        }
    }

    $null = $rtfBuilder.Append("\cf$colorIndex $wordBuilder")
}

# Generate an HTML span and append it to HTML string builder
#
function Append-HtmlSpan ($block, $tokenColor)
{
  if ($tokenColor -eq 'NewLine')
  {
    $null = $htmlBuilder.Append("<br>")
  }
  else
  {
    $block = $block.Replace('&','&amp;').Replace('>','&gt;').Replace('<','&lt;')
    if (-not $block.Trim())
    {
        $block = $block.Replace(' ', '&nbsp;')
    }
    $htmlColor = $psise.Options.TokenColors[$tokenColor].ToString().Replace('#FF', '#')
    $null = $htmlBuilder.Append("<span style='color:$htmlColor'>$block</span>")
  }
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
    $fontName = 'Malgun Gothic'
    $fontSize = 10

    # Initialize HTML builder.
    $htmlBuilder = new-object system.text.stringbuilder
    $null = $htmlBuilder.AppendLine("<p style='MARGIN: 0in 10pt 0in;font-family:$fontname;font-size:$fontSize`pt'>")

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
            Append-HtmlSpan $block $tokenColor
        }
        
        $block = $text.Substring($token.Start, $token.Length)
        $tokenColor = $token.Type.ToString()
        Append-RtfBlock $block $tokenColor
        Append-HtmlSpan $block $tokenColor
        
        $position = $token.Start + $token.Length
    }

    # Append HTML ending tag.
    $null = $htmlBuilder.Append("</p>")

    # Append RTF ending brace.
    $null = $rtfBuilder.Append('}')

    # Copy console screen buffer contents to clipboard in three formats - text, HTML and RTF.
    #
    $dataObject = New-Object Windows.DataObject

    # $dataObject.SetText([string]$text, [Windows.TextDataFormat]"UnicodeText")

    $rtf = $rtfBuilder.ToString()
    $dataObject.SetText([string]$rtf, [Windows.TextDataFormat]"Rtf")

# This prevent from pasting into Powerpoint
#    $html = $htmlBuilder.ToString()
#    $dataObject.SetText([string]$html, [Windows.TextDataFormat]"Html")

    [Windows.Clipboard]::SetDataObject($dataObject, $true)
    # 'The script has been copied to clipboard.'
}

Copy-Script
