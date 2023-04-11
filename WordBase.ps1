###
# Description:
#   Class to create Microsoft(r) Word documents.
#   - You must have Microsoft(r) Word installed.
#   - The yaml file follows the schema described below.
#
# See yaml_resume_to_word.ps1 for usage examples.
class WordBase
{

    [object]$word;
    [object]$doc;
    [object]$selection;
    [string]$baseFontName;
    [int]$baseFontSize;
    [hashtable]$styles;
    [hashtable]$wdParagraphAlignment = @{
        wdAlignParagraphCenter = 1;
        wdAlignParagraphDistribute = 4;
        wdAlignParagraphJustify = 3;
        wdAlignParagraphJustifyHi = 7;
        wdAlignParagraphJustifyLow = 8;
        wdAlignParagraphJustifyMed = 5;
        wdAlignParagraphLeft = 0;
        wdAlignParagraphRight = 2;
        wdAlignParagraphThaiJustify = 9;
    };
    [hashtable]$wdListGalleryType = @{
        wdBulletGallery = 1;
        wdNumberGallery = 2;
        wdOutlineNumberGallery = 3;
    }
    [hashtable]$wdListApplyTo = @{
        wdListApplyToWholeList = 0;
        wdListApplyToThisPointForward = 1;
        wdListApplyToSelection = 2;
    };
    [hashtable]$wdDefaultListBehavior = @{
        wdWord10ListBehavior = 2;
        wdWord8ListBehavior = 0;
        wdWord9ListBehavior = 1;
    };
    [hashtable]$wdNumberType = @{
        wdNumberParagraph = 1;
        wdNumberListNum = 2;
        wdNumberAllNumbers = 3;
    };
    [hashtable]$wdColor = @{
        wdColorBlack = 0;
        wdColorDarkGreen = 13056;
        wdColorDarkBlue = 8388608;
        wdColorDarkRed = 128;
        wdColorGray50 = 8421504;
        wdColorGreen = 32768;
        wdColorBlue = 16711680;
        wdColorRed = 255;
    };
    WordBase()
    {
        $this.word = New-Object -ComObject Word.Application;
        $this.word.visible = $true;
        $this.baseFontName = "Trebuchet MS";
        $this.baseFontSize = 11;
        $this.doc = $this.word.Documents.Add();
        $this.selection = $this.word.selection;
    }

    [void]Save([string]$file_path)
    {
        $this.doc.SaveAs([ref]$file_path)
    }

    [void]SelectStyle([string]$styleName)
    {
        $this.selection.Style = $this.doc.Styles[$styleName];
    }

    [void]SelectAlignment([int]$alignment)
    {
        $this.selection.ParagraphFormat.Alignment = $alignment;
    }

    [void]ChangeStyle([string]$styleName, [string]$fontName = $this.baseFontName, [int]$fontSize = $this.$baseFontSize, [bool]$italic = $false, [bool]$bold = $false, [int]$fontColor = 0, [int]$paragraphAlignment = 0)
    {
        $this.doc.Styles[$styleName].Font.Name = $fontName;
        $this.doc.Styles[$styleName].Font.Size = $fontSize;
        $this.doc.Styles[$styleName].Font.Bold = if ($bold) { 1 } else { 0 };;
        $this.doc.Styles[$styleName].Font.Italic =if ($italic) { 1 } else { 0 };
        $this.doc.Styles[$styleName].Font.Name = $fontName;
        $this.doc.Styles[$styleName].Font.Color = $fontColor;
        $this.doc.Styles[$styleName].ParagraphFormat.Alignment = $paragraphAlignment;
    }

    [void]SetItalic([bool]$italic)
    {
        $this.selection.font.italic = if ($italic) { 1 } else { 0 };
    }

    [void]SetBold([bool]$bold)
    {
        $this.selection.font.bold = if ($bold) { 1 } else { 0 };
    }

    [void]SetFontSize([int]$fontSize)
    {
        $this.selection.font.size = $fontSize;
    }    

    [void]SetFontName([string]$fontName)
    {
        $this.selection.font.name = $fontName;
    }    

    [void]BulletList([string[]]$textList)
    {
        $this.selection.Range.ListFormat.ApplyBulletDefault();
        $this.ListText($textList)
    }

    [void]ListText([string[]]$textList)
    {
        foreach ($text in $textList)
        {
            $this.selection.TypeText($text);
            $this.NewParagraph();
        }
        $this.selection.Range.ListFormat.RemoveNumbers($this.wdNumberType['wdNumberParagraph']);
    }

    [void]NewParagraph()
    {
        $this.selection.TypeParagraph();
    }

    [void]StyledText([string]$text, [string]$styleName, [int]$alignment)
    {
        $this.SelectStyle($styleName);
        $this.SelectAlignment($alignment);
        $this.selection.TypeText($text);
        $this.NewParagraph();
    }

    [void]Text([string]$text, [string]$fontName = $this.baseFontName, [int]$fontSize = $this.$baseFontSize, [bool]$italic = $false, [bool]$bold = $false)
    {
        $this.SetFontName($fontName);
        $this.SetFontSize($fontSize);
        $this.SetBold($bold);
        $this.SetItalic($italic);
        $this.selection.TypeText($text);
    }

    [void]Close()
    {
        $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$this.word)
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()        
    }
}

