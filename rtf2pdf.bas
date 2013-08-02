Attribute VB_Name = "NewMacros"
Sub rtf2pdf()
    Dim vDirectory As String
    Dim oDoc As Document
    
    ActiveWindow.ActivePane.DisplayRulers = Not ActiveWindow.ActivePane.DisplayRulers
    Application.ScreenUpdating = False
    
    vDirectory = "H:\bbb\"
    vFile = Dir(vDirectory & "*.rtf")
    
    Do While vFile <> ""
        Set oDoc = Documents.Open(FileName:=vDirectory & vFile, ConfirmConversions:=False)

        strOutFile = vDirectory & Left(vFile, Len(vFile) - 4) & ".pdf"
    
        oDoc.ExportAsFixedFormat OutputFileName:= _
        strOutFile, ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
        wdExportAllDocument, From:=1, To:=1, Item:=wdExportDocumentContent, _
        IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:= _
        wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
        True, UseISO19005_1:=False

        oDoc.Close SaveChanges:=False
        vFile = Dir
    Loop
    
    Application.ScreenUpdating = True
    ActiveWindow.ActivePane.DisplayRulers = Not ActiveWindow.ActivePane.DisplayRulers
End Sub
