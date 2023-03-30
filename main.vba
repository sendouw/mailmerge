Sub MailMergeAndSaveSeparateFiles()
    Dim MainDoc As Document, TargetDoc As Document
    Dim MergedRecords As MailMerge.DataSource
    Dim FilePath As String, FileName As String
    Dim i As Long
    
    FilePath = "C:\MailMerge\" 
    If Right(FilePath, 1) <> "\" Then FilePath = FilePath & "\"
    
    ' Check if the folder exists, and create it if not
    If Dir(FilePath, vbDirectory) = "" Then
        MkDir FilePath
    End If
    
    Set MainDoc = ActiveDocument
    With MainDoc.MailMerge
        .Destination = wdSendToNewDocument
        .SuppressBlankLines = True
        
        With .DataSource
            .FirstRecord = wdDefaultFirstRecord
            .LastRecord = wdDefaultLastRecord
        End With
        
        For i = 1 To .DataSource.RecordCount
            .DataSource.ActiveRecord = i
            .Execute Pause:=False
            
            Set TargetDoc = ActiveDocument
            FileName = FilePath & "Property Number " & .DataSource.DataFields("PropertyNumber").Value & ".docx"
            
            With TargetDoc
                .SaveAs FileName, FileFormat:=wdFormatDocumentDefault
                .Close SaveChanges:=False
            End With
        Next i
    End With
End Sub
