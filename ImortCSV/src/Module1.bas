Attribute VB_Name = "Module1"
Option Explicit

Sub CSVファイルを開く()

    Const OF_TITLE = "CSVファイルを開く"
    Const OF_FILE_FILTER = "CSV File(*.csv),*.csv,All Files(*.*),*.*"

    Dim SelectedFiles
    SelectedFiles = Application.GetOpenFilename(FileFilter:=OF_FILE_FILTER, TITLE:=OF_TITLE, MultiSelect:=True)

    If VarType(SelectedFiles) = vbBoolean Then
        Exit Sub
    End If

    Call OpenCSVFiles(SelectedFiles)

End Sub

Function OpenCSVFiles(CSVFiles) As Workbook

    Dim Files()
    Files = ToVariantArray(CSVFiles)

    Dim Book As Workbook
    Set Book = CreateWorkbook(1)

    Dim i As Long
    Dim Sheet As Worksheet
    Dim ResultRange As Range

    For i = LBound(Files) To UBound(Files)

        If i = LBound(Files) Then
            Set Sheet = Book.Worksheets(1)
        Else
            Set Sheet = Book.Worksheets.Add
        End If

        Set ResultRange = ImportCsvAsRange( _
                            Destination:=Sheet.Cells(1), _
                            FilePath:=CStr(Files(i)), _
                            StartRow:=1, _
                            ColumnDataType:=xlTextFormat, _
                            TextQualifier:=xlTextQualifierNone, _
                            CodePage:=msoEncodingJapaneseShiftJIS _
                        )

        Call ApplyStyle(ResultRange)
    Next

    Set OpenCSVFiles = Book

End Function

Function ToVariantArray(Source) As Variant()
    Dim Tmp(), i As Long
    If (VarType(Source) And vbArray) <> 0 Then
        If VarType(Source) = (vbArray Or vbVariant) Then
            Tmp = Source
        Else
            ReDim Tmp(LBound(Source) To UBound(Source))
            For i = LBound(Source) To UBound(Source)
                Tmp(i) = Source(i)
            Next
        End If
    Else
        ReDim Tmp(1 To 1)
        Tmp(1) = Source
    End If
    ToVariantArray = Tmp
End Function

Function CreateWorkbook(Optional Template, Optional SheetsInNewWorkbook) As Workbook

    Dim OldValue As Long
    Dim Book As Workbook

    If IsMissing(SheetsInNewWorkbook) Then
        Set Book = Workbooks.Add(Template)
    Else
        OldValue = Application.SheetsInNewWorkbook
        Application.SheetsInNewWorkbook = SheetsInNewWorkbook
        Set Book = Workbooks.Add(Template)
        Application.SheetsInNewWorkbook = OldValue
    End If

    Set CreateWorkbook = Book

End Function

Sub ApplyStyle(ResultRange As Range)

    'ResultRange

End Sub

