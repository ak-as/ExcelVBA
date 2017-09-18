Attribute VB_Name = "ImpCsvCore"
Option Explicit

Function ImportCsvAsRange( _
        Destination As Range, _
        FilePath As String, _
        Optional StartRow As Long = 1, _
        Optional ColumnDataType As XlColumnDataType = xlGeneralFormat, _
        Optional TextQualifier As XlTextQualifier = xlTextQualifierDoubleQuote, _
        Optional CodePage As MsoEncoding = msoEncodingJapaneseShiftJIS _
    ) As Range

    Dim QueryTable As QueryTable
    Set QueryTable = ImportCsvAsQueryTable(Destination, FilePath, StartRow, ColumnDataType, TextQualifier, CodePage)
    Set ImportCsvAsRange = QueryTable.ResultRange
    Call QueryTable.Delete

End Function

Function ImportCsvAsQueryTable( _
        Destination As Range, _
        FilePath As String, _
        Optional StartRow As Long = 1, _
        Optional ColumnDataType As XlColumnDataType = xlGeneralFormat, _
        Optional TextQualifier As XlTextQualifier = xlTextQualifierDoubleQuote, _
        Optional CodePage As MsoEncoding = msoEncodingJapaneseShiftJIS _
    ) As QueryTable

    'コードページ→文字セット名の変換
    Dim CharSet As String
    CharSet = CodePageToCharSet(CodePage)

    If CharSet = "" Then
        CharSet = "_autodetect"
    End If

    'カラム数の取得
    Dim ColumnCount As Long
    ColumnCount = EstimateColumnCount(FilePath, StartRow, CharSet)

    'カラムのデータ型の指定
    Dim DataTypes()
    Call ReallocAndFill(DataTypes, 1, ColumnCount, ColumnDataType)

    'クエリーテーブルの作成
    Dim QueryTable As QueryTable
    Set QueryTable = Destination.Parent.QueryTables.Add( _
            Connection:="TEXT;" & FilePath, _
            Destination:=Destination.Cells(1) _
        )

    'インポート設定と実行
    With QueryTable
        .FieldNames = False
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = False
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = CodePage
        .TextFileStartRow = StartRow
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = TextQualifier
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = DataTypes
        .TextFileTrailingMinusNumbers = True
        Call .Refresh(BackgroundQuery:=False)
    End With

    Set ImportCsvAsQueryTable = QueryTable

End Function

Private Function CodePageToCharSet(CodePage As Long) As String

    Dim StdRegProv As Object, Value

    Set StdRegProv = GetObject("winmgmts:\\.\root\default:StdRegProv")
    Call StdRegProv.GetStringValue(&H80000000, "MIME\Database\CodePage\" & CodePage, "WebCharset", Value)
    If IsNull(Value) Then
        Call StdRegProv.GetStringValue(&H80000000, "MIME\Database\CodePage\" & CodePage, "BodyCharset", Value)
        If IsNull(Value) Then
            CodePageToCharSet = ""
        Else
            CodePageToCharSet = Value
        End If
    Else
        CodePageToCharSet = Value
    End If

End Function

Private Function EstimateColumnCount( _
        FilePath As String, _
        Optional StartRow As Long = 1, _
        Optional CharSet As String = "shift_jis", _
        Optional LineSeparator As Long = -1, _
        Optional MaxScanRows As Long = 8 _
    ) As Long

    Dim CurrRow As Long, LastRow As Long
    Dim Upper As Long, MaxUpper As Long
    Dim Stream As Object

    CurrRow = 0
    LastRow = StartRow + MaxScanRows - 1
    MaxUpper = 0
    Set Stream = CreateObject("ADODB.Stream")

    With Stream
        Call .Open
        .Position = 0
        .Type = 2 'adTypeText:2
        .CharSet = CharSet
        .LineSeparator = LineSeparator
        Call .LoadFromFile(FilePath)

        Do Until .EOS
            CurrRow = CurrRow + 1
            If CurrRow > LastRow Then
                Exit Do
            End If
            If CurrRow < StartRow Then
                Call Stream.SkipLine
            Else
                Upper = UBound(Split(.ReadText(-2), ","))
                MaxUpper = IIf(MaxUpper < Upper, Upper, MaxUpper)
            End If
        Loop

        Call .Close
    End With

    EstimateColumnCount = MaxUpper + 1

End Function

Private Sub ReallocAndFill(SourceArray(), Lower As Long, Upper As Long, Value)

    ReDim SourceArray(Lower To Upper)
    Dim i As Long
    For i = Lower To Upper
        SourceArray(i) = Value
    Next

End Sub
