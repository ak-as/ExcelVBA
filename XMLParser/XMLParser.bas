Attribute VB_Name = "XMLParser"
Option Explicit

Sub XMLファイル解析()

    Dim FilePath, ReservedAttrNames As String
    Dim doc As MSXML2.DOMDocument60
    Dim Root As MSXML2.IXMLDOMElement
    Dim AttrNames As VBA.Collection
    Dim Cell As Range, ResultRange As Range

    FilePath = Application.GetOpenFilename(FileFilter:="XML File (*.xml),*.xml,All Files,*", Title:="XMLファイル解析")

    If VarType(FilePath) = vbBoolean Then
        Exit Sub
    End If

    ReservedAttrNames = InputBox("優先する属性名を入力して下さい。", "属性名の入力", "name id value")

    If ReservedAttrNames = "" Then
        Exit Sub
    End If

    Set Cell = ActiveCell

    Set doc = New MSXML2.DOMDocument60
    doc.async = False

    If Not doc.Load(FilePath) Then
        Call MsgBox(doc.parseError.reason)
        Exit Sub
    End If

    Set Root = doc.DocumentElement

    Set AttrNames = GetAttrNames(Root, ReservedAttrNames)

    Call WriteHeader(AttrNames, Cell)
    Set Cell = Cell.Offset(1, 0)

    Set ResultRange = WriteNode(Root, AttrNames, Cell)

    If Not ResultRange Is Nothing Then
        ResultRange.Select
    End If

End Sub

Private Function GetAttrNames(Node As MSXML2.IXMLDOMNode, Optional ReservedAttrName As String) As VBA.Collection

    Dim Dict As New Scripting.Dictionary
    Dim Reserved, i As Long
    Dim Arr() As String, Key
    Dim AttrNames As New VBA.Collection

    Reserved = Split(Trim(ReservedAttrName), " ")
    For i = LBound(Reserved) To UBound(Reserved)
        If Reserved(i) <> "" Then
            If Not Dict.Exists(Reserved(i)) Then
                Call Dict.Add(Reserved(i), Dict.Count)
            End If
        End If
    Next

    Call CreateAttrNameDict(Node, Dict)

    ReDim Arr(0 To Dict.Count - 1) As String
    For Each Key In Dict.Keys
        Arr(Dict(Key)) = Key
    Next

    For i = LBound(Arr) To UBound(Arr)
        Call AttrNames.Add(Arr(i))
    Next

    Set GetAttrNames = AttrNames

End Function

Private Sub CreateAttrNameDict(Node As MSXML2.IXMLDOMNode, AttrNameDict As Scripting.Dictionary)

    Dim AttrNode As IXMLDOMNode
    Dim ChildNode As IXMLDOMNode

    If Not Node.Attributes Is Nothing Then
        For Each AttrNode In Node.Attributes
            If Not AttrNameDict.Exists(AttrNode.nodeName) Then
                Call AttrNameDict.Add(AttrNode.nodeName, AttrNameDict.Count)
            End If
        Next
    End If

    Set ChildNode = Node.FirstChild
    Do While Not ChildNode Is Nothing
        Call CreateAttrNameDict(ChildNode, AttrNameDict)
        Set ChildNode = ChildNode.NextSibling
    Loop

End Sub

Private Function WriteHeader(AttributeNames As VBA.Collection, Destination As Range) As Range

    Dim Cell As Range
    Set Cell = Destination.Cells(1)

    Cell.Offset(0, 0).Value = "Element Name"
    Cell.Offset(0, 1).Value = "Element Type"
    Cell.Offset(0, 2).Value = "Element Text"
    Set Cell = Cell.Offset(0, 2)

    Dim AttrName
    For Each AttrName In AttributeNames
        Set Cell = Cell.Offset(0, 1)
        Cell.Value = AttrName
    Next

End Function

Private Function WriteNode(Node As MSXML2.IXMLDOMNode, AttributeNames As VBA.Collection, Destination As Range) As Range

    Dim Cell As Range, ResultRange As Range
    Dim AttrName, AttrNode As IXMLDOMNode, ChildNode As IXMLDOMNode

    Set Cell = Destination.Cells(1)

    Cell.Value = Node.nodeName
    Set Cell = Cell.Offset(0, 1)
    Cell.Value = Node.nodeTypeString
    Set Cell = Cell.Offset(0, 1)
    Cell.Value = Node.Text

    If Node.Attributes Is Nothing Then
        Set Cell = Cell.Offset(0, AttributeNames.Count)
    Else
        For Each AttrName In AttributeNames
            Set Cell = Cell.Offset(0, 1)
            Set AttrNode = Node.Attributes.getNamedItem(AttrName)
            Call SetAttrValueToCell(AttrNode, Cell)
        Next
    End If

    Set ChildNode = Node.FirstChild
    If ChildNode Is Nothing Then
        Set WriteNode = Application.Range(Destination.Cells(1), Cell)
    Else
        Set ResultRange = Destination.Cells(1)
        Do While Not ChildNode Is Nothing
            Set Cell = ResultRange.Offset(ResultRange.Rows.Count, 0).Resize(1, 1)
            Set ResultRange = WriteNode(ChildNode, AttributeNames, Cell)
            Set ChildNode = ChildNode.NextSibling
        Loop
        Set WriteNode = Application.Range(Destination.Cells(1), ResultRange)
    End If

End Function

Private Sub SetAttrValueToCell(AttrNode As MSXML2.IXMLDOMNode, Cell As Range)
    If Not AttrNode Is Nothing Then
        Cell.NumberFormatLocal = "@"
        Cell.Value = AttrNode.NodeValue
    Else
        Cell.NumberFormatLocal = "@"
        Cell.Value = ""
    End If
End Sub
