Attribute VB_Name = "Commands"
'@Folder("VBAProject")
Option Explicit
Option Compare Text

Public Sub CommandLoadConfig(control As IRibbonControl)
    On Error GoTo Catch
     
     'Addin.L
    
Finally:
    Exit Sub
Catch:
    
    MsgBox Err.Description
    Resume Finally
End Sub

Public Sub CommandBackupConfig(control As IRibbonControl)
    On Error GoTo Catch
    
Finally:
    Exit Sub
Catch:
    
    MsgBox Err.Description
    Resume Finally
End Sub

Function GetVersion(sorf As Workbook) As String
' take version from Sheet Sofr cell A1
    On Error GoTo Catch
 GetVersion = Trim(Replace(sorf.Worksheets("SORF").Range("A1"), "ver.", ""))
Finally:
    Exit Function
Catch:
    
    MsgBox Err.Description
    Resume Finally
End Function
Public Sub CommandSyncSORF(control As IRibbonControl)
    On Error GoTo Catch
    Dim message As String
    Dim ver As String
    Dim sorfDocument As Workbook
    Set sorfDocument = ActiveWorkbook
    
    If Not CheckPrerequisite(sorfDocument, message) _
    Then Err.Raise 1000, Description:=message
   
    
    Dim rules As CustomXMLNode
    
    If Not LoadMappingRules(sorfDocument, rules, message) _
    Then Err.Raise 1000, Description:=message
    
    Dim parts As CustomXMLParts
    Set parts = sorfDocument.CustomXMLParts.SelectByNamespace("urn:bitrix:deal")
    
    Dim part As CustomXMLPart
    If parts.Count = 0 Then
        Set part = CreatePart(sorfDocument)
    Else
        Set part = parts(1)
    End If
      
    Dim dealNode  As CustomXMLNode
    Set dealNode = CreateDeal(part)
    
    Dim mapping As Collection
    Dim area As Range
    Set area = sorfDocument.Worksheets("SORF").UsedRange
    
    Set area = area.Offset(0, 1).Resize(ColumnSize:=area.Columns.Count - 1)
    
    If Not ApplyMappingRules(area, rules.SelectSingleNode("table[@name=""Deal""]"), mapping, message) _
    Then Err.Raise 1000, Description:=message
    
    If Not WriteDeal(dealNode, mapping, message) _
    Then Err.Raise 1000, Description:=message

    Set area = sorfDocument.Worksheets("Items Breakdown").UsedRange
   
    If Not ApplyMappingRules(area, rules.SelectSingleNode("table[@name=""Items""]"), mapping, message) _
    Then Err.Raise 1000, Description:=message
    
    If Not WriteItems(dealNode, mapping, message) _
    Then Err.Raise 1000, Description:=message
    
    If Not WriteItemTotals(dealNode, mapping, message) _
    Then Err.Raise 1000, Description:=message
    
    Dim chemistry As Collection
    Dim chemistrySheet As Range
    Set chemistrySheet = sorfDocument.Worksheets("Chemistry")
    If Not LoadChemistry(chemistry, chemistrySheet, message) _
    Then Err.Raise 1000, Description:=message
    
    sorfDocument.Save

Finally:

    Exit Sub
    
Catch:
    
    If Not part Is Nothing Then part.Delete

    MsgBox Err.Description
    
    Resume Finally
End Sub
Function GetFiledByName(mapping As Collection, name As String, entry As Pair) As Boolean

    Dim found As Boolean
    For Each entry In mapping
        
        found = entry.second.GetName() = name
        
        If found Then Exit For

    Next
    
    GetFiledByName = found

End Function

Function WriteItemTotals(node As CustomXMLNode, mapping As Collection, message As String) As Boolean
    On Error GoTo Catch
    
    Dim itemNumberMapping As Pair
    If Not GetFiledByName(mapping, "Item Number", itemNumberMapping) _
    Then Err.Raise 1000, Description:="В мапинге отсуствует поле номера позиции"
    
    Dim itemNumberHeaderCell As Range
    Set itemNumberHeaderCell = itemNumberMapping.first
    
    Dim totalHeaderCell As Range
    Set totalHeaderCell = itemNumberHeaderCell.EntireRow.Find("Total")
    
    If totalHeaderCell Is Nothing _
    Then Err.Raise 1000, Description:="Неудалось найти заголовок Итогов"
    
    Dim totalCell As Range
    Dim field As CField
    
    Dim entry As Pair
    For Each entry In mapping
        
        Set field = entry.second
        
        If Not field.HasTotal Then GoTo Continue
    
        Set totalCell = totalHeaderCell.Offset(entry.first.Row - totalHeaderCell.Row, 0)
               
        SetNode node, field.GetTotal(), field.GetType(), totalCell.Value
        
Continue:
    Next
    
    WriteItemTotals = True
Finally:
    
    Exit Function
    
Catch:
    message = Err.Description
    Resume Finally
End Function

Function LoadMappingRules(document As Workbook, rules As CustomXMLNode, message As String) As Boolean
    On Error GoTo Catch
    
    Dim config As CustomXMLPart
    Set config = document.CustomXMLParts.Add
    
    If Not LoadConfiguration(config) _
    Then Err.Raise 1000, Description:="Не удалось загрузить конфигурацию"
    
    Set rules = config.SelectSingleNode("//source[@version=""" & GetVersion(document) & """ and @target-namespace=""urn:bitrix:deal""]") ' changed from 6
    
    If rules Is Nothing _
    Then Err.Raise 1000, Description:="Для данной версии структуры отсуствует описание сопоставления полей"
    
    LoadMappingRules = True
Finally:
    
    If Not config Is Nothing Then config.Delete
    
    Exit Function

Catch:
    message = Err.Description
    Resume Finally

End Function

Function WriteDeal(node As CustomXMLNode, mapping As Collection, message As String) As Boolean
    
    WriteDeal = WriteRow(1, node, mapping)
    
End Function

Function WriteItems(node As CustomXMLNode, mapping As Collection, message As String) As Boolean
    
    Dim parent As CustomXMLNode
    node.AppendChildNode "ITEMS", node.NamespaceURI
    Set parent = node.LastChild
    
    Dim rightBound As Range
    Set rightBound = mapping("0").first.End(xlToRight)

    Dim i As Integer
    Dim itemNode As CustomXMLNode
    For i = 1 To rightBound.Column
        
        parent.AppendChildNode "ITEM", parent.NamespaceURI
        Set itemNode = parent.LastChild
        
        If Not WriteRow(i, itemNode, mapping) _
        Then itemNode.Delete
        
    Next
    
    WriteItems = True
End Function

Function WriteRow(index As Integer, node As CustomXMLNode, mapping As Collection) As Boolean
    
    Dim Pair As Pair
    For Each Pair In mapping
        
        Dim cell As Range
        Set cell = Pair.first.Offset(0, index)
        
        Dim field As CField
        Set field = Pair.second
        
        If field.IsDicriminator Then _
            If IsEmpty(cell.Value) Then Exit Function
    
        SetNodeFromField node, field, cell.Value
    
    Next
    
    WriteRow = True
End Function


Function LoadConfiguration(config As CustomXMLPart) As Boolean
    LoadConfiguration = config.Load(AddIns("crm").Path & "\crm-config.xml")
End Function

Function LoadChemistry(chemisrty As Collection, chemistrySheet As Worksheet, message As String) As Boolean
    
    Dim lookupCell As Range
    Dim keyCell As Range
    Dim valueCell As Range
    
    For Each lookupCell In chemistrySheet.UsedRange.Columns(1).Cells
        
        If LCase(lookupCell.Value) = "grade" _
        Then Set keyCell = lookupCell.Offset(0, 1)
        
        If InStr(1, lookupCell.Value, "Material configuration", vbTextCompare) > 0 _
        Then Set valueCell = lookupCell.Offset(0, 1)

    Next
    
    If keyCell Is Nothing Then
        message = "Не найдено ключевое поле при разборе химического состава"
        Exit Function
    End If
    
    If valueCell Is Nothing Then
        message = "Не найдено поле значения при разборе химического состава"
        Exit Function
    End If
    
   
    Do While keyCell.Value <> ""
      
        chemisrty.Add valueCell.Value.keyCell.Value
        
        Set keyCell = keyCell.Offset(0, 1)
        Set valueCell = valueCell.Offset(0, 1)
    Loop
    
    LoadChemistry = True
    
End Function
Function CreateDeal(dealPart As CustomXMLPart) As CustomXMLNode
    Dim node As CustomXMLNode
    Set node = dealPart.SelectSingleNode("/ns0:DEAL")
    
    If Not node Is Nothing Then
        Dim child As CustomXMLNode
            For Each child In node.ChildNodes
                node.RemoveChild child
            Next
    Else
        dealPart.DocumentElement.AppendChildSubtree "<DEAL xmlns=""urn:bitrix:deal"" version=""1.0""/>"
    End If
    
    Set CreateDeal = dealPart.SelectSingleNode("/ns0:DEAL")
End Function

Function CreatePart(wk As Workbook) As CustomXMLPart
    Set CreatePart = wk.CustomXMLParts.Add("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><DEAL xmlns=""urn:bitrix:deal"" version=""1.0""/>")
End Function

Function ApplyMappingRules(area As Range, rules As CustomXMLNode, mapping As Collection, message As String) As Boolean
    Dim fieldGroups As CustomXMLNodes
    Set fieldGroups = rules.SelectNodes("field-group")
    Set mapping = New Collection
    
    Dim fieldGroup As CustomXMLNode
    For Each fieldGroup In fieldGroups
    
        Dim cell As Range
        
        If Not FindFieldGroup(area, fieldGroup, cell) Then
            message = "Не найдена группа полей " & fieldGroup.Attributes(1).text
            Exit Function
        End If
        
        If Not MapFieldGroup(cell, fieldGroup, mapping, message) Then Exit Function

    Next
    
    ApplyMappingRules = True
End Function

Function MapFieldGroup(ByVal source As Range, fieldGroup As CustomXMLNode, mapping As Collection, message As String) As Boolean
    
    Dim fieldNode As CustomXMLNode

    For Each fieldNode In fieldGroup.SelectNodes("field")
        
        Dim currentField As CField
        
        Set currentField = CField.CreateFromXml(fieldNode)
                
        Dim cell As Range
        Set cell = source.Offset(0, 0)
        
        Do While Len(cell.text) <> 0
        
            If currentField.IsRelevant(cell.text) Then
                Call mapping.Add(Pair.Create(cell.Offset(0, 0), currentField), CStr(mapping.Count))
                Exit Do
            End If
            
            Set cell = cell.Offset(1, 0)
        Loop
        
    Next
    
    MapFieldGroup = True
    
End Function

Function FindFieldGroup(area As Range, fieldGroup As CustomXMLNode, cell As Range) As Boolean
    
    Dim filter, a As String
    filter = fieldGroup.Attributes(2).text
        
    For Each cell In area.Columns(1).Cells
   
        If UCase(cell.text) Like filter Then
            FindFieldGroup = True
            Exit Function
        End If
    
    Next
    
End Function

Sub SetNode(parent As CustomXMLNode, name As String, valueType As String, Value As Variant)
    
    Dim valueNode As CustomXMLNode
    parent.AppendChildNode name, parent.NamespaceURI
    Set valueNode = parent.LastChild
    
    Select Case valueType
    Case "date":
        valueNode.text = Format(Value, "yyyy-mm-dd")
    Case "datetime":
        valueNode.text = Format(Value, "yyyy-mm-dd hh:nn:ss")
    Case "float":
    ' Change Fixed on General Number for increase number of digits after point.

        valueNode.text = Replace(Format(TryParseDouble(Value), "General Number"), Application.DecimalSeparator, ".")
     Case "integer":
        valueNode.text = TryParseLong(Value)
     Case Else:
        valueNode.text = Replace(Trim(Value), vbLf, " ") '&#10;
    End Select
     
End Sub

Sub SetNodeFromField(parent As CustomXMLNode, field As CField, Value As Variant)
    
    Dim value_ As Variant
    value_ = ""
    
    If field.IsPartial Then
    
        Dim parts As Variant
        parts = Split(Value, vbLf)
            
        Dim partIndex As Integer
        partIndex = field.GetPart() - 1
        
        If UBound(parts) >= partIndex Then value_ = parts(partIndex)
    
    Else
        
        value_ = Value
    
    End If
    
    SetNode parent, field.GetTarget, field.GetType, value_

End Sub
Function TryParseDouble(Value As Variant) As Double

    On Error Resume Next
    TryParseDouble = CDbl(Value)

End Function

Function TryParseLong(Value As Variant) As Long

    On Error Resume Next
    TryParseLong = CLng(Value)

End Function

Function CheckPrerequisite(sorf As Workbook, message As String) As Boolean
    On Error Resume Next
    
    If sorf.Worksheets("SORF") Is Nothing Then
        message = "Отсуствует страница SORF"
        Exit Function
    End If
    
'    If sorf.Worksheets("Chemistry") Is Nothing Then
'        message = "Отсуствует страница Chemistry"
'        Exit Function
'    End If
    
    If sorf.Worksheets("Items Breakdown") Is Nothing Then
        message = "Отсуствует страница Items Breakdown"
        Exit Function
    End If
    
    Dim fileVersion As String
    fileVersion = sorf.Worksheets("SORF").Range("A1").Value
    
'    If fileVersion <> "ver.5" Then
'        message = "Версия файла " & fileVersion & " не поддерживается"
'        Exit Function
'    End If
    
    CheckPrerequisite = True

End Function


