Attribute VB_Name = "ExportVBA"
Sub ExportAllVBA()
    Dim component As VBComponent
    Dim exportPath As String
    
    exportPath = "C:\Users\Kruchinin_NL\source\repos\Sorf\VBA_Export\" ' Укажите путь для сохранения
    If Dir(exportPath, vbDirectory) = "" Then MkDir exportPath
    
    For Each component In ThisWorkbook.VBProject.VBComponents
        If component.Type <> vbext_ct_Document Then
            component.Export exportPath & component.name & ".bas"
        Else
            ' Для листов и ThisWorkbook используем специальные имена
            component.Export exportPath & component.name & ".cls"
        End If
    Next component
End Sub
