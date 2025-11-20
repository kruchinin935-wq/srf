VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AddinConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

Private Const NAMESPACE = "urn:config"


Public Function Load() As Boolean
    Dim xmlConfig As CustomXMLNode
    
    If Not GetConfigFromCustomPart(xmlConfig) Then Err.Raise 1000, Description:="Конфигурация отсутствует. Восстановите конфигурацию из файла."
    
    xmlConfig
    
    
    
End Function

Public Property Get Property(name As String) As Variant

End Property


Public Function Restore(xml As String) As Boolean
        Dim xmlConfig As CustomXMLNode
    
        
        
        GetConfigFromCustomPart (xmlConfig)
        
        ThisWorkbook.CustomXMLParts.Add xml
        
            
End Function

Public Function Store() As Boolean
    
End Function

Private Function GetConfigFromCustomPart(xmlConfig As CustomXMLNode) As Boolean
    Set xmlConfig = ThisWorkbook.CustomXMLParts.SelectByNamespace(NAMESPACE)
    
    If xmlConfig Is Nothing Then Exit Function
    
    GetConfigFromCustomPart = True

End Function

Private Function CreateCustomPartConfig(xmlConfig As CustomXMLNode) As Boolean
    

End Function



