VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "VBAProject.Domain"
Option Explicit

Private m_filter As String
Private m_type As String
Private m_name As String
Private m_target As String
Private m_targetTotal  As String
Private m_part As Integer
Private m_discriminator As Boolean

Public Function CreateFromXml(source As CustomXMLNode) As CField

    Set CreateFromXml = New CField
    
    CreateFromXml.Init _
    GetNodeText(source, "name"), _
    GetNodeText(source, "filter"), _
    GetNodeText(source, "type"), _
    GetNodeNumber(source, "part"), _
    GetNodeBoolean(source, "discriminator"), _
    GetNodeText(source, "target"), _
    GetNodeText(source, "target-total")
    
    
End Function

Public Sub Init(name As String, filter As String, typeName As String, part As Integer, discriminator As Boolean, target As String, targetTotal As String)

    m_name = name
    m_filter = UCase(filter)
    m_type = typeName
    m_part = part
    m_discriminator = discriminator
    m_target = target
    m_targetTotal = targetTotal

End Sub
Private Function GetNodeText(node As CustomXMLNode, tag As String) As String
    
    Dim child As CustomXMLNode
    Set child = node.SelectSingleNode(tag)
    
    If child Is Nothing Then Exit Function
    
    GetNodeText = Trim(child.text)
    
End Function

Private Function GetNodeNumber(node As CustomXMLNode, tag As String) As Double
    
    Dim child As CustomXMLNode
    Set child = node.SelectSingleNode(tag)
    
    If child Is Nothing Then Exit Function
    
    On Error Resume Next
    GetNodeNumber = CDbl(child.text)
    
End Function

Private Function GetNodeBoolean(node As CustomXMLNode, tag As String) As Boolean
    
    Dim child As CustomXMLNode
    Set child = node.SelectSingleNode(tag)
    
    If child Is Nothing Then Exit Function
    
    On Error Resume Next
    GetNodeBoolean = CBool(child.text)
    
End Function
Public Function IsRelevant(text As String) As Boolean
    IsRelevant = UCase(text) Like m_filter
End Function

Public Property Get IsPartial() As Boolean
    IsPartial = m_part > 0
End Property

Public Property Get IsDicriminator() As Boolean
    IsDicriminator = m_discriminator
End Property

Public Function GetName() As String
    GetName = m_name
End Function

Public Function GetType() As String
    GetType = m_type
End Function

Public Function GetTarget() As String
    GetTarget = m_target
End Function

Public Function GetPart() As Integer
    GetPart = m_part
End Function

Public Function HasTotal() As Boolean
    HasTotal = Len(m_targetTotal) > 0
End Function

Public Function GetTotal() As String
    GetTotal = m_targetTotal
End Function
