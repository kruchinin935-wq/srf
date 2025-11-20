VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("VBAProject.Domain")
Option Explicit

Public first As Variant
Public second As Variant

Public Function Create(first As Variant, second As Variant) As Pair
    
    Set Create = New Pair
    Create.Init first, second
    
 End Function
 
 Public Sub Init(first As Variant, second As Variant)
    If IsObject(first) Then
        Set Me.first = first
    Else
        Me.first = first
    End If
    
    If IsObject(second) Then
        Set Me.second = second
    Else
        Me.second = second
    End If
 End Sub
