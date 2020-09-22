VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXMLNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mobjParent As clsXMLNode
Private mobjChildren() As clsXMLNode
Private mlngChildCount As Long
Private mstrPath As String
Private mstrName As String

Private Sub Class_Terminate()
    On Error Resume Next
    Dim i As Long
    
    For i = 1 To mlngChildCount
        Set mobjChildren(i) = Nothing
    Next i
    Erase mobjChildren
    
    Set mobjParent = Nothing
End Sub

Public Property Get MyPath() As String
    MyPath = mstrPath
End Property

Public Property Get MyName() As String
    MyName = mstrName
End Property

Public Property Get EndNode() As Boolean
    EndNode = (mlngChildCount = 0)
End Property

Public Property Get RootNode() As Boolean
    RootNode = (StrComp(mstrPath, mstrName, vbTextCompare) = 0)
End Property

Public Property Get MyParent() As clsXMLNode
    Set MyParent = mobjParent
End Property

Public Sub Initialize(ParentNode As clsXMLNode, Path As String, Name As String)
    On Error GoTo Initialize_Err
    If ParentNode Is Nothing Then Err.Raise 1, "[clsXMLNode][Initialize]", "Invalid Parent Node Recieved!"
    
    Set mobjParent = Nothing
    Set mobjParent = ParentNode
    mstrPath = Path
    mstrName = Name
    Exit Sub
    
Initialize_Err:
    Err.Raise Err.Number, Err.Source, "[clsXMLNode][Initialize] " & Err.Description
End Sub

Public Function AddChild(Name As String) As clsXMLNode
    On Error GoTo AddChild_Err
    Dim objNewChild As clsXMLNode
    
    If Len(Name) = 0 Then Err.Raise 1, "[clsXMLNode][AddChild]", "Invalid Name Encountered!"
    
    mlngChildCount = mlngChildCount + 1
    ReDim Preserve mobjChildren(1 To mlngChildCount)
    
    Set objNewChild = New clsXMLNode
    
    Call objNewChild.Initialize(Me, mstrPath & "\" & Name, Name)
    
    Set mobjChildren(mlngChildCount) = objNewChild
    Set AddChild = objNewChild
    Set objNewChild = Nothing
    Exit Function
    
AddChild_Err:
    Set objNewChild = Nothing
    Err.Raise Err.Number, Err.Source, "[clsXMLNode][AddChild] " & Err.Description
End Function

Public Function GetChild(Name As String) As clsXMLNode
    On Error GoTo GetChild_Err
    Dim i As Long
    
    If Len(Name) = 0 Then Exit Function
    
    For i = 1 To mlngChildCount
        If Not mobjChildren(i) Is Nothing Then
            If StrComp(mobjChildren(i).MyName, Name, vbTextCompare) = 0 Then
                Set GetChild = mobjChildren(i)
                Exit For
            End If
        End If
    Next i
    Exit Function
    
GetChild_Err:
    Err.Raise Err.Number, Err.Source, "[clsXMLNode][GetChild] " & Err.Description
End Function

