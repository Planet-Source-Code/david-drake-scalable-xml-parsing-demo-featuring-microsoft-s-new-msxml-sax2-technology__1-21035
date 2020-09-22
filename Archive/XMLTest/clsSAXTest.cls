VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSAXTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents GetMyXML As clsXMLFileParser
Attribute GetMyXML.VB_VarHelpID = -1
Private msName As String
Private msID As String
Private msAddress As String
Public mlRecordCount As Long

Public Sub GetXML(strFile As String)
    Dim sPaths(1 To 3) As String
    mlRecordCount = 0
    
    Set GetMyXML = New clsXMLFileParser
    sPaths(1) = "XMLROOT\XML\ID"
    sPaths(2) = "XMLROOT\XML\FIRSTNAME"
    sPaths(3) = "XMLROOT\XML\ADDRESS"
    
    GetMyXML.ParseXMLFile strFile, sPaths
    
    Set GetMyXML = Nothing
End Sub

Private Sub GetMyXML_XMLNode(XMLPath As String, XMLContent As String)
    Select Case XMLPath
        Case "XMLROOT\XML\ID"
            mlRecordCount = mlRecordCount + 1
            msID = XMLContent
        Case "XMLROOT\XML\FIRSTNAME"
            msName = XMLContent
        Case "XMLROOT\XML\ADDRESS"
            msAddress = XMLContent
    End Select
End Sub
