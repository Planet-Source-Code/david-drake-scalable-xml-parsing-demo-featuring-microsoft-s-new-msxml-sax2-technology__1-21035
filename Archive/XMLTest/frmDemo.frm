VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAX2 XML Parsing Demo"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboSize 
      Height          =   315
      ItemData        =   "frmDemo.frx":0000
      Left            =   120
      List            =   "frmDemo.frx":0010
      TabIndex        =   0
      Top             =   120
      Width           =   2475
   End
   Begin VB.CommandButton cmdSax 
      Caption         =   "Parse XML using SAX2"
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   2475
   End
   Begin VB.CommandButton CmdDOM 
      Caption         =   "Parse XML using the DOM"
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   1140
      Width           =   2475
   End
   Begin VB.CommandButton cmdVB 
      Caption         =   "Parse XML using VB"
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2475
   End
   Begin VB.Frame Frame1 
      Caption         =   "Testing Results:"
      Height          =   2595
      Left            =   2760
      TabIndex        =   5
      Top             =   60
      Width           =   3915
      Begin VB.Label lblSax2 
         Caption         =   "Parsing File using SAX2:"
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   1740
         Width           =   3735
      End
      Begin VB.Label lblDOM 
         Caption         =   "Parsing File using DOM:"
         Height          =   555
         Left            =   120
         TabIndex        =   7
         Top             =   1020
         Width           =   3735
      End
      Begin VB.Label lblVB 
         Caption         =   "Parsing File using VB:"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   2220
      Width           =   2475
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'FOR SETUP INSTRUCTIONS, PLEASE REVIEW THE INSTRUCTIONS.TXT FILE INCLUDED WITH THIS PROJECT
'COPYRIGHT DAVID DRAKE 2001

Private mlRecordCount As Long

Private Sub cmdAbout_Click()
    MsgBox "This demo will provide performance testing statistics for parsing" & vbLf & "large XML recordsets using a variety of methods. Most importantly," & vbLf & "this demo will prove the scalability benefits of using Microsoft SAX2." & vbLf & vbLf & "Results should indicate that SAX2 will grow linearly with load, while" & vbLf & "VB and DOM parsing grow exponentially. Therefore SAX2 is" & vbLf & " the scalable solution.", vbOKOnly + vbInformation, "About this demo..."
End Sub

Private Sub cmdVb_Click()
    Dim StopWatch As New StopWatch
    Dim XML As String
    
    Screen.MousePointer = vbHourglass
    StopWatch.Start
    
    Select Case cboSize.ListIndex
        Case 0
            Call GetStringValue(App.Path & "\SmallData.xml")
        Case 1
            Call GetStringValue(App.Path & "\MediumData.xml")
        Case 2
            Call GetStringValue(App.Path & "\LargeData.xml")
        Case 3
            Call GetStringValue(App.Path & "\VeryLargeData.xml")
    End Select
    
    StopWatch.Finish
    Screen.MousePointer = vbDefault
    
    lblVB.Caption = "Parsing File using VB: " & vbLf & "     " & CStr(mlRecordCount) & " Records in " & StopWatch.ElaspsedTime & " seconds"
    Set StopWatch = Nothing
End Sub

Private Sub CmdDOM_Click()
    Dim StopWatch As New StopWatch
    
    Screen.MousePointer = vbHourglass
    StopWatch.Start
    
    Select Case cboSize.ListIndex
        Case 0
            Call GetDOMValue(App.Path & "\SmallData.xml")
        Case 1
            Call GetDOMValue(App.Path & "\MediumData.xml")
        Case 2
            Call GetDOMValue(App.Path & "\LargeData.xml")
        Case 3
            Call GetDOMValue(App.Path & "\VeryLargeData.xml")
    End Select
    
    StopWatch.Finish
    Screen.MousePointer = vbDefault

    lblDOM.Caption = "Parsing File using DOM: " & vbLf & "     " & CStr(mlRecordCount) & " Records in " & StopWatch.ElaspsedTime & " seconds"
    Set StopWatch = Nothing
End Sub

Private Sub cmdSax_Click()
    Dim StopWatch As New StopWatch
    Dim MSXML As New clsSAXTest
    
    Screen.MousePointer = vbHourglass
    StopWatch.Start
    
    Select Case cboSize.ListIndex
        Case 0
            Call MSXML.GetXML(App.Path & "\SmallData.xml")
        Case 1
            Call MSXML.GetXML(App.Path & "\MediumData.xml")
        Case 2
            Call MSXML.GetXML(App.Path & "\LargeData.xml")
        Case 3
            Call MSXML.GetXML(App.Path & "\VeryLargeData.xml")
    End Select
    
    StopWatch.Finish
    Screen.MousePointer = vbDefault
    
    lblSax2.Caption = "Parsing File using SAX2: " & vbLf & "     " & CStr(MSXML.mlRecordCount) & " Records in " & StopWatch.ElaspsedTime & " seconds"
    
    Set MSXML = Nothing
    Set StopWatch = Nothing
End Sub

Private Function parseXMLVal(ByVal strXML As String, ByVal strNode As String) As String
    On Error GoTo ErrHandler
    
    Dim lngStartPos As Long
    Dim lngEndPos As Long
    
    lngStartPos = InStr(1, strXML, "<" & strNode & ">", vbTextCompare)
    lngStartPos = lngStartPos + Len(strNode) + 2
    lngEndPos = InStr(1, strXML, "</" & strNode & ">", vbTextCompare)
    If lngStartPos > 0 And lngStartPos < lngEndPos Then
        parseXMLVal = Mid$(strXML, lngStartPos, lngEndPos - lngStartPos)
    Else
        parseXMLVal = vbNullString
    End If
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[parseXMLVal]" & Err.Description
End Function

Private Sub GetStringValue(sFile As String)
    Dim lngStartPos As Long
    Dim lngEndPos As Long
    Dim strXML As String
    Dim strXMLNode As String
    Dim objFile As TextStream
    Dim objFS As New FileSystemObject
    Dim sFirstName As String
    Dim sAddress As String
    Dim sID As String
    mlRecordCount = 0
    
    Set objFile = objFS.OpenTextFile(sFile, ForReading, False)
    
    strXML = objFile.ReadAll
    
    'Parse Root Value
    strXML = parseXMLVal(strXML, "XMLROOT")
    
    'Parse XML tags
    lngEndPos = 1
    Do
        lngStartPos = InStr(lngEndPos, strXML, "<XML>", vbTextCompare)
        If lngStartPos <= 0 Then GoTo Finished
        
        lngStartPos = lngStartPos + 5
        lngEndPos = InStr(lngStartPos, strXML, "</XML>", vbTextCompare)
        If lngStartPos > 0 And lngStartPos < lngEndPos Then
            strXMLNode = Mid$(strXML, lngStartPos, lngEndPos - lngStartPos)
            mlRecordCount = mlRecordCount + 1
            
            'Parse record elements as in other tests
            sID = parseXMLVal(strXMLNode, "ID")
            sFirstName = parseXMLVal(strXMLNode, "FirstName")
            sAddress = parseXMLVal(strXMLNode, "Address")
            lngEndPos = lngEndPos + 6
        Else
            GoTo Finished
        End If
    Loop
Finished:
    Set objFile = Nothing
    Set objFS = Nothing
End Sub

Private Sub GetDOMValue(sFile As String)
    Dim objDOM As New DOMDocument
    Dim objNode As IXMLDOMNode
    Dim objNodes As IXMLDOMNodeList
    Dim sID As String
    Dim sFirstName As String
    Dim sAddress As String
    
    mlRecordCount = 0
    Call objDOM.Load(sFile)
    Set objNodes = objDOM.selectNodes("//xmlroot/xml")
    
    For Each objNode In objNodes
        mlRecordCount = mlRecordCount + 1
        sID = objNode.selectSingleNode("id").Text
        sFirstName = objNode.selectSingleNode("firstname").Text
        sAddress = objNode.selectSingleNode("address").XML
    Next objNode
    
    Set objNode = Nothing
    Set objDOM = Nothing
End Sub

Private Sub Form_Load()
    cboSize.ListIndex = 0
End Sub
