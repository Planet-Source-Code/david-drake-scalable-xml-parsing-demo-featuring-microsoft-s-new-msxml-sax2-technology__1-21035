VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXMLFileParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event XMLNode(XMLPath As String, XMLContent As String)
Implements IVBSAXContentHandler
Implements IVBSAXErrorHandler

Private mbInNode As Boolean
Private mbWaitingForNode As Boolean
Private msCurrentNode As String
Private msCurrentPath As String
Private msXMLContent As String

Private mobjXMLPathRoot As clsXMLNode
Private mobjXMLPathCurrent As clsXMLNode
Private mobjXMLPathTemp As clsXMLNode

Public Function ParseXMLFile(FilePath As String, XMLPaths() As String) As Boolean
    On Error GoTo ParseXMLFile_Err
    Dim objSAX As SAXXMLReader30
    Dim objXMLNode As clsXMLNode
    Dim lXMLPathStart As Long
    Dim lXMLPathIndex As Long
    Dim sXMLNode As String
    Dim lIndex As Long
    Dim lMaxPaths As Long
    Dim lMinPaths As Long
    Dim sXMLPath As String
    
    'Validate Parameters
    If Not FileExists(FilePath) Then Err.Raise 1, "[clsXMLFileParser][ParseXMLFile]", "File '" & FilePath & "' does not exists."
    If Not IsArray(XMLPaths) Then Err.Raise 1, "[clsXMLFileParser][ParseXMLFile]", "Invalid XML Path Array Encountered."
    
    lMinPaths = LBound(XMLPaths)
    lMaxPaths = UBound(XMLPaths)
    If lMinPaths < 0 Or lMaxPaths < 0 Then Err.Raise 1, "[clsXMLFileParser][ParseXMLFile]", "Invalid XML Path Array Encountered."
    
    For lIndex = lMinPaths To lMaxPaths
        sXMLPath = XMLPaths(lIndex)
        
        'Validate path
        If Len(Trim$(sXMLPath)) = 0 Then Err.Raise 1, "[clsXMLFileParser][ParseXMLFile]", "Invalid XML Path."
        If Left$(sXMLPath, 1) = "\" Then Err.Raise 1, "[clsXMLFileParser][ParseXMLFile]", "Invalid XML Path Encountered."
        If Right$(sXMLPath, 1) = "\" Then Err.Raise 1, "[clsXMLFileParser][ParseXMLFile]", "Invalid XML Path Encountered."
        If InStr(1, sXMLPath, "\\") > 0 Then Err.Raise 1, "[clsXMLFileParser][ParseXMLFile]", "Invalid XML Path Encountered."
        
        'Initialize XML Parsing Linked List
        lXMLPathStart = 1
            
        'Get Root Node
        lXMLPathIndex = InStr(lXMLPathStart, sXMLPath, "\", vbTextCompare)
        If lXMLPathIndex = 0 Then lXMLPathIndex = Len(sXMLPath) + 1
        sXMLNode = Mid$(sXMLPath, lXMLPathStart, lXMLPathIndex - lXMLPathStart)
        If Len(Trim$(sXMLNode)) = 0 Then Err.Raise 1, "[clsXMLFileParser][ParseXMLFile]", "Invalid XML Path Encountered."
            
        'On first Path, handle Root node specially
        If lIndex = lMinPaths Then
            'Initialize Root
            Set mobjXMLPathRoot = Nothing
            Set mobjXMLPathRoot = New clsXMLNode
            Call mobjXMLPathRoot.Initialize(mobjXMLPathRoot, sXMLNode, sXMLNode)
        
        'Validate that all paths have the same root
        ElseIf StrComp(mobjXMLPathRoot.MyName, sXMLNode, vbTextCompare) <> 0 Then
            Err.Raise 1, "[clsXMLFileParser][ParseXMLFile]", "XML Paths Contain different Roots."
        End If
        
        Set mobjXMLPathCurrent = Nothing
        Set mobjXMLPathCurrent = mobjXMLPathRoot
        
        'Build Children Nodes
        lXMLPathStart = lXMLPathIndex + 1
        lXMLPathIndex = InStr(lXMLPathStart, sXMLPath, "\", vbTextCompare)
        If lXMLPathIndex = 0 Then
            If lXMLPathStart < Len(sXMLPath) Then
                lXMLPathIndex = Len(sXMLPath) + 1
                sXMLNode = Mid$(sXMLPath, lXMLPathStart, lXMLPathIndex - lXMLPathStart)
            Else
                sXMLNode = ""
            End If
        Else
            sXMLNode = Mid$(sXMLPath, lXMLPathStart, lXMLPathIndex - lXMLPathStart)
        End If
        
        While Len(sXMLNode) > 0
            If Len(Trim$(sXMLNode)) = 0 Then Err.Raise 1, "[clsXMLFileParser][ParseXMLFile]", "Invalid XML Path Encountered."
            Set objXMLNode = mobjXMLPathCurrent.GetChild(sXMLNode)
            If objXMLNode Is Nothing Then _
                Set objXMLNode = mobjXMLPathCurrent.AddChild(sXMLNode)
            Set mobjXMLPathCurrent = Nothing
            Set mobjXMLPathCurrent = objXMLNode
            Set objXMLNode = Nothing
        
            'Get Next Node
            lXMLPathStart = lXMLPathIndex + 1
            lXMLPathIndex = InStr(lXMLPathStart, sXMLPath, "\", vbTextCompare)
            If lXMLPathIndex = 0 Then
                If lXMLPathStart < Len(sXMLPath) Then
                    lXMLPathIndex = Len(sXMLPath) + 1
                    sXMLNode = Mid$(sXMLPath, lXMLPathStart, lXMLPathIndex - lXMLPathStart)
                Else
                    sXMLNode = ""
                End If
            Else
                sXMLNode = Mid$(sXMLPath, lXMLPathStart, lXMLPathIndex - lXMLPathStart)
            End If
        Wend
        
        Set mobjXMLPathCurrent = Nothing
    Next lIndex
    
    Set mobjXMLPathCurrent = mobjXMLPathRoot
    
    Set objSAX = New SAXXMLReader30
    Set objSAX.contentHandler = Me
    Set objSAX.errorHandler = Me
    
    Set objXMLNode = Nothing
    
    'Start XML Parsing
    On Error Resume Next
    objSAX.parseURL FilePath
    
    On Error GoTo ParseXMLFile_Err
    If Err Then Err.Raise Err.Number, "[clsXMLFileParser][ParseXMLFile]", "Malformed XML File Encountered! " & FilePath
    
ParseXMLFile_Exit:
    Set objXMLNode = Nothing
    Set objSAX = Nothing
    Exit Function
    
ParseXMLFile_Err:
    Set objXMLNode = Nothing
    Set objSAX = Nothing
    Err.Raise Err.Number, Err.Source, "[clsXMLFileParser][ParseXMLFile]" & Err.Description
End Function

Private Sub Class_Terminate()
    Set mobjXMLPathRoot = Nothing
    Set mobjXMLPathCurrent = Nothing
End Sub

Private Function FileExists(FileName As String) As Boolean
    On Error GoTo FileExistsExit
    If Len(Trim$(FileName)) > 0 And Len(Dir$(FileName)) > 0 Then FileExists = True
FileExistsExit:
End Function

Private Sub IVBSAXContentHandler_characters(strChars As String)
    On Error GoTo ISAXContentHandler_Characters_Err
    
    'Validate Parsing state
    If Not mbInNode Or mbWaitingForNode Then Exit Sub
    
    If StrComp(strChars, vbCrLf, vbTextCompare) = 0 Then Exit Sub
    msXMLContent = msXMLContent & strChars
    Exit Sub

ISAXContentHandler_Characters_Err:
    Err.Raise Err.Number, Err.Source, "[clsXMLFileParser][IVBSAXContentHandler_characters] " & Err.Description
End Sub

Private Sub IVBSAXContentHandler_endElement(strNamespaceURI As String, strLocalName As String, strQName As String)
    On Error GoTo ISAXContentHandler_EndElement_Err
    
    'Validate Parsing state
    If Not mbInNode And Not mbWaitingForNode Then
        'Move to previous node if they match
        If Not mobjXMLPathCurrent.RootNode Then
            If StrComp(strLocalName, mobjXMLPathCurrent.MyName, vbTextCompare) = 0 Then
                Set mobjXMLPathTemp = mobjXMLPathCurrent.MyParent
                Set mobjXMLPathCurrent = Nothing
                Set mobjXMLPathCurrent = mobjXMLPathTemp
                Set mobjXMLPathTemp = Nothing
                With mobjXMLPathCurrent
                    msCurrentNode = .MyName
                    msCurrentPath = .MyPath
                End With
            End If
        End If
        
        Exit Sub
    End If
    
    'Process XML Node
    If mbWaitingForNode Then
        If StrComp(strLocalName, msCurrentNode, vbTextCompare) = 0 Then
            mbWaitingForNode = False
        Else
            Exit Sub
        End If
    Else
        If StrComp(strLocalName, msCurrentNode, vbTextCompare) = 0 Then
            mbInNode = False
            'Raise Event to return XML Content
            RaiseEvent XMLNode(msCurrentPath, msXMLContent)
            
            'Move to previous
            If Not mobjXMLPathCurrent.RootNode Then
                Set mobjXMLPathTemp = mobjXMLPathCurrent.MyParent
                Set mobjXMLPathCurrent = Nothing
                Set mobjXMLPathCurrent = mobjXMLPathTemp
                Set mobjXMLPathTemp = Nothing
                msCurrentNode = mobjXMLPathCurrent.MyName
                msCurrentPath = mobjXMLPathCurrent.MyPath
            End If
            
            msXMLContent = ""
        Else
            msXMLContent = msXMLContent & "</" & strLocalName & ">"
        End If
    End If
    Exit Sub

ISAXContentHandler_EndElement_Err:
    Err.Raise Err.Number, Err.Source, "[clsXMLFileParser][IVBSAXContentHandler_endElement] " & Err.Description
End Sub


Private Sub IVBSAXContentHandler_endPrefixMapping(strPrefix As String)

End Sub


Private Sub IVBSAXContentHandler_ignorableWhitespace(strChars As String)

End Sub

Private Sub IVBSAXContentHandler_processingInstruction(strTarget As String, strData As String)

End Sub

Private Sub IVBSAXContentHandler_skippedEntity(strName As String)

End Sub

Private Sub IVBSAXContentHandler_startDocument()

End Sub

Private Property Set IVBSAXContentHandler_documentLocator(ByVal RHS As MSXML2.IVBSAXLocator)

End Property

Private Sub IVBSAXContentHandler_endDocument()

End Sub

Private Sub IVBSAXContentHandler_startElement(strNamespaceURI As String, strLocalName As String, strQName As String, ByVal oAttributes As MSXML2.IVBSAXAttributes)
    On Error GoTo ISAXContentHandler_StartElement_Err
    If mbWaitingForNode = True Then Exit Sub
    
    'Process XML Node
    If mbInNode Then
        msXMLContent = msXMLContent & "<" & strLocalName & ">"
    Else
        If StrComp(mobjXMLPathCurrent.MyName, strLocalName, vbTextCompare) = 0 Then
            Set mobjXMLPathTemp = mobjXMLPathCurrent
        Else
            Set mobjXMLPathTemp = mobjXMLPathCurrent.GetChild(strLocalName)
        End If
        
        'If child does not exist then wait for end of this node
        If mobjXMLPathTemp Is Nothing Then
            mbWaitingForNode = True
            msCurrentNode = strLocalName
        
        'If child does exist then determine if this is an end node
        Else
            Set mobjXMLPathCurrent = Nothing
            Set mobjXMLPathCurrent = mobjXMLPathTemp
            Set mobjXMLPathTemp = Nothing
            
            'If this is an end node then capture XML
            If mobjXMLPathCurrent.EndNode Then
                mbInNode = True
                msCurrentPath = mobjXMLPathCurrent.MyPath
                msCurrentNode = strLocalName
            End If
        End If
    End If
    Exit Sub

ISAXContentHandler_StartElement_Err:
    Err.Raise Err.Number, Err.Source, "[clsXMLFileParser][IVBSAXContentHandler_startElement] " & Err.Description
End Sub

Private Sub IVBSAXContentHandler_startPrefixMapping(strPrefix As String, strURI As String)

End Sub

Private Sub IVBSAXErrorHandler_error(ByVal oLocator As MSXML2.IVBSAXLocator, strErrorMessage As String, ByVal nErrorCode As Long)
    Err.Raise nErrorCode, "Line: " & CStr(oLocator.lineNumber) & " Column: " & CStr(oLocator.columnNumber), "[clsXMLFileParser][IVBSAXErrorHandler_error] " & strErrorMessage
End Sub

Private Sub IVBSAXErrorHandler_fatalError(ByVal oLocator As MSXML2.IVBSAXLocator, strErrorMessage As String, ByVal nErrorCode As Long)
    Err.Raise nErrorCode, "Line: " & CStr(oLocator.lineNumber) & " Column: " & CStr(oLocator.columnNumber), "[clsXMLFileParser][IVBSAXErrorHandler_error] " & strErrorMessage
End Sub

Private Sub IVBSAXErrorHandler_ignorableWarning(ByVal oLocator As MSXML2.IVBSAXLocator, strErrorMessage As String, ByVal nErrorCode As Long)

End Sub

Private Sub IVBSAXErrorHandler_warning(ByVal oLocator As MSXML2.IVBSAXLocator, ByVal strError As String, ByVal nErrorCode As Long)
    Err.Raise nErrorCode, "Line: " & CStr(oLocator.lineNumber) & " Column: " & CStr(oLocator.columnNumber), "[clsXMLFileParser][IVBSAXErrorHandler_error] " & strError
End Sub
