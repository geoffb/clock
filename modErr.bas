Attribute VB_Name = "modErr"
Option Explicit

Public Function ErrHandler( _
        ByVal strFunc As String, ByVal strDesc As String, _
        Optional ByVal bNotify As Boolean = False)
    
    On Error Resume Next
    Dim intFile         As Integer
    Dim strLine         As String
    
    intFile = FreeFile
    
    Open GetPath("error.log") For Append As intFile
    
    strLine = "Timestamp: " & Now & vbCrLf _
            & "Error: " & strDesc & vbCrLf _
            & "Location: " & strFunc & vbCrLf
            
    Print #intFile, strLine
    
    Close intFile
    
    If bNotify Then
        
        MsgBox "An unexpected error ocurred." & vbCrLf & vbCrLf _
                & strLine, vbCritical, "Error"
        
    End If

End Function

