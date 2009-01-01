VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmAlert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alert"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "frmAlert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2115
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4665
      ExtentX         =   8229
      ExtentY         =   3731
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error GoTo Err_Form_Load
    
    WebBrowser1.Navigate2 GetPath("ui_alert.htm")
    
     Do While WebBrowser1.ReadyState <> READYSTATE_COMPLETE
        DoEvents
    Loop
    
    If Options(CO_PLAYSOUNDONALERT) = 1 Then
    
        Call PlayWav(Options(CO_WAVFILENAME))
    
    End If
    
Exit_Form_Load:
    Exit Sub
    
Err_Form_Load:
    Call ErrHandler("frmAlert.Form_Load", Err.Description)
    Resume Exit_Form_Load
    
End Sub

Public Property Let AlertText(ByVal strText As String)
    WebBrowser1.Document.All.spnAlertText.innerText = strText
End Property

Private Sub WebBrowser1_BeforeNavigate2( _
        ByVal pDisp As Object, _
        URL As Variant, _
        Flags As Variant, _
        TargetFrameName As Variant, _
        PostData As Variant, _
        Headers As Variant, _
        Cancel As Boolean)
    
    On Error GoTo Err_WebBrowser1_BeforeNavigate2
     
    Select Case URL

        Case GetPath("ui_alert.htm")
            'loading the user interface
            'let it occur

        Case "cp://closealert/"
            Unload Me
            
        Case Else
            Cancel = True
            pDisp.Stop
            
    End Select

Exit_WebBrowser1_BeforeNavigate2:
    Exit Sub
    
Err_WebBrowser1_BeforeNavigate2:
    Call ErrHandler("frmAlert.WebBrowser1_BeforeNavigate2", Err.Description)
    Resume Exit_WebBrowser1_BeforeNavigate2

End Sub
