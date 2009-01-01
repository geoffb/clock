VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clock"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog ComDialog 
      Left            =   450
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrClock 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5190
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5265
      ExtentX         =   9287
      ExtentY         =   9155
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
   Begin VB.Menu RCPopup 
      Caption         =   "RCPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuRCPRestore 
         Caption         =   "Restore"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    If Not App.PrevInstance Then
    
        Call LoadClock
        
    Else
        
        MsgBox "Another instance of this program is already running.", vbInformation, "Clock"
        
        Unload Me
    
    End If
     
End Sub

Public Sub SysTrayMouseEventHandler()
    
    SetForegroundWindow Me.hwnd
    PopupMenu RCPopup, vbPopupMenuRightButton
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    UnloadForms
    
End Sub

Private Sub mnuRCPExit_Click()
    
    Unload Me
    
    End
    
End Sub

Private Sub mnuRCPRestore_Click()
    
    Unhook    ' Return event control to windows
        
        
        
    Me.Show

    RemoveIconFromTray
    
End Sub

Private Sub tmrClock_Timer()
    
    Call UpdateClock
    
    
End Sub

Private Sub WebBrowser1_BeforeNavigate2( _
        ByVal pDisp As Object, URL As Variant, Flags As Variant, _
        TargetFrameName As Variant, PostData As Variant, _
        Headers As Variant, Cancel As Boolean)
    
    On Error GoTo Err_WebBrowser1_BeforeNavigate2
    
    'if navigate is to a help topic or email link
    'then let it occur.
    If URL Like "mailto:*" _
            Or URL Like "*help_jump_*" Then
        GoTo Exit_WebBrowser1_BeforeNavigate2
    End If
    
    Select Case URL
        Case GetPath("ui_main.htm")
            'loading the user interface
            'let it occur
            GoTo Exit_WebBrowser1_BeforeNavigate2

        Case "cp://alerts/":            ShowSection "alerts"
        Case "cp://options/":           ShowSection "options"
        Case "cp://about/":             ShowSection "about"
        Case "cp://help/":              ShowSection "help"
        Case "cp://mintosystray/":      MinToSysTray
        Case "cp://saveoptions/":       SaveOptions
        Case "cp://browse/":            DoBrowse "txtWavFile", "*.wav|*.wav"
        Case "cp://newalert/":          ShowSection "newalert"
        Case "cp://savealert/":         SaveAlert
        Case "cp://deletealert/":       DeleteAlert
    End Select
    
    Cancel = True
    pDisp.Stop
    
Exit_WebBrowser1_BeforeNavigate2:
    Exit Sub
    
Err_WebBrowser1_BeforeNavigate2:
    Call ErrHandler("frmMain.WebBrowser1_BeforeNavigate2", Err.Description)
    Resume Exit_WebBrowser1_BeforeNavigate2

End Sub

