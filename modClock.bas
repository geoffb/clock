Attribute VB_Name = "modClock"
Option Explicit

Public Declare Function sndPlaySound Lib "winmm.dll" Alias _
       "sndPlaySoundA" (ByVal lpszSoundName As String, _
       ByVal uFlags As Long) As Long

Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10

Public Enum ClockOptsEnum
    CO_TIMEFORMAT '0 = Standard, 1 = Military
    CO_STARTONWINLOAD '0 = No, 1 = Yes
    CO_PLAYSOUNDONALERT '0 = No, 1 = Yes
    CO_WAVFILENAME 'path to wav file
    CO_COUNT
End Enum

Private Type Alert
    Month As Integer '1 = Jan, 2 = Feb, etc...
    Day As Integer
    Year As Integer
    Hour As Integer
    Minute As Integer
    AmPm As String
    Type As Integer '0 = Single, 1 = Recurring
    RecurInterval As Integer '0 = Daily, 1 = Weekly, 2 = Monthly, 4 = Yearly
    Text As String
End Type

Private Const MenuHighlightColor = "yellow"

Private mstrOpts(CO_COUNT - 1)      As String 'Options
Private mlngAlerts                  As Long 'Alert Count
Private mobjAlerts()                As Alert 'All alerts
Private mbNoAlerts                  As Boolean

Public Property Get Options(ByVal lngOpt As ClockOptsEnum) As String
    
    If IsNumeric(mstrOpts(lngOpt)) Then
        
        Options = Val(mstrOpts(lngOpt))
            
    Else
        
        Options = mstrOpts(lngOpt)
    
    End If

End Property

Public Property Let Options( _
        ByVal lngOpt As ClockOptsEnum, ByVal strValue As String)
        
    mstrOpts(lngOpt) = strValue

End Property

Public Function GetPath(ByVal strPath As String) As String
    Dim strAppPath      As String
    
    strAppPath = App.Path
    
    If Right(strAppPath, 1) = "\" Then
        GetPath = strAppPath & strPath
    
    Else
        GetPath = strAppPath & "\" & strPath
    
    End If

End Function

Private Sub WaitForBrowser()
    
    Do While frmMain.WebBrowser1.ReadyState <> READYSTATE_COMPLETE
        DoEvents
    Loop
    
End Sub

Public Function LoadClock()
    On Error GoTo Err_LoadClock
    
    frmMain.WebBrowser1.Navigate2 GetPath("ui_main.htm")
    
    Call WaitForBrowser
     
    Call LoadOptions
    
    Call LoadAlerts
    
    Call UpdateClock(True)
    
    ShowSection "alerts"
    
Exit_LoadClock:
    Exit Function
    
Err_LoadClock:
    Call ErrHandler("modClock.LoadClock", Err.Description)
    Resume Exit_LoadClock

End Function

Public Sub PlayWav(ByVal strFilename As String)

    Call sndPlaySound(strFilename, SND_ASYNC Or SND_NODEFAULT)

End Sub

Private Function LoadOptions( _
        Optional ByVal bDefault As Boolean = False)

    On Error GoTo Err_LoadOptions
    Dim intFile             As Integer
    Dim strOpString         As String
    Dim strOpts()           As String
    Dim i                   As Integer
    
    'load options
    If Len(Dir(GetPath("clock.dat"))) > 0 And Not bDefault Then
        
        intFile = FreeFile
        
        Open GetPath("clock.dat") For Input As intFile
        
        Line Input #intFile, strOpString
        
        strOpts = Split(strOpString, "|")
        
        For i = 0 To UBound(strOpts)
        
            mstrOpts(i) = strOpts(i)
        
        Next
        
        Close intFile
    
    Else
    
        Options(CO_TIMEFORMAT) = 0
        Options(CO_STARTONWINLOAD) = 1
        Options(CO_PLAYSOUNDONALERT) = 1
        Options(CO_WAVFILENAME) = GetPath("matrix20.wav")
        
    End If
    
    'update the options sections
    With frmMain.WebBrowser1.Document.All
    
        .optFormat(0).Checked = (Options(CO_TIMEFORMAT) = 0)
        .optFormat(1).Checked = (Options(CO_TIMEFORMAT) = 1)
        '.chkStartOnWinLoad.Checked = (Options(CO_STARTONWINLOAD) = 1)
        .chkPlaySound.Checked = (Options(CO_PLAYSOUNDONALERT) = 1)
        .txtWavFile.Value = Options(CO_WAVFILENAME)
        
    End With
    
Exit_LoadOptions:
    Exit Function

Err_LoadOptions:
    Call ErrHandler("modClock.LoadOptions", Err.Description)
    Resume Exit_LoadOptions
    
End Function

Public Function SaveOptions()
    On Error GoTo Err_SaveOptions
    
    With frmMain.WebBrowser1.Document.All
    
        Options(CO_TIMEFORMAT) = IIf(.optFormat(1).Checked, 1, 0)
        'Options(CO_STARTONWINLOAD) = IIf(.chkStartOnWinLoad.Checked, 1, 0)
        Options(CO_PLAYSOUNDONALERT) = IIf(.chkPlaySound.Checked, 1, 0)
        Options(CO_WAVFILENAME) = .txtWavFile.Value
    
    End With
    
    Call SaveOptionsToFile
    
Exit_SaveOptions:
    Exit Function

Err_SaveOptions:
    Call ErrHandler("modClock.SaveOptions", Err.Description)
    Resume Exit_SaveOptions

End Function

Public Function SaveAlert()
    On Error GoTo Err_SaveAlert
    Dim objAlert        As Alert
    Dim intAlerts       As Integer
    
    With frmMain.WebBrowser1.Document.All.frmNewAlert
    
        objAlert.Month = Val(.cmbMonth.Value)
        objAlert.Day = Val(.cmbDay.Value)
        objAlert.Year = Val(.txtYear.Value)
        objAlert.Hour = Val(.cmbHour.Value)
        objAlert.Minute = Val(.cmbMinute.Value)
        objAlert.AmPm = .cmbAmPm.Value
        objAlert.Type = IIf(.optAlertType(1).Checked, 1, 0)
        objAlert.RecurInterval = Val(.cmbRecur.Value)
        objAlert.Text = .txtAlertText.Value
    
        'Reset new alert form elements by simulating a
        'click on a reset input element
        .Reset
    
    End With
    
    'save alert
    ReDim Preserve mobjAlerts(mlngAlerts) As Alert

    mobjAlerts(mlngAlerts) = objAlert

    mlngAlerts = mlngAlerts + 1
    
    mbNoAlerts = False
    
    Call SaveAlertsToFile
    
    Call DrawAlerts
    
    Call ShowSection("alerts")

Exit_SaveAlert:
    Exit Function

Err_SaveAlert:
    Call ErrHandler("modClock.SaveAlert", Err.Description)
    Resume Exit_SaveAlert

End Function

Public Function LoadAlerts()
    On Error GoTo Err_LoadAlerts
    Dim intFile         As Integer
    Dim strLine         As String
    Dim aAlert()        As String
    Dim strDate         As String
    
    mbNoAlerts = True
    
    If Len(Dir(GetPath("alerts.dat"))) > 0 Then
        
        intFile = FreeFile
        
        Open GetPath("alerts.dat") For Input As intFile
        
        Do While Not EOF(intFile)
        
            Line Input #intFile, strLine
            
            aAlert = Split(strLine, "|")
            
            ReDim Preserve mobjAlerts(mlngAlerts) As Alert

            strDate = Format(aAlert(0), "m/dd/yyyy hh:mm:ss ampm")
            
            With mobjAlerts(mlngAlerts)
                .Month = Month(strDate)
                .Day = Day(strDate)
                .Year = Year(strDate)
                .Hour = Hour(strDate)
                .Minute = Minute(strDate)
                .AmPm = IIf(strDate Like "*AM*", "AM", "PM")
                .Type = Val(aAlert(1))
                .RecurInterval = Val(aAlert(2))
                .Text = aAlert(3)
            End With

            mlngAlerts = mlngAlerts + 1
            
            mbNoAlerts = False
            
        Loop
        
        Close intFile
            
    Else
        ReDim mobjAlerts(0) As Alert
      
    End If
    
    Call DrawAlerts
    
Exit_LoadAlerts:
    Exit Function

Err_LoadAlerts:
    Call ErrHandler("modClock.LoadAlerts", Err.Description)
    Resume Exit_LoadAlerts

End Function

Public Function DrawAlerts()
    On Error GoTo Err_DrawAlerts
    Dim objE                As Object
    Dim objAlert            As Alert
    Dim i                   As Integer
    Dim j                   As Integer
    Dim objAlert1           As Alert
    Dim objAlert2           As Alert
    
    If Not mbNoAlerts Then
        
        'sort the array by date
        For i = 0 To UBound(mobjAlerts)
        
            For j = 0 To ((UBound(mobjAlerts) - i) - 1)
                
                objAlert1 = mobjAlerts(i)
                objAlert2 = mobjAlerts(i + 1)
                
                If DateDiff("n", GetAlertTimestamp(objAlert1), GetAlertTimestamp(objAlert2)) > 0 Then
                            
                    mobjAlerts(i) = objAlert2
                    mobjAlerts(i + 1) = objAlert1
                            
                End If
                
            Next
                
        Next
        
    End If
    
    With frmMain.WebBrowser1.Document.All
        
        'delete all the rows in the alerts table
        Do While (.tblAlerts.rows.length > 1)
            Call .tblAlerts.deleteRow(1)
        Loop
        
        If Not mbNoAlerts Then
        
            For i = 0 To UBound(mobjAlerts)
            
                'Make a copy of the hidden row template
                Set objE = .tblAlerts.rows(0).cloneNode(True)
            
                'Since this is a copy of a hidden element, it's
                'hidden as well. Need to make it visible.
                objE.Style.display = ""
        
                'Insert the copied row into the table after the hidden rowtemplate
                Call .tblAlerts.rows(0).insertAdjacentElement("afterEnd", objE)
                
                'Grab a reference to the newly inserted rows so that the
                'cell collection can be accessed
                Set objE = .tblAlerts.rows(1)
                 
                'Put Alert info into cells
                objE.cells(1).innerHtml = GetAlertTimestamp(mobjAlerts(i))
                
                If Len(mobjAlerts(i).Text) > 50 Then
                    objE.cells(2).innerText = Left(mobjAlerts(i).Text, 50) & "..."
                Else
                    objE.cells(2).innerText = mobjAlerts(i).Text
                End If
        
            Next
            
        End If
        
        .spnNoAlerts.Style.display = IIf(mbNoAlerts, "", "none")
    
    End With
    
Exit_DrawAlerts:
    Set objE = Nothing
    Exit Function

Err_DrawAlerts:
    Call ErrHandler("modClock.DrawAlerts", Err.Description)
    Resume Exit_DrawAlerts

End Function

Public Function SaveOptionsToFile()
    On Error GoTo Err_SaveOptionsToFile
    Dim strOpString         As String
    Dim intFile             As Integer
    
    strOpString = Join(mstrOpts, "|")
    
    intFile = FreeFile
    
    Open GetPath("clock.dat") For Output As intFile
    
    Print #intFile, strOpString
    
    Close intFile
    
Exit_SaveOptionsToFile:
    Exit Function

Err_SaveOptionsToFile:
    Call ErrHandler("modClock.SaveOptionsToFile", Err.Description)
    Resume Exit_SaveOptionsToFile

End Function

Public Function SaveAlertsToFile()
    On Error GoTo Err_SaveAlertsToFile
    Dim strLine             As String
    Dim intFile             As Integer
    Dim i                   As Integer
    Dim strDate             As String
    
    intFile = FreeFile
    
    Open GetPath("alerts.dat") For Output As intFile
    
    If Not mbNoAlerts Then
    
        For i = 0 To UBound(mobjAlerts)
            
            strDate = GetAlertTimestamp(mobjAlerts(i))
            
            With mobjAlerts(i)
            
                strLine = strDate & "|" _
                        & .Type & "|" _
                        & .RecurInterval & "|" _
                        & .Text
                    
            End With
            
            Print #intFile, strLine
            
        Next
        
    End If
    
    Close intFile
    
Exit_SaveAlertsToFile:
    Exit Function

Err_SaveAlertsToFile:
    Call ErrHandler("modClock.SaveAlertsToFile", Err.Description)
    Resume Exit_SaveAlertsToFile

End Function

Public Function UpdateClock( _
        Optional ByVal bSupressAlerts As Boolean = False)
        
    'Update the html elements with the date and time.
    'Check to see if any alerts need to be shown, if so
    'then display.
    
    On Error GoTo Err_UpdateClock
    Dim objE                As Object
    Dim strLongDate         As String
    Dim strTime             As String

    'update the displayed date
    Set objE = frmMain.WebBrowser1.Document.All.spnDate
    
    If Not objE Is Nothing Then
    
        strLongDate = WeekdayName(Weekday(Date)) & Chr(32) _
                & MonthName(Month(Date)) & Chr(32) _
                & DatePart("d", Date) & "," & Chr(32) _
                & DatePart("yyyy", Date)
        
        objE.innerText = strLongDate
        
    End If

    'update the displayed time
    Set objE = frmMain.WebBrowser1.Document.All.spnTime
    
    If Not objE Is Nothing Then
         
        If Val(Options(CO_TIMEFORMAT)) = 1 Then
            'Military time
            strTime = Format(Time, "HH:nn:ss")
        
        Else
            'standard time
            strTime = Time
        
        End If
        
        objE.innerText = strTime
        
        'set time to form caption so that the user can see it in the taskbar
        frmMain.Caption = Chr(32) & strTime
        
    End If
    
    'handle alerts
    If Not bSupressAlerts Then ProcessAlerts
         
Exit_UpdateClock:
    Set objE = Nothing
    Exit Function
    
Err_UpdateClock:
    Call ErrHandler("modClock.UpdateClock", Err.Description)
    Resume Exit_UpdateClock

End Function

Private Function ProcessAlerts()
    On Error GoTo Err_ProcessAlerts
    Dim frmA                    As frmAlert
    Dim i                       As Integer
    Dim lngDiff                 As Long
    Dim strDate                 As String
    Dim bRedraw                 As Boolean
    Dim bProcessingAlerts       As Boolean
    
    If mbNoAlerts Then Exit Function
    
    bProcessingAlerts = True
    
    Do While bProcessingAlerts
    
        For i = 0 To UBound(mobjAlerts)
            
            'If we are at the end of the array then we don't
            'need to continue looping
            If i = UBound(mobjAlerts) Then bProcessingAlerts = False
            
            'find the difference in minutes between the alert and now
            lngDiff = DateDiff("n", Now, GetAlertTimestamp(mobjAlerts(i)))
            
            If lngDiff <= 0 And lngDiff > -5 Then
                'if the difference in minutes is between nothing and
                '5 minutes in the past, then display the alert
            
                Set frmA = New frmAlert
                
                Call frmA.Show
                
                frmA.AlertText = GetAlertTimestamp(mobjAlerts(i)) & _
                        vbCrLf & vbCrLf & mobjAlerts(i).Text
                
                If mobjAlerts(i).Type = 0 Then
                    'it's a single alert; delete it
                    Call RemoveAlert(i)
                    bRedraw = True
                    'exit for loop so that it will begin again
                    'now that the array count has changed
                    Exit For
                    
                Else
                    'it's a recurring alert; update its timestamp
                    strDate = GetAlertTimestamp(mobjAlerts(i))
                    
                    Select Case mobjAlerts(i).RecurInterval
                
                        Case 0 'Daily
                            strDate = DateAdd("d", 1, strDate)
                        
                        Case 1 'Weekly
                            strDate = DateAdd("ww", 1, strDate)
                            
                        Case 2 'Monthly
                            strDate = DateAdd("m", 1, strDate)
                        
                        Case 3 'Yearly
                            strDate = DateAdd("yyyy", 1, strDate)
                    
                    End Select
                    
                    With mobjAlerts(i)
                        .Month = Month(strDate)
                        .Day = Day(strDate)
                        .Year = Year(strDate)
                    End With
                
                End If
                
                bRedraw = True
                
            End If
            
        Next
        
    Loop
    
    If bRedraw Then
        'If the number of properties of alerts has changed
        'then save the current alert list and redraw the
        'alerts section
            
        Call SaveAlertsToFile
        
        Call DrawAlerts
        
    End If
    
Exit_ProcessAlerts:
    Exit Function
    
Err_ProcessAlerts:
    Call ErrHandler("modClock.ProcessAlerts", Err.Description)
    Resume Exit_ProcessAlerts
    
End Function

Public Function ShowSection( _
        ByVal strSection As String)
        
    On Error GoTo Err_ShowSection

    With frmMain.WebBrowser1.Document.All
        'set toolbar highlights
        .aAlerts.Style.Color = IIf(strSection = "alerts" Or strSection = "newalert", MenuHighlightColor, "")
        .aOpts.Style.Color = IIf(strSection = "options", MenuHighlightColor, "")
        .aAbout.Style.Color = IIf(strSection = "about", MenuHighlightColor, "")
        '.aHelp.Style.Color = IIf(strSection = "help", MenuHighlightColor, "")
        'show/hide sections
        .divAlerts.Style.display = IIf(strSection = "alerts", "", "none")
        .divOpts.Style.display = IIf(strSection = "options", "", "none")
        .divAbout.Style.display = IIf(strSection = "about", "", "none")
        '.divHelp.Style.display = IIf(strSection = "help", "", "none")
        .divEditAlert.Style.display = IIf(strSection = "newalert", "", "none")
    End With
    
Exit_ShowSection:
    Exit Function

Err_ShowSection:
    Call ErrHandler("modClock.ShowSection", Err.Description)
    Resume Exit_ShowSection

End Function

Public Sub MinToSysTray()
    On Error GoTo Err_MinToSysTray
    
    'minimizes app to the system tray
    
    With frmMain
    
        Hook .hwnd ' Set up our handler
        
        Call AddIconToTray( _
                .hwnd, .Icon, .Icon.Handle, "Clock")
        
        .Hide
    
    End With

Exit_MinToSysTray:
    Exit Sub

Err_MinToSysTray:
    Call ErrHandler("modClock.MinToSysTray", Err.Description)
    Resume Exit_MinToSysTray

End Sub

Public Function DoBrowse( _
        ByVal strTargetTextbox As String, _
        Optional ByVal strFilter As String = "*.*|*.*")
    
    On Error GoTo Err_DoBrowse
    Dim objE        As Object
    
    With frmMain.ComDialog
        
        .Filter = strFilter
        
        .ShowOpen
        
        Select Case strTargetTextbox
        
            Case "txtWavFile"
                Set objE = frmMain.WebBrowser1.Document.All.txtWavFile
                
        End Select
    
        objE.Value = .FileName
        
    End With

Exit_DoBrowse:
    Set objE = Nothing
    Exit Function

Err_DoBrowse:
    Call ErrHandler("modClock.DoBrowse", Err.Description)
    Resume Exit_DoBrowse

End Function

Private Function GetAlertTimestamp(ByRef objAlert As Alert) As String
    
    With objAlert
    
        GetAlertTimestamp = Format(.Month & "/" _
                & .Day & "/" _
                & .Year & Chr(32) _
                & .Hour & ":" _
                & .Minute & ":00" & Chr(32) _
                & .AmPm, _
                "m/dd/yyyy " & _
                IIf(Options(CO_TIMEFORMAT) = 1, "HH:mm:ss", "h:mm:ss ampm"))

    End With

End Function

Private Function RemoveAlert(ByVal intIndex As Integer)
    On Error GoTo Err_RemoveAlert
    Dim i       As Integer
    
    If intIndex = UBound(mobjAlerts) Then
        
        If intIndex > 0 Then
            ReDim Preserve mobjAlerts(intIndex - 1) As Alert
            
        Else
            mbNoAlerts = True
        
        End If
        
        mlngAlerts = intIndex
        
    Else
    
        For i = (intIndex + 1) To UBound(mobjAlerts)
        
            mobjAlerts(i - 1) = mobjAlerts(i)
            
        Next
        
        ReDim Preserve mobjAlerts(UBound(mobjAlerts) - 1) As Alert
        
        mlngAlerts = (mlngAlerts - 1)
        
    End If
    
Exit_RemoveAlert:
    Exit Function
    
Err_RemoveAlert:
    Call ErrHandler("modClock.RemoveAlert", Err.Description)
    Resume Exit_RemoveAlert
     
End Function

Public Function DeleteAlert()
    On Error GoTo Err_DeleteAlert
    Dim intRemoved      As Integer
    Dim i               As Integer
    
    If Not mbNoAlerts Then
    
        With frmMain.WebBrowser1.Document.All
        
            For i = 1 To (.tblAlerts.rows.length - 1)
                
                If .tblAlerts.rows(i).cells(0).All(0).Checked Then
                
                    Call RemoveAlert(UBound(mobjAlerts) - ((i - intRemoved) - 1))
                    
                    intRemoved = intRemoved + 1
                    
                End If
                
            Next
        
        End With
        
        If intRemoved > 0 Then
        
            Call SaveAlertsToFile
            
            Call DrawAlerts
        
        End If
        
    End If
    
Exit_DeleteAlert:
    Exit Function
    
Err_DeleteAlert:
    Call ErrHandler("modClock.DeleteAlert", Err.Description)
    Resume Exit_DeleteAlert
    
End Function

Public Function UnloadForms()
    Dim frmForm As Form

    For Each frmForm In Forms

        Unload frmForm

    Next

End Function
