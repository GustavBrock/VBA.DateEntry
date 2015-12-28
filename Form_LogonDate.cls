VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LogonDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

    Dim DefaultDate         As Date
    Dim DefaultFormat       As String
    Dim DefaultInputMask    As String

Private Sub Form_Error(DataErr As Integer, Response As Integer)
  
    Dim ctl                 As Control
    
    Dim SelStart            As Integer
    
    On Error Resume Next
    
    Set ctl = Screen.ActiveControl
    
    Select Case ctl.Name
        Case "Logon"
            SelStart = ctl.SelStart
            ' Clear deleted digits by resetting the input mask.
            ctl.InputMask = DefaultInputMask
            ctl.SelStart = SelStart
            ctl.SelLength = 1
            Response = acDataErrContinue
    End Select

    Set ctl = Nothing

End Sub

Private Sub Form_Load()
    
    Dim InitialDate         As Date
    
    ' Set initial date.
    InitialDate = Date
    
    ' Format and length of DefaultFormat and
    ' first part of DefaultInputMask must match.
    DefaultFormat = "yyyy/mm/dd"
    DefaultInputMask = "0000/00/00;1;0"
    
    Me!Logon.Format = DefaultFormat
    Me!Logon.InputMask = DefaultInputMask
    Me!Logon.ShowDatePicker = False
    SetDefaultDate InitialDate

End Sub

Private Sub Logon_AfterUpdate()

    With Me!Logon
        If IsNull(.Value) Then
            ' Rem this line out to allow the textbox to be cleared.
            .Value = DefaultDate
        ElseIf .Value < DateSerial(9999, 12, 31) Then
            SetDefaultDate DateAdd("d", 1, .Value)
        Else
            SetDefaultDate .Value
        End If
    End With
    
End Sub

Private Sub Logon_Click()

    Dim SelStart            As Integer
    
    With Me!Logon
        If .SelStart = 4 Or .SelStart = 7 Then
            ' Move the cursor off the separator (slash)
            ' to the first digit of months or days.
            .SelStart = .SelStart + 1
        End If
        SelStart = .SelStart
        .SelStart = SelStart
        .SelLength = 1
    End With

End Sub

Private Sub Logon_Enter()
    
    With Me!Logon
        If IsNull(.Value) Then
            .Value = DefaultDate
        End If
    End With

End Sub

Private Sub Logon_KeyPress(KeyAscii As Integer)
  
    Dim Text                As String
    Dim Length              As Integer
    Dim SelStart            As Integer
    
    With Me!Logon
        Select Case KeyAscii
            Case vbKeyBack, vbKeyTab, Asc(vbLf), vbKeyReturn, vbKeyEscape, vbKeyF16
                ' Allow navigation etc. with
                ' BackSpace, Tab, Ctrl+Enter, Enter, Escape, Ctrl+BackSpace
            Case Is > 0
                Text = .Text
                Length = Len(Text)
                SelStart = .SelStart
                
                If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
                    ' Replace any invalid entry with a zero.
                    KeyAscii = vbKey0
                End If
                
                If SelStart < Length Then
                    Select Case SelStart
                        
                        ' Year part.
                        Case Is = 0
                            ' First digit of year.
                            If KeyAscii = vbKey0 Then
                                ' No year before 1000.
                                KeyAscii = vbKey1
                            End If
                            
                        ' Month part.
                        Case Is = 5
                            ' First digit of month.
                            If KeyAscii > vbKey1 Then
                                ' No month with tens beyond 1.
                                KeyAscii = vbKey1
                            End If
                        Case Is = 6
                            ' Second digit of month.
                            Select Case Val(Mid(.Text, 6, 1))
                                Case Is = 0
                                    ' Month is < 10.
                                    If KeyAscii = vbKey0 Then
                                        ' Month cannot be 00.
                                        KeyAscii = vbKey1
                                    End If
                                Case Is > 0
                                    ' Month is 10+.
                                    If KeyAscii > vbKey2 Then
                                        ' No month beyond 12.
                                        KeyAscii = vbKey2
                                    End If
                            End Select
                        
                        ' Day part.
                        Case Is = 8
                            ' First digit of day.
                            Select Case Val(Mid(.Text, 6, 2))
                                Case Is = 2
                                    ' Month is February.
                                    If KeyAscii > vbKey2 Then
                                        ' No day with tens beyond 2 for February.
                                        KeyAscii = vbKey2
                                    End If
                                Case Else
                                    If KeyAscii > vbKey3 Then
                                        ' No day with tens beyond 3.
                                        KeyAscii = vbKey3
                                    End If
                            End Select
                        Case Is = 9
                            ' Second digit of day.
                            Select Case Mid(.Text, 9, 1)
                                Case Is = 3
                                    ' Days of 30.
                                    Select Case Val(Mid(.Text, 6, 2))
                                        Case 1, 3, 5, 7, 8, 10, 12
                                            If KeyAscii > vbKey1 Then
                                                ' No day beyond 31.
                                                KeyAscii = vbKey1
                                            End If
                                        Case 4, 6, 9, 11
                                            If KeyAscii > vbKey0 Then
                                                ' No day beyond 30.
                                                KeyAscii = vbKey0
                                            End If
                                    End Select
                                Case Is = 2
                                    ' Days of 20.
                                    Select Case Val(Mid(.Text, 6, 2))
                                        Case 2
                                            If KeyAscii = vbKey9 Then
                                                ' Check for leap year.
                                                If Month(DateAdd("d", 1, DateSerial(Val(Mid(.Text, 1, 4)), 2, 28))) = 3 Then
                                                    ' Not a leap year.
                                                    KeyAscii = vbKey8
                                                End If
                                            End If
                                    End Select
                                Case Is = 0
                                    ' Days of 00.
                                    If KeyAscii = vbKey0 Then
                                        ' No day of 00.
                                        KeyAscii = vbKey1
                                    End If
                            End Select
                            
                    End Select
                End If
                
        End Select
    End With

End Sub

Private Sub SetDefaultDate(ThisDate As Date)

    DefaultDate = ThisDate
    Me!Logon.DefaultValue = Format(ThisDate, "\#yyyy\/mm\/dd\#")

End Sub
