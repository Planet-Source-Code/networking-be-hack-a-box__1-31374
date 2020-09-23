VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Hack a box"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optInput 
      Caption         =   "InputBox"
      Height          =   375
      Left            =   3960
      TabIndex        =   15
      Top             =   1680
      Width           =   2055
   End
   Begin VB.OptionButton optMsgBox 
      Caption         =   "MsgBox"
      Height          =   375
      Left            =   3960
      TabIndex        =   14
      Top             =   120
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Text            =   "Yup, I did it again :)"
      Top             =   360
      Width           =   3495
   End
   Begin VB.TextBox txtButton 
      Height          =   285
      Index           =   3
      Left            =   720
      TabIndex        =   9
      Text            =   "Button 4"
      Top             =   3720
      Width           =   2895
   End
   Begin VB.TextBox txtButton 
      Height          =   285
      Index           =   2
      Left            =   720
      TabIndex        =   8
      Text            =   "Button 3"
      Top             =   3360
      Width           =   2895
   End
   Begin VB.TextBox txtButton 
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   7
      Text            =   "Button 2"
      Top             =   3000
      Width           =   2895
   End
   Begin VB.TextBox txtButton 
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   6
      Text            =   "Button 1"
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CheckBox chkHelp 
      Caption         =   "Show help button"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   1080
      Width           =   3495
   End
   Begin VB.ComboBox lstButtons 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   4440
      List            =   "frmMain.frx":0016
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   3495
   End
   Begin VB.TextBox txtMessage 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmMain.frx":0083
      Top             =   960
      Width           =   3495
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   3720
      X2              =   3720
      Y1              =   120
      Y2              =   4080
   End
   Begin VB.Label Label5 
      Caption         =   "Title"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "1        2        3        4"
      Height          =   1335
      Left            =   480
      TabIndex        =   11
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Button text"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Style"
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Message"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdShow_Click()
    
    Dim hInst As Long
    Dim Thread As Long
    
    Dim T As Integer
    Dim Buttons As Integer
    
    ' if none selected, we just keep 0 (vbOkOnly)
    If lstButtons.Text <> "" Then Buttons = Left(lstButtons.Text, 1)
    
    ' do we need to show a help button?
    If chkHelp.Value = vbChecked Then Buttons = Buttons + vbMsgBoxHelpButton
    
    ' fill array with button text
    For T = 0 To 3
        ButtonText(T) = txtButton(T).Text
    Next T
    
    'Set up the CBT hook
    hInst = GetWindowLong(Me.hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    hHook = SetWindowsHookEx(WH_CBT, AddressOf Manipulate, hInst, _
                             Thread)
    
    'Display the box
    Dim Retval As Variant
    If optMsgBox.Value Then
        Retval = MsgBox(txtMessage.Text, Buttons, txtTitle.Text)
        Select Case Retval
        Case vbYes
            Retval = Retval & " - Yes"
        Case vbNo
            Retval = Retval & " - No"
        Case vbCancel
            Retval = Retval & " - Cancel"
        Case vbIgnore
            Retval = Retval & " - Ignore"
        Case vbAbort
            Retval = Retval & " - Abort"
        Case vbRetry
            Retval = Retval & " - Retry"
        Case vbOK
            Retval = Retval & " - Ok"
        End Select
    Else
        Retval = InputBox(txtMessage.Text, txtTitle)
    End If

    MsgBox "The previous box returned: " & vbCrLf & Retval

End Sub
