VERSION 5.00
Begin VB.Form frmMXRecord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "InforMás MXRecord Class Test"
   ClientHeight    =   3612
   ClientLeft      =   42
   ClientTop       =   336
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3612
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Enter email or Domain to search, press enter"
      Height          =   795
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   5535
      Begin VB.TextBox txtDomain 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Text            =   "txtDomain"
         Top             =   300
         Width           =   5175
      End
   End
   Begin VB.ListBox lstMXInfo 
      Height          =   2576
      IntegralHeight  =   0   'False
      Left            =   60
      TabIndex        =   0
      Top             =   960
      Width           =   5535
   End
End
Attribute VB_Name = "frmMXRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' Form          : frmMXRecord
' Date          : 29/Ago/2005 23:40
' Author        : Jorge Colaccini (JRC) <software(AT)informas.com>
' Purpose       : To test imMXRecord Class
'---------------------------------------------------------------------------------------
'

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
        lParam As Any) As Long

Private Const LB_SETTABSTOPS = &H192

Private Sub Form_Load()


    With Me
        .Left = (Screen.Width - .Width) / 2
        .Top = (Screen.Height - .Height) / 2
        .txtDomain.Text = "elserver.com.ar" ' vbNullString
        .lstMXInfo.Clear
        .Show
        .Refresh
    End With
    
    
    SetLBTabs lstMXInfo, 15, 64, 190
    
    txtDomain.SetFocus

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub txtDomain_KeyPress(KeyAscii As Integer)

    Dim i   As Integer
    Dim SMXdomain As String
    Dim sAllMXDomains() As String
    
    On Error GoTo Err_MXQuery
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    

    lstMXInfo.Clear
    Screen.MousePointer = vbHourglass
    
    Dim cMXRecord As imMXRecord
    Set cMXRecord = New imMXRecord
    
    
    If cMXRecord.Initialized Then
        cMXRecord.Timeout = 1500 ' 1.5 seconds for timeout
        'Retrieve the best MX Domain
        SMXdomain = cMXRecord.MXRecord(txtDomain.Text)
        lstMXInfo.AddItem "MXDomain: " & SMXdomain & " - IP: " & cMXRecord.GetIPfromHostname(SMXdomain)
        lstMXInfo.AddItem ""
        sAllMXDomains = cMXRecord.MXRecordList(txtDomain.Text)
        
        
        If cMXRecord.Count > 0 Then
          For i = 0 To cMXRecord.Count - 1
            lstMXInfo.AddItem "MXDomain: (" & Format$(i, "00") & ") " & sAllMXDomains(0, i) & " - IP: " & cMXRecord.GetIPfromHostname(sAllMXDomains(0, i)) & " - pref: " & sAllMXDomains(1, i)
          Next
        End If
    End If
Exit_KeyPress:

    
    Set cMXRecord = Nothing
    
    Screen.MousePointer = vbNormal

    Exit Sub

Err_MXQuery:

    Screen.MousePointer = vbNormal
    MsgBox cMXRecord.LastErrorMsg   'Err.Description
    GoTo Exit_KeyPress
End Sub


Private Sub SetLBTabs(LB As ListBox, ParamArray TabStops())


    Dim aNewTabs()      As Long
    Dim lCtr            As Long
    Dim lTabs           As Long
    Dim lRet            As Long

    On Local Error GoTo Err_LBTabs:

    ReDim aNewTabs(UBound(TabStops)) As Long

    For lCtr = 0 To UBound(TabStops)
        aNewTabs(lCtr) = TabStops(lCtr)
    Next

    lTabs = UBound(aNewTabs) + 1

    LB.SetFocus
    
    lRet = SendMessage(LB.hwnd, LB_SETTABSTOPS, lTabs, aNewTabs(0))
  
    Exit Sub

Err_LBTabs::

End Sub

