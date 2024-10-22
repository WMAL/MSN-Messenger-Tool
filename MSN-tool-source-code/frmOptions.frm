VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                        password options"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Password sets"
      Height          =   3105
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   3465
      Begin VB.CheckBox chkSmall 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Small letters"
         Height          =   345
         Left            =   180
         TabIndex        =   6
         Top             =   480
         Width           =   2235
      End
      Begin VB.CheckBox chkBig 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Capital letters"
         Height          =   345
         Left            =   180
         TabIndex        =   5
         Top             =   930
         Width           =   2235
      End
      Begin VB.CheckBox chkNumbers 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Numbers"
         Height          =   345
         Left            =   180
         TabIndex        =   4
         Top             =   1380
         Width           =   2235
      End
      Begin VB.CheckBox chkMix 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Mix"
         Height          =   345
         Left            =   180
         TabIndex        =   3
         Top             =   1830
         Width           =   2235
      End
      Begin prjMSN.lvButtons_H cmdAdd 
         Height          =   300
         Left            =   150
         TabIndex        =   7
         Top             =   2610
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "&add password"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
   End
   Begin prjMSN.lvButtons_H cmdOk 
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   3300
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      Caption         =   "&ok"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin prjMSN.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   2610
      TabIndex        =   1
      Top             =   3300
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      Caption         =   "&cancel"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'!!!!!***************!!!!!!!!!******************!!!!!!!!!!!!**********!
'Please read before making use of this code!
'Disclaimer: This is illegal if executed on real victims and could land you in prison for sure.
'This is intended for educational purposes only. We take no responsibility at all for your actions.
'This code is provided by EEEDS Eagle Eye Digital Security (Oman) for education purpose only.
'For more educational source codes please visit us http://www.digi77.com
'Author of this code W. Al Maawali Founder of  Eagle Eye Digital Solutions and Oman0.net can be reached via warith@digi77.com .

'Sharing knowledge is not about giving people something, or getting something from them.
'That is only valid for information sharing.
'Sharing knowledge occurs when people are genuinely interested in helping one another develop new capacities for action;
'it is about creating learning processes.
'Peter Senge
'!!!!!***************!!!!!!!!!******************!!!!!!!!!!!!**********!

Dim TemBig, TemSmall, TemNumbers, TemMix As String
Dim rsSmall As Recordset
Dim rsBig As Recordset
Dim rsNumbers As Recordset
Dim rsMix As Recordset

Private Sub chkBig_Click()
    SaveSetting "MSN Tool", "Options", "Big", chkBig.Value
End Sub

Private Sub chkMix_Click()
    SaveSetting "MSN Tool", "Options", "Mix", chkMix.Value
End Sub

Private Sub chkNumbers_Click()
    SaveSetting "MSN Tool", "Options", "Numbers", chkNumbers.Value
End Sub

Private Sub chkSmall_Click()
    SaveSetting "MSN Tool", "Options", "Small", chkSmall.Value
End Sub

Private Sub cmdAdd_Click()
    Dim temp As String
    
    temp$ = InputBox("Enter new password (will be added to mix list):", "Add new password")
    If temp = "" Then Exit Sub
    rsMix.AddNew
        rsMix!Password = temp$
    rsMix.Update
End Sub

Private Sub cmdCancel_Click()
    'Undo
    chkSmall.Value = TemSmall
    chkBig.Value = TemBig
    chkNumbers.Value = TemNumbers
    chkMix.Value = TemMix
    
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim rsFixed As Recordset
    
    If chkSmall.Value <> TemSmall Or _
       chkBig.Value <> TemBig Or _
       chkNumbers.Value <> TemNumbers Or _
       chkMix.Value <> TemMix Then
       
       Dim Resp As VbMsgBoxResult
       Resp = MsgBox("Apply password options, are you sure?", vbExclamation + vbYesNo)

       If Resp = vbNo Then Exit Sub
    End If
    
    'Clearing progress:
    SaveSetting "MSN Tool", "Data", "Password progress", "0"
    SaveSetting "MSN Tool", "Data", "Address progress", "0"
    
    If PAS.RecordCount = 0 Then GoTo STARTCOPY
    PAS.MoveFirst
    
    'Clearing tblPassword list:
    Do Until PAS.EOF
        PAS.Delete
        PAS.MoveNext
    Loop
    
    'Readding fixed passwords:
    Set rsFixed = db.OpenRecordset("tblFixed", dbOpenDynaset)
    rsFixed.MoveFirst
    
    Do Until rsFixed.EOF
        PAS.AddNew
            PAS!Password = rsFixed!Password
        PAS.Update
        rsFixed.MoveNext
    Loop
    
STARTCOPY:
    'Copying selected tables to tblPassword:
    If chkSmall.Value = 1 Then
        rsSmall.MoveFirst
        Do Until rsSmall.EOF
            PAS.AddNew
                PAS!Password = rsSmall!Password
            PAS.Update
            rsSmall.MoveNext
        Loop
    End If
    
    If chkBig.Value = 1 Then
        rsBig.MoveFirst
        Do Until rsBig.EOF
            PAS.AddNew
                PAS!Password = rsBig!Password
            PAS.Update
            rsBig.MoveNext
        Loop
    End If
    
    If chkNumbers.Value = 1 Then
        rsNumbers.MoveFirst
        Do Until rsNumbers.EOF
            PAS.AddNew
                PAS!Password = rsNumbers!Password
            PAS.Update
            rsNumbers.MoveNext
        Loop
    End If
    
    If chkMix.Value = 1 Then
        rsMix.MoveFirst
        Do Until rsMix.EOF
            PAS.AddNew
                PAS!Password = rsMix!Password
            PAS.Update
            rsMix.MoveNext
        Loop
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()

    Set rsSmall = db.OpenRecordset("tblSmall", dbOpenDynaset)
    Set rsBig = db.OpenRecordset("tblbig", dbOpenDynaset)
    Set rsNumbers = db.OpenRecordset("tblNumbers", dbOpenDynaset)
    Set rsMix = db.OpenRecordset("tblmix", dbOpenDynaset)

    If rsSmall.RecordCount = 0 Then chkSmall.Enabled = False
    If rsBig.RecordCount = 0 Then chkBig.Enabled = False
    If rsNumbers.RecordCount = 0 Then chkNumbers.Enabled = False
    If rsMix.RecordCount = 0 Then chkMix.Enabled = False
    
    'loading settings from registry
    TemSmall = GetSetting("MSN Tool", "Options", "Small", 0)
    TemBig = GetSetting("MSN Tool", "Options", "Big", 0)
    TemNumbers = GetSetting("MSN Tool", "Options", "Numbers", 0)
    TemMix = GetSetting("MSN Tool", "Options", "Mix", 0)
    
    chkSmall.Value = TemSmall
    chkBig.Value = TemBig
    chkNumbers.Value = TemNumbers
    chkMix.Value = TemMix
End Sub

