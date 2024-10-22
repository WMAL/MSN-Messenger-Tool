VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmFound 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                                  cracked accounts"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   300
      Top             =   3660
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prjMSN.lvButtons_H cmdClose 
      Height          =   360
      Left            =   3690
      TabIndex        =   0
      Top             =   4440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   635
      Caption         =   "&close"
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
   Begin ComctlLib.ListView lv 
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   7435
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   12582912
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin prjMSN.lvButtons_H cmdExport 
      Default         =   -1  'True
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   635
      Caption         =   "&export"
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
End
Attribute VB_Name = "frmFound"
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

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()
    
    'Initializing dialog box:
    comDlg.DialogTitle = "Export address"
    comDlg.Filter = "Text files (*.txt)|*.txt"
    comDlg.ShowSave
    
    'File name can not be blank:
    If comDlg.FileName <> "" Then
        Open comDlg.FileName For Output As 1
            rsFound.MoveFirst
            Do Until rsFound.EOF
                Print #1, rsFound!address & "," & rsFound!Password
                rsFound.MoveNext
            Loop
        Close #1
        MsgBox "Exporting complete!", vbInformation
    End If
End Sub

Private Sub Form_Load()
    With lv.ColumnHeaders
        .Add , , "Address", (lv.Width / 2)
        .Add , , "Password", (lv.Width / 2) - 640
    End With
    
    rsFound.MoveFirst
    Do Until rsFound.EOF
        With lv.ListItems
            .Add , , rsFound!address
            rsFound.MoveNext
        End With
    Loop
    
    rsFound.MoveFirst
        Do Until rsFound.EOF
            lv.ListItems(rsFound.AbsolutePosition + 1).SubItems(1) = rsFound!Password
            rsFound.MoveNext
        Loop
End Sub
