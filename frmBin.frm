VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WinBin 1.0 by Mischa Balen"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOverWrite 
      Caption         =   "OverWrite the File"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Width           =   4455
   End
   Begin VB.TextBox txtOvr 
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Text            =   "100"
      Top             =   3480
      Width           =   735
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Top             =   4920
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14111
            MinWidth        =   14111
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   4080
      TabIndex        =   5
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox txtFile 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   3855
   End
   Begin VB.CommandButton cmdGenBin 
      Caption         =   "Generate Random Binary Data"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   4455
   End
   Begin VB.TextBox txtLength 
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Text            =   "100"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtBinary 
      Height          =   1695
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "For this option, the length of the random binary data will equal the length of the file's contents."
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Overwrite the above file how many times  ::"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   3735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      X1              =   120
      X2              =   4560
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   4560
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label1 
      Caption         =   "Generate random binary data of length       ::"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   4560
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   3
      X1              =   120
      X2              =   4560
      Y1              =   2880
      Y2              =   2880
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NumberOfTimes As Long
'number of times we should overwrite the file

Private Sub cmdBrowse_Click()

Dim File1 As String

CD1.ShowOpen
File1 = FreeFile
    
If CD1.FileName <> "" Then 'if file name is true
    File1 = CD1.FileName 'return file path
    txtFile.Text = File1
ElseIf CD1.FileName = "" Then 'if file name is false
    File1 = ""
    Exit Sub
End If

File1 = ""
CD1.FileName = ""

End Sub

Private Sub cmdGenBin_Click()
'generate random binary data
SB1.Panels(1).Text = "Generating..."
    txtBinary.Text = RndBin(Val(txtLength.Text))
SB1.Panels(1).Text = "Finished"
End Sub

Private Sub cmdOverWrite_Click()
'overwrite file

On Error GoTo ErrSub
Dim i, x As Integer
Dim f As Long

NumberOfTimes = Val(txtOvr.Text) 'number of times to overwrite, specified by user

    Open txtFile.Text For Binary As #1 'open the file

    For i = 1 To NumberOfTimes 'loop until numberoftimes is full
        SB1.Panels(1).Text = "OverWriting... " & i & " of " & NumberOfTimes

            For x = 1 To LOF(1)
                Put #1, x, RndBin(1)
                FlushFileBuffers (1)
            Next x

    Next i

Close #1

SB1.Panels(1).Text = "Finished OverWriting - file is safe"

ErrSub:

If Err.Number <> 0 Then
    MsgBox ("Error: " & Err.Number & vbCrLf & Err.Description), vbCritical + vbOKOnly, "Error"
End If

End Sub
