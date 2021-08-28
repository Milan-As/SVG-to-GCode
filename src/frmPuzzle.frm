VERSION 5.00
Begin VB.Form frmPuzzle 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Puzzle size"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Clasic design"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtPiecesH 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      TabIndex        =   7
      Text            =   "4"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtPiecesW 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      TabIndex        =   6
      Text            =   "4"
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtHeight 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "5"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtWidth 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Text            =   "5"
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Cube size:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "pieces"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   5
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "pieces"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Height"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Width"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmPuzzle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
txtWidth.Text = ""
frmPuzzle.Hide
End Sub

Private Sub cmdOK_Click()

frmPuzzle.Hide

End Sub

Private Sub Form_Load()
Label4 = frmInterface.cmdMesure.Caption

Label5 = frmInterface.cmdMesure.Caption
Label7 = Round(txtWidth.Text / txtPiecesW.Text, 2) & " x " & Round(txtHeight.Text / txtPiecesH.Text, 2) & " " & Label4
'txtWidth = "1"
'txtHeight = "1"
'txtPiecesW = "1"
'txtPiecesH = "1"
 

End Sub

Private Sub txtHeight_Change()

If txtHeight <> "" Then
If txtHeight > 0.5 Then
 Label7 = Round(txtWidth.Text / txtPiecesW.Text, 2) & " x " & Round(txtHeight.Text / txtPiecesH.Text, 2) & " " & Label4
End If
End If
End Sub

Private Sub txtPiecesH_Change()

If txtPiecesH <> "" Then
If txtPiecesH > 1 Then
Label7 = Round(txtWidth.Text / txtPiecesW.Text, 2) & " x " & Round(txtHeight.Text / txtPiecesH.Text, 2) & " " & Label4
End If
End If
End Sub

Private Sub txtPiecesW_Change()
 
If txtPiecesW <> "" Then
If txtPiecesW > 1 Then
Label7 = Round(txtWidth.Text / txtPiecesW.Text, 2) & " x " & Round(txtHeight.Text / txtPiecesH.Text, 2) & " " & Label4
End If
End If
End Sub

Private Sub txtWidth_Change()

If txtWidth <> "" Then
If txtWidth > 0.5 Then
Label7 = Round(txtWidth.Text / txtPiecesW.Text, 2) & " x " & Round(txtHeight.Text / txtPiecesH.Text, 2) & " " & Label4
End If
End If
End Sub
