VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Source Code Genarator : Rajneesh Noonia"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOutPut 
      Height          =   465
      Left            =   1920
      TabIndex        =   4
      Text            =   "c:\Output.txt"
      Top             =   1290
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Code"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   1860
      Width           =   6255
   End
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "C:\Program Files\Microsoft SQL Server\80\Tools\Binn\sqldmo.dll"
      Top             =   660
      Width           =   3615
   End
   Begin VB.CommandButton cmdLocate 
      Caption         =   "Locate Component Binary"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   660
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog CDOPen 
      Height          =   480
      Left            =   4740
      TabIndex        =   6
      Top             =   180
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Binary COM component for which you want to generate source code"
      Height          =   465
      Left            =   120
      TabIndex        =   5
      Top             =   150
      Width           =   3645
   End
   Begin VB.Label Label1 
      Caption         =   "Output File :"
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   1260
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLocate_Click()
    
    Dim pFileName As String
    CDOPen.ShowOpen
    If (CDOPen.FileName <> "") Then
        pFileName = CDOPen.FileName
        txtPath.Text = pFileName
    End If
End Sub

Private Sub Command1_Click()
    Dim pCodeGen As New CodeGen
    On Error GoTo errortrap
    If (txtPath.Text <> "") Then
        pCodeGen.OutPutFileName = txtOutPut.Text
        Call pCodeGen.CodeGenerate(txtPath.Text)
    End If
    MsgBox "Done"
    Exit Sub
errortrap:
    MsgBox Err.Description
End Sub
