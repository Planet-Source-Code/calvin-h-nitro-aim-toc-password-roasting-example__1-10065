VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TOC Sample #1 - Roasting Passwords"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLog 
      Height          =   3255
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   1680
      Width           =   5055
   End
   Begin VB.TextBox txtRoastingString 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Text            =   "Tic/Toc"
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Roast"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtRoastedPassword 
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "password"
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Caption         =   "This sample project shows how to roast an AOL Instant Messenger password to be sent to a TOC server."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  txtRoastedPassword.Text = ""
  txtLog.Text = ""
  txtLog.Text = txtLog.Text & "txtRoastedPassword cleared" & vbCrLf
  For a = 1 To Len(txtPassword.Text)
    b = b + 1
    If b > Len(txtRoastingString.Text) Then b = 1
    txtLog.Text = txtLog.Text & "Password String Index: " & b & vbCrLf
    txtLog.Text = txtLog.Text & "Roasting String Index: " & a & vbCrLf
    intPW = Asc(Mid(txtPassword.Text, a, 1))
    txtLog.Text = txtLog.Text & "Ascii index of current password character : " & intPW & "(" & Chr(intPW) & ")" & vbCrLf
    intRS = Asc(Mid(txtRoastingString.Text, b, 1))
    txtLog.Text = txtLog.Text & "Ascii index of current roasting string character : " & intRS & "(" & Chr(intRS) & ")" & vbCrLf
    intXOR = intPW Xor intRS
    txtLog.Text = txtLog.Text & intPW & " XOR " & intRS & " = " & intXOR & vbCrLf
    strXOR = Hex(intXOR)
    If Len(strXOR) = 1 Then strXOR = "0" & strXOR
    txtLog.Text = txtLog.Text & "Hex value of " & intXOR & " = " & strXOR & vbCrLf
    txtRoastedPassword.Text = txtRoastedPassword.Text & strXOR
  Next a
  txtRoastedPassword.Text = "0x" & txtRoastedPassword.Text
  txtLog.Text = txtLog.Text & "0x prefix added."
End Sub

Private Sub txtLog_Change()
  txtLog.SelStart = Len(txtLog.Text)
End Sub
