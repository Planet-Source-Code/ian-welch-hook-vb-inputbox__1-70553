VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "InputboxEx Demo"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3045
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLimit 
      Caption         =   "Limit To 20 Characters"
      Height          =   465
      Left            =   255
      TabIndex        =   2
      Top             =   1545
      Width           =   2475
   End
   Begin VB.CommandButton cmdPassword 
      Caption         =   "Password"
      Height          =   465
      Left            =   255
      TabIndex        =   1
      Top             =   892
      Width           =   2475
   End
   Begin VB.CommandButton cmdNumbersOnly 
      Caption         =   "Numbers Only"
      Height          =   465
      Left            =   255
      TabIndex        =   0
      Top             =   240
      Width           =   2475
   End
   Begin VB.Label lblResult 
      BackStyle       =   0  'Transparent
      Caption         =   "Result:"
      Height          =   795
      Left            =   285
      TabIndex        =   3
      Top             =   2640
      Width           =   2460
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLimit_Click()
Dim s As String
Dim bCanc As Boolean

s = InputBoxEx("Enter anything (20 characters max)", "InputboxEx", , , , , , 20, , , bCanc)

If bCanc Then
    lblResult.Caption = "Result:" & vbCrLf & "Cancel button pressed"
ElseIf s = "" Then
    lblResult.Caption = "Result:" & vbCrLf & "No text entered"
Else
    lblResult.Caption = "Result:" & vbCrLf & s
End If
End Sub

Private Sub cmdNumbersOnly_Click()
Dim s As String
Dim bCanc As Boolean

s = InputBoxEx("Enter a number (10 digits max)", "InputboxEx", , , , , , 10, , True, bCanc)

If bCanc Then
    lblResult.Caption = "Result:" & vbCrLf & "Cancel button pressed"
ElseIf s = "" Then
    lblResult.Caption = "Result:" & vbCrLf & "No text entered"
Else
    lblResult.Caption = "Result:" & vbCrLf & s
End If
End Sub

Private Sub cmdPassword_Click()
Dim s As String
Dim bCanc As Boolean

s = InputBoxEx("Enter password (8 characters max)", "InputboxEx", , , , , , 8, "*", , bCanc)

If bCanc Then
    lblResult.Caption = "Result:" & vbCrLf & "Cancel button pressed"
ElseIf s = "" Then
    lblResult.Caption = "Result:" & vbCrLf & "No text entered"
Else
    lblResult.Caption = "Result:" & vbCrLf & s
End If

End Sub
