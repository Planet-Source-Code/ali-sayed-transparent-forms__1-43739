VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C00000&
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   315
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   402
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Transparent "
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_Trans As New ClsTrans
'********************************************
' Project:  Trans Class   v.1.0             *
' Author:   Ali Sayed                       *
' E-Mail:   AliSayed_7@Yahoo.com            *
' Date:     04/03/2003                      *
' Copyright © 2003 Ali Sayed                *
' Please let me know if you like it.        *
' For more information mail me.             *
'********************************************

Private Sub Command1_Click()
m_Trans.SubClass Me, Normal
End Sub

