VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2280
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2415
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   2415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save image as .JPG"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1830
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   90
      Picture         =   "Form1.frx":0000
      Top             =   30
      Width           =   2250
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'[Variables]

  

'[Screen Capture]
  Private Sub cmdSave_Click()
    Dim jDIB As cDIBSection
  ' obtain picture from handle
    Set jDIB = New cDIBSection
    jDIB.CreateFromPicture Image1.Picture
  ' save it
    If SaveJPG(jDIB, App.Path & "\Daughter.jpg", 80) Then
    ' success
    Else
      MsgBox "Failed to save picture."
    End If
  ' clean up
    jDIB.ClearUp
    Set jDIB = Nothing
  End Sub

