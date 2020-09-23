VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "DLL Testing Form"
   ClientHeight    =   945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2745
   LinkTopic       =   "Form1"
   ScaleHeight     =   945
   ScaleWidth      =   2745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCallFunction 
      Caption         =   "Test"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*/*/*/* PLEASE RATE */*/*/*'
'*/*/*/* PLEASE RATE */*/*/*'
'*/*/*/* PLEASE RATE */*/*/*'

'********************************************
'1st: You must add the DLL to the project. To
'do  that click on the  Menu  [Project]  then
'[Reference] choose the DLL project 'PrjName'
'********************************************

'Now: Create a new object from the DLL
Private objTest As New ClsName

Private Sub cmdCallFunction_Click()
    '*/*/*/* PLEASE RATE */*/*/*'
    '*/*/*/* PLEASE RATE */*/*/*'
    '*/*/*/* PLEASE RATE */*/*/*'
    
    'Use the command button to call the function.
    MsgBox objTest.AWX(4.2)
End Sub

