VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Покрытие выделенных поверхностей"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5385
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Run()
    txtResult.Text = GetResult(GetSelectedArea(curDoc), chkAuger.Value)
End Sub

Private Sub btnClose_Click()
    ExitApp
End Sub

Private Sub chkAuger_Click()
    Run
End Sub

Private Sub UserForm_Initialize()
    Me.labPriming.Caption = _
        "Расход грунта:" & vbNewLine & _
        "стандартный " & baseParams(0) & " г/м2" & vbNewLine & _
        "для шнеков " & augerParams(0) & " г/м2"
    Me.labEnamel.Caption = _
        "Расход эмали (2 слоя):" & vbNewLine & _
        "стандартный " & baseParams(1) & " г/м2" & vbNewLine & _
        "для шнеков " & augerParams(1) & " г/м2"
    Run
End Sub
