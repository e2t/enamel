VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "�������� ���������� ������������"
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
        "������ ������:" & vbNewLine & _
        "����������� " & baseParams(0) & " �/�2" & vbNewLine & _
        "��� ������ " & augerParams(0) & " �/�2"
    Me.labEnamel.Caption = _
        "������ ����� (2 ����):" & vbNewLine & _
        "����������� " & baseParams(1) & " �/�2" & vbNewLine & _
        "��� ������ " & augerParams(1) & " �/�2"
    Run
End Sub
