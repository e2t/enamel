Attribute VB_Name = "Main"
Option Explicit

''' 0 - priming, 1 - enamel
Public baseParams(1) As Double
Public augerParams(1) As Double

Dim swApp As Object
Public curDoc As ModelDoc2

Sub Main()
    Dim doc As ModelDoc2
    
    Set swApp = Application.SldWorks
    Set curDoc = swApp.ActiveDoc
    If curDoc Is Nothing Then Exit Sub
    Init
    MainForm.Show
End Sub

Function Init() ' mask for button
    baseParams(0) = 180 'g/m2
    baseParams(1) = 250 'g/m2
    augerParams(0) = 360 'g/m2
    augerParams(1) = 500 'g/m2
End Function

Function GetSelectedArea(doc As ModelDoc2) As Double
    Dim area As Double
    Dim mgr As SelectionMgr
    Dim face As Face2
    Dim i As Integer
    
    Set mgr = doc.SelectionManager
    area = 0
    For i = 1 To mgr.GetSelectedObjectCount2(-1)
        If mgr.GetSelectedObjectType3(i, -1) = swSelFACES Then
            Set face = mgr.GetSelectedObject6(i, -1)
            area = area + face.GetArea
        End If
    Next
    GetSelectedArea = area
End Function

Function GetResult(area As Double, isForAuger As Boolean) As String
    Dim ground As Double
    Dim enamel_2_layers As Double
    Dim params() As Double
    
    If area > 0 Then
        params = IIf(isForAuger, augerParams, baseParams)
        ground = area * params(0)
        enamel_2_layers = area * params(1)
        GetResult = "Грунт - " & Format(ground, "0.0") & " г" & vbNewLine & _
                    "Эмаль (2 слоя) - " & Format(enamel_2_layers, "0.0") & " г"
    Else
        GetResult = "Нет выделенных поверхностей."
    End If
End Function

Function ExitApp()  'mask for button
    Unload MainForm
    End
End Function
