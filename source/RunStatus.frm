VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RunStatus 
   Caption         =   "Formula Modeler ラン状況"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6180
   OleObjectBlob   =   "RunStatus.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "RunStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'    Copyright (c) 2014, Fumito Hamamura
'    All rights reserved.
'
'    Redistribution and use in source and binary forms, with or without
'    modification, are permitted provided that the following conditions are met:
'
'    1. Redistributions of source code must retain the above copyright notice,
'       this list of conditions and the following disclaimer.
'    2. Redistributions in binary form must reproduce the above copyright notice,
'       this list of conditions and the following disclaimer in the documentation
'       and/or other materials provided with the distribution.
'
'    THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
'    ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
'    WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
'    DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR
'    ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
'    (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
'    LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
'    ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
'    (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
'    SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'

Public CancelButtonClicked As Boolean


Private Sub UserForm_Initialize()

    CancelButtonClicked = False
    Me.Caption = "Formula Modeler" + " " + VERSION_STRING + " ラン状況"

End Sub


Private Sub CancelButton_Click()

    CancelButtonClicked = True
    'Unload Me

End Sub



Public Sub UpdateStatus()

    Me.ElapsedTime = CDate(Date + Time - g_CurrModel.StartTime)
    Me.CurrentFormula = g_CurrModel.CurrentFormula
    'Me.CallCount = g_CurrModel.CallCount(g_CurrModel.CurrentFormula)
    'Me.CurrItrLevel = g_CurrModel.CurrentItrLevel
    
    If g_CurrModel.ItrLevel >= 1 Then Me.CallCount1 = g_CurrModel.ItrCount(1) Else Me.CallCount1 = Empty
    If g_CurrModel.ItrLevel >= 2 Then Me.CallCount2 = g_CurrModel.ItrCount(2) Else Me.CallCount2 = Empty
    If g_CurrModel.ItrLevel >= 3 Then Me.CallCount3 = g_CurrModel.ItrCount(3) Else Me.CallCount3 = Empty
    If g_CurrModel.ItrLevel >= 4 Then Me.CallCount4 = g_CurrModel.ItrCount(4) Else Me.CallCount4 = Empty
    If g_CurrModel.ItrLevel >= 5 Then Me.CallCount5 = g_CurrModel.ItrCount(5) Else Me.CallCount5 = Empty
    If g_CurrModel.ItrLevel >= 6 Then Me.CallCount6 = g_CurrModel.ItrCount(6) Else Me.CallCount6 = Empty

End Sub



