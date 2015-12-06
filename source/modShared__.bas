Attribute VB_Name = "modShared__"
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

Option Explicit
Option Private Module

Public Const PROJ_NAME = "FML" 'Must be Project Name
Public Const MAJOR_VER_NO = 0
Public Const MINOR_VER_NO = 5
Public Const REVISION_NO = 0

Public Const VERSION_STRING = "0.5.0"
'Public Const LICENSE_EXPIRE_ON = #12/21/2016#


Public Const MAX_DIM_COUNT = 10
Public Const MAX_FORMULA_ID = 99
Public Const MAX_CALL_STACK = 99

'--- Types Used in Model Class -----------------------------------
' Defined in this Module to avoid being exposed in Object browser

Public Enum fmlIteratorType

    fmlItrNull = 0
    fmlItrBoundary
    fmlItrArray
    fmlItrCSV

End Enum

Public Enum array_arg_type

    arg_is_paramarray
    arg_is_variant

End Enum

Public Enum ModelState_t

    state_null = 0
    state_in_startup
    state_in_formula
    state_in_model

End Enum


Public Type FormulaWrapper_t

    Formula As FML.IFormula
    Exist As Boolean
    CallCount As Long
    RunTime As Double  'In seconds
    StatupDone As Boolean

End Type

'--- Call Stack ---
Public Type FormulaCall_t

    ID As Long
    TimeStampSec As Double
    TimeStampDate As Long

End Type


'--- Iterator Implementation ---
Public Type Itr_t

    ItrType As fmlIteratorType
    IsEndOfItr As Boolean
    NextItr As Variant
    ItrCount As Long
    ItrCountUnit As Long
    
    '--- For Index Itr ---
    StartItr As Variant
    EndItr As Variant
    SkipListItr As Variant
    SkipListCount As Long
    
    '--- For Array Itr ---
    ArrayItr As Variant
    
    '--- For Fiele Itr ---
    FileItr As Long

End Type


'--- File I/O Implementation ---
Public Type FileIO_t

    FileNo As Long
    IsOpen As Boolean
    FilePath As String
    SkipCount As Long
    FieldLBound As Long
    FieldTypes As Variant   'Array of Arrays (Column, Type)

End Type
'----------------------------------

'Public Type RunInfo_t
'
'    StartTime As Date
'    ElapsedTime As Date
'    CurrFormula As Long
'    CallCount As Long
'    CurrItrLvel As Long
'
'End Type
'
'Public RunInfo As RunInfo_t

Public g_CurrModel As FModel
