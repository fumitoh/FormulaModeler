Attribute VB_Name = "modErrInfo__"
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

Public errstr_FormulaNotFound As String
Public errstr_MainFormulaCalled As String
Public errstr_CallStackTooDeep As String
Public errstr_InvalidArgument As String
Public errstr_NotExpectedType As String
Public errstr_BeyondSystemLimit As String
Public errstr_WrongContext As String
Public errstr_FileError As String
Public errstr_InvalidLicense As String

Public Enum function_id

    id_VarToCol__ = 100
    id_VarToRow__
    id_NewModel__
    id_BreakItr__
    id_SetIndexItr__
    id_SetArrayItr__
    id_SetFileItr__
    id_GetItr__
    id_GetItrA__
    id_NextItr__
    id_NextItrA__
    id_SetItrSkip__
    id_PrefetchItr__
    id_FileInput__
    id_FileOutput__
    id_VarToFile__
    id_VecToLine__
    id_Model_Startup__
    
    id_MultDimArray__
    id_EmbedArray__
    id_ExtendUBound__
    id_Max__
    id_Min__
    id_UnNestArray__
    id_NestArray__
    id_ResizeNestedArray__
    id_NewArray__
    id_NewArray_Test__
    id_ReorderDim__
    id_AppendToVar__
    id_NewNestedArray__
    id_NewJaggedArray__
    id_NewLookupTable__
    id_LookupExact__
    id_LookupMatch__
    id_ClearVar__
    id_GetVar__
    id_SetNumber__
    id_FindVal__
    id_SumVarArray__
    id_SumVarArraySub__
    id_NumEprToArray__
    id_SplitFileExt__
    
    id_Startup__
    id_Formula__
        
    function_id_end
        
End Enum

Public FuncID As Long

Private m_FunctionNameString(100 To function_id_end) As String

Public Sub set_err_msg()

    FuncID = 0
    
    errstr_CallStackTooDeep = "Call Stack too deep: "
    errstr_MainFormulaCalled = "Main Formula called: "
    errstr_FormulaNotFound = "Formula not found: "
    errstr_InvalidArgument = "Invalid argument: "
    errstr_NotExpectedType = "Type not expected: "
    errstr_BeyondSystemLimit = "Beyond system limit: "
    errstr_WrongContext = "Wrong context: "
    errstr_FileError = "File Error: "
    errstr_InvalidLicense = "ライセンスが無効です。"
    
End Sub

