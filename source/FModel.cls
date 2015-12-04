VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Public DebugMode As Boolean

Private m_Param As Variant
Private m_RetVal As Variant
Private m_Var As Variant
Private m_Formula(MAX_FORMULA_ID) As FormulaWrapper_t
Private m_StartTime As Date
Private m_ModelState As ModelState_t

'--- Runtime Report --------------------------------
Private m_CurrCall As Long
Private m_CallStack(MAX_CALL_STACK) As FormulaCall_t
'___________________________________________________


'--- Itration --------------------------------------
Private Const MAX_ITERATION_LEVEL As Long = 100
Private Const SKIP_LIST_MAX_SIZE As Long = 100
Private m_CurrItrLevel As Long
'Private m_CurrentSkipListCount As Long
Private m_Itr(0 To MAX_ITERATION_LEVEL) As Itr_t
'---------------------------------------------------


'--- File IO ---------------------------------------
Private Const MAX_FILE_ID = 99
Private m_FileIn(0 To MAX_FILE_ID) As FileIO_t
Private m_FileOut(0 To MAX_FILE_ID) As FileIO_t
Private m_FileInUsed(0 To MAX_FILE_ID) As Boolean
Private m_FileOutUsed(0 To MAX_FILE_ID) As Boolean
'---------------------------------------------------


Private Sub Class_Initialize()

    m_Param = Empty
    m_RetVal = Empty

    'Initialize members 'TODO Initialize Global vars, Close Files

End Sub


Public Function Run(ParamArray Parameter() As Variant) As Variant
    
    Dim i As Long
    
    On Error GoTo HandleError:
    
    Call modErrInfo__.set_err_msg
    
'    If LICENSE_EXPIRE_ON < Date Then
'
'        Err.Raise Number:=fmlMainFormulaCalled, _
'                 Source:=PROJ_NAME, _
'                 Description:=modErrInfo__.errstr_InvalidLicense + "(使用期限 " + CStr(LICENSE_EXPIRE_ON) + ")"
'
'    End If
    
    If IsMissing(Parameter) Then m_Param = Empty Else m_Param = Parameter
    
    Set g_CurrModel = Me
    m_StartTime = Date + Time
    
    Load RunStatus
    RunStatus.StartTime = m_StartTime
    RunStatus.Show vbModeless
      
'    For i = 0 To MAX_FORMULA_ID
'
'        If m_Formula(i).Exist Then
'
'            m_Formula(i).Formula.Startup Me, m_Var
'
'        End If
'
'    Next i
    
    'm_ModelState = state_after_startup
    
    'On Error GoTo 0

    'm_Formula(0).Formula Me, v0, v1, v2, v3, v4, v5, v6, v7, v8, v9, v10
    
    'm_Formula(0).Formula.Startup Me, m_Var
    Main 0, m_Var
    
    Run = m_RetVal  'Set Return Value

    run_time_report

    Close
    Unload RunStatus
    
    'End
    
    Exit Function

HandleError:
    
    Close
    
    Debug.Print "Startup Error: " + CStr(Err.Number) + " " + Err.Description
    Debug.Print "Startup: " + CStr(i)
    Debug.Print "Source: " + CStr(Err.Source)

    If Err.Source = PROJ_NAME Then     'Excel.Application.VBE.ActiveVBProject.Name Then

        If modErrInfo__.FuncID <> 0 Then

            Debug.Print "Function: " + CStr(modErrInfo__.FuncID)

        End If

    End If
    
    MsgBox "Error: " + CStr(Err.Number) + " " + Err.Description
    
    End

End Function



Public Sub Main( _
    FormulaID As Long, _
    Optional ByRef v0 = Empty, _
    Optional ByRef v1 = Empty, _
    Optional ByRef v2 = Empty, _
    Optional ByRef v3 = Empty, _
    Optional ByRef v4 = Empty, _
    Optional ByRef v5 = Empty, _
    Optional ByRef v6 = Empty, _
    Optional ByRef v7 = Empty, _
    Optional ByRef v8 = Empty, _
    Optional ByRef v9 = Empty, _
    Optional ByRef v10 = Empty)
    
    Dim curr_sec As Double
    Dim curr_date As Long
    Dim date_count As Long

    CheckEvents
    
    On Error GoTo HandleError:
    
    If m_ModelState = state_in_formula And FormulaID = 0 Then
    
        Err.Raise Number:=fmlMainFormulaCalled, _
                 Source:=PROJ_NAME, _
                 Description:=modErrInfo__.errstr_MainFormulaCalled + "Formula " + CStr(FormulaID)
    
    
    ElseIf Not m_Formula(FormulaID).Exist Then
    
        Err.Raise Number:=fmlFormulaNotFound, _
                 Source:=PROJ_NAME, _
                 Description:=modErrInfo__.errstr_FormulaNotFound + "Formula " + CStr(FormulaID)
                 
    ElseIf m_CurrCall + 1 > MAX_CALL_STACK Then
    
        Err.Raise Number:=fmlCallStackTooDeep, _
                 Source:=PROJ_NAME, _
                 Description:=modErrInfo__.errstr_CallStackTooDeep + "Formula " + CStr(FormulaID)
                 
    End If
    
    m_ModelState = state_in_model
    
    Call stack_push(FormulaID)
    
    If Not m_Formula(FormulaID).StatupDone Then
        
        m_ModelState = state_in_startup
        m_Formula(FormulaID).Formula.Startup Me, v0, v1, v2, v3, v4, v5, v6, v7, v8, v9, v10
        m_Formula(FormulaID).StatupDone = True
        m_ModelState = state_in_model
    
    End If
        
    m_ModelState = state_in_formula
    m_Formula(FormulaID).Formula.Main Me, v0, v1, v2, v3, v4, v5, v6, v7, v8, v9, v10
    m_Formula(FormulaID).CallCount = m_Formula(FormulaID).CallCount + 1
    m_ModelState = state_in_model
        
    Call stack_pop
    
    Exit Sub

HandleError:

    Debug.Print "Error: " + CStr(Err.Number) + " " + Err.Description
    Debug.Print "Driver: " + CStr(FormulaID)
    Debug.Print "Iteration Level: " + CStr(m_CurrItrLevel)
    Debug.Print "Next Iterator: " + get_nextitr_str()
    Debug.Print "Source: " + CStr(Err.Source)
    
    If Err.Source = PROJ_NAME Then  'Excel.Application.VBE.ActiveVBProject.Name Then
    
        If modErrInfo__.FuncID <> 0 Then
        
            Debug.Print "Function: " + CStr(modErrInfo__.FuncID)
        
        End If
    
    End If
    
    MsgBox "Error: " + CStr(Err.Number) + " " + Err.Description
    
    End 'Terminate Run
    
End Sub


Private Sub stack_push(FormulaID As Long)

    'Run Time Tracking
    'Stop calling formulas clock
    If Not FormulaID = 0 Then Call add_time
        
    'Push the current formula onto the CallStack
    m_CurrCall = m_CurrCall + 1
    m_CallStack(m_CurrCall).ID = FormulaID
    m_CallStack(m_CurrCall).TimeStampSec = CDbl(Timer)
    m_CallStack(m_CurrCall).TimeStampDate = CLng(Date)


End Sub


Private Sub stack_pop()

    'Run Time Tracking
    Call add_time
        
    'Resume calling formulas clock
    m_CurrCall = m_CurrCall - 1
    m_CallStack(m_CurrCall).TimeStampSec = CDbl(Timer)
    m_CallStack(m_CurrCall).TimeStampDate = CLng(Date)


End Sub



'Explicit Startup Call
Public Sub Startup( _
    FormulaID As Long, _
    Optional ByRef v0 = Empty, _
    Optional ByRef v1 = Empty, _
    Optional ByRef v2 = Empty, _
    Optional ByRef v3 = Empty, _
    Optional ByRef v4 = Empty, _
    Optional ByRef v5 = Empty, _
    Optional ByRef v6 = Empty, _
    Optional ByRef v7 = Empty, _
    Optional ByRef v8 = Empty, _
    Optional ByRef v9 = Empty, _
    Optional ByRef v10 = Empty)
    
    Dim prev_model_state As ModelState_t

    CheckEvents
    
    On Error GoTo HandleError:

    If m_ModelState <> state_in_startup Then
    
        Err.Raise Number:=fmlMainFormulaCalled, _
                 Source:=PROJ_NAME, _
                 Description:=modErrInfo__.errstr_MainFormulaCalled + "Formula " + CStr(FormulaID)
    
    
    ElseIf Not m_Formula(FormulaID).Exist Then
    
        Err.Raise Number:=fmlFormulaNotFound, _
                 Source:=PROJ_NAME, _
                 Description:=modErrInfo__.errstr_FormulaNotFound + "Formula " + CStr(FormulaID)
                 
    ElseIf m_CurrCall + 1 > MAX_CALL_STACK Then
    
        Err.Raise Number:=fmlCallStackTooDeep, _
                 Source:=PROJ_NAME, _
                 Description:=modErrInfo__.errstr_CallStackTooDeep + "Formula " + CStr(FormulaID)
                 
    End If

    prev_model_state = m_ModelState
    m_ModelState = state_in_model
    
    Call stack_push(FormulaID)
    
    If Not m_Formula(FormulaID).StatupDone Then
        
        m_ModelState = state_in_startup
        m_Formula(FormulaID).Formula.Startup Me, v0, v1, v2, v3, v4, v5, v6, v7, v8, v9, v10
        m_Formula(FormulaID).StatupDone = True
        m_ModelState = state_in_model
    
    End If
                
    Call stack_pop

    m_ModelState = prev_model_state
    
    Exit Sub

HandleError:

    modErrInfo__.FuncID = id_Model_Startup__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext


End Sub


Private Function ind2str(ind As Variant) As String

    Dim ind_lb As Long, ind_ub As Long
    Dim i As Long
    Dim s As String
    
    If IsEmpty(ind) Then
    
        ind2str = ""
        Exit Function
    
    End If

    ind_lb = LBound(ind)
    ind_ub = UBound(ind)

    s = CStr(ind(ind_lb))
    For i = ind_lb + 1 To ind_ub

        s = s + ", " + CStr(ind(i))
        
    Next i
    
    ind2str = s

End Function


Public Property Get CurrentFormula() As Long

    CurrentFormula = m_CallStack(m_CurrCall).ID
    
End Property

Public Property Get CallCount(FormulaID As Long) As Long

    CallCount = m_Formula(FormulaID).CallCount

End Property

'Public Property Get CurrentItrLevel() As Long
'
'    CurrentItrLevel = m_CurrItrLevel
'
'End Property

Public Property Get StartTime() As Date

    StartTime = m_StartTime

End Property


Public Property Set Formula(Optional FormulaID As Long = 0, f As FML.IFormula)

    Dim fml_wrapper As FormulaWrapper_t
    
    If Not (0 <= FormulaID And FormulaID <= MAX_FORMULA_ID) Then
    
        'TODO: Throw Error
    
    ElseIf m_Formula(FormulaID).Exist Then
    
        'TODO: Throw Error
    
    End If
    
    Set fml_wrapper.Formula = f
    fml_wrapper.Exist = True
    
    m_Formula(FormulaID) = fml_wrapper
    
    'Set m_Formula(FormulaID).Formula = f
    'm_Formula(FormulaID).Exist = True

End Property


Public Property Get Formula(Optional FormulaID As Long = 0) As FML.IFormula
    
    If Not (0 <= FormulaID And FormulaID <= MAX_FORMULA_ID) Then
    
        'TODO: Throw Error
    
    ElseIf Not m_Formula(FormulaID).Exist Then
    
        'TODO: Throw Error
    
    End If
    
    Set Formula = m_Formula(FormulaID).Formula

End Property



'------------- Iterator Methods ------------------------------
'   SetIndexItr
'       set_itr_from
'       set_itr_to
'   SetArrayItr
'       set_itr_array
'   NextItr
'   GetItr
'   GetItrAsArray
'   SetIndexItrSkip
'   PrefetchItr



Public Sub SetIndexItr(ParamArray Index() As Variant)

    Dim arg_ub As Long
    Dim i As Long
    Dim itr_from As Variant, itr_to As Variant
    Dim ind_len As Long
        
    On Error GoTo HandleError:
    
    ind_len = UBound(Index) - LBound(Index) + 1
    
    If ind_len <= 0 Then
    
        ReDim itr_from(0)
        ReDim itr_to(0)
    
        itr_from(0) = 0
        itr_to(0) = 2147483647
        
    ElseIf ind_len = 1 Then
    
        ReDim itr_from(0)
        ReDim itr_to(0)
    
        itr_from(0) = Index(0)
        itr_to(0) = 2147483647

    'Arg is a single pair of From-To
    ElseIf Not IsArray(Index(0)) Then
    
        If ind_len = 2 Then
        
            ReDim itr_from(0)
            ReDim itr_to(0)
            
            itr_from(0) = Index(0)
            itr_to(0) = Index(1)
            
        Else
            
            Err.Raise Number:=fmlInvalidArgument, _
                      Source:=PROJ_NAME, _
                      Description:=modErrInfo__.errstr_InvalidArgument + "Index"
            
        End If
        
    Else    'Arg is an array of From-To arrays
    
        arg_ub = UBound(Index)
        ReDim itr_from(arg_ub)
        ReDim itr_to(arg_ub)
        
        For i = 0 To arg_ub
       
           itr_from(i) = Index(i)(LBound(Index(i)))
           itr_to(i) = Index(i)(LBound(Index(i)) + 1)

        Next i
    
    End If
    
    set_itr_from itr_from
    set_itr_to itr_to
              
    Exit Sub

HandleError:
    
    modErrInfo__.FuncID = id_SetIndexItr__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext


End Sub


Public Sub SetArrayItr(ByVal Source As Variant)
    
    
    On Error GoTo HandleError:
    
    'Copy to var for the case of Source being Range
    If IsObject(Source) Then
    
        Source = Source
    
    End If
    
    If Not modArrSpt__.IsArrayAllocated(Source) Then
    
        Err.Raise Number:=fmlInvalidArgument, _
                  Source:=PROJ_NAME, _
                  Description:=modErrInfo__.errstr_InvalidArgument + "Source"
    End If
    
    If m_CurrItrLevel + 1 > MAX_ITERATION_LEVEL Then
    
        Err.Raise Number:=fmlBeyondSystemLimit, _
                  Source:=PROJ_NAME, _
                  Description:=modErrInfo__.errstr_BeyondSystemLimit + "MAX_ITERATION_LEVEL"
    End If
    
    m_CurrItrLevel = m_CurrItrLevel + 1
    m_Itr(m_CurrItrLevel).ItrType = fmlItrArray
    m_Itr(m_CurrItrLevel).ArrayItr = Source
    m_Itr(m_CurrItrLevel).StartItr = LBound(Source)
    m_Itr(m_CurrItrLevel).NextItr = LBound(Source)
    m_Itr(m_CurrItrLevel).EndItr = UBound(Source)
    m_Itr(m_CurrItrLevel).ItrCountUnit = 1
    
    Exit Sub
                
HandleError:
    
    modErrInfo__.FuncID = id_SetArrayItr__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext
                
End Sub

Private Function set_itr_from(Index As Variant) As Boolean
'Index as Array(Long)

    Dim ind_len As Long, i As Long
    Dim type_check As Boolean
        
    set_itr_from = False
    
    ind_len = UBound(Index)
            
    For i = 0 To ind_len
        
        If Not modUtil__.is_whole_number(Index(i)) Then
        
            Err.Raise Number:=fmlInvalidArgument, _
                  Source:=PROJ_NAME, _
                  Description:=modErrInfo__.errstr_InvalidArgument + "Index"
            
        End If
            
    Next i
        
    If m_CurrItrLevel + 1 > MAX_ITERATION_LEVEL Then
    
        Err.Raise Number:=fmlBeyondSystemLimit, _
                  Source:=PROJ_NAME, _
                  Description:=modErrInfo__.errstr_BeyondSystemLimit + "MAX_ITERATION_LEVEL"
    End If
    
    m_CurrItrLevel = m_CurrItrLevel + 1
    m_Itr(m_CurrItrLevel).ItrType = fmlItrBoundary
    m_Itr(m_CurrItrLevel).StartItr = Index
    m_Itr(m_CurrItrLevel).NextItr = Index
    m_Itr(m_CurrItrLevel).ItrCountUnit = 1

    set_itr_from = True

End Function

Private Function set_itr_to(Index As Variant) As Boolean
'Not to be exposed outside the project

    Dim i As Long
    Dim ind_len As Long
    
    set_itr_to = False
    
'    If m_CurrItrLevel < 1 Or _
'       Not IsEmpty(m_Itr(m_CurrItrLevel).ArrayItr) Then
'
'        Err.Raise Number:=fmlWrongContext, _
'            Source:=PROJ_NAME, _
'            Description:=modErrInfo__.errstr_WrongContext + "SetIndexItrTo"
'
'    End If
'
'    ind_len = UBound(index)
'    If ind_len <> UBound(m_Itr(m_CurrItrLevel).StartItr) Then
'
'        Err.Raise Number:=fmlInvalidArgument, _
'            Source:=PROJ_NAME, _
'            Description:=modErrInfo__.errstr_InvalidArgument + "Index"
'
'    End If
    
    
    For i = 0 To ind_len
        
        If Not modUtil__.is_whole_number(Index(i)) Then
        
            Err.Raise Number:=fmlInvalidArgument, _
                  Source:=PROJ_NAME, _
                  Description:=modErrInfo__.errstr_InvalidArgument + "Arg"
            
        End If
            
    Next i
    
    
    m_Itr(m_CurrItrLevel).EndItr = Index
    set_itr_to = True


End Function


Public Sub SetFileItr(FileID As Long, Optional CountUnit As Long = 1)

    Dim skip_count As Long
    Dim temp As Variant
    
    On Error GoTo HandleError:
    
    If Not m_FileInUsed(FileID) Then
    
        Err.Raise Number:=fmlFileError, _
              Source:=PROJ_NAME, _
              Description:=modErrInfo__.errstr_FileError + "File ID Not Assigned."
              
    ElseIf m_FileIn(FileID).IsOpen Then
    
        Err.Raise Number:=fmlFileError, _
              Source:=PROJ_NAME, _
              Description:=modErrInfo__.errstr_FileError + "Already in Use."
    
    Else
    
        m_FileIn(FileID).FileNo = FreeFile
        Open m_FileIn(FileID).FilePath For Input As m_FileIn(FileID).FileNo
        m_FileIn(FileID).IsOpen = True
    
    End If
    
    'Push to Iteration Stack
    m_CurrItrLevel = m_CurrItrLevel + 1
    m_Itr(m_CurrItrLevel).ItrType = fmlItrCSV
    m_Itr(m_CurrItrLevel).FileItr = FileID
    m_Itr(m_CurrItrLevel).ItrCountUnit = CountUnit
    
    '--- Skip Header ---
    skip_count = m_FileIn(FileID).SkipCount
    Do While (skip_count > 0) And (Not EOF(m_FileIn(FileID).FileNo))
    
        Line Input #m_FileIn(FileID).FileNo, temp
        
        skip_count = skip_count - 1
    Loop
    
    If EOF(m_FileIn(FileID).FileNo) Then
    
        m_Itr(m_CurrItrLevel).IsEndOfItr = True
    
    End If

    Exit Sub
    
HandleError:
    
    modErrInfo__.FuncID = id_SetFileItr__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext
    
End Sub


Public Function NextItr(ParamArray Index() As Variant) As Boolean

    Dim ind_buf As Variant
    Dim ind_dummy(0) As Variant 'Updated on 2013-10-14
    Dim i As Long

    On Error GoTo HandleError:


    If UBound(Index) < LBound(Index) Then
    
        NextItr = next_itr(arg_is_paramarray, ind_dummy) 'Updated on 2013-10-14
    
    Else
        
        ind_buf = Index
        NextItr = next_itr(arg_is_paramarray, ind_buf)
        
        For i = 0 To UBound(ind_buf)

            Index(i) = ind_buf(i)

        Next i
        
    End If

Exit Function
    

HandleError:
    
    modErrInfo__.FuncID = id_NextItr__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext

End Function

Public Function NextItrA(Optional Index As Variant, Optional ByVal LBoundOverride As Long = 0) As Boolean

    Dim ind_dummy(0) As Variant 'Updated on 2013-10-14
    
    On Error GoTo HandleError:

    If IsMissing(Index) Then
    
        NextItrA = next_itr(arg_is_variant, ind_dummy) 'Updated on 2013-10-14
    
    Else
        
        NextItrA = next_itr(arg_is_variant, Index, LBoundOverride)
        
    End If

Exit Function
                
HandleError:
    
    modErrInfo__.FuncID = id_NextItrA__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext

End Function



Private Function next_itr(arg_type As array_arg_type, Optional Index As Variant, Optional ByVal IndexLBound As Long = 0) As Boolean
'NextItr and NextItrA are almost the same except that NextItr
'returns results in Paramarray but NextItr does in Variant Array.

    Dim skip_flag As Boolean
    Dim skip_list_count As Long
    Dim skip_list_ub As Long, skip_list_lb As Long
    Dim i As Long
    Dim is_end_of_etr As Boolean
    Dim ind_len As Long, lb As Long, ub_ind As Long, lb_ind As Long
    Dim str_buf As Variant
    Dim var_buf As Variant
    Dim fld_t As Variant

    'Error if there's no iteration.
    If m_CurrItrLevel < 1 Then
    
        Err.Raise Number:=fmlWrongContext, _
                  Source:=PROJ_NAME, _
                  Description:=modErrInfo__.errstr_WrongContext
    End If

    Do
        'If End of Current Iteration, Clean up the current level in Stack
        If m_Itr(m_CurrItrLevel).IsEndOfItr Then
            
            Call clean_up_curr_level
                        
            next_itr = False
            Exit Function
    
        End If
        
        'Check whether to skip new counter
        skip_flag = False
'        If Not IsEmpty(m_Itr(m_CurrItrLevel).SkipListItr) Then
        If m_Itr(m_CurrItrLevel).SkipListCount > 0 Then
        
            For i = 0 To m_Itr(m_CurrItrLevel).SkipListCount - 1
            
                If m_Itr(m_CurrItrLevel).ItrType = fmlItrBoundary Then
                
                    skip_flag = skip_flag Or modUtil__.compare_index(m_Itr(m_CurrItrLevel).NextItr, m_Itr(m_CurrItrLevel).SkipListItr(i))
                    
                ElseIf m_Itr(m_CurrItrLevel).ItrType = fmlItrArray Then
                
                    skip_flag = skip_flag Or modUtil__.compare_index(m_Itr(m_CurrItrLevel).ArrayItr(m_Itr(m_CurrItrLevel).NextItr), m_Itr(m_CurrItrLevel).SkipListItr(i))
                
                End If
            
            Next i
         
        End If

        
        'Check Arg & decide length of index
        If IsMissing(Index) Then
            ind_len = 0
        
        ElseIf arg_type = arg_is_paramarray Then
            
            'Get Index
            ub_ind = UBound(Index)
            lb_ind = LBound(Index)
            ind_len = ub_ind - lb_ind + 1 'Could be 0
            
        Else    'Create Index Array
                
            If m_Itr(m_CurrItrLevel).ItrType = fmlItrBoundary Then

                ReDim Index(IndexLBound To UBound(m_Itr(m_CurrItrLevel).StartItr) + IndexLBound)
            
                'Get Index
                ub_ind = UBound(Index)
                lb_ind = LBound(Index)
                ind_len = ub_ind - lb_ind + 1 'Could be 0
                        
            ElseIf m_Itr(m_CurrItrLevel).ItrType = fmlItrArray Then 'lb_ind + Original Array LB
            
                ReDim Index(LBound(m_Itr(m_CurrItrLevel).ArrayItr, 2) + IndexLBound To UBound(m_Itr(m_CurrItrLevel).ArrayItr, 2) + IndexLBound)
                
                'Get Index
                ub_ind = UBound(Index)
                lb_ind = LBound(Index)
                ind_len = ub_ind - lb_ind + 1 'Could be 0
                
                
            ElseIf m_Itr(m_CurrItrLevel).ItrType = fmlItrCSV Then
            
                'Decide length of Index later
                Index = Array(0) 'Dummy
                
                lb_ind = IndexLBound + m_FileIn(m_Itr(m_CurrItrLevel).FileItr).FieldLBound
                ind_len = 1
                
            
            Else
                Err.Raise Number:=fmlInvalidArgument, _
                          Source:=PROJ_NAME, _
                          Description:=modErrInfo__.errstr_InvalidArgument + "IterationType"
            End If
            
            
        End If
        
        
        'This part should be the similar to that in GetItr.
        'Duplicated because ParmArray cannot be passed.
      
        If Not skip_flag And ind_len > 0 Then
        
            If m_Itr(m_CurrItrLevel).ItrType = fmlItrBoundary Then
            
                If ub_ind > UBound(m_Itr(m_CurrItrLevel).StartItr) + lb_ind Then
                
                    Err.Raise Number:=fmlInvalidArgument, _
                              Source:=PROJ_NAME, _
                              Description:=modErrInfo__.errstr_InvalidArgument + "Index length is too long."
                End If
                    
                For i = lb_ind To ub_ind
            
                    Index(i) = m_Itr(m_CurrItrLevel).NextItr(i - lb_ind)
                
                Next i
                
            ElseIf m_Itr(m_CurrItrLevel).ItrType = fmlItrArray Then
            
                lb = LBound(m_Itr(m_CurrItrLevel).ArrayItr, 2)
                
                If ind_len > UBound(m_Itr(m_CurrItrLevel).ArrayItr, 2) - lb + 1 Then
                
                    Err.Raise Number:=fmlInvalidArgument, _
                              Source:=PROJ_NAME, _
                              Description:=modErrInfo__.errstr_InvalidArgument + "Index length is too long."
                End If
        
                For i = lb_ind To ub_ind
                
                    Index(i) = m_Itr(m_CurrItrLevel).ArrayItr(m_Itr(m_CurrItrLevel).NextItr, i + lb - lb_ind)
                    
                Next i
            
            ElseIf m_Itr(m_CurrItrLevel).ItrType = fmlItrCSV Then
            
                Line Input #m_FileIn(m_Itr(m_CurrItrLevel).FileItr).FileNo, str_buf
                
                str_buf = Split(str_buf, ",")
                
                ReDim var_buf(lb_ind To UBound(str_buf) + lb_ind)
                
                If Not IsEmpty(m_FileIn(m_Itr(m_CurrItrLevel).FileItr).FieldTypes) Then
                
                    For Each fld_t In m_FileIn(m_Itr(m_CurrItrLevel).FileItr).FieldTypes
                    
                        If fld_t(0) >= m_FileIn(m_Itr(m_CurrItrLevel).FileItr).FieldLBound And _
                           fld_t(0) <= UBound(str_buf) + m_FileIn(m_Itr(m_CurrItrLevel).FileItr).FieldLBound Then
                            
                            Select Case fld_t(1)
                            
                            Case VbVarType.vbBoolean
                                
                                var_buf(fld_t(0) + IndexLBound) = CBool(str_buf(fld_t(0)))
                            
                            Case VbVarType.vbDate
                            
                                var_buf(fld_t(0) + IndexLBound) = CDate(str_buf(fld_t(0)))
                            
                            Case VbVarType.vbLong
                            
                                var_buf(fld_t(0) + IndexLBound) = CLng(str_buf(fld_t(0)))
                            
                            Case VbVarType.vbDouble
                            
                                var_buf(fld_t(0) + IndexLBound) = CDbl(str_buf(fld_t(0)))
                            
                            Case VbVarType.vbString
                            
                                var_buf(fld_t(0) + IndexLBound) = CStr(str_buf(fld_t(0)))
                            
                            Case Else
                                'Error
                                                    
                            End Select
                            
                        End If
                    
                    Next fld_t
                    
                End If
                
                For i = lb_ind To UBound(var_buf)
                
                    If IsEmpty(var_buf(i)) Then
                        
                        var_buf(i) = modUtil__.str2var(str_buf(i - lb_ind))
                    
                    End If
                
                Next i
                
                If arg_type = arg_is_variant Then
                
                    Index = var_buf
                
                Else    'arg_is_param_array
                
                    For i = lb_ind To ub_ind
                    
                        Index(i - lb_ind) = var_buf(i - lb_ind)
                    
                    Next i
                
                End If
                
            Else
            
                Err.Raise Number:=fmlInvalidArgument, _
                          Source:=PROJ_NAME, _
                          Description:=modErrInfo__.errstr_InvalidArgument + "IterationType"
                
            End If
        
        End If  'End of Getting Index
                
        
        'Increment Current Iterator
        If m_Itr(m_CurrItrLevel).ItrType = fmlItrBoundary Then
               
            m_Itr(m_CurrItrLevel).IsEndOfItr = Not modUtil__.increment_index(m_Itr(m_CurrItrLevel).NextItr, m_Itr(m_CurrItrLevel).StartItr, m_Itr(m_CurrItrLevel).EndItr)
            
        ElseIf m_Itr(m_CurrItrLevel).ItrType = fmlItrArray Then
        
            If m_Itr(m_CurrItrLevel).NextItr = m_Itr(m_CurrItrLevel).EndItr Then
            
                m_Itr(m_CurrItrLevel).IsEndOfItr = True
            
            Else
                m_Itr(m_CurrItrLevel).NextItr = m_Itr(m_CurrItrLevel).NextItr + 1
            
            End If
            
        ElseIf m_Itr(m_CurrItrLevel).ItrType = fmlItrCSV Then
        
            If EOF(m_FileIn(m_Itr(m_CurrItrLevel).FileItr).FileNo) Then
            
                m_Itr(m_CurrItrLevel).IsEndOfItr = True
            
            End If
            
        End If
                
        m_Itr(m_CurrItrLevel).ItrCount = m_Itr(m_CurrItrLevel).ItrCount + 1
        
    Loop While skip_flag
          
    next_itr = True
    
End Function


Public Sub BreakItr()

    On Error GoTo HandleError:

    'Error if there's no iteration.
    If m_CurrItrLevel < 1 Then
    
        Err.Raise Number:=fmlWrongContext, _
                  Source:=PROJ_NAME, _
                  Description:=modErrInfo__.errstr_WrongContext
    End If
    
    Call clean_up_curr_level
    
    Exit Sub
        
HandleError:
    
    modErrInfo__.FuncID = id_BreakItr__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext

End Sub


Private Function clean_up_curr_level() As Boolean

    Dim empty_itr As Itr_t
    
    If m_Itr(m_CurrItrLevel).ItrType = fmlItrCSV Then
    
        If m_FileIn(m_Itr(m_CurrItrLevel).FileItr).IsOpen Then
        
            Close #m_FileIn(m_Itr(m_CurrItrLevel).FileItr).FileNo
            m_FileIn(m_Itr(m_CurrItrLevel).FileItr).IsOpen = False
        
        End If
        
        m_Itr(m_CurrItrLevel).FileItr = 0
    
    End If
    
'    m_Itr(m_CurrItrLevel).ItrType = fmlItrNull
'    m_Itr(m_CurrItrLevel).NextItr = Empty
'    m_Itr(m_CurrItrLevel).StartItr = Empty
'    m_Itr(m_CurrItrLevel).EndItr = Empty
'    m_Itr(m_CurrItrLevel).ArrayItr = Empty
'    m_Itr(m_CurrItrLevel).SkipListItr = Empty
'    m_Itr(m_CurrItrLevel).IsEndOfItr = False
    
    m_Itr(m_CurrItrLevel) = empty_itr
    
    'm_CurrentSkipListCount = 0
    m_CurrItrLevel = m_CurrItrLevel - 1

    clean_up_curr_level = True

End Function



Public Sub SetItrSkip(ParamArray Index() As Variant)

    Dim skip_list_size As Long
    Dim skip_list As Variant
    Dim skip_flag As Boolean
    Dim i As Long
    Dim flag As Boolean
    Dim skip_list_count As Long
        
    On Error GoTo HandleError:
        
    'Error if there's no iteration.
    If m_CurrItrLevel < 1 Then
    
        Err.Raise Number:=fmlWrongContext, _
          Source:=PROJ_NAME, _
          Description:=modErrInfo__.errstr_WrongContext + "Outside iterations."
          
    
    'Check the given length is the same as the current iterator.
    ElseIf UBound(Index) - LBound(Index) + 1 _
        <> UBound(m_Itr(m_CurrItrLevel).NextItr) _
        - LBound(m_Itr(m_CurrItrLevel).NextItr) + 1 Then
    
        Err.Raise Number:=fmlInvalidArgument, _
                  Source:=PROJ_NAME, _
                  Description:=modErrInfo__.errstr_InvalidArgument + "Index"
    End If

    skip_list_count = m_Itr(m_CurrItrLevel).SkipListCount
    
    If skip_list_count = 0 Then
    
        ReDim skip_list(0 To SKIP_LIST_MAX_SIZE) As Variant
        m_Itr(m_CurrItrLevel).SkipListItr = skip_list
    
    End If
    
    For i = 0 To UBound(Index)
    
        m_Itr(m_CurrItrLevel).SkipListItr(skip_list_count + i) = Index(i)
        
    Next i
    
    m_Itr(m_CurrItrLevel).SkipListCount = skip_list_count + UBound(Index) + 1
        
    Exit Sub

HandleError:
    
    modErrInfo__.FuncID = id_SetItrSkip__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext

End Sub



Public Function GetItr(ParamArray Index() As Variant) As Boolean

    Dim ind_buf As Variant
    Dim i As Long

    On Error GoTo HandleError:


    If UBound(Index) < LBound(Index) Then
    
        GetItr = get_itr(arg_is_paramarray)
    
    Else
        
        ind_buf = Index
        GetItr = get_itr(arg_is_paramarray, ind_buf)
        
        For i = 0 To UBound(ind_buf)

            Index(i) = ind_buf(i)

        Next i
        
    End If

Exit Function
                
HandleError:
    
    modErrInfo__.FuncID = id_GetItr__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext
        

End Function

Public Function GetItrA(Optional Index As Variant) As Boolean

    
    On Error GoTo HandleError:

    If IsMissing(Index) Then
    
        GetItrA = get_itr(arg_is_variant)
    
    Else
        
        GetItrA = get_itr(arg_is_variant, Index)
        
    End If

Exit Function
                
HandleError:
    
    modErrInfo__.FuncID = id_GetItrA__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext


End Function


Private Function get_itr(arg_type As array_arg_type, Optional Index As Variant) As Boolean

    Dim skip_flag As Boolean
    Dim ind_len As Long, ub_ind As Long, lb_ind As Long
    Dim i As Long
    Dim lb As Long
    
    Dim next_itr As Variant     'local version of m_Itr(m_CurrItrLevel).NextItr
    Dim is_end_of_itr As Boolean 'local version of m_Itr(m_CurrItrLevel).IsEndOfItr
    
                    
    get_itr = False
    
    'Error if there's no iteration.
    If m_CurrItrLevel < 1 Then
    
        Err.Raise Number:=fmlWrongContext, _
          Source:=PROJ_NAME, _
          Description:=modErrInfo__.errstr_WrongContext + "Outside iteration."

    End If
        
    'Copy to local vars
    next_itr = m_Itr(m_CurrItrLevel).NextItr
    is_end_of_itr = m_Itr(m_CurrItrLevel).IsEndOfItr
    
    Do
        'Check if End of Itr
        If is_end_of_itr Then
        
            get_itr = False
            Exit Function
            
        End If
        
        'Check whether to skip new counter
        skip_flag = False
        If Not IsEmpty(m_Itr(m_CurrItrLevel).SkipListItr) Then
            
            'Check whether to skip new counter
            If m_Itr(m_CurrItrLevel).ItrType = fmlItrBoundary Then
            
                skip_flag = skip_flag Or modUtil__.compare_index(next_itr, m_Itr(m_CurrItrLevel).SkipListItr(i))
                
            ElseIf m_Itr(m_CurrItrLevel).ItrType = fmlItrArray Then
            
                skip_flag = skip_flag Or modUtil__.compare_index(m_Itr(m_CurrItrLevel).ArrayItr(next_itr), m_Itr(m_CurrItrLevel).SkipListItr(i))
            
            End If
            
        End If

        'Check Arg
        If IsMissing(Index) Then
            ind_len = 0
        
        ElseIf arg_type = arg_is_paramarray Then
            
            'Get Index
            ub_ind = UBound(Index)
            lb_ind = LBound(Index)
            ind_len = ub_ind - lb_ind + 1 'Could be 0
            
        Else    'Create Index Array
        
            If m_Itr(m_CurrItrLevel).ItrType = fmlItrBoundary Then

                ReDim Index(0 To UBound(m_Itr(m_CurrItrLevel).StartItr))
                
            ElseIf m_Itr(m_CurrItrLevel).ItrType = fmlItrArray Then
            
                ReDim Index(LBound(m_Itr(m_CurrItrLevel).ArrayItr, 2) To UBound(m_Itr(m_CurrItrLevel).ArrayItr, 2))
                
            Else
                Err.Raise Number:=fmlInvalidArgument, _
                          Source:=PROJ_NAME, _
                          Description:=modErrInfo__.errstr_InvalidArgument + "IterationType"
            End If
            
            'Get Index
            ub_ind = UBound(Index)
            lb_ind = LBound(Index)
            ind_len = ub_ind - lb_ind + 1 'Could be 0
            
        End If
        
        
        'Get index
        'This part should be similar to that in NextItr/NextItrA.
        If ind_len > 0 And Not skip_flag Then
        
            If m_Itr(m_CurrItrLevel).ItrType = fmlItrBoundary Then
            
                If ub_ind > UBound(m_Itr(m_CurrItrLevel).StartItr) Then
                
                    Err.Raise Number:=fmlInvalidArgument, _
                              Source:=PROJ_NAME, _
                              Description:=modErrInfo__.errstr_InvalidArgument + "Index length is too long."
                End If
                    
                For i = 0 To ub_ind
            
                    Index(i) = next_itr(i)
                
                Next i
                
            ElseIf m_Itr(m_CurrItrLevel).ItrType = fmlItrArray Then
            
                
                lb = LBound(m_Itr(m_CurrItrLevel).ArrayItr, 2)
                
                If ind_len > UBound(m_Itr(m_CurrItrLevel).ArrayItr, 2) - lb + 1 Then
                
                    Err.Raise Number:=fmlInvalidArgument, _
                              Source:=PROJ_NAME, _
                              Description:=modErrInfo__.errstr_InvalidArgument + "Index length is too long."
                End If
        
                For i = lb_ind To ub_ind
                
                    Index(i) = m_Itr(m_CurrItrLevel).ArrayItr(next_itr, i + lb - lb_ind)
                    
                Next i
            
            Else
            
                Err.Raise Number:=fmlInvalidArgument, _
                          Source:=PROJ_NAME, _
                          Description:=modErrInfo__.errstr_InvalidArgument + "IterationType"
                
            End If
        
        End If
        'End of Getting Index
        
        'Increment next_itr
        
    Loop While skip_flag

    get_itr = True
    
                
End Function


Public Property Get ItrLevel() As Long

    ItrLevel = m_CurrItrLevel

End Property


Public Property Get ItrCount(Optional ItrLevel As Variant) As Long

    Dim itr_lvl As Long
    
    If IsMissing(ItrLevel) Then itr_lvl = m_CurrItrLevel Else itr_lvl = ItrLevel

    If itr_lvl > 0 And itr_lvl <= m_CurrItrLevel Then

        ItrCount = m_Itr(itr_lvl).ItrCount / m_Itr(itr_lvl).ItrCountUnit

    Else
    
        ItrCount = 0
    
    End If

End Property

Private Function get_nextitr_str() As String
'Used inside Error Handler. Must not throw.

    Dim i As Long

    If m_CurrItrLevel = 0 Then
    
        get_nextitr_str = "N/A"
        
    ElseIf m_Itr(m_CurrItrLevel).IsEndOfItr Then
    
        get_nextitr_str = "End Of Iteration"
    
    ElseIf IsNumeric(m_Itr(m_CurrItrLevel).NextItr) Then
    
        get_nextitr_str = CStr(m_Itr(m_CurrItrLevel).NextItr)
    
    ElseIf IsArray(m_Itr(m_CurrItrLevel).NextItr) Then
    
        get_nextitr_str = CStr(m_Itr(m_CurrItrLevel).NextItr(LBound(m_Itr(m_CurrItrLevel).NextItr)))
    
        For i = LBound(m_Itr(m_CurrItrLevel).NextItr) + 1 To UBound(m_Itr(m_CurrItrLevel).NextItr)
        
            If IsNumeric(m_Itr(m_CurrItrLevel).NextItr(i)) Then
        
                get_nextitr_str = get_nextitr_str & "," & CStr(m_Itr(m_CurrItrLevel).NextItr(i))
                
            Else
            
                get_nextitr_str = get_nextitr_str & ", Unknown"
            
            End If
        
        Next i
    
    Else
    
        get_nextitr_str = "Unknown"
    
    End If
    
End Function

'Public Function PrefetchItr(ByVal PrefetchCount As Long, Index As Variant) As Boolean
'
'    Dim skip_flag As Boolean
'    Dim skip_list_count As Long
'    Dim skip_list_ub As Long, skip_list_lb As Long
'    Dim i As Long
'
'    Dim curr_itr As Variant, prev_itr As Variant
'
'    On Error GoTo HandleError:
'
'    prefetch_itr = False
'    curr_itr = m_Itr(m_CurrItrLevel).NextItr
'    prev_itr = curr_itr
'
'    Do While PrefetchCount >= 0
'
'        Do
'            'Increment Current Iterator
'            If m_Itr(m_CurrItrLevel).ItrType = fmlItrBoundary Then
'
'                prefetch_itr = modUtil__.increment_index(curr_itr, m_Itr(m_CurrItrLevel).StartItr, m_Itr(m_CurrItrLevel).EndItr)
'
'            ElseIf m_Itr(m_CurrItrLevel).ItrType = fmlItrArray Then
'
'                If curr_itr = m_Itr(m_CurrItrLevel).EndItr Then
'
'                    prefetch_itr = False
'
'                Else
'                    curr_itr = curr_itr + 1
'                    prefetch_itr = True
'
'                End If
'
'            End If
'
'            If Not prefetch_itr Then Exit Do
'
'            'Check whether to skip new counter
'            If IsEmpty(m_Itr(m_CurrItrLevel).SkipListItr) Then
'
'                skip_flag = False
'
'            Else
'                skip_flag = False
'                For i = 0 To m_CurrentSkipListCount - 1
'
'                    skip_flag = skip_flag Or modUtil__.compare_index(curr_itr, m_Itr(m_CurrItrLevel).SkipListItr(i))
'
'                Next i
'
'            End If
'
'        Loop While skip_flag
'
'        If Not prefetch_itr Then
'
'            Exit Do
'
'        Else
'            prev_itr = curr_itr
'
'        End If
'
'        PrefetchCount = PrefetchCount - 1
'
'    Loop
'
'    Index = prev_itr
'
'    Exit Function
'
'HandleError:
'
'    modErrInfo__.FuncID = id_PrefetchItr__
'    Err.Raise Number:=Err.Number, _
'                Source:=Err.Source, _
'                Description:=Err.Description, _
'                HelpFile:=Err.HelpFile, _
'                HelpContext:=Err.HelpContext
'
'End Function


Private Sub run_time_report()

    Dim i As Long
    Dim total_run_time As Double
    
    Debug.Print "=========== Formula Modeler ランログ ==========="
    Debug.Print "バージョン: " + VERSION_STRING
    Debug.Print "ライセンス: " + "BSD 2-Clause License(Open Source)"
    'Debug.Print "使用期限: " + CStr(LICENSE_EXPIRE_ON) + " (" + CStr(LICENSE_EXPIRE_ON - Date) + "日)"
    Debug.Print "開始時刻: " + CStr(m_StartTime)
    
    For i = 0 To MAX_FORMULA_ID
    
        If m_Formula(i).Exist Then

            total_run_time = total_run_time + m_Formula(i).RunTime
            
        End If
        
    Next i
    
    Debug.Print "合計ランタイム: " + Format(total_run_time, "0.0") + "秒"
    Debug.Print "------ フォーミュラ毎ランログ ------"
    Debug.Print "ID: 呼出回数, ランタイム(秒)"
    
    For i = 0 To MAX_FORMULA_ID
    
        If m_Formula(i).Exist Then

'            Debug.Print "Formula " + Format(i, "#0") + ": 呼出回数 " + Format(m_Formula(i).CallCount, "######0") + ", ランタイム " + Format(m_Formula(i).RunTime, "0.0")
            Debug.Print Format(i, "00: ") + Format(m_Formula(i).CallCount, "######0") + ", " + Format(m_Formula(i).RunTime, "0.0")
            
        End If
    
    Next i

End Sub


Private Sub add_time()

    Dim curr_sec As Double
    Dim curr_date As Long
    Dim date_count As Long

    'Run Time Tracking
    curr_sec = CDbl(Timer)
    curr_date = CLng(Date)
    
    date_count = curr_date - m_CallStack(m_CurrCall).TimeStampDate
    
    'Add time
    If date_count = 0 Then
    
        m_Formula(m_CallStack(m_CurrCall).ID).RunTime _
            = m_Formula(m_CallStack(m_CurrCall).ID).RunTime + curr_sec - m_CallStack(m_CurrCall).TimeStampSec
    
    Else    'Past midnight
    
        m_Formula(m_CallStack(m_CurrCall).ID).RunTime _
            = m_Formula(m_CallStack(m_CurrCall).ID).RunTime + curr_sec - m_CallStack(m_CurrCall).TimeStampSec _
            + date_count * 24 * 60 * 60
    
    End If

End Sub


Public Property Let FileInput(FileID As Long, Optional ByVal SkipCount As Long = 0, Optional ByVal FieldLBound As Long = 0, Optional FieldTypes As Variant, FilePath As String)


    On Error GoTo HandleError:


    If m_FileInUsed(FileID) Then
    
        Err.Raise Number:=fmlFileError, _
          Source:=PROJ_NAME, _
          Description:=modErrInfo__.errstr_FileError + "File ID is already assigned."
        
    Else
    
        m_FileIn(FileID).FilePath = FilePath
        m_FileInUsed(FileID) = True
        m_FileIn(FileID).FieldLBound = FieldLBound
        m_FileIn(FileID).SkipCount = SkipCount
    
    End If
    
    If Not IsMissing(FieldTypes) Then
    
        set_field_types FileID, FieldTypes
        
    End If
    
    Exit Property

HandleError:
    
    modErrInfo__.FuncID = id_FileInput__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext

End Property

Private Sub set_field_types(file_id As Long, field_types As Variant)
    
    Dim i As Long
    Dim ft As Variant
    Dim ft_len As Long
    Dim field_types_buf As Variant
    Dim field_types_lb As Long, field_types_ub As Long
    Dim type_id As VbVarType
    
    If Not IsArray(field_types) Then
    
        Err.Raise Number:=fmlInvalidArgument, _
          Source:=PROJ_NAME, _
          Description:=modErrInfo__.errstr_InvalidArgument + "Invalid Field Types."
    
    End If
    
    field_types_lb = LBound(field_types)
    field_types_ub = UBound(field_types)
    
    ReDim field_types_buf(0 To field_types_ub - field_types_lb) 'Shift Index
    
    i = 0
    For Each ft In field_types
    
        If Not UBound(ft) - LBound(ft) = 1 Then
        
            Err.Raise Number:=fmlInvalidArgument, _
              Source:=PROJ_NAME, _
              Description:=modErrInfo__.errstr_InvalidArgument + "Invalid Field Types."
        
        End If
        
        If Not modUtil__.is_whole_number(ft(LBound(ft))) Then
        
            Err.Raise Number:=fmlInvalidArgument, _
              Source:=PROJ_NAME, _
              Description:=modErrInfo__.errstr_InvalidArgument + "Invalid Field Types."
            
        End If
        
        type_id = ft(UBound(ft))
        
        If Not (type_id = vbBoolean _
                Or type_id = vbDate _
                Or type_id = vbDouble _
                Or type_id = vbString _
                Or type_id = vbLong) Then
                
            Err.Raise Number:=fmlInvalidArgument, _
              Source:=PROJ_NAME, _
              Description:=modErrInfo__.errstr_InvalidArgument + "Field Type Not Supported."
              
              
        End If
        
        field_types_buf(i) = Array(ft(LBound(ft)), type_id)
        i = i + 1
    
    Next ft
    

    m_FileIn(file_id).FieldTypes = field_types_buf

End Sub


Public Property Let FileOutput(FileID As Long, FilePath As String)

    On Error GoTo HandleError:

    If m_FileOutUsed(FileID) Then
    
        Err.Raise Number:=fmlFileError, _
          Source:=PROJ_NAME, _
          Description:=modErrInfo__.errstr_FileError + "File ID is already assigned."
        
    Else
    
        m_FileOut(FileID).FilePath = FilePath
        m_FileOutUsed(FileID) = True
    
    End If

    Exit Property

HandleError:
    
    modErrInfo__.FuncID = id_FileOutput__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext

End Property



Private Function var2file(file_num As Long, input_var As Variant) As Boolean

    Dim row_size As Long, col_size As Long
    Dim dim_size As Long
    Dim out_var As Variant
    Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long
    Dim i As Long, j As Long
    Dim temp As String
    
    '--- Single Value ---
    If Not IsArray(input_var) Then
            
        Print #file_num, modUtil__.var2line(input_var)
        var2file = True
        Exit Function
    
    End If
    
    dim_size = modArrSpt__.NumberOfArrayDimensions(input_var)
    
    Select Case dim_size
        
    Case 1
            
        Print #file_num, modUtil__.var2line(input_var)
        var2file = True
        Exit Function
    
    Case 2
        
        lb1 = LBound(input_var)
        ub1 = UBound(input_var)
        lb2 = LBound(input_var, 2)
        ub2 = UBound(input_var, 2)
        
        ReDim out_var(lb2 To ub2)
        
        For i = lb1 To ub1
        
            For j = lb2 To ub2
            
                out_var(j) = input_var(i, j)
        
            Next j
                        
            Print #file_num, modUtil__.var2line(out_var)
            
        Next i
               
        var2file = True
        Exit Function
        
    Case Else
    
        var2file = False
        Exit Function
        
    End Select
    
    
End Function

Public Sub VarToFile(FileID As Long, ParamArray Source())    'Cannot be a function due to VBA bug
    
    Dim file_num As Long
    Dim v As Variant
    
    On Error GoTo HandleError:
    
    If UBound(Source) < LBound(Source) Then
    
        Exit Sub
    
    End If
    
    '--- Open if not opened ---
    If Not m_FileOut(FileID).IsOpen Then
    
        m_FileOut(FileID).FileNo = FreeFile
        Open m_FileOut(FileID).FilePath For Output Access Write As #m_FileOut(FileID).FileNo
                
        m_FileOut(FileID).IsOpen = True
    
    End If
    
    file_num = m_FileOut(FileID).FileNo
 
    For Each v In Source
 
        If Not var2file(file_num, v) Then
                         
            Err.Raise Number:=fmlFileError, _
              Source:=PROJ_NAME, _
              Description:=modErrInfo__.errstr_FileError
        
        End If
        
    Next v
        
    Exit Sub
    
HandleError:

    Close
    modErrInfo__.FuncID = id_VarToFile__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext
                
End Sub

Public Sub VecToLine(FileID As Long, Header As Variant, Source As Variant)
    
    Dim file_num As Long
    Dim hdr_len As Long, src_len As Long
    Dim v As Variant
    Dim i As Long
    
    On Error GoTo HandleError:
    
    '--- Check if Source is a vector ---
    
    If IsArray(Source) Then
    
        src_len = UBound(Source) - LBound(Source) + 1
    
        If src_len <= 0 Then
        
            Err.Raise Number:=fmlInvalidArgument, _
              Source:=PROJ_NAME, _
              Description:=modErrInfo__.errstr_InvalidArgument + "Source must be a vector or value."
            
        ElseIf modArrSpt__.NumberOfArrayDimensions(Source) > 1 Then
        
            Err.Raise Number:=fmlInvalidArgument, _
              Source:=PROJ_NAME, _
              Description:=modErrInfo__.errstr_InvalidArgument + "Source must be a vector or value."
                
        End If
        
    Else
        
        src_len = 1
                
    End If
    
    '--- Check if Header is a vector ---
    
    If IsArray(Header) Then
    
        hdr_len = UBound(Header) - LBound(Header) + 1
    
        If hdr_len <= 0 Then
        
            Err.Raise Number:=fmlInvalidArgument, _
              Source:=PROJ_NAME, _
              Description:=modErrInfo__.errstr_InvalidArgument + "Header must be a vector or value."
            
        ElseIf modArrSpt__.NumberOfArrayDimensions(Header) > 1 Then
        
            Err.Raise Number:=fmlInvalidArgument, _
              Source:=PROJ_NAME, _
              Description:=modErrInfo__.errstr_InvalidArgument + "Header must be a vector or value."
                
        End If
        
    ElseIf IsEmpty(Header) Or IsNull(Header) Then
    
        hdr_len = 0
        
    Else
    
        hdr_len = 1
    
    End If
    
    '--- Open if not opened ---
    
    If Not m_FileOut(FileID).IsOpen Then
    
        m_FileOut(FileID).FileNo = FreeFile
        Open m_FileOut(FileID).FilePath For Output Access Write As #m_FileOut(FileID).FileNo
                
        m_FileOut(FileID).IsOpen = True
    
    End If
    
    file_num = m_FileOut(FileID).FileNo
 
 
    '--- Create Header + Source array ---

    ReDim v(0 To hdr_len + src_len - 1)
    
    '--- Copy Header ---
    
    If IsArray(Header) Then
    
        For i = LBound(Header) To UBound(Header)
        
            v(i - LBound(Header)) = Header(i)
        
        Next i
    
    ElseIf IsEmpty(Header) Or IsNull(Header) Then
    
        'Do Nothing
    
    Else 'Header is a value.
    
        v(0) = Header
        
    End If

    '--- Copy Source ---
    
    If IsArray(Source) Then
    
        For i = LBound(Source) To UBound(Source)
        
            v(i - LBound(Source) + hdr_len) = Source(i)
        
        Next i
    
    Else 'Header is a value.
    
        v(hdr_len) = Source
        
    End If
    
     
    '--- Output to file ---
    
    If Not var2file(file_num, v) Then
                     
        Err.Raise Number:=fmlFileError, _
          Source:=PROJ_NAME, _
          Description:=modErrInfo__.errstr_FileError
    
    End If
        
        
    Exit Sub
    
HandleError:

    Close
    modErrInfo__.FuncID = id_VecToLine__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext
                
End Sub


Public Property Get RunParam() As Variant

    RunParam = m_Param

End Property


Public Property Let RetVal(v As Variant)

    m_RetVal = v

End Property



