Attribute VB_Name = "Utility"
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


Public Enum fmlRangeDim    'TODO: Come up with Better naming

    fmlRow
    fmlCol

End Enum

Public Enum fmlCompareOp

    fmlEqualTo '= 0 * 2 ^ 1 + 1 * 2 ^ 0
    fmlLessThan '= 1 * 2 ^ 1 + 0 * 2 ^ 0
    fmlNoMoreThan '= 1 * 2 ^ 1 + 1 * 2 ^ 0
    
End Enum

Public Enum fmlFindOption

    fmlFirstAsMoreThan = 1
    fmlFirstAsEqualTo = 2
    fmlLastAsElse = 4

End Enum


'Public Enum fmlArrayMappingType
'
'    fmlMapIndex
'    fmlMapUBound
'
'End Enum

Public Sub SumVarArray(Target As Variant, Source As Variant, Optional Parameter As Variant = 1, Optional ByVal Index1 As Variant = Empty, Optional ByVal Index2 As Variant = Empty)

    Dim trg_lb As Long, trg_ub As Long
    Dim src_lb As Long, src_ub As Long
    Dim lb As Long, ub As Long
    
    Dim trg_elm_lb As Long, trg_elm_ub As Long
    Dim src_elm_lb As Long, src_elm_ub As Long
    Dim elm_lb As Long, elm_ub As Long
    
    Dim i As Long, j As Long

    On Error GoTo HandleError:
     
    src_lb = LBound(Source)
    src_ub = UBound(Source)
    
    If IsEmpty(Index1) Then
                
        lb = src_lb
        ub = src_ub
        
    Else
    
        lb = Index1(LBound(Index1))
        ub = Index1(LBound(Index1) + 1)
       
    End If
    
    
    If IsEmpty(Target) Then
            
        ReDim Target(lb To ub)
          
        If IsEmpty(Index2) Then
        
            For i = LBound(Target) To UBound(Target)
            
                create_new_array Target(i), Array(LBound(Source(i))), Array(UBound(Source(i)))
        
                For j = LBound(Source(i)) To UBound(Source(i))
        
                    Target(i)(j) = Parameter * Source(i)(j)
                    
                Next j
            
            Next i
                        
        Else
        
            trg_elm_lb = Index2(LBound(Index2))
            trg_elm_ub = Index2(LBound(Index2) + 1)
        
            For i = LBound(Target) To UBound(Target)
        
                create_new_array Target(i), Array(trg_elm_lb), Array(trg_elm_ub)
            
                For j = LBound(Source(i)) To UBound(Source(i))
        
                    Target(i)(j) = Parameter * Source(i)(j)
                    
                Next j
        
            Next i
            
        End If
        
        
    ElseIf IsArray(Target) Then
    
        trg_lb = LBound(Target)
        trg_ub = UBound(Target)
    
        
        If trg_lb > lb Or trg_ub < ub Then
        
            Err.Raise Number:=fmlInvalidArgument, _
                      Source:=PROJ_NAME, _
                      Description:=modErrInfo__.errstr_InvalidArgument + "Souce LBound and Ubound must be within that of Target."
        
        End If
        
        If IsEmpty(Index2) Then
                                     
            For i = lb To ub
            
                trg_elm_lb = LBound(Target(i))
                trg_elm_ub = UBound(Target(i))
                
                src_elm_lb = LBound(Source(i))
                src_elm_ub = UBound(Source(i))
                
                If src_elm_lb < trg_elm_lb Then
                
                    Err.Raise Number:=fmlInvalidArgument, _
                              Source:=PROJ_NAME, _
                              Description:=modErrInfo__.errstr_InvalidArgument + "Souce element's LBound must be no smaller than Target element's  Lbound."
                
                End If
                
                If trg_elm_ub < src_elm_ub Then
                
                    modUtil__.redim_nested_array Target(i), src_elm_ub, True
                            
                End If
                
                For j = src_elm_lb To src_elm_ub
                
                    Target(i)(j) = Target(i)(j) + Parameter * Source(i)(j)
                
                Next j
            
            Next i
            
        Else
        
            elm_lb = Index2(LBound(Index2))
            elm_ub = Index2(LBound(Index2) + 1)
            
            For i = lb To ub
            
                trg_elm_lb = LBound(Target(i))
                trg_elm_ub = UBound(Target(i))
                
                src_elm_lb = LBound(Source(i))
                src_elm_ub = UBound(Source(i))
                
                If elm_lb < trg_elm_lb Then
                
                    Err.Raise Number:=fmlInvalidArgument, _
                              Source:=PROJ_NAME, _
                              Description:=modErrInfo__.errstr_InvalidArgument + "Souce element's LBound must be no smaller than Target element's  Lbound."
                
                End If
                
                If trg_elm_ub < elm_ub Then
                
                    modUtil__.redim_nested_array Target(i), elm_ub, True
                            
                End If
                
                For j = elm_lb To elm_ub
                
                    Target(i)(j) = Target(i)(j) + Parameter * Source(i)(j)
                
                Next j
            
            Next i
        
        End If
            
    Else
    
        Err.Raise Number:=fmlInvalidArgument, _
                  Source:=PROJ_NAME, _
                  Description:=modErrInfo__.errstr_InvalidArgument + "Target must be Array or Empty"
        
    End If
    
    Exit Sub
    
HandleError:
    
    modErrInfo__.FuncID = id_SumVarArraySub__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext
End Sub




Public Sub NewJaggedArray(Target As Variant, Source As Variant)
    
    Dim src_buf As Variant
    Dim table_rows As Long, table_cols As Long
    Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long
    Dim row As Long, col As Long
    Dim ind As Variant
    
    On Error GoTo HandleError:

    src_buf = Source
    
    If Not modArrSpt__.IsArrayAllocated(src_buf) Or _
        modArrSpt__.NumberOfArrayDimensions(src_buf) <> 2 Then
    
        Err.Raise Number:=fmlInvalidArgument, _
                  Source:=PROJ_NAME, _
                  Description:=modErrInfo__.errstr_InvalidArgument + "Source"
                                                
    End If

    lb1 = LBound(src_buf, 1)
    ub1 = UBound(src_buf, 1)
    lb2 = LBound(src_buf, 2)
    ub2 = UBound(src_buf, 2)
    
    table_rows = ub1 - lb1 + 1
    table_cols = ub2 - lb2 + 1
    
    ReDim ind(0 To table_cols - 1)
    
    For row = lb1 To ub1
        
        For col = lb2 To ub2
            
            ind(col - lb2) = src_buf(row, col)
            
        Next col
    
        set_to_nested_array Target, ind, Empty
    
    Next row
            
    Exit Sub
        
HandleError:
    
    modErrInfo__.FuncID = id_NewJaggedArray__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext

End Sub

Public Sub NewLookupTable(Target As Variant, Source As Variant, KeyLength As Long, Optional VarLBound As Long = 0)

'/// Lookup Tabel Internal Implementation
'/// Internal representation of LookupTable
'/// A LookupTable is implemented as a jagged array, i.e. nested 1 dim arrays.
'/// The first array is a list of variables.
'/// For each element of the first array, an array tree each of whose nodes are Array(0, 1).

    Dim src_buf As Variant
    Dim table_rows As Long, table_cols As Long
    Dim var_ind As Long, key_ind As Long, key_elm_ind As Long
    Dim var_count As Long
    Dim non_empty_key_count As Long
    Dim internal_key As Variant, internal_key_head As Variant, internal_key_body As Variant
    Dim extl_key As Variant
    Dim i As Long, j As Long
    Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long
    Dim test As Boolean
    
    On Error GoTo HandleError:
    
    src_buf = Source
    
    If Not modArrSpt__.IsArrayAllocated(src_buf) Or _
        modArrSpt__.NumberOfArrayDimensions(src_buf) <> 2 Then
    
        Err.Raise Number:=fmlInvalidArgument, _
                  Source:=PROJ_NAME, _
                  Description:=modErrInfo__.errstr_InvalidArgument + "Source"
                                                
    End If

    lb1 = LBound(src_buf, 1)
    ub1 = UBound(src_buf, 1)
    lb2 = LBound(src_buf, 2)
    ub2 = UBound(src_buf, 2)
    
    table_rows = ub1 - lb1 + 1
    table_cols = ub2 - lb2 + 1
    
    If KeyLength < 0 Or KeyLength >= table_cols Then
    
         Err.Raise Number:=fmlInvalidArgument, _
                  Source:=PROJ_NAME, _
                  Description:=modErrInfo__.errstr_InvalidArgument + "KeyLength"
    
    End If
    
    var_count = table_cols - KeyLength
    ReDim Target(VarLBound To VarLBound + var_count - 1) As Variant
    
    For var_ind = VarLBound To VarLBound + var_count - 1
        
        For key_ind = lb1 To ub1
        
            If modUtil__.is_numeric(src_buf(key_ind, var_ind - VarLBound + lb2 + KeyLength)) Then
                            
                            
                ReDim extl_key(0 To KeyLength - 1)
                
                For i = 0 To KeyLength - 1
                
                    extl_key(i) = src_buf(key_ind, lb2 + i)
                
                Next i
                
                get_internal_key internal_key, extl_key, KeyLength
                If Not set_to_nested_array(Target(var_ind), internal_key, src_buf(key_ind, lb2 + KeyLength + var_ind - VarLBound)) Then
                
                    Err.Raise Number:=fmlInvalidArgument, _
                             Source:=PROJ_NAME, _
                             Description:=modErrInfo__.errstr_InvalidArgument + "変数Index" + CStr(var_ind) + "に重複したキー(行Index" + CStr(key_ind) + ")があります。"
                
                
                End If
            
            End If
            
        Next key_ind
    
    Next var_ind
    
    Exit Sub

HandleError:
    
    modErrInfo__.FuncID = id_NewLookupTable__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext

End Sub



Private Function get_internal_key(internal_key As Variant, extl_key As Variant, key_len As Long) As Boolean

    Dim internal_key_head As Variant, internal_key_body As Variant
    Dim non_empty_key_count As Long
    Dim lb As Long
    Dim i As Long, j As Long


    ReDim internal_key_head(0 To key_len - 1) As Variant
    ReDim internal_key_body(0 To key_len - 1) As Variant

    lb = LBound(extl_key) 'Must be always 0.
    non_empty_key_count = 0

    'Create for internal key head and body.
    For i = lb To lb + key_len - 1

        'Integer in a Excel cell is stored as a double number in Variant array.
        If modUtil__.is_whole_number(extl_key(i)) Then
        
            non_empty_key_count = non_empty_key_count + 1
            internal_key_head(i - lb) = 1
            internal_key_body(i - lb) = extl_key(i)
            
        Else
            internal_key_head(i - lb) = 0
        
        End If
    
    Next i
    
    ReDim internal_key(0 To key_len + non_empty_key_count - 1)
    
    'Create internal key.
    j = key_len
    For i = 0 To key_len - 1
        
        If internal_key_head(i) = 1 Then
        
            internal_key(i) = 1
            internal_key(j) = internal_key_body(i)
            j = j + 1
            
        Else
            internal_key(i) = 0

        End If
        
    Next i


End Function


Private Function set_to_nested_array(targ_array As Variant, Key As Variant, arg_val As Variant) As Boolean

    Dim key_elm As Long
    Dim key_len As Long
    Dim next_key As Variant
    Dim i As Long
    Dim ub As Long, lb As Long
    Dim new_array As Variant
    Dim key_elm_exists As Boolean
    
    key_len = UBound(Key) - LBound(Key) + 1
    key_elm = Key(LBound(Key))
    
    If Not modArrSpt__.IsArrayAllocated(targ_array) Then
    
        ReDim targ_array(key_elm To key_elm) As Variant
    
    ElseIf key_elm < LBound(targ_array) Then
    
        lb = LBound(targ_array)
        ub = UBound(targ_array)
    
        ReDim new_array(key_elm To ub) As Variant
        
        For i = lb To ub
        
            new_array(i) = targ_array(i)
            
        Next i
        
        new_array(key_elm) = arg_val
        targ_array = new_array
        
    
    ElseIf UBound(targ_array) < key_elm Then
        
        ReDim Preserve targ_array(LBound(targ_array) To key_elm)
        
    Else
    
        key_elm_exists = True
    
    End If
        
    
    If key_len = 1 Then
    
        If key_elm_exists Then
        
            set_to_nested_array = False
            Exit Function
                     
        End If

        targ_array(key_elm) = arg_val
        set_to_nested_array = True
        Exit Function
        
    Else
        
        ReDim next_key(LBound(Key) + 1 To UBound(Key)) As Variant
        For i = LBound(Key) + 1 To UBound(Key)
    
            next_key(i) = Key(i)
        
        Next i
        
        set_to_nested_array = set_to_nested_array(targ_array(key_elm), next_key, arg_val)
        
    End If
        
    'set_to_nested_array = False
  

End Function



Public Function LookupExact(Result As Variant, Table As Variant, ByVal VarIndex As Long, ParamArray Key() As Variant) As Boolean

    Dim key_len As Long
    Dim i As Long
    Dim internal_key As Variant
    Dim external_key As Variant

    On Error GoTo HandleError:

    If UBound(Key) < LBound(Key) Then
    
        Err.Raise Number:=fmlInvalidArgument, _
                 Source:=PROJ_NAME, _
                 Description:=modErrInfo__.errstr_InvalidArgument + "Key"
        
    ElseIf VarIndex < LBound(Table) Or VarIndex > UBound(Table) Then
        
        Err.Raise Number:=fmlInvalidArgument, _
                 Source:=PROJ_NAME, _
                 Description:=modErrInfo__.errstr_InvalidArgument + "VarIndex " + CStr(VarIndex)
                
    End If
        
    external_key = Key
    key_len = UBound(Key) - LBound(Key) + 1
    get_internal_key internal_key, external_key, key_len
    LookupExact = get_from_nested_array(Result, Table(VarIndex), internal_key)
        
    Exit Function

HandleError:
    
    modErrInfo__.FuncID = id_LookupExact__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext

End Function


Private Function LookupMatchBAK(Result As Variant, Table As Variant, ByVal VarIndex As Long, ParamArray Key() As Variant) As Boolean

    Dim key_len As Long
    Dim i As Long
    Dim internal_key As Variant
    Dim external_key As Variant
    
    Dim key_lb As Long, key_ub As Long

    On Error GoTo HandleError:
    
    key_ub = UBound(Key)
    key_lb = LBound(Key)

    If key_ub < key_lb Then
    
        Err.Raise Number:=fmlInvalidArgument, _
                 Source:=PROJ_NAME, _
                 Description:=modErrInfo__.errstr_InvalidArgument + "Key"
        
    ElseIf VarIndex < LBound(Table) Or VarIndex > UBound(Table) Then
        
        Err.Raise Number:=fmlInvalidArgument, _
                 Source:=PROJ_NAME, _
                 Description:=modErrInfo__.errstr_InvalidArgument + "VarIndex " + CStr(VarIndex)
                
    End If
        
    key_len = key_ub - key_lb + 1
    external_key = Key
    get_internal_key internal_key, external_key, key_len
    LookupMatchBAK = get_from_nested_array(Result, Table(VarIndex), internal_key)

    If LookupMatchBAK Then Exit Function

    i = key_ub
    
    Do While i >= key_lb
    
        If modUtil__.is_numeric(external_key(i)) Then
    
           external_key(i) = Null
           get_internal_key internal_key, external_key, key_len
           LookupMatchBAK = get_from_nested_array(Result, Table(VarIndex), internal_key)
           If LookupMatchBAK Then Exit Function
    
        End If
        
        i = i - 1
    Loop
        
    Exit Function

HandleError:
    
    modErrInfo__.FuncID = id_LookupMatch__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext

End Function

Public Function LookupMatch(Result As Variant, Table As Variant, ByVal VarIndex As Long, ParamArray Key() As Variant) As Boolean

    Dim key_len As Long, key_len_exnill As Long
    Dim i As Long, j As Long
    Dim internal_key As Variant
    Dim external_key As Variant, external_key_masked As Variant
    Dim ex_nil2inc_nill As Variant 'Array Index map, keys excluding nill to including nill
    Dim lex_order As Variant    'True/False series
    
    Dim key_lb As Long, key_ub As Long

    On Error GoTo HandleError:
    
    key_ub = UBound(Key)
    key_lb = LBound(Key)

    If key_ub < key_lb Then
    
        Err.Raise Number:=fmlInvalidArgument, _
                 Source:=PROJ_NAME, _
                 Description:=modErrInfo__.errstr_InvalidArgument + "Key"
        
    ElseIf VarIndex < LBound(Table) Or VarIndex > UBound(Table) Then
        
        Err.Raise Number:=fmlInvalidArgument, _
                 Source:=PROJ_NAME, _
                 Description:=modErrInfo__.errstr_InvalidArgument + "VarIndex " + CStr(VarIndex)
                
    End If
        
    key_len = key_ub - key_lb + 1
    external_key = Key
    
    ReDim ex_nil2inc_nill(0 To key_ub - key_lb) As Long
    
    For i = key_lb To key_ub
    
        If modUtil__.is_numeric(external_key(i)) Then
        
            ex_nil2inc_nill(key_len_exnill) = i
            key_len_exnill = key_len_exnill + 1
        
        End If
    
    Next i
    
    'Loop over Number of Matching elements
    For i = key_len_exnill To 0 Step -1
    
        ReDim lex_order(0 To key_len_exnill - 1) As Boolean
        
        'Initialize to True sequence
        For j = 0 To i - 1
    
            lex_order(j) = True
        
        Next j
                
        Do
            'Create Masked external key based on lex_order
            external_key_masked = external_key
            
            For j = 0 To key_len_exnill - 1
            
                If lex_order(j) = False Then
                
                    external_key_masked(ex_nil2inc_nill(j)) = Null
            
                End If
            
            Next j
                            
            get_internal_key internal_key, external_key_masked, key_len
            LookupMatch = get_from_nested_array(Result, Table(VarIndex), internal_key)
            
            If LookupMatch Then Exit Function
                
        Loop While incr_lexico_order(lex_order)
    
    Next i
        
    Exit Function

HandleError:
    
    modErrInfo__.FuncID = id_LookupMatch__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext

End Function

Private Function incr_lexico_order(v As Variant) As Boolean

    'v: Array of True/False
    
    Dim v_lb As Long, v_ub As Long, i As Long, j As Long
    Dim right_empty As Boolean
    Dim true_count As Long
    
    v_lb = LBound(v)
    v_ub = UBound(v)

    i = v_ub

    Do While i >= v_lb
    
        If v(i) = True And right_empty Then

            true_count = true_count + 1
            v(i) = False
            
            For j = i + 1 To v_ub
            
            
                If j <= i + true_count Then
                
                    v(j) = True
                
                Else
                
                    v(j) = False
                
                End If
            
            Next j
            
            incr_lexico_order = True
            Exit Function
        
        ElseIf v(i) = True And Not right_empty Then
        
            true_count = true_count + 1
        
        ElseIf v(i) = False Then
        
            right_empty = True
        
        Else
        
            'Error
        
        End If
        
        i = i - 1
        
    Loop
    
    incr_lexico_order = False
    Exit Function

End Function


Private Function get_from_nested_array(ret_val As Variant, src_var As Variant, ByVal Key As Variant) As Boolean

    Dim key_elm As Long
    Dim key_len As Long
    Dim next_key As Variant
    Dim i As Long
    
    key_len = UBound(Key) - LBound(Key) + 1
           
    key_elm = Key(LBound(Key))
    
    If key_elm < LBound(src_var) Or UBound(src_var) < key_elm Then
    
        get_from_nested_array = False
        Exit Function
    
    End If
        
    If key_len = 1 Then

        If Not IsArray(src_var(key_elm)) _
            And Not IsEmpty(src_var(key_elm)) Then
        
            ret_val = src_var(key_elm)
            get_from_nested_array = True
            Exit Function
            
        Else
        
            get_from_nested_array = False
            Exit Function
        
        End If
        
    Else
    
        If IsArray(src_var(key_elm)) Then
        
            ReDim next_key(LBound(Key) + 1 To UBound(Key)) As Variant
            For i = LBound(Key) + 1 To UBound(Key)
        
                next_key(i) = Key(i)
            
            Next i
            
            get_from_nested_array = get_from_nested_array(ret_val, src_var(key_elm), next_key)
            Exit Function
            
        Else
        
            get_from_nested_array = False
            Exit Function
        
'            Err.Raise Number:=fmlInvalidArgument, _
'                     Source:=PROJ_NAME, _
'                     Description:=modErrInfo__.errstr_InvalidArgument + "Key"
        
        End If
        
    End If


End Function



Private Function get_recursive(Val As Variant, Var As Variant, ByVal ind As Variant) As Boolean

    'Assumes Ubound(ind) = 0

    Dim dim_size As Long
    Dim ind_len As Long
    Dim ind_next As Variant
    Dim i As Long

    dim_size = modArrSpt__.NumberOfArrayDimensions(Var)
    ind_len = UBound(ind) + 1
    
    'Index is too short
    If ind_len < dim_size Then
        
        Val = Empty
        get_recursive = False
        Exit Function
        
    ElseIf ind_len > dim_size Then
    
        ReDim ind_next(ind_len - dim_size - 1) As Variant
        
        For i = 0 To ind_len - dim_size - 1
        
            ind_next(i) = ind(i + dim_size)
        
        Next i
        
    Else    'ind_len = dim_size
        
    End If
    
        
    Select Case dim_size
    
    Case 0
    
        Val = Var
        get_recursive = True
        Exit Function
    
    Case 1
    
        If Not IsArray(Var(ind(0))) Then
        
            get_recursive = True
            Val = Var(ind(0))
            Exit Function
            
        ElseIf ind_len = dim_size And IsArray(Var(ind(0))) Then
        
            get_recursive = False
            Val = Empty
            Exit Function
            
        Else
        
            get_recursive = get_recursive(Val, Var(ind(0)), ind_next)
            Exit Function
            
        End If
        
    
    Case 2
    
        If Not IsArray(Var(ind(0), ind(1))) Then
        
            get_recursive = True
            Val = Var(ind(0), ind(1))
            Exit Function
            
        ElseIf ind_len = dim_size And IsArray(Var(ind(0), ind(1))) Then
        
            get_recursive = False
            Val = Empty
            Exit Function
            
        Else
        
            get_recursive = get_recursive(Val, Var(ind(0), ind(1)), ind_next)
            Exit Function
            
        End If
    
    Case 3
    
        If Not IsArray(Var(ind(0), ind(1), ind(2))) Then
        
            get_recursive = True
            Val = Var(ind(0), ind(1), ind(2))
            Exit Function
            
        ElseIf ind_len = dim_size And IsArray(Var(ind(0), ind(1), ind(2))) Then
        
            get_recursive = False
            Val = Empty
            Exit Function
            
        Else
        
            get_recursive = get_recursive(Val, Var(ind(0), ind(1), ind(2)), ind_next)
            Exit Function
            
        End If
    
    Case 4
    
        If Not IsArray(Var(ind(0), ind(1), ind(2), ind(3))) Then
        
            get_recursive = True
            Val = Var(ind(0), ind(1), ind(2), ind(3))
            Exit Function
            
        ElseIf ind_len = dim_size And IsArray(Var(ind(0), ind(1), ind(2), ind(3))) Then
        
            get_recursive = False
            Val = Empty
            Exit Function
            
        Else
        
            get_recursive = get_recursive(Val, Var(ind(0), ind(1), ind(2), ind(3)), ind_next)
            Exit Function
            
        End If
    
    Case 5
    
        If Not IsArray(Var(ind(0), ind(1), ind(2), ind(3), ind(4))) Then
        
            get_recursive = True
            Val = Var(ind(0), ind(1), ind(2), ind(3), ind(4))
            Exit Function
            
        ElseIf ind_len = dim_size And IsArray(Var(ind(0), ind(1), ind(2), ind(3), ind(4))) Then
        
            get_recursive = False
            Val = Empty
            Exit Function
            
        Else
        
            get_recursive = get_recursive(Val, Var(ind(0), ind(1), ind(2), ind(3), ind(4)), ind_next)
            Exit Function
            
        End If
    
    Case 6
    
        If Not IsArray(Var(ind(0), ind(1), ind(2), ind(3), ind(4), ind(5))) Then
        
            get_recursive = True
            Val = Var(ind(0), ind(1), ind(2), ind(3), ind(4), ind(5))
            Exit Function
            
        ElseIf ind_len = dim_size And IsArray(Var(ind(0), ind(1), ind(2), ind(3), ind(4), ind(5))) Then
        
            get_recursive = False
            Val = Empty
            Exit Function
            
        Else
        
            get_recursive = get_recursive(Val, Var(ind(0), ind(1), ind(2), ind(3), ind(4), ind(5)), ind_next)
            Exit Function
            
        End If
    
    Case 7
    
        If Not IsArray(Var(ind(0), ind(1), ind(2), ind(3), ind(4), ind(5), ind(6))) Then
        
            get_recursive = True
            Val = Var(ind(0), ind(1), ind(2), ind(3), ind(4), ind(5), ind(6))
            Exit Function
            
        ElseIf ind_len = dim_size And IsArray(Var(ind(0), ind(1), ind(2), ind(3), ind(4), ind(5), ind(6))) Then
        
            get_recursive = False
            Val = Empty
            Exit Function
            
        Else
        
            get_recursive = get_recursive(Val, Var(ind(0), ind(1), ind(2), ind(3), ind(4), ind(5), ind(6)), ind_next)
            Exit Function
            
        End If
        
    Case 8
    
        If Not IsArray(Var(ind(0), ind(1), ind(2), ind(3), ind(4), ind(5), ind(6), ind(7))) Then
        
            get_recursive = True
            Val = Var(ind(0), ind(1), ind(2), ind(3), ind(4), ind(5), ind(6), ind(7))
            Exit Function
            
        ElseIf ind_len = dim_size And IsArray(Var(ind(0), ind(1), ind(2), ind(3), ind(4), ind(5), ind(6), ind(7))) Then
        
            get_recursive = False
            Val = Empty
            Exit Function
            
        Else
        
            get_recursive = get_recursive(Val, Var(ind(0), ind(1), ind(2), ind(3), ind(4), ind(5), ind(6), ind(7)), ind_next)
            Exit Function
            
        End If
    
    Case 9
        
        If Not IsArray(Var(ind(0), ind(1), ind(2), ind(3), ind(4), ind(5), ind(6), ind(7), ind(8))) Then
        
            get_recursive = True
            Val = Var(ind(0), ind(1), ind(2), ind(3), ind(4), ind(5), ind(6), ind(7), ind(8))
            Exit Function
            
        ElseIf ind_len = dim_size And IsArray(Var(ind(0), ind(1), ind(2), ind(3), ind(4), ind(5), ind(6), ind(7), ind(8))) Then
        
            get_recursive = False
            Val = Empty
            Exit Function
            
        Else
        
            get_recursive = get_recursive(Val, Var(ind(0), ind(1), ind(2), ind(3), ind(4), ind(5), ind(6), ind(7), ind(8)), ind_next)
            Exit Function
            
        End If
    
    Case 10
        
        If Not IsArray(Var(ind(0), ind(1), ind(2), ind(3), ind(4), ind(5), ind(6), ind(7), ind(8), ind(9))) Then
        
            get_recursive = True
            Val = Var(ind(0), ind(1), ind(2), ind(3), ind(4), ind(5), ind(6), ind(7), ind(8), ind(9))
            Exit Function
            
        ElseIf ind_len = dim_size And IsArray(Var(ind(0), ind(1), ind(2), ind(3), ind(4), ind(5), ind(6), ind(7), ind(8), ind(9))) Then
        
            get_recursive = False
            Val = Empty
            Exit Function
            
        Else
        
            get_recursive = get_recursive(Val, Var(ind(0), ind(1), ind(2), ind(3), ind(4), ind(5), ind(6), ind(7), ind(8), ind(9)), ind_next)
            Exit Function
            
        End If
        
    Case Else
    
        get_recursive = False
        Val = Empty
        Exit Function
    
    End Select
        
End Function


'Public Function SetNumber(Var As Variant, ByVal value As Variant) As Boolean
'
'    Dim var_type As VbVarType
'
'    On Error GoTo HandleError:
'
'    var_type = VarType(value)
'
'    If var_type = VbVarType.vbBoolean _
'        Or var_type = VbVarType.vbInteger _
'        Or var_type = VbVarType.vbLong _
'        Or var_type = VbVarType.vbSingle _
'        Or var_type = VbVarType.vbDouble Then
'
'        Var = value
'        SetNumber = True
'        Exit Function
'
'    Else
'
'        SetNumber = False
'        Exit Function
'
'    End If
'
'    Exit Function
'
'HandleError:
'
'    modErrInfo__.FuncID = id_SetNumber__
'    Err.Raise Number:=Err.Number, _
'                Source:=Err.Source, _
'                Description:=Err.Description, _
'                HelpFile:=Err.HelpFile, _
'                HelpContext:=Err.HelpContext
'
'
'End Function


Public Sub EmbedArray(Target As Variant, Source As Variant)     ', _
                        Optional AllowExtraDims As Boolean = True, _
                        Optional MappingType As acArrayMappingType = fmlMapIndex, _
                        Optional IndexOffset As Variant = 0, _
                        Optional ExpandDim As Boolean = True) As Boolean
    
    Dim trg_dim_count As Long
    Dim src_dim_count As Long
    Dim lb_trg() As Long, ub_trg() As Long
    Dim lb_src() As Long, ub_src() As Long
    Dim lb_extra() As Long, ub_extra() As Long
    Dim src_ind() As Long, ind_extra() As Long, trg_ind() As Long
    Dim extra_dim_count As Long
    Dim i As Long

    On Error GoTo HandleError:

    If Not IsArray(Target) Then Exit Sub
    If Not IsArray(Source) Then Exit Sub

    trg_dim_count = modArrSpt__.NumberOfArrayDimensions(Target)
    src_dim_count = modArrSpt__.NumberOfArrayDimensions(Source)
    
    If src_dim_count > trg_dim_count Then Exit Sub  'TODO Throw Error
    extra_dim_count = trg_dim_count - src_dim_count
    
    If extra_dim_count > 0 Then
    
        ReDim lb_extra(0 To extra_dim_count - 1) As Long
        ReDim ub_extra(0 To extra_dim_count - 1) As Long
    
        For i = 1 To extra_dim_count
        
            lb_extra(i - 1) = LBound(Target, src_dim_count + i)
            ub_extra(i - 1) = UBound(Target, src_dim_count + i)
            
        Next i
    
    End If
    
    modUtil__.get_bound lb_src, ub_src, Source, src_dim_count
    src_ind = lb_src
    Do
        
        If extra_dim_count > 0 Then
        
            ind_extra = lb_extra
            Do
                trg_ind = src_ind
                modArrSpt__.ConcatenateArrays trg_ind, ind_extra
                
                set_array_element Target, trg_ind, get_array_element(Source, src_ind)
            
            Loop While modUtil__.increment_index(ind_extra, lb_extra, ub_extra)
            
        Else
            
            set_array_element Target, src_ind, get_array_element(Source, src_ind)
        
        End If
        
    Loop While modUtil__.increment_index(src_ind, lb_src, ub_src)
            
    Exit Sub
    
HandleError:
    
    modErrInfo__.FuncID = id_EmbedArray__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext
    
End Sub



'Public Sub ExtendUBound(Target As Variant, NthDim As Long, UBoundIndex As Long)
'
'    Dim dim_count  As Long
'    Dim lb As Variant, ub As Variant, ub_orig As Variant, idx As Variant
'    Dim new_array As Variant
'    Dim i As Long, j As Long
'
'    On Error GoTo HandleError:
'
'    If Not IsArray(Target) Then Exit Sub
'
'    dim_count = modArrSpt__.NumberOfArrayDimensions(Target)
'
'    If dim_count <= 0 Or NthDim > dim_count Then Exit Sub   'TODO Throw Error
'
'    If UBoundIndex <= UBound(Target, NthDim) Then Exit Sub  'TODO Throw Error
'
'    modUtil__.get_bound lb, ub, Target, dim_count
'
'    ub_orig = ub
'    ub(LBound(ub) + NthDim - 1) = UBoundIndex
'
'    create_mult_dim_array new_array, lb, ub
'
'    idx = lb
'    Do
'        set_array_element new_array, idx, get_array_element(Target, idx)
'
'    Loop While modUtil__.increment_index(idx, lb, ub_orig)
'
'    Target = new_array
'
'    Exit Sub
'
'HandleError:
'
'    modErrInfo__.FuncID = id_ExtendUBound__
'    Err.Raise Number:=Err.Number, _
'                Source:=Err.Source, _
'                Description:=Err.Description, _
'                HelpFile:=Err.HelpFile, _
'                HelpContext:=Err.HelpContext
'
'End Sub


Public Sub ReorderDim(Target As Variant, Source As Variant, ParamArray DimOrder() As Variant)

    Dim dim_count As Long
    Dim lb As Variant, ub As Variant
    Dim i As Long, j As Long
    Dim ind As Variant, ind_lb As Variant, ind_ub As Variant, ind_new As Variant
    Dim out_array As Variant
    Dim lb_dim_od As Long

    On Error GoTo HandleError:
    
    If Not IsArray(Source) Then Exit Sub            'TODO Throw Error

    dim_count = modArrSpt__.NumberOfArrayDimensions(Source)
    
    If UBound(DimOrder) - LBound(DimOrder) + 1 <> dim_count Then Exit Sub
    
    ReDim lb(1 To dim_count) As Long
    ReDim ub(1 To dim_count) As Long
    
    lb_dim_od = LBound(DimOrder)
    
    For i = 1 To dim_count
        
        If DimOrder(lb_dim_od - 1 + i) > dim_count Then Exit Sub
    
        lb(i) = LBound(Source, DimOrder(lb_dim_od - 1 + i))
        ub(i) = UBound(Source, DimOrder(lb_dim_od - 1 + i))
    
    Next i
    
    create_mult_dim_array out_array, lb, ub
    
    ReDim ind(1 To dim_count) As Long
    ReDim ind_lb(1 To dim_count) As Long
    ReDim ind_ub(1 To dim_count) As Long
    ReDim ind_new(1 To dim_count) As Long
    
    For i = 1 To dim_count
        
        ind(i) = LBound(Source, i)
        ind_lb(i) = ind(i)
        ind_ub(i) = UBound(Source, i)
    
    Next i
    
    Do
        For i = 1 To dim_count
        
            ind_new(DimOrder(lb_dim_od - 1 + i)) = ind(i)
            
        Next i
    
        set_array_element out_array, ind_new, get_array_element(Source, ind)
    
    
    Loop While modUtil__.increment_index(ind, ind_lb, ind_ub)
    
    Target = out_array

    Exit Sub

HandleError:
    
    modErrInfo__.FuncID = id_ReorderDim__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext
                
End Sub



Public Function Max(ParamArray Arg()) As Variant

    Dim max_arg As Long
    Dim i
    
    On Error GoTo HandleError:
    
    max_arg = UBound(Arg)
    
    Max = Arg(0)
    For i = 1 To max_arg
    
        If Arg(i) > Max Then
        
            Max = Arg(i)
        
        End If
    Next i

    Exit Function

HandleError:
    
    modErrInfo__.FuncID = id_Max__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext

End Function

Public Function Min(ParamArray Arg()) As Variant

    Dim max_arg As Long
    Dim i
    
    On Error GoTo HandleError:
    
    max_arg = UBound(Arg)
    
    Min = Arg(0)
    For i = 1 To max_arg
    
        If Arg(i) < Min Then
        
            Min = Arg(i)
        
        End If
    Next i

    Exit Function
    
HandleError:
    
    modErrInfo__.FuncID = id_Min__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext
    
End Function


Public Sub UnNestArray(Target As Variant, ByVal Source As Variant)

    Dim i As Long
    
    Dim upper_dim As Long, lower_dim As Long, total_dim As Long
    Dim ind_lb As Long, ind_ub As Long, ind As Long
    
    Dim ind_lb2, ind_ub2, ind2, ind_lb2_min, ind_ub2_max, ind_lb2_min_ext, ind_ub2_max_ext
    Dim total_ind, total_ind_lb, total_ind_ub
    
    upper_dim = modArrSpt__.NumberOfArrayDimensions(Source)
    
    If upper_dim <> 1 Then
    
        Err.Raise Number:=fmlInvalidArgument, _
                 Source:=PROJ_NAME, _
                 Description:=modErrInfo__.errstr_InvalidArgument
    
    End If
    
    ind_lb = LBound(Source)
    ind_ub = UBound(Source)
    
    '--- Get lower_dim ---
    For ind = ind_lb To ind_ub
            
        If IsArray(Source(ind)) Then
        
            lower_dim = modArrSpt__.NumberOfArrayDimensions(Source(ind))
            
        End If
        
        If IsArray(Source(ind)) And ind > ind_lb Then
        
            If lower_dim <> modArrSpt__.NumberOfArrayDimensions(Source(ind)) Then
            
                Err.Raise Number:=fmlInvalidArgument, _
                        Source:=PROJ_NAME, _
                        Description:=modErrInfo__.errstr_InvalidArgument
                          
            End If
        
        End If
        
                
    Next ind
    
    '--- Get lower ub lb  ---
    For ind = ind_lb To ind_ub
    
        If IsArray(Source(ind)) Then
        
            get_bound ind_lb2, ind_ub2, Source(ind), lower_dim
        
            If IsEmpty(ind_lb2_min) Then
                
                ind_lb2_min = ind_lb2
                ind_ub2_max = ind_ub2
                
            Else
            
                For i = LBound(ind_lb2) To UBound(ind_ub2)
                
                    ind_lb2_min(i) = Min(ind_lb2(i), ind_lb2_min(i))
                    ind_ub2_max(i) = Max(ind_ub2(i), ind_ub2_max(i))
                
                Next i
            
            End If
        
        End If
                
    Next ind

    ReDim ind_lb2_min_ext(lower_dim)
    ReDim ind_ub2_max_ext(lower_dim)
    
    For i = 0 To lower_dim
    
        If i = 0 Then
            
            ind_lb2_min_ext(i) = ind_lb
            ind_ub2_max_ext(i) = ind_ub
            
        Else
        
            ind_lb2_min_ext(i) = ind_lb2_min(i - 1)
            ind_ub2_max_ext(i) = ind_ub2_max(i - 1)
        
        End If
       
    Next i
    
    '--- Create Target ----
    new_array_lbub Target, ind_lb2_min_ext, ind_ub2_max_ext, lower_dim + 1
    
    '--- Copy Values ---
    For ind = ind_lb To ind_ub
    
        If Not IsEmpty(Source(ind)) Then
    
            get_bound ind_lb2, ind_ub2, Source(ind), lower_dim
            
            ind2 = ind_lb2
            
            Do
            
                connect_index total_ind, 1, lower_dim, Array(ind), ind2
                set_array_element Target, total_ind, get_array_element(Source(ind), ind2)
                
            Loop While increment_index(ind2, ind_lb2, ind_ub2)
            
        End If
    
    Next ind

    Exit Sub

HandleError:
    
    modErrInfo__.FuncID = id_UnNestArray__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext


End Sub


Private Sub new_array_lbub(trg, lb_ind, ub_ind, dim_size As Long)

    Select Case dim_size
    
    Case 1
    
        ReDim trg(lb_ind(0) To ub_ind(0))
    
    Case 2
    
        ReDim trg(lb_ind(0) To ub_ind(0), _
                    lb_ind(1) To ub_ind(1))
    
    Case 3
    
        ReDim trg(lb_ind(0) To ub_ind(0), _
                    lb_ind(1) To ub_ind(1), _
                    lb_ind(2) To ub_ind(2))
    
    Case 4
    
        ReDim trg(lb_ind(0) To ub_ind(0), _
                    lb_ind(1) To ub_ind(1), _
                    lb_ind(2) To ub_ind(2), _
                    lb_ind(3) To ub_ind(3))
    
    Case 5
    
        ReDim trg(lb_ind(0) To ub_ind(0), _
                    lb_ind(1) To ub_ind(1), _
                    lb_ind(2) To ub_ind(2), _
                    lb_ind(3) To ub_ind(3), _
                    lb_ind(4) To ub_ind(4))
    
    Case 6
    
        ReDim trg(lb_ind(0) To ub_ind(0), _
                    lb_ind(1) To ub_ind(1), _
                    lb_ind(2) To ub_ind(2), _
                    lb_ind(3) To ub_ind(3), _
                    lb_ind(4) To ub_ind(4), _
                    lb_ind(5) To ub_ind(5))
    
    Case 7
    
        ReDim trg(lb_ind(0) To ub_ind(0), _
                    lb_ind(1) To ub_ind(1), _
                    lb_ind(2) To ub_ind(2), _
                    lb_ind(3) To ub_ind(3), _
                    lb_ind(4) To ub_ind(4), _
                    lb_ind(5) To ub_ind(5), _
                    lb_ind(6) To ub_ind(6))
    
    Case 8
    
        ReDim trg(lb_ind(0) To ub_ind(0), _
                    lb_ind(1) To ub_ind(1), _
                    lb_ind(2) To ub_ind(2), _
                    lb_ind(3) To ub_ind(3), _
                    lb_ind(4) To ub_ind(4), _
                    lb_ind(5) To ub_ind(5), _
                    lb_ind(6) To ub_ind(6), _
                    lb_ind(7) To ub_ind(7))
    
    Case 9
    
        ReDim trg(lb_ind(0) To ub_ind(0), _
                    lb_ind(1) To ub_ind(1), _
                    lb_ind(2) To ub_ind(2), _
                    lb_ind(3) To ub_ind(3), _
                    lb_ind(4) To ub_ind(4), _
                    lb_ind(5) To ub_ind(5), _
                    lb_ind(6) To ub_ind(6), _
                    lb_ind(7) To ub_ind(7), _
                    lb_ind(8) To ub_ind(8))
    
    
    Case 10
       
        ReDim trg(lb_ind(0) To ub_ind(0), _
                    lb_ind(1) To ub_ind(1), _
                    lb_ind(2) To ub_ind(2), _
                    lb_ind(3) To ub_ind(3), _
                    lb_ind(4) To ub_ind(4), _
                    lb_ind(5) To ub_ind(5), _
                    lb_ind(6) To ub_ind(6), _
                    lb_ind(7) To ub_ind(7), _
                    lb_ind(8) To ub_ind(8), _
                    lb_ind(9) To ub_ind(9))
       
    
    End Select
    

End Sub

Private Sub connect_index(total_ind, size1 As Long, size2 As Long, ind1, ind2)

    Dim i As Long

    ReDim total_ind(size1 + size2 - 1)

    For i = 0 To size1 - 1
    
        total_ind(i) = ind1(0)
    
    Next i
    
    
    For i = 0 To size2 - 1
    
        total_ind(i + size1) = ind2(i)
    
    Next i


End Sub


Public Sub NestArray(Target As Variant, ByVal Source As Variant, NestPosition As Long)

    Dim orig_dim_count As Long, upper_dim_count As Long, lower_dim_count As Long
    Dim lb_upper As Variant, ub_upper As Variant
    Dim lb_lower As Variant, ub_lower As Variant
    Dim upper As Variant, lower As Variant
    Dim ind_upper As Variant, ind_lower As Variant, ind As Variant
    Dim i As Long
    
    On Error GoTo HandleError:
    
    If Not IsArray(Source) Then
    
        Err.Raise Number:=fmlInvalidArgument, _
                 Source:=PROJ_NAME, _
                 Description:=modErrInfo__.errstr_InvalidArgument
            
    End If
    
    orig_dim_count = modArrSpt__.NumberOfArrayDimensions(Source)
    
    If NestPosition < 1 Or NestPosition >= orig_dim_count Then
    
        Err.Raise Number:=fmlInvalidArgument, _
                 Source:=PROJ_NAME, _
                 Description:=modErrInfo__.errstr_InvalidArgument
    
    End If
    
    ReDim lb_upper(1 To NestPosition)
    ReDim ub_upper(1 To NestPosition)
    ReDim lb_lower(NestPosition + 1 To orig_dim_count)
    ReDim ub_lower(NestPosition + 1 To orig_dim_count)
    
    For i = 1 To NestPosition
    
        lb_upper(i) = LBound(Source, i)
        ub_upper(i) = UBound(Source, i)
    
    Next i
    
    For i = NestPosition + 1 To orig_dim_count
    
        lb_lower(i) = LBound(Source, i)
        ub_lower(i) = UBound(Source, i)
    
    Next i
    
    create_new_array upper, lb_upper, ub_upper
    
    ind_upper = lb_upper
    Do
        create_new_array lower, lb_lower, ub_lower
        
        ind_lower = lb_lower
        Do
            ind = ind_upper
            modArrSpt__.ConcatenateArrays ind, ind_lower
            set_array_element lower, ind_lower, get_array_element(Source, ind)
            
        Loop While modUtil__.increment_index(ind_lower, lb_lower, ub_lower)
        
        set_array_element upper, ind_upper, lower
        
    Loop While modUtil__.increment_index(ind_upper, lb_upper, ub_upper)
    
    Target = upper
    
    Exit Sub
    
HandleError:
    
    modErrInfo__.FuncID = id_NestArray__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext
    


End Sub



Private Sub set_array_element(Arr As Variant, ArrayIndex As Variant, elm As Variant)

    Dim dim_count As Long
    Dim ub_ind As Long

    ub_ind = LBound(ArrayIndex)
    dim_count = UBound(ArrayIndex) - LBound(ArrayIndex) + 1
    
    Select Case dim_count
    
    Case 1
        Arr(ArrayIndex(ub_ind)) = elm
    
    Case 2
        Arr( _
            ArrayIndex(ub_ind), _
            ArrayIndex(ub_ind + 1)) = elm
    
    Case 3
        Arr( _
            ArrayIndex(ub_ind), _
            ArrayIndex(ub_ind + 1), _
            ArrayIndex(ub_ind + 2)) = elm
    Case 4
        Arr( _
            ArrayIndex(ub_ind), _
            ArrayIndex(ub_ind + 1), _
            ArrayIndex(ub_ind + 2), _
            ArrayIndex(ub_ind + 3)) = elm
    Case 5
        Arr( _
            ArrayIndex(ub_ind), _
            ArrayIndex(ub_ind + 1), _
            ArrayIndex(ub_ind + 2), _
            ArrayIndex(ub_ind + 3), _
            ArrayIndex(ub_ind + 4)) = elm
    
    Case 6
        Arr( _
            ArrayIndex(ub_ind), _
            ArrayIndex(ub_ind + 1), _
            ArrayIndex(ub_ind + 2), _
            ArrayIndex(ub_ind + 3), _
            ArrayIndex(ub_ind + 4), _
            ArrayIndex(ub_ind + 5)) = elm
            
    Case 7
        Arr( _
            ArrayIndex(ub_ind), _
            ArrayIndex(ub_ind + 1), _
            ArrayIndex(ub_ind + 2), _
            ArrayIndex(ub_ind + 3), _
            ArrayIndex(ub_ind + 4), _
            ArrayIndex(ub_ind + 5), _
            ArrayIndex(ub_ind + 6)) = elm
    
    Case 8
        Arr( _
            ArrayIndex(ub_ind), _
            ArrayIndex(ub_ind + 1), _
            ArrayIndex(ub_ind + 2), _
            ArrayIndex(ub_ind + 3), _
            ArrayIndex(ub_ind + 4), _
            ArrayIndex(ub_ind + 5), _
            ArrayIndex(ub_ind + 6), _
            ArrayIndex(ub_ind + 7)) = elm
    
    Case 9
        Arr( _
            ArrayIndex(ub_ind), _
            ArrayIndex(ub_ind + 1), _
            ArrayIndex(ub_ind + 2), _
            ArrayIndex(ub_ind + 3), _
            ArrayIndex(ub_ind + 4), _
            ArrayIndex(ub_ind + 5), _
            ArrayIndex(ub_ind + 6), _
            ArrayIndex(ub_ind + 7), _
            ArrayIndex(ub_ind + 8)) = elm
            
    Case 10
         Arr( _
            ArrayIndex(ub_ind), _
            ArrayIndex(ub_ind + 1), _
            ArrayIndex(ub_ind + 2), _
            ArrayIndex(ub_ind + 3), _
            ArrayIndex(ub_ind + 4), _
            ArrayIndex(ub_ind + 5), _
            ArrayIndex(ub_ind + 6), _
            ArrayIndex(ub_ind + 7), _
            ArrayIndex(ub_ind + 8), _
            ArrayIndex(ub_ind + 9)) = elm
            
    End Select

End Sub

Private Function is_elm_array(InputArray As Variant, ArrayIndex As Variant, Optional ind_len_cap As Long = MAX_DIM_COUNT) As Long

    Dim lb_ind As Long
    Dim i As Long
    Dim dim_count As Long
    
    lb_ind = LBound(ArrayIndex)
    dim_count = IIf(ind_len_cap < UBound(ArrayIndex), ind_len_cap, UBound(ArrayIndex)) - lb_ind + 1

    Select Case dim_count
    
    Case 1
        is_elm_array = IsArray(InputArray( _
                            ArrayIndex(lb_ind)))
        Exit Function
    
    Case 2
        is_elm_array = IsArray(InputArray( _
                            ArrayIndex(lb_ind), _
                            ArrayIndex(lb_ind + 1)))
        Exit Function
    
    Case 3
        is_elm_array = IsArray(InputArray( _
                            ArrayIndex(lb_ind), _
                            ArrayIndex(lb_ind + 1), _
                            ArrayIndex(lb_ind + 2)))
        Exit Function
    
    Case 4
        is_elm_array = IsArray(InputArray( _
                            ArrayIndex(lb_ind), _
                            ArrayIndex(lb_ind + 1), _
                            ArrayIndex(lb_ind + 2), _
                            ArrayIndex(lb_ind + 3)))
        Exit Function
    
    Case 5
        is_elm_array = IsArray(InputArray( _
                            ArrayIndex(lb_ind), _
                            ArrayIndex(lb_ind + 1), _
                            ArrayIndex(lb_ind + 2), _
                            ArrayIndex(lb_ind + 3), _
                            ArrayIndex(lb_ind + 4)))
        Exit Function
    
    Case 6
        is_elm_array = IsArray(InputArray( _
                            ArrayIndex(lb_ind), _
                            ArrayIndex(lb_ind + 1), _
                            ArrayIndex(lb_ind + 2), _
                            ArrayIndex(lb_ind + 3), _
                            ArrayIndex(lb_ind + 4), _
                            ArrayIndex(lb_ind + 5)))
        Exit Function
    
    Case 7
        is_elm_array = IsArray(InputArray( _
                            ArrayIndex(lb_ind), _
                            ArrayIndex(lb_ind + 1), _
                            ArrayIndex(lb_ind + 2), _
                            ArrayIndex(lb_ind + 3), _
                            ArrayIndex(lb_ind + 4), _
                            ArrayIndex(lb_ind + 5), _
                            ArrayIndex(lb_ind + 6)))
        Exit Function
    
    Case 8
        is_elm_array = IsArray(InputArray( _
                            ArrayIndex(lb_ind), _
                            ArrayIndex(lb_ind + 1), _
                            ArrayIndex(lb_ind + 2), _
                            ArrayIndex(lb_ind + 3), _
                            ArrayIndex(lb_ind + 4), _
                            ArrayIndex(lb_ind + 5), _
                            ArrayIndex(lb_ind + 6), _
                            ArrayIndex(lb_ind + 7)))
        Exit Function
    
    Case 9
        is_elm_array = IsArray(InputArray( _
                            ArrayIndex(lb_ind), _
                            ArrayIndex(lb_ind + 1), _
                            ArrayIndex(lb_ind + 2), _
                            ArrayIndex(lb_ind + 3), _
                            ArrayIndex(lb_ind + 4), _
                            ArrayIndex(lb_ind + 5), _
                            ArrayIndex(lb_ind + 6), _
                            ArrayIndex(lb_ind + 7), _
                            ArrayIndex(lb_ind + 8)))
        Exit Function
    
    Case 10
        is_elm_array = IsArray(InputArray( _
                            ArrayIndex(lb_ind), _
                            ArrayIndex(lb_ind + 1), _
                            ArrayIndex(lb_ind + 2), _
                            ArrayIndex(lb_ind + 3), _
                            ArrayIndex(lb_ind + 4), _
                            ArrayIndex(lb_ind + 5), _
                            ArrayIndex(lb_ind + 6), _
                            ArrayIndex(lb_ind + 7), _
                            ArrayIndex(lb_ind + 8), _
                            ArrayIndex(lb_ind + 9)))
        Exit Function
    
    Case Else
        is_elm_array = 0
        Exit Function
    
    End Select

End Function

Private Function get_array_element(InputArray As Variant, ArrayIndex As Variant, Optional ind_len_cap As Long = MAX_DIM_COUNT) As Variant

    Dim lb_ind As Long
    Dim i As Long
    Dim dim_count As Long
    
    lb_ind = LBound(ArrayIndex)
    dim_count = IIf(ind_len_cap < UBound(ArrayIndex), ind_len_cap, UBound(ArrayIndex)) - lb_ind + 1

    Select Case dim_count
    
    Case 1
        get_array_element = InputArray( _
                            ArrayIndex(lb_ind))
        Exit Function
    
    Case 2
        get_array_element = InputArray( _
                            ArrayIndex(lb_ind), _
                            ArrayIndex(lb_ind + 1))
        Exit Function
    
    Case 3
        get_array_element = InputArray( _
                            ArrayIndex(lb_ind), _
                            ArrayIndex(lb_ind + 1), _
                            ArrayIndex(lb_ind + 2))
        Exit Function
    
    Case 4
        get_array_element = InputArray( _
                            ArrayIndex(lb_ind), _
                            ArrayIndex(lb_ind + 1), _
                            ArrayIndex(lb_ind + 2), _
                            ArrayIndex(lb_ind + 3))
        Exit Function
    
    Case 5
        get_array_element = InputArray( _
                            ArrayIndex(lb_ind), _
                            ArrayIndex(lb_ind + 1), _
                            ArrayIndex(lb_ind + 2), _
                            ArrayIndex(lb_ind + 3), _
                            ArrayIndex(lb_ind + 4))
        Exit Function
    
    Case 6
        get_array_element = InputArray( _
                            ArrayIndex(lb_ind), _
                            ArrayIndex(lb_ind + 1), _
                            ArrayIndex(lb_ind + 2), _
                            ArrayIndex(lb_ind + 3), _
                            ArrayIndex(lb_ind + 4), _
                            ArrayIndex(lb_ind + 5))
        Exit Function
    
    Case 7
        get_array_element = InputArray( _
                            ArrayIndex(lb_ind), _
                            ArrayIndex(lb_ind + 1), _
                            ArrayIndex(lb_ind + 2), _
                            ArrayIndex(lb_ind + 3), _
                            ArrayIndex(lb_ind + 4), _
                            ArrayIndex(lb_ind + 5), _
                            ArrayIndex(lb_ind + 6))
        Exit Function
    
    Case 8
        get_array_element = InputArray( _
                            ArrayIndex(lb_ind), _
                            ArrayIndex(lb_ind + 1), _
                            ArrayIndex(lb_ind + 2), _
                            ArrayIndex(lb_ind + 3), _
                            ArrayIndex(lb_ind + 4), _
                            ArrayIndex(lb_ind + 5), _
                            ArrayIndex(lb_ind + 6), _
                            ArrayIndex(lb_ind + 7))
        Exit Function
    
    Case 9
        get_array_element = InputArray( _
                            ArrayIndex(lb_ind), _
                            ArrayIndex(lb_ind + 1), _
                            ArrayIndex(lb_ind + 2), _
                            ArrayIndex(lb_ind + 3), _
                            ArrayIndex(lb_ind + 4), _
                            ArrayIndex(lb_ind + 5), _
                            ArrayIndex(lb_ind + 6), _
                            ArrayIndex(lb_ind + 7), _
                            ArrayIndex(lb_ind + 8))
        Exit Function
    
    Case 10
        get_array_element = InputArray( _
                            ArrayIndex(lb_ind), _
                            ArrayIndex(lb_ind + 1), _
                            ArrayIndex(lb_ind + 2), _
                            ArrayIndex(lb_ind + 3), _
                            ArrayIndex(lb_ind + 4), _
                            ArrayIndex(lb_ind + 5), _
                            ArrayIndex(lb_ind + 6), _
                            ArrayIndex(lb_ind + 7), _
                            ArrayIndex(lb_ind + 8), _
                            ArrayIndex(lb_ind + 9))
        Exit Function
    
    Case Else
        get_array_element = Empty
        Exit Function
    
    End Select

End Function


Private Sub create_mult_dim_array(out_arr As Variant, lb As Variant, ub As Variant)

    Dim dim_count As Long
    Dim i As Long
    Dim lb_lb As Long, ub_lb As Long

    dim_count = UBound(lb) - LBound(lb) + 1
    lb_lb = LBound(lb)
    ub_lb = LBound(ub)

    Select Case dim_count
    
    Case 1
        ReDim out_arr(lb(lb_lb) To ub(ub_lb)) As Variant
    
    Case 2
        ReDim out_arr(lb(lb_lb) To ub(ub_lb), _
                      lb(lb_lb + 1) To ub(ub_lb + 1)) As Variant
        
    Case 3
        ReDim out_arr(lb(lb_lb) To ub(ub_lb), _
                      lb(lb_lb + 1) To ub(ub_lb + 1), _
                      lb(lb_lb + 2) To ub(ub_lb + 2)) As Variant
        
    Case 4
        ReDim out_arr(lb(lb_lb) To ub(ub_lb), _
                      lb(lb_lb + 1) To ub(ub_lb + 1), _
                      lb(lb_lb + 2) To ub(ub_lb + 2), _
                      lb(lb_lb + 3) To ub(ub_lb + 3)) As Variant
        
    Case 5
        ReDim out_arr(lb(lb_lb) To ub(ub_lb), _
                      lb(lb_lb + 1) To ub(ub_lb + 1), _
                      lb(lb_lb + 2) To ub(ub_lb + 2), _
                      lb(lb_lb + 3) To ub(ub_lb + 3), _
                      lb(lb_lb + 4) To ub(ub_lb + 4)) As Variant
        
    Case 6
        ReDim out_arr(lb(lb_lb) To ub(ub_lb), _
                      lb(lb_lb + 1) To ub(ub_lb + 1), _
                      lb(lb_lb + 2) To ub(ub_lb + 2), _
                      lb(lb_lb + 3) To ub(ub_lb + 3), _
                      lb(lb_lb + 4) To ub(ub_lb + 4), _
                      lb(lb_lb + 5) To ub(ub_lb + 5)) As Variant
        
    Case 7
        ReDim out_arr(lb(lb_lb) To ub(ub_lb), _
                      lb(lb_lb + 1) To ub(ub_lb + 1), _
                      lb(lb_lb + 2) To ub(ub_lb + 2), _
                      lb(lb_lb + 3) To ub(ub_lb + 3), _
                      lb(lb_lb + 4) To ub(ub_lb + 4), _
                      lb(lb_lb + 5) To ub(ub_lb + 5), _
                      lb(lb_lb + 6) To ub(ub_lb + 6)) As Variant
        
    Case 8
        ReDim out_arr(lb(lb_lb) To ub(ub_lb), _
                      lb(lb_lb + 1) To ub(ub_lb + 1), _
                      lb(lb_lb + 2) To ub(ub_lb + 2), _
                      lb(lb_lb + 3) To ub(ub_lb + 3), _
                      lb(lb_lb + 4) To ub(ub_lb + 4), _
                      lb(lb_lb + 5) To ub(ub_lb + 5), _
                      lb(lb_lb + 6) To ub(ub_lb + 6), _
                      lb(lb_lb + 7) To ub(ub_lb + 7)) As Variant
        
    Case 9
        ReDim out_arr(lb(lb_lb) To ub(ub_lb), _
                      lb(lb_lb + 1) To ub(ub_lb + 1), _
                      lb(lb_lb + 2) To ub(ub_lb + 2), _
                      lb(lb_lb + 3) To ub(ub_lb + 3), _
                      lb(lb_lb + 4) To ub(ub_lb + 4), _
                      lb(lb_lb + 5) To ub(ub_lb + 5), _
                      lb(lb_lb + 6) To ub(ub_lb + 6), _
                      lb(lb_lb + 7) To ub(ub_lb + 7), _
                      lb(lb_lb + 8) To ub(ub_lb + 8)) As Variant
        
    Case 10
        ReDim out_arr(lb(lb_lb) To ub(ub_lb), _
                      lb(lb_lb + 1) To ub(ub_lb + 1), _
                      lb(lb_lb + 2) To ub(ub_lb + 2), _
                      lb(lb_lb + 3) To ub(ub_lb + 3), _
                      lb(lb_lb + 4) To ub(ub_lb + 4), _
                      lb(lb_lb + 5) To ub(ub_lb + 5), _
                      lb(lb_lb + 6) To ub(ub_lb + 6), _
                      lb(lb_lb + 7) To ub(ub_lb + 7), _
                      lb(lb_lb + 8) To ub(ub_lb + 8), _
                      lb(lb_lb + 9) To ub(ub_lb + 9)) As Variant

        
    End Select

End Sub


Public Function FindVal( _
    Result As Variant, _
    ByVal Key As Variant, _
    ByVal Var As Variant, _
    ColIndex As Long, _
    Optional FindOperator As fmlCompareOp = fmlEqualTo, _
    Optional FindOption As fmlFindOption = 0) As Boolean
    
    Dim key_count As Long
    Dim current_key() As Variant
    Dim i As Long, i_last As Long, i_first As Long
    Dim compare_result As Boolean
    Dim arr_in As Variant

    On Error GoTo HandleError:
    
    FindVal = False
    
    arr_in = Var
    
    If Not IsArray(arr_in) Then Exit Function   'TODO Throw an Error
    If modArrSpt__.NumberOfArrayDimensions(arr_in) <> 2 Then Exit Function  'TODO Throw an Error
    
    key_count = arrayalize_scalar(Key)
    
    i = LBound(arr_in, 1)
    i_last = UBound(arr_in, 1)
    
    Do While i <= i_last
    
        modArrSpt__.GetRow arr_in, current_key, i
            
        Select Case FindOperator
        Case fmlEqualTo
            
            If i <> i_last And compare_keys(Key, current_key, fmlEqualTo) Then
            
                Result = arr_in(i, ColIndex)
                FindVal = True
                Exit Function
                
            ElseIf i = i_last And _
                ((FindOption And fmlLastAsElse) = fmlLastAsElse) Then
                
                Result = arr_in(i, ColIndex)
                FindVal = True
                Exit Function
        
            ElseIf i = i_last And compare_keys(Key, current_key, fmlEqualTo) Then
                
                Result = arr_in(i, ColIndex)
                FindVal = True
                Exit Function
                
            Else
            
                FindVal = False
                Exit Function
           
            End If
        
        Case fmlLessThan, fmlNoMoreThan
        
            If i = i_first Then
          
                If ((FindOption And fmlFirstAsMoreThan) = fmlFirstAsMoreThan) Then
                
                    If compare_keys(Key, current_key, fmlLessThan) Then
                    
                        FindVal = False
                        Exit Function
                        
                    End If
                    
                End If
                    
                If ((FindOption And fmlFirstAsEqualTo) = fmlFirstAsEqualTo) Then
                
                    If compare_keys(Key, current_key, fmlEqualTo) Then
                        
                        Result = arr_in(i, ColIndex)
                        FindVal = True
                        Exit Function
                        
                    End If
    
                End If
                
                If Not (FindOption And fmlFirstAsMoreThan) = fmlFirstAsMoreThan _
                    And Not (FindOption And fmlFirstAsEqualTo) = fmlFirstAsEqualTo Then
                
                    If compare_keys(Key, current_key, FindOperator) Then
                    
                        Result = arr_in(i, ColIndex)
                        FindVal = True
                        Exit Function
                        
                    End If
                
                End If
            
            ElseIf i = i_last Then
            
                If ((FindOption And fmlLastAsElse) = fmlLastAsElse) Then
                    
                    Result = arr_in(i, ColIndex)
                    FindVal = True
                    Exit Function
                
                Else
                    
                    compare_result = compare_keys(Key, current_key, FindOperator)
                    If compare_result Then
                    
                        Result = arr_in(i, ColIndex)
                        FindVal = True
                        Exit Function
                        
                    End If
                    
                End If
            
            Else
                
                If compare_keys(Key, current_key, FindOperator) Then
                
                    Result = arr_in(i, ColIndex)
                    FindVal = True
                    Exit Function
                    
                End If
            
            End If
                
        Case Else
        
            'TODO Throw Error
        
        End Select
        i = i + 1
    Loop
    
    FindVal = False
    Exit Function

HandleError:
    
    modErrInfo__.FuncID = id_FindVal__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext

End Function


Private Function compare_keys( _
    a_key As Variant, _
    b_key As Variant, _
    compare_type As fmlCompareOp) As Boolean
    'Compare keys only to the length of a_key

    Dim key_len As Long
    Dim i_min As Long
    Dim i As Long, j As Long
    
    key_len = modArrSpt__.NumElements(a_key)
    
    i_min = LBound(a_key)
    i = i_min
    j = LBound(b_key)
    
    compare_keys = False
    Do While i < i_min + key_len
    
        Select Case compare_type
    
        Case fmlCompareOp.fmlEqualTo
            If a_key(i) <> b_key(j) Then Exit Function
        
        Case fmlCompareOp.fmlLessThan
            If Not (a_key(i) < b_key(j)) Then Exit Function
        
        Case fmlCompareOp.fmlNoMoreThan
            If Not (a_key(i) <= b_key(j)) Then Exit Function
        
        End Select
    
        i = i + 1: j = j + 1
    Loop

    compare_keys = True

End Function


Public Sub NewArray(Target As Variant, ParamArray Index() As Variant)

    Dim i As Long
    Dim lb_array() As Long, ub_array() As Long
    Dim ind_ub As Long
    Dim dim_count As Long

    On Error GoTo HandleError:
    
    ind_ub = UBound(Index)
    
    If ind_ub = 1 Then
    
        If Not IsArray(Index(0)) _
            And Not IsArray(Index(1)) Then
    
            ReDim lb_array(0) As Long
            ReDim ub_array(0) As Long
            
            lb_array(0) = Index(0)
            ub_array(0) = Index(1)
            
            dim_count = 1
            
        End If
        
    End If
    
    If dim_count = 0 Then 'Index passed as Array(s)
    
        ReDim lb_array(ind_ub) As Long
        ReDim ub_array(ind_ub) As Long
        
        For i = 0 To ind_ub
        
            lb_array(i) = Index(i)(0)
            ub_array(i) = Index(i)(1)
        
        Next i
    
        dim_count = ind_ub + 1
        
    End If
    
    
    Select Case dim_count
    
    Case 1
        ReDim Target(lb_array(0) To ub_array(0))
    
    Case 2
        ReDim Target( _
            lb_array(0) To ub_array(0), _
            lb_array(1) To ub_array(1))
    
    Case 3
        ReDim Target( _
            lb_array(0) To ub_array(0), _
            lb_array(1) To ub_array(1), _
            lb_array(2) To ub_array(2))
    
    Case 4
        ReDim Target( _
            lb_array(0) To ub_array(0), _
            lb_array(1) To ub_array(1), _
            lb_array(2) To ub_array(2), _
            lb_array(3) To ub_array(3))
    
    Case 5
        ReDim Target( _
            lb_array(0) To ub_array(0), _
            lb_array(1) To ub_array(1), _
            lb_array(2) To ub_array(2), _
            lb_array(3) To ub_array(3), _
            lb_array(4) To ub_array(4))
        
    Case 6
        ReDim Target( _
            lb_array(0) To ub_array(0), _
            lb_array(1) To ub_array(1), _
            lb_array(2) To ub_array(2), _
            lb_array(3) To ub_array(3), _
            lb_array(4) To ub_array(4), _
            lb_array(5) To ub_array(5))
    
    Case 7
        ReDim Target( _
            lb_array(0) To ub_array(0), _
            lb_array(1) To ub_array(1), _
            lb_array(2) To ub_array(2), _
            lb_array(3) To ub_array(3), _
            lb_array(4) To ub_array(4), _
            lb_array(5) To ub_array(5), _
            lb_array(6) To ub_array(6))
    
    Case 8
        ReDim Target( _
            lb_array(0) To ub_array(0), _
            lb_array(1) To ub_array(1), _
            lb_array(2) To ub_array(2), _
            lb_array(3) To ub_array(3), _
            lb_array(4) To ub_array(4), _
            lb_array(5) To ub_array(5), _
            lb_array(6) To ub_array(6), _
            lb_array(7) To ub_array(7))
    
    Case 9
        ReDim Target( _
            lb_array(0) To ub_array(0), _
            lb_array(1) To ub_array(1), _
            lb_array(2) To ub_array(2), _
            lb_array(3) To ub_array(3), _
            lb_array(4) To ub_array(4), _
            lb_array(5) To ub_array(5), _
            lb_array(6) To ub_array(6), _
            lb_array(7) To ub_array(7), _
            lb_array(8) To ub_array(8))
            
    Case 10
        ReDim Target( _
            lb_array(0) To ub_array(0), _
            lb_array(1) To ub_array(1), _
            lb_array(2) To ub_array(2), _
            lb_array(3) To ub_array(3), _
            lb_array(4) To ub_array(4), _
            lb_array(5) To ub_array(5), _
            lb_array(6) To ub_array(6), _
            lb_array(7) To ub_array(7), _
            lb_array(8) To ub_array(8), _
            lb_array(9) To ub_array(9))
    
    End Select
    
    Exit Sub

HandleError:
    
    modErrInfo__.FuncID = id_NewArray__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext


End Sub




Public Sub NewNestedArray(Target As Variant, ParamArray Index() As Variant)


    On Error GoTo HandleError:


    If UBound(Index) < LBound(Index) Then
        
        Err.Raise Number:=fmlInvalidArgument, _
                  Source:=PROJ_NAME, _
                  Description:=modErrInfo__.errstr_InvalidArgument
    
    End If

    new_nested_array Target, Index

    Exit Sub

HandleError:
    
    modErrInfo__.FuncID = id_NewNestedArray__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext


End Sub


Private Sub new_nested_array(trg As Variant, ByVal ind)

    'ind is an array of arrays. each elemet array is either array(LB, UB) or Array(array(lb, ub), ...)

    Dim trg_lb, trg_ub, trg_ind As Variant
    Dim ind_len As Long
    Dim next_ind As Variant
    Dim i As Long, j As Long, ind_lb As Long
    Dim trg_dim As Long
    
    ind_lb = LBound(ind)
    ind_len = UBound(ind) - ind_lb
    
    'Create Array from the 1st ind element
    If Not IsArray(ind(ind_lb)(LBound(ind(ind_lb)))) Then
    
        trg_dim = 0
        
        ReDim trg_lb(0)
        ReDim trg_ub(0)
        ReDim trg_ind(0)
    
        trg_lb(0) = ind(ind_lb)(LBound(ind(ind_lb)))
        trg_ub(0) = ind(ind_lb)(LBound(ind(ind_lb)) + 1)
        
        ReDim trg(trg_lb(0) To trg_ub(0))
    
    Else
    
        trg_dim = UBound(ind(ind_lb)) - LBound(ind(ind_lb))
        
        ReDim trg_lb(trg_dim)
        ReDim trg_ub(trg_dim)
    
        For i = 0 To trg_dim
        
            j = LBound(ind(ind_lb)) + i
            
            trg_lb(i) = ind(ind_lb)(j)(LBound(ind(ind_lb)(j)))
            trg_ub(i) = ind(ind_lb)(j)(LBound(ind(ind_lb)(j)) + 1)
        
        Next i
        
        create_mult_dim_array trg, trg_lb, trg_ub
            
    End If
    
    'Create Next Index if not end
    If ind_len = 0 Then
    
        Exit Sub
    
    Else
    
        ReDim next_ind(ind_len - 1)
        
        For i = 0 To ind_len - 1
        
            next_ind(i) = ind(i + 1)
        
        Next i
    
    End If
            
    trg_ind = trg_lb
       
    Do
        
        Select Case trg_dim
        
        Case 0
            new_nested_array trg(trg_ind(0)), next_ind
        
        Case 1
            new_nested_array trg(trg_ind(0), trg_ind(1)), next_ind
        
        Case 2
            new_nested_array trg(trg_ind(0), trg_ind(1), trg_ind(2)), next_ind
        
        Case 3
            new_nested_array trg(trg_ind(0), trg_ind(1), trg_ind(2), trg_ind(3)), next_ind
        
        Case 4
            new_nested_array trg(trg_ind(0), trg_ind(1), trg_ind(2), trg_ind(3), trg_ind(4)), next_ind
        
        Case 5
            new_nested_array trg(trg_ind(0), trg_ind(1), trg_ind(2), trg_ind(3), trg_ind(4), trg_ind(5)), next_ind
        
        Case 6
            new_nested_array trg(trg_ind(0), trg_ind(1), trg_ind(2), trg_ind(3), trg_ind(4), trg_ind(5), trg_ind(6)), next_ind
        
        Case 7
            new_nested_array trg(trg_ind(0), trg_ind(1), trg_ind(2), trg_ind(3), trg_ind(4), trg_ind(5), trg_ind(6), trg_ind(7)), next_ind
        
        Case 8
            new_nested_array trg(trg_ind(0), trg_ind(1), trg_ind(2), trg_ind(3), trg_ind(4), trg_ind(5), trg_ind(6), trg_ind(7), trg_ind(8)), next_ind
        
        Case 9
            new_nested_array trg(trg_ind(0), trg_ind(1), trg_ind(2), trg_ind(3), trg_ind(4), trg_ind(5), trg_ind(6), trg_ind(7), trg_ind(8), trg_ind(9)), next_ind
        
        Case Else
            'TODO Error
            
        End Select
            
                
    Loop While increment_index(trg_ind, trg_lb, trg_ub)

End Sub


Private Function create_new_array( _
    ArrayToCreate As Variant, _
    LBoundArray As Variant, _
    UBoundArray As Variant) As Boolean

    Dim i As Long
    Dim dim_count As Long
    Dim lblb As Long, lbub As Long
    Dim ublb As Long, ubub As Long

    create_new_array = False

    'TODO: Check LB <= UB
    'TODO: Check LB and UB are the same size
    'TODO: Allow LB and UB to be scalars
        
    'If Not dim_count_lb = dim_count_ub Then Exit Function

    lblb = LBound(LBoundArray)
    lbub = UBound(LBoundArray)
    ublb = LBound(UBoundArray)
    ubub = UBound(UBoundArray)
    
    
    dim_count = lbub - lblb + 1
    
    Select Case dim_count
    
    Case 1
        ReDim ArrayToCreate(LBoundArray(lblb) To UBoundArray(ublb))
    
    Case 2
        ReDim ArrayToCreate(LBoundArray(lblb) To UBoundArray(ublb), _
            LBoundArray(lblb + 1) To UBoundArray(ublb + 1))
    
    Case 3
        ReDim ArrayToCreate(LBoundArray(lblb) To UBoundArray(ublb), _
            LBoundArray(lblb + 1) To UBoundArray(ublb + 1), _
            LBoundArray(lblb + 2) To UBoundArray(ublb + 2))
    
    Case 4
        ReDim ArrayToCreate(LBoundArray(lblb) To UBoundArray(ublb), _
            LBoundArray(lblb + 1) To UBoundArray(ublb + 1), _
            LBoundArray(lblb + 2) To UBoundArray(ublb + 2), _
            LBoundArray(lblb + 3) To UBoundArray(ublb + 3))
    
    Case 5
        ReDim ArrayToCreate(LBoundArray(lblb) To UBoundArray(ublb), _
            LBoundArray(lblb + 1) To UBoundArray(ublb + 1), _
            LBoundArray(lblb + 2) To UBoundArray(ublb + 2), _
            LBoundArray(lblb + 3) To UBoundArray(ublb + 3), _
            LBoundArray(lblb + 4) To UBoundArray(ublb + 4))
        
    Case 6
        ReDim ArrayToCreate(LBoundArray(lblb) To UBoundArray(ublb), _
            LBoundArray(lblb + 1) To UBoundArray(ublb + 1), _
            LBoundArray(lblb + 2) To UBoundArray(ublb + 2), _
            LBoundArray(lblb + 3) To UBoundArray(ublb + 3), _
            LBoundArray(lblb + 4) To UBoundArray(ublb + 4), _
            LBoundArray(lblb + 5) To UBoundArray(ublb + 5))
    
    Case 7
        ReDim ArrayToCreate(LBoundArray(lblb) To UBoundArray(ublb), _
            LBoundArray(lblb + 1) To UBoundArray(ublb + 1), _
            LBoundArray(lblb + 2) To UBoundArray(ublb + 2), _
            LBoundArray(lblb + 3) To UBoundArray(ublb + 3), _
            LBoundArray(lblb + 4) To UBoundArray(ublb + 4), _
            LBoundArray(lblb + 5) To UBoundArray(ublb + 5), _
            LBoundArray(lblb + 6) To UBoundArray(ublb + 6))
    
    Case 8
        ReDim ArrayToCreate(LBoundArray(lblb) To UBoundArray(ublb), _
            LBoundArray(lblb + 1) To UBoundArray(ublb + 1), _
            LBoundArray(lblb + 2) To UBoundArray(ublb + 2), _
            LBoundArray(lblb + 3) To UBoundArray(ublb + 3), _
            LBoundArray(lblb + 4) To UBoundArray(ublb + 4), _
            LBoundArray(lblb + 5) To UBoundArray(ublb + 5), _
            LBoundArray(lblb + 6) To UBoundArray(ublb + 6), _
            LBoundArray(lblb + 7) To UBoundArray(ublb + 7))
    
    Case 9
        ReDim ArrayToCreate(LBoundArray(lblb) To UBoundArray(ublb), _
            LBoundArray(lblb + 1) To UBoundArray(ublb + 1), _
            LBoundArray(lblb + 2) To UBoundArray(ublb + 2), _
            LBoundArray(lblb + 3) To UBoundArray(ublb + 3), _
            LBoundArray(lblb + 4) To UBoundArray(ublb + 4), _
            LBoundArray(lblb + 5) To UBoundArray(ublb + 5), _
            LBoundArray(lblb + 6) To UBoundArray(ublb + 6), _
            LBoundArray(lblb + 7) To UBoundArray(ublb + 7), _
            LBoundArray(lblb + 8) To UBoundArray(ublb + 8))
            
    Case 10
        ReDim ArrayToCreate(LBoundArray(lblb) To UBoundArray(ublb), _
            LBoundArray(lblb + 1) To UBoundArray(ublb + 1), _
            LBoundArray(lblb + 2) To UBoundArray(ublb + 2), _
            LBoundArray(lblb + 3) To UBoundArray(ublb + 3), _
            LBoundArray(lblb + 4) To UBoundArray(ublb + 4), _
            LBoundArray(lblb + 5) To UBoundArray(ublb + 5), _
            LBoundArray(lblb + 6) To UBoundArray(ublb + 6), _
            LBoundArray(lblb + 7) To UBoundArray(ublb + 7), _
            LBoundArray(lblb + 8) To UBoundArray(ublb + 8), _
            LBoundArray(lblb + 9) To UBoundArray(ublb + 9))
    
    End Select
    
    create_new_array = True

End Function

Public Sub ResizeNestedArray(Target As Variant, ParamArray Index() As Variant)

    On Error GoTo HandleError:

    If UBound(Index) < LBound(Index) Then
        
        Err.Raise Number:=fmlInvalidArgument, _
                  Source:=PROJ_NAME, _
                  Description:=modErrInfo__.errstr_InvalidArgument
    
    End If

    If IsArray(Target) Then

        resize_nested_array_recursive Target, Index
        
    End If

    Exit Sub

HandleError:
    
    modErrInfo__.FuncID = id_ResizeNestedArray__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext

End Sub


Private Function resize_nested_array_recursive(trg, ByVal ubs) As Boolean

    Dim i As Long
    Dim ub As Long, lb As Long, size As Long
    Dim ubs_lb As Long, ubs_ub As Long
    Dim ubs_next
    Dim indice_lb, indice_ub, ind_next
        
    ubs_lb = LBound(ubs)
    ubs_ub = UBound(ubs)
    
    size = modArrSpt__.NumberOfArrayDimensions(trg)
    
    If size = 0 Then
    
        resize_nested_array_recursive = True
        Exit Function
                
    Else
    
        '--- Create Next ubs ---
        If ubs_lb < ubs_ub Then
        
            ReDim ubs_next(ubs_lb + 1 To ubs_ub)
        
            For i = ubs_lb + 1 To ubs_ub
            
                ubs_next(i) = ubs(i)
            
            Next i
            
        End If
    
        Select Case size
        
        Case 1
        
            modUtil__.get_bound indice_lb, indice_ub, trg, size
            
            ind_next = indice_lb
            
            If ubs_lb < ubs_ub Then
    
                Do
                    If IsArray(trg) Then
                    
                        resize_nested_array_recursive trg(ind_next(0)), ubs_next
                    
                    End If
                
                Loop While modUtil__.increment_index(ind_next, indice_lb, indice_ub)
                
            End If
                            
            ReDim Preserve trg(LBound(trg, 1) To ubs(ubs_lb))
            
        
        Case 2
        
            modUtil__.get_bound indice_lb, indice_ub, trg, size
            
            ind_next = indice_lb
    
            If ubs_lb < ubs_ub Then
    
                Do
                    If IsArray(trg) Then
                    
                        resize_nested_array_recursive trg(ind_next(0), _
                                                          ind_next(1)), ubs_next
                    
                    End If
                
                Loop While modUtil__.increment_index(ind_next, indice_lb, indice_ub)
        
            End If
        
            ReDim Preserve trg(LBound(trg, 1) To UBound(trg, 1), _
                               LBound(trg, 2) To ubs(ubs_lb))
                
        Case 3
            
            modUtil__.get_bound indice_lb, indice_ub, trg, size
            
            ind_next = indice_lb
    
            If ubs_lb < ubs_ub Then
    
                Do
                    If IsArray(trg) Then
                    
                        resize_nested_array_recursive trg(ind_next(0), _
                                                          ind_next(1), _
                                                          ind_next(2)), ubs_next
                    
                    End If
                
                Loop While modUtil__.increment_index(ind_next, indice_lb, indice_ub)
            
            End If
            
            ReDim Preserve trg(LBound(trg, 1) To UBound(trg, 1), _
                               LBound(trg, 2) To UBound(trg, 2), _
                               LBound(trg, 3) To ubs(ubs_lb))
        
        Case 4
        
            
            modUtil__.get_bound indice_lb, indice_ub, trg, size
            
            ind_next = indice_lb
    
            If ubs_lb < ubs_ub Then
    
                Do
                    If IsArray(trg) Then
                    
                        resize_nested_array_recursive trg(ind_next(0), _
                                                          ind_next(1), _
                                                          ind_next(2), _
                                                          ind_next(3)), ubs_next
                    
                    End If
                
                Loop While modUtil__.increment_index(ind_next, indice_lb, indice_ub)
            
            End If
            
            ReDim Preserve trg(LBound(trg, 1) To UBound(trg, 1), _
                               LBound(trg, 2) To UBound(trg, 2), _
                               LBound(trg, 3) To UBound(trg, 3), _
                               LBound(trg, 4) To ubs(ubs_lb))
        
        Case 5
        
            modUtil__.get_bound indice_lb, indice_ub, trg, size
            
            ind_next = indice_lb
    
            If ubs_lb < ubs_ub Then
    
                Do
                    If IsArray(trg) Then
                    
                        resize_nested_array_recursive trg(ind_next(0), _
                                                          ind_next(1), _
                                                          ind_next(2), _
                                                          ind_next(3), _
                                                          ind_next(4)), ubs_next
                    
                    End If
                
                Loop While modUtil__.increment_index(ind_next, indice_lb, indice_ub)
        
            End If
        
            ReDim Preserve trg(LBound(trg, 1) To UBound(trg, 1), _
                               LBound(trg, 2) To UBound(trg, 2), _
                               LBound(trg, 3) To UBound(trg, 3), _
                               LBound(trg, 4) To UBound(trg, 4), _
                               LBound(trg, 5) To ubs(ubs_lb))
        
        Case 6
        
            modUtil__.get_bound indice_lb, indice_ub, trg, size
            
            ind_next = indice_lb
    
            If ubs_lb < ubs_ub Then
    
                Do
                    If IsArray(trg) Then
                    
                        resize_nested_array_recursive trg(ind_next(0), _
                                                          ind_next(1), _
                                                          ind_next(2), _
                                                          ind_next(3), _
                                                          ind_next(4), _
                                                          ind_next(5)), ubs_next
                    
                    End If
                
                Loop While modUtil__.increment_index(ind_next, indice_lb, indice_ub)
            
            End If
        
            ReDim Preserve trg(LBound(trg, 1) To UBound(trg, 1), _
                               LBound(trg, 2) To UBound(trg, 2), _
                               LBound(trg, 3) To UBound(trg, 3), _
                               LBound(trg, 4) To UBound(trg, 4), _
                               LBound(trg, 5) To UBound(trg, 5), _
                               LBound(trg, 6) To ubs(ubs_lb))
        
        Case 7
        
            
            modUtil__.get_bound indice_lb, indice_ub, trg, size
            
            ind_next = indice_lb
            
            If ubs_lb < ubs_ub Then
    
                Do
                    If IsArray(trg) Then
                    
                        resize_nested_array_recursive trg(ind_next(0), _
                                                          ind_next(1), _
                                                          ind_next(2), _
                                                          ind_next(3), _
                                                          ind_next(4), _
                                                          ind_next(5), _
                                                          ind_next(6)), ubs_next
                    
                    End If
                
                Loop While modUtil__.increment_index(ind_next, indice_lb, indice_ub)
            
            End If
            
            ReDim Preserve trg(LBound(trg, 1) To UBound(trg, 1), _
                               LBound(trg, 2) To UBound(trg, 2), _
                               LBound(trg, 3) To UBound(trg, 3), _
                               LBound(trg, 4) To UBound(trg, 4), _
                               LBound(trg, 5) To UBound(trg, 5), _
                               LBound(trg, 6) To UBound(trg, 6), _
                               LBound(trg, 7) To ubs(ubs_lb))
        
        Case 8
        
            modUtil__.get_bound indice_lb, indice_ub, trg, size
            
            ind_next = indice_lb
    
            If ubs_lb < ubs_ub Then
    
                Do
                    If IsArray(trg) Then
                    
                        resize_nested_array_recursive trg(ind_next(0), _
                                                          ind_next(1), _
                                                          ind_next(2), _
                                                          ind_next(3), _
                                                          ind_next(4), _
                                                          ind_next(5), _
                                                          ind_next(6), _
                                                          ind_next(7)), ubs_next
                    
                    End If
                
                Loop While modUtil__.increment_index(ind_next, indice_lb, indice_ub)
        
            End If
        
            ReDim Preserve trg(LBound(trg, 1) To UBound(trg, 1), _
                               LBound(trg, 2) To UBound(trg, 2), _
                               LBound(trg, 3) To UBound(trg, 3), _
                               LBound(trg, 4) To UBound(trg, 4), _
                               LBound(trg, 5) To UBound(trg, 5), _
                               LBound(trg, 6) To UBound(trg, 6), _
                               LBound(trg, 7) To UBound(trg, 7), _
                               LBound(trg, 8) To ubs(ubs_lb))
        
        Case 9
        
            modUtil__.get_bound indice_lb, indice_ub, trg, size
            
            ind_next = indice_lb
    
            If ubs_lb < ubs_ub Then
    
                Do
                    If IsArray(trg) Then
                    
                        resize_nested_array_recursive trg(ind_next(0), _
                                                          ind_next(1), _
                                                          ind_next(2), _
                                                          ind_next(3), _
                                                          ind_next(4), _
                                                          ind_next(5), _
                                                          ind_next(6), _
                                                          ind_next(7), _
                                                          ind_next(8)), ubs_next
                    
                    End If
                
                Loop While modUtil__.increment_index(ind_next, indice_lb, indice_ub)
        
            End If
        
            ReDim Preserve trg(LBound(trg, 1) To UBound(trg, 1), _
                               LBound(trg, 2) To UBound(trg, 2), _
                               LBound(trg, 3) To UBound(trg, 3), _
                               LBound(trg, 4) To UBound(trg, 4), _
                               LBound(trg, 5) To UBound(trg, 5), _
                               LBound(trg, 6) To UBound(trg, 6), _
                               LBound(trg, 7) To UBound(trg, 7), _
                               LBound(trg, 8) To UBound(trg, 8), _
                               LBound(trg, 9) To ubs(ubs_lb))
        
        
        Case 10
        
            modUtil__.get_bound indice_lb, indice_ub, trg, size
            
            ind_next = indice_lb
    
            If ubs_lb < ubs_ub Then
    
                Do
                    If IsArray(trg) Then
                    
                        resize_nested_array_recursive trg(ind_next(0), _
                                                          ind_next(1), _
                                                          ind_next(2), _
                                                          ind_next(3), _
                                                          ind_next(4), _
                                                          ind_next(5), _
                                                          ind_next(6), _
                                                          ind_next(7), _
                                                          ind_next(8), _
                                                          ind_next(9)), ubs_next
                    
                    End If
                
                Loop While modUtil__.increment_index(ind_next, indice_lb, indice_ub)
        
            End If
        
            ReDim Preserve trg(LBound(trg, 1) To UBound(trg, 1), _
                               LBound(trg, 2) To UBound(trg, 2), _
                               LBound(trg, 3) To UBound(trg, 3), _
                               LBound(trg, 4) To UBound(trg, 4), _
                               LBound(trg, 5) To UBound(trg, 5), _
                               LBound(trg, 6) To UBound(trg, 6), _
                               LBound(trg, 7) To UBound(trg, 7), _
                               LBound(trg, 8) To UBound(trg, 8), _
                               LBound(trg, 9) To UBound(trg, 9), _
                               LBound(trg, 10) To ubs(ubs_lb))
        
        
        Case Else
        
            resize_nested_array_recursive = False
            Exit Function
        
        End Select
        
        
    End If

End Function


Public Sub MultDimArray( _
    Target As Variant, _
    ByVal Source As Variant, _
    FirstDim As fmlRangeDim, _
    ParamArray Index() As Variant)

    'Example of Index(): 0, Array(1, 10), 0, Array(1, 2)

    Dim ind As Variant
    Dim arr_in As Variant, arr_out As Variant
    Dim row_first_count As Long, row_last_count As Long
    Dim col_first_count As Long, col_last_count As Long
    Dim src_dim_count As Long    '=2
    Dim src_row_size As Long, src_col_size As Long
    Dim i As Long, j As Long
    Dim first_lb() As Long, first_ub() As Long
    Dim sec_lb() As Long, sec_ub() As Long
    
    Dim ind_ub As Long, sec_ind0 As Long
    
    Dim flag As Boolean
    
    On Error GoTo HandleError:
 
    ReDim first_lb(9) As Long
    ReDim first_ub(9) As Long
    ReDim sec_lb(9) As Long
    ReDim sec_ub(9) As Long
 
    
    If IsObject(Source) Then
    
        Source = Source
    
    End If
    
    'Check if Source is Array
    If Not IsArray(Source) Then
    
        Err.Raise Number:=fmlInvalidArgument, _
                  Source:=PROJ_NAME, _
                  Description:=modErrInfo__.errstr_InvalidArgument
        
    End If
    
    'Check if Source is Allocated
    If Not modArrSpt__.IsArrayAllocated(Source) Then
    
        Err.Raise Number:=fmlInvalidArgument, _
                  Source:=PROJ_NAME, _
                  Description:=modErrInfo__.errstr_InvalidArgument
        
    End If
    
    src_dim_count = modArrSpt__.NumberOfArrayDimensions(Source)
    
    'Set src sizes
    If src_dim_count = 1 Then
        
        src_row_size = UBound(Source, 1) - LBound(Source, 1) + 1
    
    ElseIf src_dim_count = 2 Then
    
        src_row_size = UBound(Source, 1) - LBound(Source, 1) + 1
        src_col_size = UBound(Source, 2) - LBound(Source, 2) + 1
    
    Else
    
        Err.Raise Number:=fmlInvalidArgument, _
                  Source:=PROJ_NAME, _
                  Description:=modErrInfo__.errstr_InvalidArgument
    
    End If
    
    
    'Check FirstDim
    If (Not FirstDim = fmlRow) And (Not FirstDim = fmlCol) Then
        
        Err.Raise Number:=fmlInvalidArgument, _
                  Source:=PROJ_NAME, _
                  Description:=modErrInfo__.errstr_InvalidArgument
    
    End If
    
    'If all indexes are omitted.
    If UBound(Index) < LBound(Index) Then
    
        If src_dim_count = 1 Then
        
            ind = Array(LBound(Source))
        
        Else
    
            If FirstDim = fmlRow Then
                ind = Array(LBound(Source, 1), LBound(Source, 2))
                
            Else
                ind = Array(LBound(Source, 2), LBound(Source, 1))
                
            End If
            
        End If
        
    Else
    
        ind = Index
        
    End If
    
    ind_ub = UBound(ind)
    
    '1st param in Index()
    If IsArray(ind(0)) Then
    
        Err.Raise Number:=fmlInvalidArgument, _
                  Source:=PROJ_NAME, _
                  Description:=modErrInfo__.errstr_InvalidArgument
    
    Else
                
        first_lb(0) = ind(0)
         
    End If
    
    'Analyze Index
    For i = 1 To ind_ub
    
        If Not IsArray(ind(i)) Then   'Fist of the second
        
            If Not sec_ind0 = 0 Then
            
                Err.Raise Number:=fmlInvalidArgument, _
                        Source:=PROJ_NAME, _
                        Description:=modErrInfo__.errstr_InvalidArgument
            Else
            
                sec_ind0 = i
                sec_lb(i - sec_ind0) = ind(i)
            
            End If
        
        Else
        
            If sec_ind0 = 0 Then
            
                first_lb(i) = ind(i)(0)
                first_ub(i) = ind(i)(1)
            
            Else
        
                sec_lb(i - sec_ind0) = ind(i)(0)
                sec_ub(i - sec_ind0) = ind(i)(1)

            End If
        
        End If
        
    Next i
    
    If sec_ind0 = 0 Then    'Either Row or Col Indexes are specified
    
        ReDim Preserve first_lb(UBound(ind)) As Long
        ReDim Preserve first_ub(UBound(ind)) As Long

        If src_dim_count = 2 Then 'If second indexes are omitted for 2dim source
        
            ReDim Preserve sec_lb(0) As Long
            ReDim Preserve sec_ub(0) As Long
            
            
            If FirstDim = fmlRow Then   'Use Source L/Ubounds for the second index.
            
                sec_lb(0) = LBound(Source, 2)
                sec_ub(0) = UBound(Source, 2)
                
            Else
                 
                sec_lb(0) = LBound(Source, 1)
                sec_ub(0) = UBound(Source, 1)
            
            End If
        
        End If
        
    Else    'Both Row and Col indexes are specified
    
        ReDim Preserve first_lb(sec_ind0 - 1) As Long
        ReDim Preserve first_ub(sec_ind0 - 1) As Long
        ReDim Preserve sec_lb(ind_ub - sec_ind0) As Long
        ReDim Preserve sec_ub(ind_ub - sec_ind0) As Long
    
    End If
    
    
    If src_dim_count = 1 Then
    
        If sec_ind0 <> 0 Then
        
            Err.Raise Number:=fmlInvalidArgument, _
                      Source:=PROJ_NAME, _
                      Description:=modErrInfo__.errstr_InvalidArgument
        
        End If
        
        If Not set_1st_ub(src_row_size, first_lb, first_ub) Then
        
            Err.Raise Number:=fmlInvalidArgument, _
                      Source:=PROJ_NAME, _
                      Description:=modErrInfo__.errstr_InvalidArgument
                      
        End If
        
        If Not restruct_vec2mult(Target, Source, first_lb, first_ub) Then
        
            Err.Raise Number:=fmlInvalidArgument, _
                      Source:=PROJ_NAME, _
                      Description:=modErrInfo__.errstr_InvalidArgument
        
        End If
    
    ElseIf src_dim_count = 2 Then
                                       
        If FirstDim = fmlRow Then
        
            If src_row_size <= 1 Then
        
                Err.Raise Number:=fmlInvalidArgument, _
                          Source:=PROJ_NAME, _
                          Description:=modErrInfo__.errstr_InvalidArgument
        
            End If
        
            flag = True
            flag = flag And set_1st_ub(src_row_size, first_lb, first_ub)
            flag = flag And set_1st_ub(src_col_size, sec_lb, sec_ub)
            flag = flag And restruct2mult_dim(Source, Target, first_lb, first_ub, sec_lb, sec_ub, FirstDim)
                
        Else
        
            If src_col_size <= 1 Then
        
                Err.Raise Number:=fmlInvalidArgument, _
                          Source:=PROJ_NAME, _
                          Description:=modErrInfo__.errstr_InvalidArgument
        
            End If
        
            flag = True
            flag = flag And set_1st_ub(src_row_size, sec_lb, sec_ub)
            flag = flag And set_1st_ub(src_col_size, first_lb, first_ub)
            flag = flag And restruct2mult_dim(Source, Target, sec_lb, sec_ub, first_lb, first_ub, FirstDim)
            
        End If
        
        If Not flag Then
        
            Err.Raise Number:=fmlInvalidArgument, _
                      Source:=PROJ_NAME, _
                      Description:=modErrInfo__.errstr_InvalidArgument
                
        End If
                
    Else    'Must not reach here.
    
        Err.Raise Number:=fmlInvalidArgument, _
                  Source:=PROJ_NAME, _
                  Description:=modErrInfo__.errstr_InvalidArgument
    
    End If
    
    Exit Sub
    
HandleError:
    
    modErrInfo__.FuncID = id_MultDimArray__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext
    
    
End Sub

Private Function set_1st_ub(total_size As Long, ind_lb As Variant, ByRef ind_ub As Variant) As Boolean

    Dim i As Long, j As Long
    Dim ind_count As Long

    ind_count = UBound(ind_lb) - LBound(ind_lb) + 1
    
    j = 1
    For i = 1 To ind_count - 1
    
        j = j * (ind_ub(i) - ind_lb(i) + 1)
        
    Next i
    
    If (total_size \ j < 1) Or (total_size Mod j <> 0) Then
        
        set_1st_ub = False
        Exit Function
    
    Else
    
        j = total_size \ j
    
    End If
    
    
    ind_ub(0) = j + ind_lb(0) - 1
    set_1st_ub = True

End Function


Private Function get_2dim_index( _
    rc1_lb As Variant, rc1_ub As Variant, _
    rc2_lb As Variant, rc2_ub As Variant, _
    ind_mult_dim As Variant, _
    ind_2dim As Variant) As Boolean

    '/// Take multi dim array index Array(i1, i2, ..., in)
    '/// Conver it to 2dim array index
    '/// Based on UB, LB info passed as other arguments.

    Dim dim_count1 As Long, dim_count2 As Long
    Dim i As Long, j As Long
    Dim dim_size As Long
    
    ReDim ind_2dim(1)
    
    dim_count1 = UBound(rc1_lb) - LBound(rc1_lb) + 1
    
    If IsEmpty(rc2_lb) Then
        dim_count2 = 0
        ind_2dim(1) = 0
    
    Else
        dim_count2 = UBound(rc2_lb) - LBound(rc2_lb) + 1
        
    End If
    
    For i = 0 To dim_count1 - 1
    
        If i = 0 Then
            dim_size = 1
        Else
            dim_size = dim_size * (rc1_ub(dim_count1 - i) - rc1_lb(dim_count1 - i) + 1)
        End If
        
        ind_2dim(0) = ind_2dim(0) + (ind_mult_dim(dim_count1 - i - 1) - rc1_lb(dim_count1 - i - 1)) * dim_size
        
    Next i
    
    For i = 0 To dim_count2 - 1
    
        If i = 0 Then
            dim_size = 1
        Else
            dim_size = dim_size * (rc2_ub(dim_count2 - i) - rc2_lb(dim_count2 - i) + 1)
        End If
        
        ind_2dim(1) = ind_2dim(1) + (ind_mult_dim(dim_count1 + dim_count2 - i - 1) - rc2_lb(dim_count2 - i - 1)) * dim_size
        
    Next i

End Function


Private Function arrayalize_scalar(ByRef ind As Variant) As Long
'
'   Change ind into array with 1 element if it is a scalar.
'   retun the length of the array.
'
    If IsEmpty(ind) Then
        arrayalize_scalar = 0

    ElseIf IsArray(ind) Then
        arrayalize_scalar = UBound(ind) - LBound(ind) + 1
    
    Else
        ind = Array(ind)
        arrayalize_scalar = 1
    End If

End Function

Private Function restruct_vec2mult(arr_out As Variant, arr_in As Variant, ind_lb As Variant, ind_ub As Variant) As Boolean

    Dim dim_count As Long, arr_in_lb As Long
    Dim i As Long
    Dim i0 As Long, i1 As Long, i2 As Long, i3 As Long, i4 As Long, i5 As Long, i6 As Long, i7 As Long, i8 As Long, i9 As Long
    Dim dim_size() As Long

    dim_count = UBound(ind_lb) + 1
    arr_in_lb = LBound(arr_in)
    
    ReDim dim_size(dim_count - 1)
    
    'Set dim_size()
    For i = dim_count - 1 To 0 Step -1
    
        If i = dim_count - 1 Then
        
            dim_size(i) = 1
            
        Else
    
            dim_size(i) = dim_size(i + 1) * (ind_ub(i + 1) - ind_lb(i + 1) + 1)
            
        End If
    
    Next i
    
    'Set arr_out()
    Select Case dim_count
    
    Case 1
    
        ReDim arr_out(ind_lb(0) To ind_ub(0))
        
        For i0 = ind_lb(0) To ind_ub(0)
            
            arr_out(i0) = arr_in(arr_in_lb + _
                                (i0 - ind_lb(0)))
        
        Next i0
        
    Case 2
    
        ReDim arr_out(ind_lb(0) To ind_ub(0), _
                      ind_lb(1) To ind_ub(1))
    
        For i0 = ind_lb(0) To ind_ub(0)
        For i1 = ind_lb(1) To ind_ub(1)
            
            arr_out(i0, i1) = arr_in(arr_in_lb + _
                            dim_size(0) * (i0 - ind_lb(0)) + _
                            dim_size(1) * (i1 - ind_lb(1)))
        
        Next i1
        Next i0
   
    
    Case 3
    
        ReDim arr_out(ind_lb(0) To ind_ub(0), _
                      ind_lb(1) To ind_ub(1), _
                      ind_lb(2) To ind_ub(2))
    
        For i0 = ind_lb(0) To ind_ub(0)
        For i1 = ind_lb(1) To ind_ub(1)
        For i2 = ind_lb(2) To ind_ub(2)
            
            arr_out(i0, i1, i2) = arr_in(arr_in_lb + _
                            dim_size(0) * (i0 - ind_lb(0)) + _
                            dim_size(1) * (i1 - ind_lb(1)) + _
                            dim_size(2) * (i2 - ind_lb(2)))
        
        Next i2
        Next i1
        Next i0
    
    
    Case 4
    
        ReDim arr_out(ind_lb(0) To ind_ub(0), _
                      ind_lb(1) To ind_ub(1), _
                      ind_lb(2) To ind_ub(2), _
                      ind_lb(3) To ind_ub(3))
    
        For i0 = ind_lb(0) To ind_ub(0)
        For i1 = ind_lb(1) To ind_ub(1)
        For i2 = ind_lb(2) To ind_ub(2)
        For i3 = ind_lb(3) To ind_ub(3)
            
            arr_out(i0, i1, i2, i3) = arr_in(arr_in_lb + _
                            dim_size(0) * (i0 - ind_lb(0)) + _
                            dim_size(1) * (i1 - ind_lb(1)) + _
                            dim_size(2) * (i2 - ind_lb(2)) + _
                            dim_size(3) * (i3 - ind_lb(3)))
        Next i3
        Next i2
        Next i1
        Next i0
    
    
    Case 5
    
        ReDim arr_out(ind_lb(0) To ind_ub(0), _
                      ind_lb(1) To ind_ub(1), _
                      ind_lb(2) To ind_ub(2), _
                      ind_lb(3) To ind_ub(3), _
                      ind_lb(4) To ind_ub(4))
    
        For i0 = ind_lb(0) To ind_ub(0)
        For i1 = ind_lb(1) To ind_ub(1)
        For i2 = ind_lb(2) To ind_ub(2)
        For i3 = ind_lb(3) To ind_ub(3)
        For i4 = ind_lb(4) To ind_ub(4)
            
            arr_out(i0, i1, i2, i3, i4) = arr_in(arr_in_lb + _
                            dim_size(0) * (i0 - ind_lb(0)) + _
                            dim_size(1) * (i1 - ind_lb(1)) + _
                            dim_size(2) * (i2 - ind_lb(2)) + _
                            dim_size(3) * (i3 - ind_lb(3)) + _
                            dim_size(4) * (i4 - ind_lb(4)))
        Next i4
        Next i3
        Next i2
        Next i1
        Next i0
    
    
    Case 6
    
        ReDim arr_out(ind_lb(0) To ind_ub(0), _
                      ind_lb(1) To ind_ub(1), _
                      ind_lb(2) To ind_ub(2), _
                      ind_lb(3) To ind_ub(3), _
                      ind_lb(4) To ind_ub(4))
    
        For i0 = ind_lb(0) To ind_ub(0)
        For i1 = ind_lb(1) To ind_ub(1)
        For i2 = ind_lb(2) To ind_ub(2)
        For i3 = ind_lb(3) To ind_ub(3)
        For i4 = ind_lb(4) To ind_ub(4)
        For i5 = ind_lb(5) To ind_ub(5)
            
            arr_out(i0, i1, i2, i3, i4, i5) = arr_in(arr_in_lb + _
                            dim_size(0) * (i0 - ind_lb(0)) + _
                            dim_size(1) * (i1 - ind_lb(1)) + _
                            dim_size(2) * (i2 - ind_lb(2)) + _
                            dim_size(3) * (i3 - ind_lb(3)) + _
                            dim_size(4) * (i4 - ind_lb(4)) + _
                            dim_size(5) * (i5 - ind_lb(5)))
        Next i5
        Next i4
        Next i3
        Next i2
        Next i1
        Next i0
    
    Case 7
    
        ReDim arr_out(ind_lb(0) To ind_ub(0), _
                      ind_lb(1) To ind_ub(1), _
                      ind_lb(2) To ind_ub(2), _
                      ind_lb(3) To ind_ub(3), _
                      ind_lb(4) To ind_ub(4), _
                      ind_lb(5) To ind_ub(5))
    
        For i0 = ind_lb(0) To ind_ub(0)
        For i1 = ind_lb(1) To ind_ub(1)
        For i2 = ind_lb(2) To ind_ub(2)
        For i3 = ind_lb(3) To ind_ub(3)
        For i4 = ind_lb(4) To ind_ub(4)
        For i5 = ind_lb(5) To ind_ub(5)
        For i6 = ind_lb(6) To ind_ub(6)
            
            arr_out(i0, i1, i2, i3, i4, i5, i6) = arr_in(arr_in_lb + _
                            dim_size(0) * (i0 - ind_lb(0)) + _
                            dim_size(1) * (i1 - ind_lb(1)) + _
                            dim_size(2) * (i2 - ind_lb(2)) + _
                            dim_size(3) * (i3 - ind_lb(3)) + _
                            dim_size(4) * (i4 - ind_lb(4)) + _
                            dim_size(5) * (i5 - ind_lb(5)) + _
                            dim_size(6) * (i6 - ind_lb(6)))
        Next i6
        Next i5
        Next i4
        Next i3
        Next i2
        Next i1
        Next i0
    
    Case 8
    
        ReDim arr_out(ind_lb(0) To ind_ub(0), _
                      ind_lb(1) To ind_ub(1), _
                      ind_lb(2) To ind_ub(2), _
                      ind_lb(3) To ind_ub(3), _
                      ind_lb(4) To ind_ub(4), _
                      ind_lb(5) To ind_ub(5), _
                      ind_lb(6) To ind_ub(6), _
                      ind_lb(7) To ind_ub(7))
    
        For i0 = ind_lb(0) To ind_ub(0)
        For i1 = ind_lb(1) To ind_ub(1)
        For i2 = ind_lb(2) To ind_ub(2)
        For i3 = ind_lb(3) To ind_ub(3)
        For i4 = ind_lb(4) To ind_ub(4)
        For i5 = ind_lb(5) To ind_ub(5)
        For i6 = ind_lb(6) To ind_ub(6)
        For i7 = ind_lb(7) To ind_ub(7)
            
            arr_out(i0, i1, i2, i3, i4, i5, i6, i7) = arr_in(arr_in_lb + _
                            dim_size(0) * (i0 - ind_lb(0)) + _
                            dim_size(1) * (i1 - ind_lb(1)) + _
                            dim_size(2) * (i2 - ind_lb(2)) + _
                            dim_size(3) * (i3 - ind_lb(3)) + _
                            dim_size(4) * (i4 - ind_lb(4)) + _
                            dim_size(5) * (i5 - ind_lb(5)) + _
                            dim_size(6) * (i6 - ind_lb(6)) + _
                            dim_size(7) * (i7 - ind_lb(7)))
        Next i7
        Next i6
        Next i5
        Next i4
        Next i3
        Next i2
        Next i1
        Next i0
    
    Case 9
    
        ReDim arr_out(ind_lb(0) To ind_ub(0), _
                      ind_lb(1) To ind_ub(1), _
                      ind_lb(2) To ind_ub(2), _
                      ind_lb(3) To ind_ub(3), _
                      ind_lb(4) To ind_ub(4), _
                      ind_lb(5) To ind_ub(5), _
                      ind_lb(6) To ind_ub(6), _
                      ind_lb(7) To ind_ub(7), _
                      ind_lb(8) To ind_ub(8))
    
        For i0 = ind_lb(0) To ind_ub(0)
        For i1 = ind_lb(1) To ind_ub(1)
        For i2 = ind_lb(2) To ind_ub(2)
        For i3 = ind_lb(3) To ind_ub(3)
        For i4 = ind_lb(4) To ind_ub(4)
        For i5 = ind_lb(5) To ind_ub(5)
        For i6 = ind_lb(6) To ind_ub(6)
        For i7 = ind_lb(7) To ind_ub(7)
        For i8 = ind_lb(8) To ind_ub(8)
            
            arr_out(i0, i1, i2, i3, i4, i5, i6, i7, i8) = arr_in(arr_in_lb + _
                            dim_size(0) * (i0 - ind_lb(0)) + _
                            dim_size(1) * (i1 - ind_lb(1)) + _
                            dim_size(2) * (i2 - ind_lb(2)) + _
                            dim_size(3) * (i3 - ind_lb(3)) + _
                            dim_size(4) * (i4 - ind_lb(4)) + _
                            dim_size(5) * (i5 - ind_lb(5)) + _
                            dim_size(6) * (i6 - ind_lb(6)) + _
                            dim_size(7) * (i7 - ind_lb(7)) + _
                            dim_size(8) * (i8 - ind_lb(8)))
        Next i8
        Next i7
        Next i6
        Next i5
        Next i4
        Next i3
        Next i2
        Next i1
        Next i0
    
    Case 10
    
        ReDim arr_out(ind_lb(0) To ind_ub(0), _
                      ind_lb(1) To ind_ub(1), _
                      ind_lb(2) To ind_ub(2), _
                      ind_lb(3) To ind_ub(3), _
                      ind_lb(4) To ind_ub(4), _
                      ind_lb(5) To ind_ub(5), _
                      ind_lb(6) To ind_ub(6), _
                      ind_lb(7) To ind_ub(7), _
                      ind_lb(8) To ind_ub(8), _
                      ind_lb(9) To ind_ub(9))
    
        For i0 = ind_lb(0) To ind_ub(0)
        For i1 = ind_lb(1) To ind_ub(1)
        For i2 = ind_lb(2) To ind_ub(2)
        For i3 = ind_lb(3) To ind_ub(3)
        For i4 = ind_lb(4) To ind_ub(4)
        For i5 = ind_lb(5) To ind_ub(5)
        For i6 = ind_lb(6) To ind_ub(6)
        For i7 = ind_lb(7) To ind_ub(7)
        For i8 = ind_lb(8) To ind_ub(8)
        For i9 = ind_lb(9) To ind_ub(9)
            
            arr_out(i0, i1, i2, i3, i4, i5, i6, i7, i8, i9) = arr_in(arr_in_lb + _
                            dim_size(0) * (i0 - ind_lb(0)) + _
                            dim_size(1) * (i1 - ind_lb(1)) + _
                            dim_size(2) * (i2 - ind_lb(2)) + _
                            dim_size(3) * (i3 - ind_lb(3)) + _
                            dim_size(4) * (i4 - ind_lb(4)) + _
                            dim_size(5) * (i5 - ind_lb(5)) + _
                            dim_size(6) * (i6 - ind_lb(6)) + _
                            dim_size(7) * (i7 - ind_lb(7)) + _
                            dim_size(8) * (i8 - ind_lb(8)) + _
                            dim_size(9) * (i9 - ind_lb(9)))
        Next i9
        Next i8
        Next i7
        Next i6
        Next i5
        Next i4
        Next i3
        Next i2
        Next i1
        Next i0
                      
    Case Else
        
        restruct_vec2mult = False
        Exit Function

    End Select

    restruct_vec2mult = True
    
End Function

Private Function restruct2mult_dim( _
    ByRef arr_in As Variant, _
    ByRef arr_out As Variant, _
    row_ind_lb As Variant, _
    row_ind_ub As Variant, _
    col_ind_lb As Variant, _
    col_ind_ub As Variant, _
    dim_order As fmlRangeDim) As Boolean
    
    'Dim dim_ind(10) As Long
    
    Dim i0 As Long, _
        i1 As Long, _
        i2 As Long, _
        i3 As Long, _
        i4 As Long, _
        i5 As Long, _
        i6 As Long, _
        i7 As Long, _
        i8 As Long, _
        i9 As Long
        
    Dim ind_2d As Variant
    Dim dim_count As Long, row_dim_count As Long, col_dim_count As Long
    Dim ind_lb As Variant, ind_ub As Variant
    Dim first_ind_lb As Variant, sec_ind_lb As Variant
    Dim first_ind_ub As Variant, sec_ind_ub As Variant
    
    restruct2mult_dim = False
    
    row_dim_count = UBound(row_ind_lb) - LBound(row_ind_lb) + 1
    col_dim_count = UBound(col_ind_lb) - LBound(col_ind_lb) + 1
    
    If LBound(arr_in) = UBound(arr_in) _
        And LBound(arr_in, 2) = UBound(arr_in, 2) Then
        
        Exit Function
    
    ElseIf LBound(arr_in) = UBound(arr_in) Then 'Single Row Vector
    
        dim_count = col_dim_count
    
        ind_lb = col_ind_lb
        ind_ub = col_ind_ub

        first_ind_lb = col_ind_lb
        first_ind_ub = col_ind_ub

        sec_ind_lb = Empty
        sec_ind_ub = Empty
    
    ElseIf LBound(arr_in, 2) = UBound(arr_in, 2) Then 'Single Col Vector
    
        dim_count = row_dim_count
    
        ind_lb = row_ind_lb
        ind_ub = row_ind_ub
        
        first_ind_lb = row_ind_lb
        first_ind_ub = row_ind_ub
                
        sec_ind_lb = Empty
        sec_ind_ub = Empty
        
    ElseIf dim_order = fmlRow Then
    
        dim_count = row_dim_count + col_dim_count
        
        ind_lb = row_ind_lb
        ind_ub = row_ind_ub
        
        first_ind_lb = row_ind_lb
        first_ind_ub = row_ind_ub
        
        sec_ind_lb = col_ind_lb
        sec_ind_ub = col_ind_ub
        
        If Not modArrSpt__.ConcatenateArrays(ind_lb, col_ind_lb, True) Then Exit Function
        If Not modArrSpt__.ConcatenateArrays(ind_ub, col_ind_ub, True) Then Exit Function
        
    ElseIf dim_order = fmlCol Then
    
        dim_count = row_dim_count + col_dim_count
        
        ind_lb = col_ind_lb
        ind_ub = col_ind_ub
        
        first_ind_lb = col_ind_lb
        first_ind_ub = col_ind_ub
        
        sec_ind_lb = row_ind_lb
        sec_ind_ub = row_ind_ub
        
        If Not modArrSpt__.ConcatenateArrays(ind_lb, row_ind_lb, True) Then Exit Function
        If Not modArrSpt__.ConcatenateArrays(ind_ub, row_ind_ub, True) Then Exit Function
    
    End If

    Select Case dim_count
    Case 1
        ReDim arr_out(ind_lb(0) To ind_ub(0))
        
        For i0 = ind_lb(0) To ind_ub(0)
        
            Call get_2dim_index(first_ind_lb, first_ind_ub, sec_ind_lb, sec_ind_ub, Array(i0, i1), ind_2d)
        
            If LBound(arr_in) = UBound(arr_in) Then
                arr_out(i0) = arr_in(LBound(arr_in, 1), ind_2d(0) + LBound(arr_in, 2))
            
            ElseIf LBound(arr_in, 2) = UBound(arr_in, 2) Then
                arr_out(i0) = arr_in(ind_2d(0) + LBound(arr_in, 1), LBound(arr_in, 2))
            
            ElseIf dim_order = fmlRow Then
                arr_out(i0) = arr_in(ind_2d(0) + LBound(arr_in, 1), ind_2d(1) + LBound(arr_in, 2))
                
            Else
                arr_out(i0) = arr_in(ind_2d(1) + LBound(arr_in, 1), ind_2d(0) + LBound(arr_in, 2))
            End If
            
        Next i0
    
    Case 2
        ReDim arr_out(ind_lb(0) To ind_ub(0), _
                  ind_lb(1) To ind_ub(1))
        
        For i0 = ind_lb(0) To ind_ub(0)
        For i1 = ind_lb(1) To ind_ub(1)
        
            Call get_2dim_index(first_ind_lb, first_ind_ub, sec_ind_lb, sec_ind_ub, Array(i0, i1), ind_2d)
            
            If LBound(arr_in) = UBound(arr_in) Then
                arr_out(i0, i1) = arr_in(LBound(arr_in, 1), ind_2d(0) + LBound(arr_in, 2))
            
            ElseIf LBound(arr_in, 2) = UBound(arr_in, 2) Then
                arr_out(i0, i1) = arr_in(ind_2d(0) + LBound(arr_in, 1), LBound(arr_in, 2))
            
            ElseIf dim_order = fmlRow Then
                arr_out(i0, i1) = arr_in(ind_2d(0) + LBound(arr_in, 1), ind_2d(1) + LBound(arr_in, 2))
                
            Else
                arr_out(i0, i1) = arr_in(ind_2d(1) + LBound(arr_in, 1), ind_2d(0) + LBound(arr_in, 2))
            End If
        
        Next i1
        Next i0
        
    Case 3
        ReDim arr_out(ind_lb(0) To ind_ub(0), _
                  ind_lb(1) To ind_ub(1), _
                  ind_lb(2) To ind_ub(2))
                  
        For i0 = ind_lb(0) To ind_ub(0)
        For i1 = ind_lb(1) To ind_ub(1)
        For i2 = ind_lb(2) To ind_ub(2)
                
            Call get_2dim_index(first_ind_lb, first_ind_ub, sec_ind_lb, sec_ind_ub, _
                Array(i0, i1, i2), ind_2d)
                
            If LBound(arr_in) = UBound(arr_in) Then
                arr_out(i0, i1, i2) = arr_in(LBound(arr_in, 1), ind_2d(0) + LBound(arr_in, 2))
            
            ElseIf LBound(arr_in, 2) = UBound(arr_in, 2) Then
                arr_out(i0, i1, i2) = arr_in(ind_2d(0) + LBound(arr_in, 1), LBound(arr_in, 2))
            
            ElseIf dim_order = fmlRow Then
                arr_out(i0, i1, i2) = arr_in(ind_2d(0) + LBound(arr_in, 1), ind_2d(1) + LBound(arr_in, 2))
                
            Else
                arr_out(i0, i1, i2) = arr_in(ind_2d(1) + LBound(arr_in, 1), ind_2d(0) + LBound(arr_in, 2))
            End If
                
        Next i2
        Next i1
        Next i0


    Case 4
        ReDim arr_out(ind_lb(0) To ind_ub(0), _
                  ind_lb(1) To ind_ub(1), _
                  ind_lb(2) To ind_ub(2), _
                  ind_lb(3) To ind_ub(3))
    
        For i0 = ind_lb(0) To ind_ub(0)
        For i1 = ind_lb(1) To ind_ub(1)
        For i2 = ind_lb(2) To ind_ub(2)
        For i3 = ind_lb(3) To ind_ub(3)
            
            Call get_2dim_index(first_ind_lb, first_ind_ub, sec_ind_lb, sec_ind_ub, _
                Array(i0, i1, i2, i3), ind_2d)
            
            If LBound(arr_in) = UBound(arr_in) Then
                arr_out(i0, i1, i2, i3) = arr_in(LBound(arr_in, 1), ind_2d(0) + LBound(arr_in, 2))
            
            ElseIf LBound(arr_in, 2) = UBound(arr_in, 2) Then
                arr_out(i0, i1, i2, i3) = arr_in(ind_2d(0) + LBound(arr_in, 1), LBound(arr_in, 2))
            
            ElseIf dim_order = fmlRow Then
                arr_out(i0, i1, i2, i3) = arr_in(ind_2d(0) + LBound(arr_in, 1), ind_2d(1) + LBound(arr_in, 2))
                
            Else
                arr_out(i0, i1, i2, i3) = arr_in(ind_2d(1) + LBound(arr_in, 1), ind_2d(0) + LBound(arr_in, 2))
            End If
       
        Next i3
        Next i2
        Next i1
        Next i0
    
    Case 5
        ReDim arr_out(ind_lb(0) To ind_ub(0), _
                  ind_lb(1) To ind_ub(1), _
                  ind_lb(2) To ind_ub(2), _
                  ind_lb(3) To ind_ub(3), _
                  ind_lb(4) To ind_ub(4))
    
        For i0 = ind_lb(0) To ind_ub(0)
        For i1 = ind_lb(1) To ind_ub(1)
        For i2 = ind_lb(2) To ind_ub(2)
        For i3 = ind_lb(3) To ind_ub(3)
        For i4 = ind_lb(4) To ind_ub(4)
       
            Call get_2dim_index(first_ind_lb, first_ind_ub, sec_ind_lb, sec_ind_ub, _
                Array(i0, i1, i2, i3, i4), ind_2d)
       
            If LBound(arr_in) = UBound(arr_in) Then
                arr_out(i0, i1, i2, i3, i4) = arr_in(LBound(arr_in, 1), ind_2d(0) + LBound(arr_in, 2))
            
            ElseIf LBound(arr_in, 2) = UBound(arr_in, 2) Then
                arr_out(i0, i1, i2, i3, i4) = arr_in(ind_2d(0) + LBound(arr_in, 1), LBound(arr_in, 2))
            
            ElseIf dim_order = fmlRow Then
                arr_out(i0, i1, i2, i3, i4) = arr_in(ind_2d(0) + LBound(arr_in, 1), ind_2d(1) + LBound(arr_in, 2))
                
            Else
                arr_out(i0, i1, i2, i3, i4) = arr_in(ind_2d(1) + LBound(arr_in, 1), ind_2d(0) + LBound(arr_in, 2))
            End If
       
        Next i4
        Next i3
        Next i2
        Next i1
        Next i0
    
    
    Case 6
        ReDim arr_out(ind_lb(0) To ind_ub(0), _
                  ind_lb(1) To ind_ub(1), _
                  ind_lb(2) To ind_ub(2), _
                  ind_lb(3) To ind_ub(3), _
                  ind_lb(4) To ind_ub(4), _
                  ind_lb(5) To ind_ub(5))
    
        For i0 = ind_lb(0) To ind_ub(0)
        For i1 = ind_lb(1) To ind_ub(1)
        For i2 = ind_lb(2) To ind_ub(2)
        For i3 = ind_lb(3) To ind_ub(3)
        For i4 = ind_lb(4) To ind_ub(4)
        For i5 = ind_lb(5) To ind_ub(5)
             
             
            Call get_2dim_index(first_ind_lb, first_ind_ub, sec_ind_lb, sec_ind_ub, _
                Array(i0, i1, i2, i3, i4, i5), ind_2d)
       
            If LBound(arr_in) = UBound(arr_in) Then
                arr_out(i0, i1, i2, i3, i4, i5) = arr_in(LBound(arr_in, 1), ind_2d(0) + LBound(arr_in, 2))
            
            ElseIf LBound(arr_in, 2) = UBound(arr_in, 2) Then
                arr_out(i0, i1, i2, i3, i4, i5) = arr_in(ind_2d(0) + LBound(arr_in, 1), LBound(arr_in, 2))
            
            ElseIf dim_order = fmlRow Then
                arr_out(i0, i1, i2, i3, i4, i5) = arr_in(ind_2d(0) + LBound(arr_in, 1), ind_2d(1) + LBound(arr_in, 2))
                
            Else
                arr_out(i0, i1, i2, i3, i4, i5) = arr_in(ind_2d(1) + LBound(arr_in, 1), ind_2d(0) + LBound(arr_in, 2))
            End If
       
        Next i5
        Next i4
        Next i3
        Next i2
        Next i1
        Next i0
    
    
    Case 7
        ReDim arr_out(ind_lb(0) To ind_ub(0), _
                  ind_lb(1) To ind_ub(1), _
                  ind_lb(2) To ind_ub(2), _
                  ind_lb(3) To ind_ub(3), _
                  ind_lb(4) To ind_ub(4), _
                  ind_lb(5) To ind_ub(5), _
                  ind_lb(6) To ind_ub(6))
    
        For i0 = ind_lb(0) To ind_ub(0)
        For i1 = ind_lb(1) To ind_ub(1)
        For i2 = ind_lb(2) To ind_ub(2)
        For i3 = ind_lb(3) To ind_ub(3)
        For i4 = ind_lb(4) To ind_ub(4)
        For i5 = ind_lb(5) To ind_ub(5)
        For i6 = ind_lb(6) To ind_ub(6)
            
            
            Call get_2dim_index(first_ind_lb, first_ind_ub, sec_ind_lb, sec_ind_ub, _
                Array(i0, i1, i2, i3, i4, i5, i6), ind_2d)
            
            If LBound(arr_in) = UBound(arr_in) Then
                arr_out(i0, i1, i2, i3, i4, i5, i6) = arr_in(LBound(arr_in, 1), ind_2d(0) + LBound(arr_in, 2))
            
            ElseIf LBound(arr_in, 2) = UBound(arr_in, 2) Then
                arr_out(i0, i1, i2, i3, i4, i5, i6) = arr_in(ind_2d(0) + LBound(arr_in, 1), LBound(arr_in, 2))
            
            ElseIf dim_order = fmlRow Then
                arr_out(i0, i1, i2, i3, i4, i5, i6) = arr_in(ind_2d(0) + LBound(arr_in, 1), ind_2d(1) + LBound(arr_in, 2))
                
            Else
                arr_out(i0, i1, i2, i3, i4, i5, i6) = arr_in(ind_2d(1) + LBound(arr_in, 1), ind_2d(0) + LBound(arr_in, 2))
            End If

        
        Next i6
        Next i5
        Next i4
        Next i3
        Next i2
        Next i1
        Next i0
    
    
    Case 8
        ReDim arr_out(ind_lb(0) To ind_ub(0), _
                  ind_lb(1) To ind_ub(1), _
                  ind_lb(2) To ind_ub(2), _
                  ind_lb(3) To ind_ub(3), _
                  ind_lb(4) To ind_ub(4), _
                  ind_lb(5) To ind_ub(5), _
                  ind_lb(6) To ind_ub(6), _
                  ind_lb(7) To ind_ub(7))
    
        For i0 = ind_lb(0) To ind_ub(0)
        For i1 = ind_lb(1) To ind_ub(1)
        For i2 = ind_lb(2) To ind_ub(2)
        For i3 = ind_lb(3) To ind_ub(3)
        For i4 = ind_lb(4) To ind_ub(4)
        For i5 = ind_lb(5) To ind_ub(5)
        For i6 = ind_lb(6) To ind_ub(6)
        For i7 = ind_lb(7) To ind_ub(7)
        
            Call get_2dim_index(first_ind_lb, first_ind_ub, sec_ind_lb, sec_ind_ub, _
                Array(i0, i1, i2, i3, i4, i5, i6, i7), ind_2d)
            
            If LBound(arr_in) = UBound(arr_in) Then
                arr_out(i0, i1, i2, i3, i4, i5, i6, i7) = arr_in(LBound(arr_in, 1), ind_2d(0) + LBound(arr_in, 2))
            
            ElseIf LBound(arr_in, 2) = UBound(arr_in, 2) Then
                arr_out(i0, i1, i2, i3, i4, i5, i6, i7) = arr_in(ind_2d(0) + LBound(arr_in, 1), LBound(arr_in, 2))
            
            ElseIf dim_order = fmlRow Then
                arr_out(i0, i1, i2, i3, i4, i5, i6, i7) = arr_in(ind_2d(0) + LBound(arr_in, 1), ind_2d(1) + LBound(arr_in, 2))
                
            Else
                arr_out(i0, i1, i2, i3, i4, i5, i6, i7) = arr_in(ind_2d(1) + LBound(arr_in, 1), ind_2d(0) + LBound(arr_in, 2))
            End If
        
        Next i7
        Next i6
        Next i5
        Next i4
        Next i3
        Next i2
        Next i1
        Next i0
    
    
    Case 9
        ReDim arr_out(ind_lb(0) To ind_ub(0), _
                  ind_lb(1) To ind_ub(1), _
                  ind_lb(2) To ind_ub(2), _
                  ind_lb(3) To ind_ub(3), _
                  ind_lb(4) To ind_ub(4), _
                  ind_lb(5) To ind_ub(5), _
                  ind_lb(6) To ind_ub(6), _
                  ind_lb(7) To ind_ub(7), _
                  ind_lb(8) To ind_ub(8))
    
        For i0 = ind_lb(0) To ind_ub(0)
        For i1 = ind_lb(1) To ind_ub(1)
        For i2 = ind_lb(2) To ind_ub(2)
        For i3 = ind_lb(3) To ind_ub(3)
        For i4 = ind_lb(4) To ind_ub(4)
        For i5 = ind_lb(5) To ind_ub(5)
        For i6 = ind_lb(6) To ind_ub(6)
        For i7 = ind_lb(7) To ind_ub(7)
        For i8 = ind_lb(8) To ind_ub(8)
        
        
            Call get_2dim_index(first_ind_lb, first_ind_ub, sec_ind_lb, sec_ind_ub, _
                Array(i0, i1, i2, i3, i4, i5, i6, i7, i8), ind_2d)
        
        
            If LBound(arr_in) = UBound(arr_in) Then
                arr_out(i0, i1, i2, i3, i4, i5, i6, i7, i8) = arr_in(LBound(arr_in, 1), ind_2d(0) + LBound(arr_in, 2))
            
            ElseIf LBound(arr_in, 2) = UBound(arr_in, 2) Then
                arr_out(i0, i1, i2, i3, i4, i5, i6, i7, i8) = arr_in(ind_2d(0) + LBound(arr_in, 1), LBound(arr_in, 2))
            
            ElseIf dim_order = fmlRow Then
                arr_out(i0, i1, i2, i3, i4, i5, i6, i7, i8) = arr_in(ind_2d(0) + LBound(arr_in, 1), ind_2d(1) + LBound(arr_in, 2))
                
            Else
                arr_out(i0, i1, i2, i3, i4, i5, i6, i7, i8) = arr_in(ind_2d(1) + LBound(arr_in, 1), ind_2d(0) + LBound(arr_in, 2))
            End If
        
        
        
        Next i8
        Next i7
        Next i6
        Next i5
        Next i4
        Next i3
        Next i2
        Next i1
        Next i0
    
    
    Case 10
        ReDim arr_out(ind_lb(0) To ind_ub(0), _
                  ind_lb(1) To ind_ub(1), _
                  ind_lb(2) To ind_ub(2), _
                  ind_lb(3) To ind_ub(3), _
                  ind_lb(4) To ind_ub(4), _
                  ind_lb(5) To ind_ub(5), _
                  ind_lb(6) To ind_ub(6), _
                  ind_lb(7) To ind_ub(7), _
                  ind_lb(8) To ind_ub(8), _
                  ind_lb(9) To ind_ub(9))
    
        For i0 = ind_lb(0) To ind_ub(0)
        For i1 = ind_lb(1) To ind_ub(1)
        For i2 = ind_lb(2) To ind_ub(2)
        For i3 = ind_lb(3) To ind_ub(3)
        For i4 = ind_lb(4) To ind_ub(4)
        For i5 = ind_lb(5) To ind_ub(5)
        For i6 = ind_lb(6) To ind_ub(6)
        For i7 = ind_lb(7) To ind_ub(7)
        For i8 = ind_lb(8) To ind_ub(8)
        For i9 = ind_lb(9) To ind_ub(9)
            
            Call get_2dim_index(first_ind_lb, first_ind_ub, sec_ind_lb, sec_ind_ub, _
                Array(i0, i1, i2, i3, i4, i5, i6, i7, i8, i9), ind_2d)
        
            If LBound(arr_in) = UBound(arr_in) Then
                arr_out(i0, i1, i2, i3, i4, i5, i6, i7, i8, i9) = arr_in(LBound(arr_in, 1), ind_2d(0) + LBound(arr_in, 2))
            
            ElseIf LBound(arr_in, 2) = UBound(arr_in, 2) Then
                arr_out(i0, i1, i2, i3, i4, i5, i6, i7, i8, i9) = arr_in(ind_2d(0) + LBound(arr_in, 1), LBound(arr_in, 2))
            
            ElseIf dim_order = fmlRow Then
                arr_out(i0, i1, i2, i3, i4, i5, i6, i7, i8, i9) = arr_in(ind_2d(0) + LBound(arr_in, 1), ind_2d(1) + LBound(arr_in, 2))
                
            Else
                arr_out(i0, i1, i2, i3, i4, i5, i6, i7, i8, i9) = arr_in(ind_2d(1) + LBound(arr_in, 1), ind_2d(0) + LBound(arr_in, 2))
            End If
     
        
        Next i9
        Next i8
        Next i7
        Next i6
        Next i5
        Next i4
        Next i3
        Next i2
        Next i1
        Next i0
    
    End Select
    
    restruct2mult_dim = True

End Function

Public Sub VarToRow(UpperLeftCorner As Excel.Range, ParamArray Source() As Variant)
    
    Dim v As Variant
    Dim r As Excel.Range
    Dim r_next As Excel.Range
        
    On Error GoTo HandleError:
    
    If IsEmpty(Source) Then Exit Sub
   
    Set r = UpperLeftCorner
        
    For Each v In Source
    
        If Not var2range(r, v, r_next, fmlRow) Then
                
            Err.Raise Number:=fmlInvalidArgument, _
                      Source:=PROJ_NAME, _
                      Description:=modErrInfo__.errstr_InvalidArgument
         
        End If
        Set r = r_next
    
    Next v
    
    Set UpperLeftCorner = r
    
    Exit Sub

HandleError:
    
    modErrInfo__.FuncID = id_VarToRow__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext

End Sub


Public Sub VarToCol(UpperLeftCorner As Excel.Range, ParamArray Source() As Variant)

    
    Dim v As Variant
    Dim r As Excel.Range
    Dim r_next As Excel.Range
    
    On Error GoTo HandleError:
    
    If IsEmpty(Source) Then Exit Sub
    
    
    Set r = UpperLeftCorner
        
    For Each v In Source
    
         If Not var2range(r, v, r_next, fmlCol) Then
         
            Err.Raise Number:=fmlInvalidArgument, _
                      Source:=PROJ_NAME, _
                      Description:=modErrInfo__.errstr_InvalidArgument
         
         End If
         Set r = r_next
    
    Next v
    
    Set UpperLeftCorner = r
    
    Exit Sub

HandleError:
    
    modErrInfo__.FuncID = id_VarToCol__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext

End Sub


Private Function var2range(UpperLeftCorner As Excel.Range, InputVar As Variant, r_next As Excel.Range, direction As fmlRangeDim) As Boolean

    Dim out_range As Excel.Range
    Dim row_size As Long, col_size As Long
    Dim dim_size As Long
    Dim out_var As Variant
    Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long
    Dim i As Long, j As Long
    
    '--- Single Value ---
    If Not IsArray(InputVar) Then
    
        UpperLeftCorner.value = InputVar
        
        If direction = fmlCol Then
        
            Set r_next = UpperLeftCorner.Offset(0, 1)
            var2range = True
            
        ElseIf direction = fmlRow Then
        
            Set r_next = UpperLeftCorner.Offset(1, 0)
            var2range = True
            
        Else
            var2range = False
        
        End If
        Exit Function
    
    End If
    
    dim_size = modArrSpt__.NumberOfArrayDimensions(InputVar)
    
    Select Case dim_size
        
    Case 1
        lb1 = LBound(InputVar)
        ub1 = UBound(InputVar)
        
        If direction = fmlCol Then
        
            ReDim out_var(lb1 To ub1, 0 To 0)
            
            For i = lb1 To ub1
                out_var(i, 0) = InputVar(i)
            Next i
        
            Excel.Range(UpperLeftCorner, _
                UpperLeftCorner.Offset(ub1 - lb1, 0)) = out_var
                
            Set r_next = UpperLeftCorner.Offset(0, 1)
            var2range = True
               
        ElseIf direction = fmlRow Then
        
            ReDim out_var(0 To 0, lb1 To ub1)
            
            For i = lb1 To ub1
                out_var(0, i) = InputVar(i)
            Next i
        
            Excel.Range(UpperLeftCorner, _
                UpperLeftCorner.Offset(0, ub1 - lb1)) = out_var
                
            Set r_next = UpperLeftCorner.Offset(1, 0)
            var2range = True
        
        Else
        
            var2range = False
        
        End If
        
    
    Case 2
        lb1 = LBound(InputVar)
        ub1 = UBound(InputVar)
        lb2 = LBound(InputVar, 2)
        ub2 = UBound(InputVar, 2)
        
        If direction = fmlCol Then
                
            ReDim out_var(lb2 To ub2, lb1 To ub1)
            
            For i = lb1 To ub1
                For j = lb2 To ub2
                    out_var(j, i) = InputVar(i, j)
                Next j
            Next i
            
            Excel.Range(UpperLeftCorner, _
                UpperLeftCorner.Offset(ub2 - lb2, ub1 - lb1)) = out_var
                
            Set r_next = UpperLeftCorner.Offset(0, ub1 - lb1 + 1)
            var2range = True
        
        ElseIf direction = fmlRow Then
        
            ReDim out_var(lb1 To ub1, lb2 To ub2)
            
            For i = lb1 To ub1
                For j = lb2 To ub2
                    out_var(i, j) = InputVar(i, j)
                Next j
            Next i
            
            Excel.Range(UpperLeftCorner, _
                UpperLeftCorner.Offset(ub1 - lb1, ub2 - lb2)) = out_var
                
            Set r_next = UpperLeftCorner.Offset(ub1 - lb1 + 1, 0)
            var2range = True
               
        Else
            var2range = False
        
        End If
        
    Case Else
    
        var2range = False
        
    End Select
    
    
End Function


Public Function NumEprToArray(ByVal NumEpr As Variant) As Variant

    Dim v As Variant, vv As Variant
    Dim ret_val As Variant
    Dim curr_ind As Long
    Dim ll As Long, lu As Long
    Dim arg_type As VbVarType

    ReDim ret_val(10) As Variant

    On Error GoTo HandleError:
    
    If IsObject(NumEpr) Then NumEpr = NumEpr
    
    arg_type = VarType(NumEpr)

    If arg_type = vbDouble _
    Or arg_type = vbLong _
    Or arg_type = vbInteger Then
    
        NumEprToArray = Array(NumEpr)
        Exit Function
        
    ElseIf Not VarType(NumEpr) = vbString Then
        
        Err.Raise Number:=fmlInvalidArgument, _
                  Source:=PROJ_NAME, _
                  Description:=modErrInfo__.errstr_InvalidArgument
    End If
    
    NumEpr = Split(NumEpr, ",")
    
    For Each v In NumEpr
    
        v = Split(v, "-")
        
        If UBound(v) = 0 Then
        
            If IsNumeric(v(0)) Then
            
                ret_val(curr_ind) = CLng(v(0))
                curr_ind = curr_ind + 1
                
                If curr_ind > UBound(ret_val) Then
                
                    ReDim Preserve ret_val(UBound(ret_val) + 10)
                
                End If
                        
            Else
            
                Err.Raise Number:=fmlInvalidArgument, _
                          Source:=PROJ_NAME, _
                          Description:=modErrInfo__.errstr_InvalidArgument
            End If
        
        
        ElseIf UBound(v) = 1 Then
        
            If IsNumeric(v(0)) And IsNumeric(v(1)) Then
            
                ll = CLng(v(0))
                lu = CLng(v(1))
                
                If ll > lu Then
                
                    Err.Raise Number:=fmlInvalidArgument, _
                              Source:=PROJ_NAME, _
                              Description:=modErrInfo__.errstr_InvalidArgument
                                
                End If
                
                vv = ll
                Do While vv <= lu
                
                    ret_val(curr_ind) = vv
                    curr_ind = curr_ind + 1
                    vv = vv + 1
                    
                    If curr_ind > UBound(ret_val) Then
                    
                        ReDim Preserve ret_val(UBound(ret_val) + 10)
                    
                    End If
                    
                Loop
            
            Else
            
                Err.Raise Number:=fmlInvalidArgument, _
                          Source:=PROJ_NAME, _
                          Description:=modErrInfo__.errstr_InvalidArgument
            
            End If
        
        Else
        
                Err.Raise Number:=fmlInvalidArgument, _
                          Source:=PROJ_NAME, _
                          Description:=modErrInfo__.errstr_InvalidArgument
                
        End If
    
    Next v
    
    ReDim Preserve ret_val(curr_ind - 1)
    
    NumEprToArray = ret_val
    
    Exit Function
      
HandleError:
    
    modErrInfo__.FuncID = id_NumEprToArray__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext

End Function



Public Function SplitFileExt(ByVal FilePath As String) As Variant

    Dim str_path As String, str_ext As String
    Dim dot_pos As Long
    

    On Error GoTo HandleError:
    
    dot_pos = InStrRev(FilePath, ".")
    
    If dot_pos = 0 Then
    
        SplitFileExt = Array(FilePath, "")
        
        Exit Function
            
    End If
    
    str_path = Left(FilePath, dot_pos - 1)
    str_ext = Right(FilePath, Len(FilePath) - InStrRev(FilePath, "."))
    
    SplitFileExt = Array(str_path, str_ext)
    
    Exit Function
      
HandleError:
    
    modErrInfo__.FuncID = id_SplitFileExt__
    Err.Raise Number:=Err.Number, _
                Source:=Err.Source, _
                Description:=Err.Description, _
                HelpFile:=Err.HelpFile, _
                HelpContext:=Err.HelpContext

End Function



