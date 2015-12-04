Attribute VB_Name = "modUtil__"
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

#If VBA7 Then
    
    Private Declare PtrSafe Function GetInputState Lib "USER32" () As Long

#Else

    Private Declare Function GetInputState Lib "USER32" () As Long

#End If

Private m_TimeLastChecked As Variant
 
Public Sub CheckEvents()
'Reference: http://www.sunvisor.net/parts/doevents

    If GetInputState() Or (DateDiff("s", m_TimeLastChecked, Time) > 1) Then
    
    'If DateDiff("s", m_TimeLastChecked, Time) > 1 Then
    
        DoEvents
        m_TimeLastChecked = Time
        RunStatus.UpdateStatus
        
        If RunStatus.CancelButtonClicked Then
        
            End
        
        End If
        
    End If
    
End Sub


Public Function redim_nested_array(a As Variant, new_ub As Long, is_preserve As Boolean) As Boolean

    If is_preserve = True Then

        ReDim Preserve a(new_ub)
        
    Else
        
        ReDim a(new_ub)
    
    End If

    redim_nested_array = True

End Function





Public Function is_whole_number(Var As Variant) As Boolean

    Dim var_type As Long

    is_whole_number = False

    var_type = VarType(Var)

    If var_type = vbBoolean Or _
        var_type = vbInteger Or _
        var_type = vbLong Then
    
        is_whole_number = True
        
    ElseIf (var_type = vbDouble Or var_type = vbSingle) Then
    
        If Var = CLng(Var) Then
           
            is_whole_number = True
        Else
            is_whole_number = False
        
        End If
    
    Else
    
        is_whole_number = False
    
    End If

End Function

Public Function is_numeric(Var As Variant) As Boolean

    Dim var_type As Long
    
    var_type = VarType(Var)
    
    If var_type = vbBoolean Or _
        var_type = vbInteger Or _
        var_type = vbLong Or _
        var_type = vbSingle Or _
        var_type = vbDouble Then
    
        is_numeric = True
        
    Else
    
        is_numeric = False
    
    End If

End Function


Public Function str2var(ByVal str As String) As Variant

    Dim buf As Variant
    
    '--- Rmove double quotes and blanks ---
    str = Trim(str)
    
    If Left(str, 1) = """" Then
    
        str = Mid(str, 2, Len(str) - 1)
    
    End If
    
    If Right(str, 1) = """" Then
    
        str = Mid(str, 1, Len(str) - 1)
    
    End If
    
    str = Trim(str)
    '---------------------------------------

    If str = "" Then
    
        str2var = Empty
    
    ElseIf StrComp("#NULL#", str, VbCompareMethod.vbTextCompare) = 0 Then
    
        str2var = Null
        
    ElseIf StrComp("#TRUE#", str, VbCompareMethod.vbTextCompare) = 0 Then
       
        str2var = True
    
    ElseIf StrComp("#FALSE#", str, VbCompareMethod.vbTextCompare) = 0 Then
    
        str2var = False
    
    ElseIf IsNumeric(str) Then
    
        buf = CDbl(str)
    
        If is_whole_number(buf) Then
        
            str2var = CLng(buf)
    
        Else
        
            str2var = buf
        
        End If
    
    ElseIf IsDate(str) Then
    
        str2var = CDate(str)
                
    Else    'String
    
        str2var = str
    
    End If
    

End Function


Public Function var2line(var_in As Variant) As String

    Dim i As Long
    Dim ub As Long, lb As Long
    Dim str As String
    
    '--- Scalar ---
    If Not IsArray(var_in) Then
    
        var2line = var2str(var_in)
        Exit Function
    
    End If
    
    
    ub = UBound(var_in)
    lb = LBound(var_in)

    var2line = var2str(var_in(lb))

    For i = lb + 1 To ub
    
        var2line = var2line + "," + var2str(var_in(i))
            
    Next i

End Function


Private Function var2str(var_in As Variant) As String

    Dim str_buf As String

    If IsEmpty(var2str) Then
    
        var2str = ""
        
    ElseIf IsNull(var2str) Then
    
        var2str = "#NULL#"
        
    ElseIf VarType(var_in) = vbBoolean Then
    
        If var_in Then var2str = "#TRUE#"
    
    ElseIf VarType(var_in) = vbBoolean Then
    
        If Not var_in Then var2str = "#FALSE#"
        
    ElseIf VarType(var_in) = vbDate Then
    
        If CLng(var_in) = 0 Then   'Time Only
        
            str_buf = Format(var_in, "hh:mm:ss")

        ElseIf CLng(var_in) = var_in Then   'Date Only
            
            str_buf = Format(var_in, "yyyy-mm-dd")
        
        Else    'Date Time
        
            str_buf = Format(var_in, "yyyy-mm-dd hh:mm:ss")
        
        End If
        
        var2str = "#" + str_buf + "#"

    Else
    
        var2str = CStr(var_in)
    
    End If

End Function



Public Function compare_index(ind As Variant, ind_match As Variant) As Boolean

    Dim ind_lb As Long, ind_ub As Long
    Dim i As Long
    Dim all_empty As Boolean
    
    ind_lb = LBound(ind)
    ind_ub = UBound(ind)

    compare_index = False

    all_empty = True
    For i = ind_lb To ind_ub

        If Not IsEmpty(ind_match(i)) Then
        
            If Not (ind(i) = ind_match(i)) Then Exit Function
            all_empty = False
        
        End If
        
    Next i
    
    compare_index = IIf(all_empty, False, True)
    
End Function

Public Function get_bound(ind_lb As Variant, ind_ub As Variant, input_array As Variant, dim_count As Long) As Boolean

    Dim i As Long
    
    ReDim ind_lb(0 To dim_count - 1) As Long
    ReDim ind_ub(0 To dim_count - 1) As Long
    
    For i = 1 To dim_count
    
        ind_lb(i - 1) = LBound(input_array, i)
        ind_ub(i - 1) = UBound(input_array, i)
    
    Next i
    
    get_bound = True
    
End Function


Public Function increment_index(ByRef ind As Variant, ind_lb As Variant, ind_ub As Variant, Optional ind_len_cap As Long = MAX_DIM_COUNT) As Boolean
'Copied To Iteration module
'Index as Long
'Index as Array(Long)
'Ignore elements after ind_len_cap

    Dim ind_len As Long
    Dim i As Long, j As Long, i_max As Long
    
    If Not IsArray(ind) Then
    
        If ind = ind_ub Then
            increment_index = False
            Exit Function
            
        Else
            ind = ind + 1
    
        End If
        
    Else
        
        i_max = IIf(ind_len_cap < UBound(ind), ind_len_cap, UBound(ind))
        i = i_max
    
        Do While ind(i) = ind_ub(i)
            
            If i = LBound(ind) Then
                increment_index = False
                Exit Function
            End If
            
            i = i - 1
        Loop
        
        ind(i) = ind(i) + 1
        
        For j = i + 1 To i_max
        
            ind(j) = ind_lb(j)
        
        Next j
        
    End If
    
    increment_index = True
    
End Function


