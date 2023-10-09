Attribute VB_Name = "Util"
Option Explicit

Function get_file_list(file_path As String, file_extension As String) As String()
    Dim fname As String: fname = Dir(file_path & "\*." & file_extension)
    
    Dim file_list() As String
    
    Dim i As Integer: i = 0
    Do While fname <> ""
        ReDim Preserve file_list(i)
        
        file_list(i) = file_path & "\" & fname
        fname = Dir()
        i = i + 1
    Loop
    
    get_file_list = file_list
    

End Function


' 配列の要素数を求める。
'
' ary：対象となる配列。
' return：配列の要素数。引数として初期化されていない配列を指定した時は-1、配列以外を指定した時は-100を返す。
Function CalcArrayLength(ary As Variant) As Integer
    If (IsArray(ary)) Then
        If (IsInitialized(ary)) Then
            CalcArrayLength = UBound(ary) - LBound(ary) + 1
        Else
            CalcArrayLength = -1
        End If
    Else
        CalcArrayLength = -100
    End If

End Function

' 配列が初期化されているかをチェックする。
'
' ary：対象となる配列。
' return：配列が初期化済みならTrue、そうでなければFalseを返す。
Function IsInitialized(ary As Variant) As Boolean
    On Error GoTo NOT_INITIALIZED_ERROR
    Dim length As Long: length = UBound(ary)    ' 動的配列が初期化されていなければ、ここでエラーが発生する。
    IsInitialized = True
    Exit Function

' 配列が初期化されていない場合はここに飛ばされる。
NOT_INITIALIZED_ERROR:
    IsInitialized = False
End Function
