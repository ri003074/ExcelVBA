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


' �z��̗v�f�������߂�B
'
' ary�F�ΏۂƂȂ�z��B
' return�F�z��̗v�f���B�����Ƃ��ď���������Ă��Ȃ��z����w�肵������-1�A�z��ȊO���w�肵������-100��Ԃ��B
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

' �z�񂪏���������Ă��邩���`�F�b�N����B
'
' ary�F�ΏۂƂȂ�z��B
' return�F�z�񂪏������ς݂Ȃ�True�A�����łȂ����False��Ԃ��B
Function IsInitialized(ary As Variant) As Boolean
    On Error GoTo NOT_INITIALIZED_ERROR
    Dim length As Long: length = UBound(ary)    ' ���I�z�񂪏���������Ă��Ȃ���΁A�����ŃG���[����������B
    IsInitialized = True
    Exit Function

' �z�񂪏���������Ă��Ȃ��ꍇ�͂����ɔ�΂����B
NOT_INITIALIZED_ERROR:
    IsInitialized = False
End Function
