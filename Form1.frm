VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows 기본값
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function getPi(Pattern As String) As Long()
Dim i&, x&, horizon&
Dim PI() As Long
Dim Temp$
ReDim PI(Len(Pattern)) As Long

PI(0) = -1
For i = 1 To Len(Pattern) '//찾을 문자열의 길이만큼 반복
    Temp = Left$(Pattern, i) '//찾을 문자열에서 한글짜씩 더해가면서 거기까지만 pi값을 찾아와야 되기에 자른다.
    
    For x = 1 To Len(Temp) - 1
        '//잘라왔던 곳에서 경계의 길이가 최대가 되는 접두부의 길이를 구한다.
        '//양쪽에서 같은 길이만큼 잘라와서 비교하면 같은건 당연지사. 그러므로 1을 빼준다.
        
        If Left$(Temp, x) = Right$(Temp, x) Then
            horizon = Len(Temp) - (Len(Left$(Temp, x)) * 2)
            '//경계의 길이는 전체길이 - 접두부의길이*2 (접두부와 접미부가 같음을 확인했기에)
            If PI(i) = 0 And horizon > 0 Then PI(i) = x
            '//경계의 길이가 0인 경우는 제외한다.
            '//for x 가 진행될수록 경계의 길이는 줄어드므로 pi(i) 값이 초기값일때 한번만! 넣어준다.
        End If
        
    Next x
Next i

getPi = PI
End Function

Private Function KMP(Str As String, Tofind As String, PI() As Long)
Dim i&, x&
Dim MatchedLength&
Dim Temp$

For x = 1 To Len(Str) '//찾을 문자열이 들어있는 문자열의 길이만큼 반복
Temp = Right$(Str, Len(Str) - x + 1) '//찾을 문자열이 오른쪽으로 이동하는것을 구현
MatchedLength = 0 '//일치 접두부의 길이가 담길 변수
For i = 1 To Len(Tofind)
    If Left$(Temp, i) = Left$(Tofind, i) Then
        MatchedLength = i '//일치 접두부의 길이를 구한다.
    End If
Next i
If MatchedLength = Len(Tofind) Then MsgBox x & " 번째에서 발견!"
'//일치 접두부의 길이가 찾을 문자열의 길이라면 일치하는 것.
x = x + MatchedLength - PI(MatchedLength) - 1
'//공식에 따라 도출된 이동거리를 더해준다. Next x 가 x에 1을 더하므로 미리 빼준다.
Next x

End Function

Private Sub Form_Load()
Dim PI() As Long
Dim Str$, Tofind$

Str = "BABAABABAABABAABBBABAABABAABABBAAB"
Tofind = "BAABAB"
PI = getPi(Tofind)
Call KMP(Str, Tofind, PI)

End
End Sub
