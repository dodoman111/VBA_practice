Attribute VB_Name = "Module1"
Sub saanalysis()
Attribute saanalysis.VB_ProcData.VB_Invoke_Func = " \n14"

Dim wrk As Workbook
Dim a As Worksheet
Set wrk = ThisWorkbook
Set a = wrk.Sheets("base")
Set b = wrk.Sheets("노드")
Set c = wrk.Sheets("resul")
Dim rse()                         As Long
Dim split_0                       As String
Dim split_1()                     As String
Dim split_2()                     As String
Dim split_3()                     As String
Dim split_4()                     As String
Dim split_5()                     As String
Dim split_o()                     As String
Dim split_d()                     As String
Dim split_1_nm                    As Long
Dim split_2_nm                    As Long
Dim split_3_nm                    As Long
Dim resul_nm                      As Long
Dim tt()                         As Long
Dim x(99999, 2, 2)                As Long
Dim resul_osa_tt(99999)            As Long

Dim resul_o_rse()                 As Long
Dim resul_o_t()                   As Long

Dim resul(100000, 3)              As Long
Dim ind(100000, 1)                As Double

Dim i As Long
Dim f1 As Long
Dim f2 As Long

ReDim tt(99999) As Long
ReDim rse(3) As Long
Path = a.Cells(1, 2)
file_nm = a.Cells(2, 2)
rse(1) = a.Cells(3, 2)
rse(2) = a.Cells(3, 3)
rse(3) = a.Cells(3, 4)
dst_1 = a.Cells(4, 3)
dst_2 = a.Cells(4, 4)
cnt = 0

For f1 = 1 To 100000
If b.Cells(f1 + 1, 1) = "" Then Exit For
    ind(b.Cells(f1 + 1, 1), 1) = b.Cells(f1 + 1, 2)
Next f1

'======================전체차량 검지시각=========
Open Path & file_nm For Input As #1
    While Not EOF(1)
        Line Input #1, split_0
        split_1 = split(split_0, ",")
        split_1_nm = UBound(split_1)
        split_2 = split(split_1(split_1_nm), "|")
        split_2_nm = UBound(split_2)
        
        
        For f1 = 2 To split_2_nm
            split_3 = split(split_2(f1 - 2), ":")
                If split_3(0) = rse(1) Then
                    split_4 = split(split_2(f1 - 1), ":")
                    If split_4(0) = rse(2) Then
                        split_5 = split(split_2(f1), ":")
                            If split_5(0) = rse(3) And split_3(2) > 0 And split_5(2) < 86400 And split_4(2) > split_3(2) And split_5(2) > split_4(2) Then
                                  
                                  For ii = 1 To split_2_nm
                                        split_o = split(split_2(ii - 1), ":")
                                        If Left(split_o(0), 2) = "10" Then Exit For
                                  Next ii
                                  For ii = 1 To split_2_nm
                                        split_d = split(split_2(split_2_nm - ii + 1), ":")
                                        If Left(split_d(0), 2) = "10" Then Exit For
                                  Next ii
                                       
                                       cnt = cnt + 1
                                       x(cnt, 1, 1) = split_o(0)
                                       x(cnt, 1, 2) = split_o(2)
                                       x(cnt, 2, 1) = split_d(0)
                                       x(cnt, 2, 2) = split_d(2)
                                       
                                       resul(cnt, 1) = split_3(2)
                                       resul(cnt, 2) = split_4(2)
                                       resul(cnt, 3) = split_5(2)
                                       
                                       c.Cells(cnt + 2, 1) = resul(cnt, 1)
                                       c.Cells(cnt + 2, 2) = resul(cnt, 2)
                                       c.Cells(cnt + 2, 3) = resul(cnt, 3)
                                       tt(cnt) = resul(cnt, 3) - resul(cnt, 2)
                                 
                          End If
                    End If
                End If
        Next f1
        
        
        
    Wend
Close #1
'======휴게소 전구간 링크 평균통행시간(ctt) 구하기======
resul_nm = cnt

    For f1 = 1 To resul_nm
        ctt = (resul(f1, 2) - resul(f1, 1)) / dst_1
        Sum = Sum + ctt
    Next f1
        ctt = Sum / resul_nm
        
'======휴게소 최초출발지-휴게소까지 통행시간구하기======
    For f0 = 1 To resul_nm
        
        resul_osa_tt(f0) = resul(f0, 2) - x(f0, 1, 2)
        c.Cells(f0 + 2, 5) = x(f0, 1, 1)
        c.Cells(f0 + 2, 6) = x(f0, 2, 1)
        c.Cells(f0 + 2, 7) = resul_osa_tt(f0)
    Next f0

'======휴게소 최초출발지-휴게소까지, 휴게소에서 최종목적지까지 통행거리======
    Dim resul_osa_dst(100000) As Double
    Dim resul_sad_dst(100000) As Double
    
   
    For f0 = 1 To resul_nm
        resul_osa_dst(f0) = Abs(a.Cells(5, 3) - ind(x(f0, 1, 1), 1))
        resul_sad_dst(f0) = Abs(ind(x(f0, 2, 1), 1) - a.Cells(5, 3))
        c.Cells(f0 + 2, 8) = resul_osa_dst(f0)
        c.Cells(f0 + 2, 9) = resul_sad_dst(f0)
        If tt(f0) / dst_2 > ctt + 10 Then
            c.Cells(f0 + 2, 10) = 1
        Else
            c.Cells(f0 + 2, 10) = 0
        End If
    Next f0
