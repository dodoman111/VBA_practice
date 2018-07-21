Attribute VB_Name = "Module1"
Dim wrk                             As Workbook
Dim a                               As Worksheet
Dim b                               As Worksheet
Dim c                               As Worksheet
Dim split_0                         As String
Dim split_1()                       As String
Dim split_2()                       As String
Dim split_3()                       As String
Dim split_4()                       As String
Dim split_5()                       As String
Dim split_6()                       As String
Dim split_7()                       As String
Dim split_8()                       As String
Dim split_1_nm                      As Long
Dim split_2_nm                      As Long
Dim split_6_nm                      As Long
Dim split_7_nm                      As Long
Dim ind(100000)                     As Double
Dim SA()                            As Long
Dim 휴게소_TT()                     As Long
Dim 이전구간_TT()                   As Long
Dim 휴게소정보()                    As Long
Dim TT()                            As Long
Dim 이용횟수()                      As Long
Dim 휴게소까지통행시간()            As Long
Dim 휴게소정보첫지점()              As Long

Sub 차량정보()

Set wrk = ThisWorkbook
Set a = wrk.Worksheets("base")
Set b = wrk.Worksheets("node")
Set c = wrk.Worksheets("result")
Set d = wrk.Worksheets("result1")
Set e = wrk.Worksheets("result2")
Set f = wrk.Worksheets("result3")
Set g = wrk.Worksheets("result4")

Car_nm = a.Cells(5, 2)
Seq = a.Cells(6, 2)


'============== rse 거리정보 저장 =========
For f1 = 1 To 100000
    If b.Cells(f1 + 1, 1) = "" Then Exit For
        ind(b.Cells(f1 + 1, 1)) = b.Cells(f1 + 1, 2)
Next f1


'=============== 휴게소 rse 저장 ===========

ReDim SA(7)
ReDim 휴게소_TT(999999)
ReDim 이전구간_TT(999999)
ReDim 휴게소정보(199999, 7, 2)              '휴게소정보(차량id, 휴게소번호, 이용여부(0:이용안함, 1:이용)
ReDim 휴게소정보첫지점(199999)

For f1 = 1 To 7
    SA(f1) = a.Cells(3, f1 + 1)
Next f1

path_nm = a.Cells(1, 2)
file_nm = a.Cells(2, 2)
n = 7                                                                                   '영동고속도로 구간내 휴게소 갯수
bv = 50                                                                                 '휴게소 이용분별 조건 계수

Open path_nm & file_nm For Input As #1
    Do While Not EOF(1)
        Line Input #1, split_0
        r = r + 1
        
        split_1 = Split(split_0, ",")
        split_1_nm = UBound(split_1)
        split_2 = Split(split_1(split_1_nm), "|")
        split_2_nm = UBound(split_2)
        
        For j = 1 To 7
            For i = 1 To split_2_nm - 1
                If Left(split_2(i), 4) = SA(j) Then
                    If Left(split_2(i + 1), 4) = SA(j) + 1 Then
                        If Left(split_2(i - 1), 4) = SA(j) - 1 Then
                            split_3 = Split(split_2(i), ":")
                            split_4 = Split(split_2(i + 1), ":")
                            split_5 = Split(split_2(i - 1), ":")
'                            If i > 3 Then
                                split_6 = Split(split_2(1), ":")
'                            Else: split_6 = Split(split_2(2), ":")
'                            End If
                            
                            cnt = cnt + 1
                            휴게소_TT(cnt) = (split_4(2) - split_3(2)) / (ind(split_4(0)) - ind(split_3(0)))                 '단위 km 당 휴게소 구간 통행시간
                            이전구간_TT(cnt) = (split_3(2) - split_5(2)) / (ind(split_3(0)) - ind(split_5(0)))               '단위 km 당 휴게소 이전구간 통행시간
                                                                                             
                            
                            '======== 휴게소 이용 조건 5분이상 쉰차량 대상 ============
                            If 휴게소_TT(cnt) > 이전구간_TT(cnt) + 300 / (ind(split_3(0)) - ind(split_5(0))) Then           '휴게소 쉰 차량
                                k = 1
                            Else: k = 0
                            End If
                            
                            If r <> rr Then
                                car = car + 1
                            End If
                            
                            휴게소정보(car, j, k) = split_4(2)                                                               '휴게소 구간 마지막 rse 검지시각
            
                            휴게소정보첫지점(car) = split_6(2)          '고속도로 진입 첫지점
                            
                            rr = r
                        End If
                    End If
                End If
            Next i
        Next j
    Loop
Close #1



'========= 출력단 ===========

For i = 1 To car
    c.Cells(i + 2, 1) = i           '차량번호ID
    For j = 1 To 7
        For k = 0 To 1
        If 휴게소정보(i, j, k) <> 0 Then                         '휴게소정보(차량id, 휴게소번호, 이용여부(0:이용안함, 1:이용) =  휴게소 나와서 첫 RSE 검지시각
            If j = 1 Then
                c.Cells(i + 2, 2) = k
                c.Cells(i + 2, 3) = 휴게소정보(i, j, k)
            End If
            If j = 2 Then
                c.Cells(i + 2, 4) = k
                c.Cells(i + 2, 5) = 휴게소정보(i, j, k)
            End If
            If j = 3 Then
                c.Cells(i + 2, 6) = k
                c.Cells(i + 2, 7) = 휴게소정보(i, j, k)
            End If
            If j = 4 Then
                c.Cells(i + 2, 8) = k
                c.Cells(i + 2, 9) = 휴게소정보(i, j, k)
            End If
            If j = 5 Then
                c.Cells(i + 2, 10) = k
                c.Cells(i + 2, 11) = 휴게소정보(i, j, k)
            End If
            If j = 6 Then
                c.Cells(i + 2, 12) = k
                c.Cells(i + 2, 13) = 휴게소정보(i, j, k)
            End If
            If j = 7 Then
                c.Cells(i + 2, 14) = k
                c.Cells(i + 2, 15) = 휴게소정보(i, j, k)
            End If
        End If
        Next k
    Next j
Next i


'=========== 휴게소 이용횟수에 따른 분류 ==================

ReDim TT(car, 7)
ReDim 이용횟수(car)
ReDim 휴게소까지통행시간(car, 7, 7)

For i = 1 To car
    For j = 1 To 7
        If c.Cells(i + 2, 2 * j) = 1 Then
            n = n + 1                                       '휴게소 이용횟수
            TT(i, n) = c.Cells(i + 2, 2 * j + 1)
        End If
    Next j
    이용횟수(i) = n                                         'i번째 차의 휴게소 이용횟수
    n = 0
Next i

''======== 버블알고리즘 오름차순 정렬 =====
'For i = 1 To car
'    For k = 1 To 6
'        m = TT(i, k)
'        If m >= TT(i, k + 1) Then
'            TT(i, k) = TT(i, k + 1)
'            TT(i, k + 1) = m
'        End If
'    Next k
'Next i


'===================== 휴게소까지 통행시간 산정 =====
For i = 1 To car
    For k = 1 To 5
        If 이용횟수(i) > 1 Then
            If k = 1 Then                                                                                   '휴게소 1번 이용하면
                휴게소까지통행시간(i, 이용횟수(i), k) = Abs((TT(i, k) - 휴게소정보첫지점(i)) / 60)          '휴게소까지 통행시간 분으로 환산
            ElseIf k > 1 Then
                휴게소까지통행시간(i, 이용횟수(i), k) = Abs((TT(i, k) - TT(i, k - 1)) / 60)                 '휴게소 2번 이상 이용하면 휴게소까지통행시간(차량id, 이용횟수) = 통행시간
            End If
        ElseIf 이용횟수(i) = 1 Then
            휴게소까지통행시간(i, 이용횟수(i), 1) = Abs(TT(i, 1) - 휴게소정보첫지점(i)) / 60
        End If
    Next k
Next i

'====== 휴게소 이용횟수에 따른 통행시간 분포 출력 =====
For i = 1 To car
    If 이용횟수(i) = 1 Then
'        If 휴게소까지통행시간(i, 이용횟수(i), 1) < 360 Then         '이상치 제거 조건 통행시간 6시간 이내
            cnt_0 = cnt_0 + 1
            d.Cells(2 + cnt_0, 2) = 휴게소까지통행시간(i, 이용횟수(i), 1)
'        End If
    End If
    If 이용횟수(i) = 2 Then
'        If 휴게소까지통행시간(i, 이용횟수(i), 2) < 360 And 휴게소까지통행시간(i, 이용횟수(i), 2) > 0 Then
            cnt_2 = cnt_2 + 1
            e.Cells(2 + cnt_2, 2) = 휴게소까지통행시간(i, 이용횟수(i), 1)
            e.Cells(2 + cnt_2, 3) = 휴게소까지통행시간(i, 이용횟수(i), 2)
'        End If
    End If
    If 이용횟수(i) = 3 Then
'        If 휴게소까지통행시간(i, 이용횟수(i), 3) < 360 And 휴게소까지통행시간(i, 이용횟수(i), 3) > 0 Then
            cnt_3 = cnt_3 + 1
            f.Cells(2 + cnt_3, 2) = 휴게소까지통행시간(i, 이용횟수(i), 1)
            f.Cells(2 + cnt_3, 3) = 휴게소까지통행시간(i, 이용횟수(i), 2)
            f.Cells(2 + cnt_3, 4) = 휴게소까지통행시간(i, 이용횟수(i), 3)
'        End If
    End If
    If 이용횟수(i) >= 4 Then
'        If 휴게소까지통행시간(i, 이용횟수(i), 4) < 300 And 휴게소까지통행시간(i, 이용횟수(i), 4) > 0 Then
            cnt_4 = cnt_4 + 1
            g.Cells(2 + cnt_4, 2) = 휴게소까지통행시간(i, 이용횟수(i), 1)
            g.Cells(2 + cnt_4, 3) = 휴게소까지통행시간(i, 이용횟수(i), 2)
            g.Cells(2 + cnt_4, 4) = 휴게소까지통행시간(i, 이용횟수(i), 3)
            g.Cells(2 + cnt_4, 5) = 휴게소까지통행시간(i, 이용횟수(i), 4)
            g.Cells(2 + cnt_4, 6) = 휴게소까지통행시간(i, 이용횟수(i), 5)
            g.Cells(2 + cnt_4, 7) = 휴게소까지통행시간(i, 이용횟수(i), 6)
            g.Cells(2 + cnt_4, 8) = 휴게소까지통행시간(i, 이용횟수(i), 7)
'        End If
    End If
Next i


End Sub
