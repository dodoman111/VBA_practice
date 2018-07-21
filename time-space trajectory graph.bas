Attribute VB_Name = "Module1"
'==================================================================밀도곡선
Dim wrk             As Workbook
Dim A               As Worksheet
Dim B               As Worksheet
'==================================================================
Dim f1              As Long
Dim f2              As Long
Dim f3              As Long
Dim f4              As Long
Dim f5              As Long
Dim f6              As Long
'==================================================================
Dim path            As String
Dim file            As String
Dim link(4)         As Long
Dim in_out(2)       As Long
Dim init            As String
Dim init_1()        As String
Dim init_1_nm       As Long
Dim init_2()        As String
Dim init_2_nm       As Long
Dim init_3()        As String
Dim init_4()        As String
Dim init_5()        As String
Dim init_6()        As String
Dim seq             As Long
Dim seq_nm()        As Long
Dim resul()         As Long
Dim resul_af()      As Long
Dim link_lane_nm    As Long
Dim direction       As Long
Dim judge()         As Long
Dim judge_direction As Long
Dim detected        As Long
Dim predic_in       As Long
Dim predic_out      As Long
Dim tm_ind          As Long
Dim cnt             As Long
Dim link_length()   As Long
Dim X               As Long
Dim Y               As Long
Dim st_tm           As Long
Dim ed_tm           As Long
Dim flag            As Long
Sub path_anal()
Set wrk = ThisWorkbook
Set A = wrk.Sheets("base")
Set B = wrk.Sheets("resul")
path = A.Cells(1, 2)
file = A.Cells(2, 2)
st_tm = A.Cells(8, 4)
ed_tm = A.Cells(8, 5)
ReDim link_length(3)
link_lane_nm = A.Cells(6, 2)
link(1) = A.Cells(3, 2)  '''대상링크 첫번째 RSE
link(2) = A.Cells(3, 3)  '''대상링크 두번째 RSE
link(3) = A.Cells(3, 4)  '''대상링크 세번째 RSE
link(4) = A.Cells(3, 5)
If link(2) - link(1) < 0 Then
    direction = -1
Else
    direction = 1
End If
ReDim resul(1000000, 4) ''' (#,1);#시퀀스내 구간진입교통량 (#,2);#시퀀스내 구간진출교통량
ReDim seq_nm(2) ''' 1;구간진입시간, 2;구간진출시간
cnt = 0
Open path & file For Input As #1
While Not EOF(1)
     Line Input #1, init
     init_1 = Split(init, ",")
     init_1_nm = UBound(init_1)
     init_2 = Split(init_1(init_1_nm), "|")
     For f1 = 1 To init_1(init_1_nm - 1) - 1
        If Left(init_2(f1 - 1), 4) = link(1) Then
            init_3 = Split(init_2(f1 - 1), ":")
            For f2 = f1 + direction To init_1(init_1_nm - 1) - 1
                If Left(init_2(f2 - 1), 4) = link(2) Then
                    init_4 = Split(init_2(f2 - 1), ":")
                    For f3 = f2 + direction To init_1(init_1_nm - 1) - 1
                        If Left(init_2(f3 - 1), 4) = link(3) Then
                            init_5 = Split(init_2(f3 - 1), ":")
                            For f4 = f3 + direction To init_1(init_1_nm - 1) - 1
                                If Left(init_2(f4 - 1), 4) = link(4) Then
                                init_6 = Split(init_2(f4 - 1), ":")
                                    If init_3(2) >= st_tm And init_3(2) <= ed_tm Then
                                        If init_4(2) >= st_tm And init_4(2) <= ed_tm Then
                                            If init_5(2) >= st_tm And init_5(2) <= ed_tm Then
                                                If init_6(2) >= st_tm And init_6(2) <= ed_tm Then
                                                    cnt = cnt + 1
                                                    resul(cnt, 1) = init_3(2) - st_tm + 1
                                                    resul(cnt, 2) = init_4(2) - st_tm + 1
                                                    resul(cnt, 3) = init_5(2) - st_tm + 1
                                                    resul(cnt, 4) = init_6(2) - st_tm + 1
                                                End If
                                            End If
                                        End If
                                    End If
                                Exit For
                                End If
                            Next f4
                            Exit For
                        End If
                    Next f3
                    Exit For
                End If
            Next f2
            Exit For
        End If
    Next f1
 Wend
Close #1
Dim X_ As Long
Dim Y_ As Long
Dim init_X As Long
Dim init_Y As Long
Dim Delta_X As Double
Dim vect As Long
Dim from_f3 As Long
Dim to_f3 As Long
Dim pong As Long
init_X = 1
Delta_X = A.Cells(7, 2)
init_Y = 1200
link_length(1) = A.Cells(5, 3) * 1000
link_length(2) = A.Cells(5, 4) * 1000
link_length(3) = A.Cells(5, 5) * 1000
For f1 = 1 To cnt
'===================1번째 구간========================================
    X = resul(f1, 1) / Delta_X
    X_ = resul(f1, 2) / Delta_X
    Y = init_Y
    Y_ = Y - link_length(1) / 20
    flag = 0
    If X_ - X <= 0 Then
        pong = 1
    Else
        pong = X_ - X
    End If
    If (Y - Y_) / (pong) > 1 Then
        vect = (Y - Y_) / (pong)
        For f2 = X To X_
            from_f3 = Y - vect * (f2 - X)
            to_f3 = Y - vect * (f2 - X) - vect
            If to_f3 <= 0 Then
                to_f3 = 1
            End If
            For f3 = from_f3 To to_f3 Step -1
                If f3 = Y_ Then
                    Exit For
                End If
                B.Cells(f3, f2).Borders(xlEdgeRight).Weight = xlThin
            Next f3
        Next f2
        flag = 1
    Else
        vect = (X_ - X) / (Y - Y_)
        For f2 = Y To Y_ Step -1
            from_f3 = X + vect * (Y - f2)
            to_f3 = X + vect * (Y - f2) + vect
            For f3 = from_f3 + 1 To to_f3 + 1
                If f3 = X_ Then
                    Exit For
                End If
                B.Cells(f2, f3).Borders(xlEdgeRight).Weight = xlThin
            Next f3
        Next f2
        flag = 2
    End If
'===================2번째 구간========================================
    If flag = 1 Then
        X = resul(f1, 2) / Delta_X
        X_ = resul(f1, 3) / Delta_X
        Y = f3 + 1
        Y_ = (init_Y - (link_length(1) + link_length(2)) / 20)
    Else
        X = f3 - 4
        X_ = resul(f1, 3) / Delta_X
        Y = init_Y - link_length(1) / 20
        Y_ = (init_Y - (link_length(1) + link_length(2)) / 20)
    End If
    flag = 0
    If X_ - X <= 0 Then
        pong = 1
    Else
        pong = X_ - X
    End If
    If (Y - Y_) / (pong) > 1 Then
        vect = (Y - Y_) / (pong)
        For f2 = X To X_
            from_f3 = Y - vect * (f2 - X)
            to_f3 = Y - vect * (f2 - X) - vect
            If to_f3 <= 0 Then
                to_f3 = 1
            End If
            For f3 = from_f3 + 1 To to_f3 Step -1
                If f3 = Y_ Then
                    Exit For
                End If
                B.Cells(f3, f2).Borders(xlEdgeBottom).Weight = xlThin
            Next f3
        Next f2
        flag = 1
    Else
        vect = (X_ - X) / (Y - Y_)
        For f2 = Y To Y_ Step -1
            from_f3 = X + vect * (Y - f2)
            to_f3 = X + vect * (Y - f2) + vect
            For f3 = from_f3 + 1 To to_f3 + 1
                If f3 = X_ Then
                    Exit For
                End If
                B.Cells(f2, f3).Borders(xlEdgeBottom).Weight = xlThin
            Next f3
        Next f2
        flag = 2
    End If
'===================3번째 구간========================================
    If flag = 1 Then
        X = resul(f1, 3) / Delta_X
        X_ = resul(f1, 4) / Delta_X
        Y = f3 + 1
        Y_ = (init_Y - (link_length(1) + link_length(2) + link_length(3)) / 20)
    Else
        X = f3 - 4
        X_ = resul(f1, 4) / Delta_X
        Y = (init_Y - (link_length(1) + link_length(2)) / 20)
        Y_ = (init_Y - (link_length(1) + link_length(2) + link_length(3)) / 20)
    End If
    flag = 0
    If X_ - X <= 0 Then
        pong = 1
    Else
        pong = X_ - X
    End If
    If (Y - Y_) / (pong) > 1 Then
        vect = (Y - Y_) / (pong)
        For f2 = X To X_
            from_f3 = Y - vect * (f2 - X)
            to_f3 = Y - vect * (f2 - X) - vect
            If to_f3 <= 0 Then
                to_f3 = 1
            End If
            For f3 = from_f3 + 1 To to_f3 Step -1
                If f3 = Y_ Then
                    Exit For
                End If
                B.Cells(f3, f2).Borders(xlEdgeRight).Weight = xlThin
            Next f3
        Next f2
    Else
        vect = (X_ - X) / (Y - Y_)
        For f2 = Y To Y_ Step -1
            from_f3 = X + vect * (Y - f2)
            to_f3 = X + vect * (Y - f2) + vect
            For f3 = from_f3 + 1 To to_f3 + 1
                If f3 = X_ Then
                    Exit For
                End If
                B.Cells(f2, f3).Borders(xlEdgeRight).Weight = xlThin
            Next f3
        Next f2
    End If
Next f1
End Sub

