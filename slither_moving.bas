Attribute VB_Name = "Module1"
Dim wbk As Workbook
Dim sht As Worksheet

Dim cur_sero(5) As Long     '지렁이 각 마디의 현재 세로좌표
Dim cur_garo(5) As Long     '지렁이 각 마디의 현재 가로좌표
Dim new_sero(5) As Long     '지렁이 각 마디의 새로운 세로좌표
Dim new_garo(5) As Long     '지렁이 각 마디의 새로운 가로좌표

Dim errcnt As Long          '진행불가능여부 판별용 변수
Dim situ, waytogo, ordr As String '다음 진행의 유형 및 방향 관련

Dim direc As Double         '이동방향을 결정할 난수값 저장 변수

Dim lft_poss_low, lft_poss_up, forw_poss_low, forw_poss_up, rght_poss_low, rght_poss_up As Double '방향별 이동확률 설정 변수

Sub Worm()


Set wbk = ThisWorkbook
Set sht = wbk.Sheets("Sheet1")

sht.Cells().Clear '해당 워크시트 내용 초기화




'---------------------- 방향별 이동확률 설정 --------------------

lft_poss_low = 0
lft_poss_up = 1 / 4
forw_poss_low = lft_poss_up
forw_poss_up = 3 / 4
rght_poss_low = forw_poss_up
rhgt_poss_up = 1




'---------------------- 최초 위치 설정 --------------------------

For i = 1 To 5 Step 1
    
    cur_sero(i) = 6 - i
    cur_garo(i) = 1
    sht.Cells(cur_sero(i), cur_garo(i)).Interior.Color = 255
        
Next i





'----------- 여기부터는 실제 이동이 수행되는 반복 부분 ----------

For jjj = 1 To 500 Step 1               ' 반복횟수 500회



    For iiii = 1 To 3000000 Step 1      ' 눈으로 이동상황을 확인하도록 다음 턴 까지의 시간을 벌기 위한 내용없는 For-Next 구문
 
    Next iiii                           ' iiii의 반복횟수를 줄이면 지렁이가 이동하는 속도가 빨라짐
 



    '============= 지렁이의 다음 턴 머리방향 판별 ==================
    
    If cur_garo(1) = cur_garo(2) Then           ' 1번 마디와 2번 마디의 가로좌표값이 같은 경우 (1, 2번이 수직으로 위치)
    
        If cur_sero(1) < cur_sero(2) Then       ' 1번 마디의 세로좌표값이 2번 마디의 세로좌표값보다 작으면
            
            situ = "↑"                         ' 다음 턴 머리방향은 ↑ 방향
            
        Else
            
            situ = "↓"                         ' 작지 않은 경우, 즉 큰 경우(같을 수는 없으므로) 다음 턴 머리방향은 ↓ 방향
            
        End If
        
    End If
    
 
    
    If cur_sero(1) = cur_sero(2) Then           ' 1번 마디와 2번 마디의 세로좌표값이 같은 경우 (1, 2번이 수평으로 위치)
        
        If cur_garo(1) < cur_garo(2) Then       ' 1번 마디의 가로좌표값이 2번 마디의 세로좌표값보다 작으면
        
            situ = "←"                         ' 다음 턴 머리방향은 ← 방향
            
        Else
        
            situ = "→"                         ' 작지 않은 경우, 즉 큰 경우(같을 수는 없으므로) 다음 턴 머리방향은 → 방향
            
        End If
            
    End If
    
    '=========== 지렁이의 다음 턴 머리방향 판별 종료 ==================
    
    
    
    
    '========= 지렁이가 실제로 이동할 지점을 결정할 난수 발생 및 이동방향 결정 ========
    
    direc = Rnd()
    
    If lft_poss_low <= direc And direc < lft_poss_up Then           ' 난수값이 좌측 이동 구간에 떨어지면

        waytogo = "좌"                                              ' 이동방향은 좌측
        
    ElseIf forw_poss_low <= direc And direc < forw_poss_up Then     ' 난수값이 전방 이동 구간에 떨어지면
        
        waytogo = "직"                                              ' 이동방향은 직진
        
    ElseIf rght_poss_low <= direc And direc < rght_poss_up Then     ' 난수값이 우측 이동 구간에 떨어지면
        
        waytogo = "우"                                              ' 이동방향은 우측
        
    End If
    
    '======= 지렁이가 실제로 이동할 지점을 결정할 난수 발생 및 이동방향 결정 종료 ======
    
    
    
    
    
    
'    MsgBox "머리가 향한 방향=" & situ & " / " & waytogo & " 방향 이동"
    ' 윗줄의 주석표시를 제거하면 메세지 박스를 통해 난수값과 실제 이동방향 일치여부를 점검할 수 있음




    '========= 머리방향 및 이동방향별 새로운 머리좌표 계산 유형판별 ========

    If situ = "←" Then                             ' 머리방향이 ←일때


        If waytogo = "좌" Then                          ' 이동방향이 좌측이라면
        
            ordr = "세로증가"                           ' 한칸 아래로 내려가므로 세로값 증가
             
        End If
    
    
        If waytogo = "직" Then                          ' 이동방향이 직진이라면
    
            ordr = "가로감소"                           ' 한칸 왼쪽으로 가므로 가로값 감소
    
        End If
    
    
        If waytogo = "우" Then                          ' 이동방향이 우측이라면
    
            ordr = "세로감소"                           ' 한칸 위로 올라가므로 세로값 감소
       
        End If


    End If                                          ' 머리방향 ←일때의 판별 종료




    If situ = "↑" Then                             ' 머리방향이 ↑일때


        If waytogo = "좌" Then                          ' 이동방향이 좌측이라면
        
            ordr = "가로감소"                           ' 한칸 왼쪽으로 가므로 가로값 감소
        
        End If
    
    
        If waytogo = "직" Then                          ' 이동방향이 직진이라면
    
            ordr = "세로감소"                           ' 한칸 위로 올라가므로 세로값 감소
           
        End If
    
    
        If waytogo = "우" Then                          ' 이동방향이 우측이라면
    
            ordr = "가로증가"                           ' 한칸 오른쪽으로 가므로 가로값 증가
        
        End If
    
    
    End If                                          ' 머리방향 ↑일때의 판별 종료




    If situ = "→" Then                             ' 머리방향이 →일때


        If waytogo = "좌" Then                          ' 이동방향이 좌측이라면
        
            ordr = "세로감소"                           ' 한칸 위로 올라가므로 세로값 감소
        
        End If
    
    
        If waytogo = "직" Then                          ' 이동방향이 직진이라면
    
            ordr = "가로증가"                           ' 한칸 오른쪽으로 가므로 가로값 증가
    
        End If
    
    
        If waytogo = "우" Then                          ' 이동방향이 우측이라면
    
            ordr = "세로증가"                           ' 한칸 아래로 내려가므로 세로값 증가
        
        End If
       
       
    End If                                          ' 머리방향 →일때의 판별 종료




    If situ = "↓" Then                             ' 머리방향이 ↓일때


        If waytogo = "좌" Then                          ' 이동방향이 좌측이라면
        
            ordr = "가로증가"                           ' 한칸 오른쪽으로 가므로 가로값 증가
        
        End If
    
    
        If waytogo = "직" Then                          ' 이동방향이 직진이라면
    
            ordr = "세로증가"                           ' 한칸 아래로 내려가므로 세로값 증가
    
        End If
    
    
        If waytogo = "우" Then                          ' 이동방향이 우측이라면
    
            ordr = "가로감소"                           ' 한칸 왼쪽으로 가므로 가로값 감소
       
        End If
            
            
    End If                                          ' 머리방향 ↓일때의 판별 종료

    '======= 머리방향 및 이동방향별 새로운 머리좌표 계산 유형판별 종료  ======




    '======= 머리좌표 실제계산 ========

    If ordr = "가로감소" Then
    
        new_sero(1) = cur_sero(1)                   ' 세로값 변화없음
        new_garo(1) = cur_garo(1) - 1               ' 한칸 왼쪽으로 가므로 새로운 머리좌표는 가로값에 1을 뺌
    
    End If
    
    
    
    If ordr = "가로증가" Then
    
        new_sero(1) = cur_sero(1)                   ' 세로값 변화없음
        new_garo(1) = cur_garo(1) + 1               ' 한칸 오른쪽으로 가므로 새로운 머리좌표는 가로값에 1을 더함
    
    End If
    
    
    
    If ordr = "세로감소" Then
    
        new_sero(1) = cur_sero(1) - 1               ' 한칸 위로 올라가므로 새로운 머리좌표는 세로값에 1을 뺌
        new_garo(1) = cur_garo(1)                   ' 가로값 변화없음
 
    End If
    
    
    
    If ordr = "세로증가" Then
    
        new_sero(1) = cur_sero(1) + 1               ' 한칸 아래로 내려가므로 새로운 머리좌표는 세로값에 1을 더함
        new_garo(1) = cur_garo(1)                   ' 가로값 변화없음

    End If

    '===== 머리좌표 실제계산 종료 =====



    
    
    '===== 지렁이는 과연 이동할 수 있을 것인가? =====


    errcnt = 0                                                                              ' 진행불가능여부변수 값 초기화



    If new_sero(1) = 0 Or new_sero(1) = 24 Or new_garo(1) = 0 Or new_garo(1) = 60 Then      ' 계산된 다음 머리좌표값이 공간범위를 넘을 경우
    
        errcnt = 1                                                                          ' 진행불가능변수 값을 1로 설정
    
    End If



    For jj = 2 To 4 Step 1                                                                  ' 2~4번 마디와의 충돌여부 판별

        If new_sero(1) = cur_sero(jj) And new_garo(1) = cur_garo(jj) Then                   ' 새로운 1번마디와 현재의 2~4번 마디가 같으면
    
            errcnt = 1                                                                      ' 진행불가능변수 값을 1로 설정
        
        End If

    Next jj






    ' ========== 실제로 지렁이를 이동시키는 부분 ================

    If errcnt = 1 Then                                                                      ' 만약 진행불가능변수 값이 1이면

                                                                                            ' 아무것도 하지 않고 다음 턴으로

    Else                                                                                    ' 진행불가능변수 값이 1이 아니면 (즉, 0이면)


        ' ********* 새로운 지렁이 위치에 대한 계산(2~5번 마디) 및 색칠 *********

        For i = 1 To 5 Step 1

            If i = 1 Then                                                                   ' 1번 마디에 대한 부분
        
                sht.Cells(new_sero(i), new_garo(i)).Interior.Color = 255                    ' 새로운 1번 마디에 색칠
                
            Else                                                                            ' Else 이하 : 2~5번 마디에 대하여 계산 및 색칠
        
                new_sero(i) = cur_sero(i - 1)                                               ' 새로운 i번 마디는 기존의 i-1번 마디를 계승
                new_garo(i) = cur_garo(i - 1)
                sht.Cells(new_sero(i), new_garo(i)).Interior.Color = 255                    ' 새로운 2~5번 마디에 색칠
            
            End If
        
        Next i

        sht.Cells(cur_sero(5), cur_garo(5)).Clear                                           ' 기존의 5번 마디는 색칠 제거

        
        ' ******* 새로운 지렁이 위치에 대한 계산(2~5번 마디) 및 색칠 종료 *******

    
    
        For i = 1 To 5 Step 1

            cur_sero(i) = new_sero(i)                                                       ' 다음 턴을 위해서 새 좌표를 "현재 좌표"로 설정
            cur_garo(i) = new_garo(i)
    
        Next i

    End If

Next jjj


End Sub


