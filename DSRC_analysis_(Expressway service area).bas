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
Dim �ްԼ�_TT()                     As Long
Dim ��������_TT()                   As Long
Dim �ްԼ�����()                    As Long
Dim TT()                            As Long
Dim �̿�Ƚ��()                      As Long
Dim �ްԼұ�������ð�()            As Long
Dim �ްԼ�����ù����()              As Long

Sub ��������()

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


'============== rse �Ÿ����� ���� =========
For f1 = 1 To 100000
    If b.Cells(f1 + 1, 1) = "" Then Exit For
        ind(b.Cells(f1 + 1, 1)) = b.Cells(f1 + 1, 2)
Next f1


'=============== �ްԼ� rse ���� ===========

ReDim SA(7)
ReDim �ްԼ�_TT(999999)
ReDim ��������_TT(999999)
ReDim �ްԼ�����(199999, 7, 2)              '�ްԼ�����(����id, �ްԼҹ�ȣ, �̿뿩��(0:�̿����, 1:�̿�)
ReDim �ްԼ�����ù����(199999)

For f1 = 1 To 7
    SA(f1) = a.Cells(3, f1 + 1)
Next f1

path_nm = a.Cells(1, 2)
file_nm = a.Cells(2, 2)
n = 7                                                                                   '������ӵ��� ������ �ްԼ� ����
bv = 50                                                                                 '�ްԼ� �̿�к� ���� ���

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
                            �ްԼ�_TT(cnt) = (split_4(2) - split_3(2)) / (ind(split_4(0)) - ind(split_3(0)))                 '���� km �� �ްԼ� ���� ����ð�
                            ��������_TT(cnt) = (split_3(2) - split_5(2)) / (ind(split_3(0)) - ind(split_5(0)))               '���� km �� �ްԼ� �������� ����ð�
                                                                                             
                            
                            '======== �ްԼ� �̿� ���� 5���̻� ������ ��� ============
                            If �ްԼ�_TT(cnt) > ��������_TT(cnt) + 300 / (ind(split_3(0)) - ind(split_5(0))) Then           '�ްԼ� �� ����
                                k = 1
                            Else: k = 0
                            End If
                            
                            If r <> rr Then
                                car = car + 1
                            End If
                            
                            �ްԼ�����(car, j, k) = split_4(2)                                                               '�ްԼ� ���� ������ rse �����ð�
            
                            �ްԼ�����ù����(car) = split_6(2)          '��ӵ��� ���� ù����
                            
                            rr = r
                        End If
                    End If
                End If
            Next i
        Next j
    Loop
Close #1



'========= ��´� ===========

For i = 1 To car
    c.Cells(i + 2, 1) = i           '������ȣID
    For j = 1 To 7
        For k = 0 To 1
        If �ްԼ�����(i, j, k) <> 0 Then                         '�ްԼ�����(����id, �ްԼҹ�ȣ, �̿뿩��(0:�̿����, 1:�̿�) =  �ްԼ� ���ͼ� ù RSE �����ð�
            If j = 1 Then
                c.Cells(i + 2, 2) = k
                c.Cells(i + 2, 3) = �ްԼ�����(i, j, k)
            End If
            If j = 2 Then
                c.Cells(i + 2, 4) = k
                c.Cells(i + 2, 5) = �ްԼ�����(i, j, k)
            End If
            If j = 3 Then
                c.Cells(i + 2, 6) = k
                c.Cells(i + 2, 7) = �ްԼ�����(i, j, k)
            End If
            If j = 4 Then
                c.Cells(i + 2, 8) = k
                c.Cells(i + 2, 9) = �ްԼ�����(i, j, k)
            End If
            If j = 5 Then
                c.Cells(i + 2, 10) = k
                c.Cells(i + 2, 11) = �ްԼ�����(i, j, k)
            End If
            If j = 6 Then
                c.Cells(i + 2, 12) = k
                c.Cells(i + 2, 13) = �ްԼ�����(i, j, k)
            End If
            If j = 7 Then
                c.Cells(i + 2, 14) = k
                c.Cells(i + 2, 15) = �ްԼ�����(i, j, k)
            End If
        End If
        Next k
    Next j
Next i


'=========== �ްԼ� �̿�Ƚ���� ���� �з� ==================

ReDim TT(car, 7)
ReDim �̿�Ƚ��(car)
ReDim �ްԼұ�������ð�(car, 7, 7)

For i = 1 To car
    For j = 1 To 7
        If c.Cells(i + 2, 2 * j) = 1 Then
            n = n + 1                                       '�ްԼ� �̿�Ƚ��
            TT(i, n) = c.Cells(i + 2, 2 * j + 1)
        End If
    Next j
    �̿�Ƚ��(i) = n                                         'i��° ���� �ްԼ� �̿�Ƚ��
    n = 0
Next i

''======== ����˰��� �������� ���� =====
'For i = 1 To car
'    For k = 1 To 6
'        m = TT(i, k)
'        If m >= TT(i, k + 1) Then
'            TT(i, k) = TT(i, k + 1)
'            TT(i, k + 1) = m
'        End If
'    Next k
'Next i


'===================== �ްԼұ��� ����ð� ���� =====
For i = 1 To car
    For k = 1 To 5
        If �̿�Ƚ��(i) > 1 Then
            If k = 1 Then                                                                                   '�ްԼ� 1�� �̿��ϸ�
                �ްԼұ�������ð�(i, �̿�Ƚ��(i), k) = Abs((TT(i, k) - �ްԼ�����ù����(i)) / 60)          '�ްԼұ��� ����ð� ������ ȯ��
            ElseIf k > 1 Then
                �ްԼұ�������ð�(i, �̿�Ƚ��(i), k) = Abs((TT(i, k) - TT(i, k - 1)) / 60)                 '�ްԼ� 2�� �̻� �̿��ϸ� �ްԼұ�������ð�(����id, �̿�Ƚ��) = ����ð�
            End If
        ElseIf �̿�Ƚ��(i) = 1 Then
            �ްԼұ�������ð�(i, �̿�Ƚ��(i), 1) = Abs(TT(i, 1) - �ްԼ�����ù����(i)) / 60
        End If
    Next k
Next i

'====== �ްԼ� �̿�Ƚ���� ���� ����ð� ���� ��� =====
For i = 1 To car
    If �̿�Ƚ��(i) = 1 Then
'        If �ްԼұ�������ð�(i, �̿�Ƚ��(i), 1) < 360 Then         '�̻�ġ ���� ���� ����ð� 6�ð� �̳�
            cnt_0 = cnt_0 + 1
            d.Cells(2 + cnt_0, 2) = �ްԼұ�������ð�(i, �̿�Ƚ��(i), 1)
'        End If
    End If
    If �̿�Ƚ��(i) = 2 Then
'        If �ްԼұ�������ð�(i, �̿�Ƚ��(i), 2) < 360 And �ްԼұ�������ð�(i, �̿�Ƚ��(i), 2) > 0 Then
            cnt_2 = cnt_2 + 1
            e.Cells(2 + cnt_2, 2) = �ްԼұ�������ð�(i, �̿�Ƚ��(i), 1)
            e.Cells(2 + cnt_2, 3) = �ްԼұ�������ð�(i, �̿�Ƚ��(i), 2)
'        End If
    End If
    If �̿�Ƚ��(i) = 3 Then
'        If �ްԼұ�������ð�(i, �̿�Ƚ��(i), 3) < 360 And �ްԼұ�������ð�(i, �̿�Ƚ��(i), 3) > 0 Then
            cnt_3 = cnt_3 + 1
            f.Cells(2 + cnt_3, 2) = �ްԼұ�������ð�(i, �̿�Ƚ��(i), 1)
            f.Cells(2 + cnt_3, 3) = �ްԼұ�������ð�(i, �̿�Ƚ��(i), 2)
            f.Cells(2 + cnt_3, 4) = �ްԼұ�������ð�(i, �̿�Ƚ��(i), 3)
'        End If
    End If
    If �̿�Ƚ��(i) >= 4 Then
'        If �ްԼұ�������ð�(i, �̿�Ƚ��(i), 4) < 300 And �ްԼұ�������ð�(i, �̿�Ƚ��(i), 4) > 0 Then
            cnt_4 = cnt_4 + 1
            g.Cells(2 + cnt_4, 2) = �ްԼұ�������ð�(i, �̿�Ƚ��(i), 1)
            g.Cells(2 + cnt_4, 3) = �ްԼұ�������ð�(i, �̿�Ƚ��(i), 2)
            g.Cells(2 + cnt_4, 4) = �ްԼұ�������ð�(i, �̿�Ƚ��(i), 3)
            g.Cells(2 + cnt_4, 5) = �ްԼұ�������ð�(i, �̿�Ƚ��(i), 4)
            g.Cells(2 + cnt_4, 6) = �ްԼұ�������ð�(i, �̿�Ƚ��(i), 5)
            g.Cells(2 + cnt_4, 7) = �ްԼұ�������ð�(i, �̿�Ƚ��(i), 6)
            g.Cells(2 + cnt_4, 8) = �ްԼұ�������ð�(i, �̿�Ƚ��(i), 7)
'        End If
    End If
Next i


End Sub
