Attribute VB_Name = "Module1"
Dim wbk As Workbook
Dim sht As Worksheet

Dim cur_sero(5) As Long     '������ �� ������ ���� ������ǥ
Dim cur_garo(5) As Long     '������ �� ������ ���� ������ǥ
Dim new_sero(5) As Long     '������ �� ������ ���ο� ������ǥ
Dim new_garo(5) As Long     '������ �� ������ ���ο� ������ǥ

Dim errcnt As Long          '����Ұ��ɿ��� �Ǻ��� ����
Dim situ, waytogo, ordr As String '���� ������ ���� �� ���� ����

Dim direc As Double         '�̵������� ������ ������ ���� ����

Dim lft_poss_low, lft_poss_up, forw_poss_low, forw_poss_up, rght_poss_low, rght_poss_up As Double '���⺰ �̵�Ȯ�� ���� ����

Sub Worm()


Set wbk = ThisWorkbook
Set sht = wbk.Sheets("Sheet1")

sht.Cells().Clear '�ش� ��ũ��Ʈ ���� �ʱ�ȭ




'---------------------- ���⺰ �̵�Ȯ�� ���� --------------------

lft_poss_low = 0
lft_poss_up = 1 / 4
forw_poss_low = lft_poss_up
forw_poss_up = 3 / 4
rght_poss_low = forw_poss_up
rhgt_poss_up = 1




'---------------------- ���� ��ġ ���� --------------------------

For i = 1 To 5 Step 1
    
    cur_sero(i) = 6 - i
    cur_garo(i) = 1
    sht.Cells(cur_sero(i), cur_garo(i)).Interior.Color = 255
        
Next i





'----------- ������ʹ� ���� �̵��� ����Ǵ� �ݺ� �κ� ----------

For jjj = 1 To 500 Step 1               ' �ݺ�Ƚ�� 500ȸ



    For iiii = 1 To 3000000 Step 1      ' ������ �̵���Ȳ�� Ȯ���ϵ��� ���� �� ������ �ð��� ���� ���� ������� For-Next ����
 
    Next iiii                           ' iiii�� �ݺ�Ƚ���� ���̸� �����̰� �̵��ϴ� �ӵ��� ������
 



    '============= �������� ���� �� �Ӹ����� �Ǻ� ==================
    
    If cur_garo(1) = cur_garo(2) Then           ' 1�� ����� 2�� ������ ������ǥ���� ���� ��� (1, 2���� �������� ��ġ)
    
        If cur_sero(1) < cur_sero(2) Then       ' 1�� ������ ������ǥ���� 2�� ������ ������ǥ������ ������
            
            situ = "��"                         ' ���� �� �Ӹ������� �� ����
            
        Else
            
            situ = "��"                         ' ���� ���� ���, �� ū ���(���� ���� �����Ƿ�) ���� �� �Ӹ������� �� ����
            
        End If
        
    End If
    
 
    
    If cur_sero(1) = cur_sero(2) Then           ' 1�� ����� 2�� ������ ������ǥ���� ���� ��� (1, 2���� �������� ��ġ)
        
        If cur_garo(1) < cur_garo(2) Then       ' 1�� ������ ������ǥ���� 2�� ������ ������ǥ������ ������
        
            situ = "��"                         ' ���� �� �Ӹ������� �� ����
            
        Else
        
            situ = "��"                         ' ���� ���� ���, �� ū ���(���� ���� �����Ƿ�) ���� �� �Ӹ������� �� ����
            
        End If
            
    End If
    
    '=========== �������� ���� �� �Ӹ����� �Ǻ� ���� ==================
    
    
    
    
    '========= �����̰� ������ �̵��� ������ ������ ���� �߻� �� �̵����� ���� ========
    
    direc = Rnd()
    
    If lft_poss_low <= direc And direc < lft_poss_up Then           ' �������� ���� �̵� ������ ��������

        waytogo = "��"                                              ' �̵������� ����
        
    ElseIf forw_poss_low <= direc And direc < forw_poss_up Then     ' �������� ���� �̵� ������ ��������
        
        waytogo = "��"                                              ' �̵������� ����
        
    ElseIf rght_poss_low <= direc And direc < rght_poss_up Then     ' �������� ���� �̵� ������ ��������
        
        waytogo = "��"                                              ' �̵������� ����
        
    End If
    
    '======= �����̰� ������ �̵��� ������ ������ ���� �߻� �� �̵����� ���� ���� ======
    
    
    
    
    
    
'    MsgBox "�Ӹ��� ���� ����=" & situ & " / " & waytogo & " ���� �̵�"
    ' ������ �ּ�ǥ�ø� �����ϸ� �޼��� �ڽ��� ���� �������� ���� �̵����� ��ġ���θ� ������ �� ����




    '========= �Ӹ����� �� �̵����⺰ ���ο� �Ӹ���ǥ ��� �����Ǻ� ========

    If situ = "��" Then                             ' �Ӹ������� ���϶�


        If waytogo = "��" Then                          ' �̵������� �����̶��
        
            ordr = "��������"                           ' ��ĭ �Ʒ��� �������Ƿ� ���ΰ� ����
             
        End If
    
    
        If waytogo = "��" Then                          ' �̵������� �����̶��
    
            ordr = "���ΰ���"                           ' ��ĭ �������� ���Ƿ� ���ΰ� ����
    
        End If
    
    
        If waytogo = "��" Then                          ' �̵������� �����̶��
    
            ordr = "���ΰ���"                           ' ��ĭ ���� �ö󰡹Ƿ� ���ΰ� ����
       
        End If


    End If                                          ' �Ӹ����� ���϶��� �Ǻ� ����




    If situ = "��" Then                             ' �Ӹ������� ���϶�


        If waytogo = "��" Then                          ' �̵������� �����̶��
        
            ordr = "���ΰ���"                           ' ��ĭ �������� ���Ƿ� ���ΰ� ����
        
        End If
    
    
        If waytogo = "��" Then                          ' �̵������� �����̶��
    
            ordr = "���ΰ���"                           ' ��ĭ ���� �ö󰡹Ƿ� ���ΰ� ����
           
        End If
    
    
        If waytogo = "��" Then                          ' �̵������� �����̶��
    
            ordr = "��������"                           ' ��ĭ ���������� ���Ƿ� ���ΰ� ����
        
        End If
    
    
    End If                                          ' �Ӹ����� ���϶��� �Ǻ� ����




    If situ = "��" Then                             ' �Ӹ������� ���϶�


        If waytogo = "��" Then                          ' �̵������� �����̶��
        
            ordr = "���ΰ���"                           ' ��ĭ ���� �ö󰡹Ƿ� ���ΰ� ����
        
        End If
    
    
        If waytogo = "��" Then                          ' �̵������� �����̶��
    
            ordr = "��������"                           ' ��ĭ ���������� ���Ƿ� ���ΰ� ����
    
        End If
    
    
        If waytogo = "��" Then                          ' �̵������� �����̶��
    
            ordr = "��������"                           ' ��ĭ �Ʒ��� �������Ƿ� ���ΰ� ����
        
        End If
       
       
    End If                                          ' �Ӹ����� ���϶��� �Ǻ� ����




    If situ = "��" Then                             ' �Ӹ������� ���϶�


        If waytogo = "��" Then                          ' �̵������� �����̶��
        
            ordr = "��������"                           ' ��ĭ ���������� ���Ƿ� ���ΰ� ����
        
        End If
    
    
        If waytogo = "��" Then                          ' �̵������� �����̶��
    
            ordr = "��������"                           ' ��ĭ �Ʒ��� �������Ƿ� ���ΰ� ����
    
        End If
    
    
        If waytogo = "��" Then                          ' �̵������� �����̶��
    
            ordr = "���ΰ���"                           ' ��ĭ �������� ���Ƿ� ���ΰ� ����
       
        End If
            
            
    End If                                          ' �Ӹ����� ���϶��� �Ǻ� ����

    '======= �Ӹ����� �� �̵����⺰ ���ο� �Ӹ���ǥ ��� �����Ǻ� ����  ======




    '======= �Ӹ���ǥ ������� ========

    If ordr = "���ΰ���" Then
    
        new_sero(1) = cur_sero(1)                   ' ���ΰ� ��ȭ����
        new_garo(1) = cur_garo(1) - 1               ' ��ĭ �������� ���Ƿ� ���ο� �Ӹ���ǥ�� ���ΰ��� 1�� ��
    
    End If
    
    
    
    If ordr = "��������" Then
    
        new_sero(1) = cur_sero(1)                   ' ���ΰ� ��ȭ����
        new_garo(1) = cur_garo(1) + 1               ' ��ĭ ���������� ���Ƿ� ���ο� �Ӹ���ǥ�� ���ΰ��� 1�� ����
    
    End If
    
    
    
    If ordr = "���ΰ���" Then
    
        new_sero(1) = cur_sero(1) - 1               ' ��ĭ ���� �ö󰡹Ƿ� ���ο� �Ӹ���ǥ�� ���ΰ��� 1�� ��
        new_garo(1) = cur_garo(1)                   ' ���ΰ� ��ȭ����
 
    End If
    
    
    
    If ordr = "��������" Then
    
        new_sero(1) = cur_sero(1) + 1               ' ��ĭ �Ʒ��� �������Ƿ� ���ο� �Ӹ���ǥ�� ���ΰ��� 1�� ����
        new_garo(1) = cur_garo(1)                   ' ���ΰ� ��ȭ����

    End If

    '===== �Ӹ���ǥ ������� ���� =====



    
    
    '===== �����̴� ���� �̵��� �� ���� ���ΰ�? =====


    errcnt = 0                                                                              ' ����Ұ��ɿ��κ��� �� �ʱ�ȭ



    If new_sero(1) = 0 Or new_sero(1) = 24 Or new_garo(1) = 0 Or new_garo(1) = 60 Then      ' ���� ���� �Ӹ���ǥ���� ���������� ���� ���
    
        errcnt = 1                                                                          ' ����Ұ��ɺ��� ���� 1�� ����
    
    End If



    For jj = 2 To 4 Step 1                                                                  ' 2~4�� ������� �浹���� �Ǻ�

        If new_sero(1) = cur_sero(jj) And new_garo(1) = cur_garo(jj) Then                   ' ���ο� 1������� ������ 2~4�� ���� ������
    
            errcnt = 1                                                                      ' ����Ұ��ɺ��� ���� 1�� ����
        
        End If

    Next jj






    ' ========== ������ �����̸� �̵���Ű�� �κ� ================

    If errcnt = 1 Then                                                                      ' ���� ����Ұ��ɺ��� ���� 1�̸�

                                                                                            ' �ƹ��͵� ���� �ʰ� ���� ������

    Else                                                                                    ' ����Ұ��ɺ��� ���� 1�� �ƴϸ� (��, 0�̸�)


        ' ********* ���ο� ������ ��ġ�� ���� ���(2~5�� ����) �� ��ĥ *********

        For i = 1 To 5 Step 1

            If i = 1 Then                                                                   ' 1�� ���� ���� �κ�
        
                sht.Cells(new_sero(i), new_garo(i)).Interior.Color = 255                    ' ���ο� 1�� ���� ��ĥ
                
            Else                                                                            ' Else ���� : 2~5�� ���� ���Ͽ� ��� �� ��ĥ
        
                new_sero(i) = cur_sero(i - 1)                                               ' ���ο� i�� ����� ������ i-1�� ���� ���
                new_garo(i) = cur_garo(i - 1)
                sht.Cells(new_sero(i), new_garo(i)).Interior.Color = 255                    ' ���ο� 2~5�� ���� ��ĥ
            
            End If
        
        Next i

        sht.Cells(cur_sero(5), cur_garo(5)).Clear                                           ' ������ 5�� ����� ��ĥ ����

        
        ' ******* ���ο� ������ ��ġ�� ���� ���(2~5�� ����) �� ��ĥ ���� *******

    
    
        For i = 1 To 5 Step 1

            cur_sero(i) = new_sero(i)                                                       ' ���� ���� ���ؼ� �� ��ǥ�� "���� ��ǥ"�� ����
            cur_garo(i) = new_garo(i)
    
        Next i

    End If

Next jjj


End Sub


