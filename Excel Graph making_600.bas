Attribute VB_Name = "Module1"
Sub �׷���()
Dim wrk As Workbook
Dim a As Worksheet
Dim sht As Worksheet
Set wrk = ThisWorkbook
n = 128         '�������
k = 1
j = 1
Set sht = wrk.Sheets("��÷��fre")


        For i = 0 To n * 6 Step 2
            sht.Shapes.AddChart.Select
               With ActiveChart
                    .ChartType = xlXYScatterSmooth                     'chart type ����(��� �ִ� �л���)
                    .SetSourceData Source:=Range("$b$3:$c$17")
                    .SetSourceData Source:=Range("b3:c17").Offset(, i)
                    .ApplyLayout (1)
                    .Axes(xlValue).AxisTitle.Select                         '��Ʈ y�� �̸� ����
                    .Axes(xlValue, xlPrimary).AxisTitle.Text = "��"       '�̸� �̰ɷ� �ض�
                    .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 8
                    .Axes(xlCategory).AxisTitle.Select                      '��Ʈ x�� �̸� ����
                    .Axes(xlCategory, xlPrimary).AxisTitle.Text = "����ð�" '�̸� �̰ɷ� �ض�
                    .Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = 8
                    .ChartTitle.Select '��Ʈ ���� ����
                    .ChartTitle.Text = sht.Cells(1, i + 3) & sht.Cells(2, i + 3) '��Ʈ ������ �� ���� �������� �ض� (1�� °, 2�� °)
                    .ChartTitle.Font.Size = 10
                    .Legend.Select
                    .Legend.Font.Size = 8
                    .ChartArea.Select
                        With Selection
                            
                            If k = i / 12 Then
                                j = i / 12 / k
                                k = k + 1
                            End If
                            
                            .Top = Range("b10").Offset(13 * k, j).Top
                            .Left = Range("b10").Offset(, 7 * j).Left
                            j = j + 1
                        End With
                End With
       Next i
'Next f1
End Sub
   
