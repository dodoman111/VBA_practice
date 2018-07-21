Attribute VB_Name = "Module1"
Sub 그래프()
Dim wrk As Workbook
Dim a As Worksheet
Dim sht As Worksheet
Set wrk = ThisWorkbook
n = 128         '정류장수
k = 1
j = 1
Set sht = wrk.Sheets("비첨두fre")


        For i = 0 To n * 6 Step 2
            sht.Shapes.AddChart.Select
               With ActiveChart
                    .ChartType = xlXYScatterSmooth                     'chart type 지정(곡선이 있는 분산형)
                    .SetSourceData Source:=Range("$b$3:$c$17")
                    .SetSourceData Source:=Range("b3:c17").Offset(, i)
                    .ApplyLayout (1)
                    .Axes(xlValue).AxisTitle.Select                         '차트 y축 이름 선택
                    .Axes(xlValue, xlPrimary).AxisTitle.Text = "빈도"       '이름 이걸로 해라
                    .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 8
                    .Axes(xlCategory).AxisTitle.Select                      '차트 x축 이름 선택
                    .Axes(xlCategory, xlPrimary).AxisTitle.Text = "통행시간" '이름 이걸로 해라
                    .Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = 8
                    .ChartTitle.Select '차트 제목 선택
                    .ChartTitle.Text = sht.Cells(1, i + 3) & sht.Cells(2, i + 3) '차트 제목은 이 셀의 내용으로 해라 (1번 째, 2번 째)
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
   
