Attribute VB_Name = "Module1"
Dim wrk             As Workbook
Dim a               As Worksheet
Dim B               As Worksheet
Dim BL()            As Long
Dim DL()            As Long
Dim C()             As Long

Sub 버블알고리듬()

Set wrk = ThisWorkbook
Set a = wrk.Sheets("Sheet1")
Set B = wrk.Sheets("Sheet2")

'========= 배열 넣기 =========

ReDim BL(100)
ReDim DL(100)

For i = 1 To 100
    If a.Cells(i + 1, 2) = "" Then Exit For
    BL(i) = a.Cells(i + 2, 2)
    DL(i) = a.Cells(i + 2, 2)
Next i

ReDim C(i)
N = UBound(C) - 2

'========= Bouble Algorithm (내림차순) =========

For k = 1 To N
    For j = 1 To N - 1
        If BL(j) < BL(j + 1) Then
            v = BL(j)
            BL(j) = BL(j + 1)
            BL(j + 1) = v
        End If
    Next j
Next k

'========= Bouble Algorithm (오름차순) =========

For k = 1 To N
    For j = 1 To N - 1
        If DL(j) > DL(j + 1) Then
            v = DL(j)
            DL(j) = DL(j + 1)
            DL(j + 1) = v
        End If
    Next j
Next k

'============ 출력단 ==========

For i = 1 To N
    a.Cells(i + 2, 5) = BL(i)
    a.Cells(i + 2, 7) = DL(i)
Next i


End Sub
