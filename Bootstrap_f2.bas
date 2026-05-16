Attribute VB_Name = "Bootstrap_f2"
'==============================================================
' Bootstrap f2 溶出曲线相似性分析宏
' Bootstrap f2 Dissolution Profile Similarity Analysis
'
' 理论依据 / Based on:
'   Shah VP et al. Pharm Res. 1998;15(6):889-896.
'   Mendyk A et al. Dissolut Technol. 2013;20(1):13-17.
'
' 使用方法 / Usage:
'   Alt+F8 -> RunBootstrap -> 运行/Run
'==============================================================
Option Explicit

'--------------------------------------------------------------
' 主入口 / Main entry point
'--------------------------------------------------------------
Public Sub RunBootstrap()

    Dim wsRef  As Worksheet, wsTst As Worksheet
    Dim wsPrm  As Worksheet, wsRes As Worksheet
    On Error GoTo ErrHandler

    Set wsRef = ThisWorkbook.Sheets("参考数据")
    Set wsTst = ThisWorkbook.Sheets("测试数据")
    Set wsPrm = ThisWorkbook.Sheets("参数设置")
    Set wsRes = ThisWorkbook.Sheets("计算结果")

    '─── 读取参数 / Read parameters ───────────────────────────
    Dim nBoot   As Long   : nBoot   = ReadLong(wsPrm, 3, 3, 5000)
    Dim ci      As Double : ci      = ReadDbl(wsPrm, 4, 3, 90)
    Dim sMode   As String : sMode   = Trim(CStr(wsPrm.Cells(5, 3).Value))
    Dim aRule   As String : aRule   = Trim(CStr(wsPrm.Cells(6, 3).Value))
    Dim rSeed   As Long   : rSeed   = ReadLong(wsPrm, 7, 3, 42)

    If nBoot < 100 Then nBoot = 100
    If nBoot > 50000 Then nBoot = 50000
    If ci <= 0 Or ci >= 100 Then ci = 90

    Randomize rSeed

    '─── 读取溶出数据 / Read dissolution data ─────────────────
    Dim nTP As Integer, nRefB As Integer, nTstB As Integer
    nTP   = CountTimePoints(wsRef)
    nRefB = CountBatches(wsRef)
    nTstB = CountBatches(wsTst)

    If nTP < 3 Then
        MsgBox "错误: 参考数据至少需要 3 个时间点！" & vbCrLf & _
               "Error: Need at least 3 time points.", vbCritical
        Exit Sub
    End If
    If nTP <> CountTimePoints(wsTst) Then
        MsgBox "错误: 参考与测试的时间点数量不一致！" & vbCrLf & _
               "Error: Reference and test have different number of time points.", vbCritical
        Exit Sub
    End If

    Dim refD() As Double : ReDim refD(1 To nTP, 1 To nRefB)
    Dim tstD() As Double : ReDim tstD(1 To nTP, 1 To nTstB)
    LoadData wsRef, refD, nTP, nRefB
    LoadData wsTst, tstD, nTP, nTstB

    '─── 原始 f1 / f2 ─────────────────────────────────────────
    Dim refMO() As Double : ReDim refMO(1 To nTP)
    Dim tstMO() As Double : ReDim tstMO(1 To nTP)
    ColMeans refD, nTP, nRefB, refMO
    ColMeans tstD, nTP, nTstB, tstMO

    Dim mask0() As Boolean : ReDim mask0(1 To nTP)
    ApplyRule refMO, tstMO, nTP, aRule, mask0

    Dim f1Orig As Double : f1Orig = CalcF1(refMO, tstMO, mask0, nTP)
    Dim f2Orig As Double : f2Orig = CalcF2(refMO, tstMO, mask0, nTP)
    Dim nPtsUsed As Integer : nPtsUsed = CountMask(mask0, nTP)

    '─── RSD 检查 ─────────────────────────────────────────────
    Dim rsdWarnRef As String, rsdWarnTst As String
    rsdWarnRef = RSDCheck(refD, nTP, nRefB)
    rsdWarnTst = RSDCheck(tstD, nTP, nTstB)

    '─── Bootstrap 循环 ───────────────────────────────────────
    Dim f2Boots() As Double : ReDim f2Boots(1 To nBoot)
    Dim bRefM()   As Double : ReDim bRefM(1 To nTP)
    Dim bTstM()   As Double : ReDim bTstM(1 To nTP)
    Dim bMask()   As Boolean: ReDim bMask(1 To nTP)

    Dim useWhole As Boolean
    useWhole = (InStr(1, LCase(sMode), "whole") > 0 Or _
                InStr(1, LCase(sMode), "vector") > 0)

    Dim i As Long, b As Integer, tp As Integer
    Application.ScreenUpdating = False

    For i = 1 To nBoot
        If useWhole Then
            BootWhole refD, nTP, nRefB, bRefM
            BootWhole tstD, nTP, nTstB, bTstM
        Else
            BootIndiv refD, nTP, nRefB, bRefM
            BootIndiv tstD, nTP, nTstB, bTstM
        End If
        ApplyRule bRefM, bTstM, nTP, aRule, bMask
        f2Boots(i) = CalcF2(bRefM, bTstM, bMask, nTP)

        If i Mod 200 = 0 Then
            Application.StatusBar = "Bootstrap 运行中... " & _
                                    Format(i / nBoot * 100, "0") & "% (" & i & "/" & nBoot & ")"
            DoEvents
        End If
    Next i

    Application.StatusBar = False
    Application.ScreenUpdating = True

    '─── 统计 / Statistics ────────────────────────────────────
    Dim sorted() As Double : ReDim sorted(1 To nBoot)
    Dim k As Long
    For k = 1 To nBoot : sorted(k) = f2Boots(k) : Next k
    QuickSort sorted, 1, nBoot

    Dim alpha As Double : alpha = (100 - ci) / 2 / 100
    Dim loIdx As Long   : loIdx = Application.WorksheetFunction.Max(1, CLng(alpha * nBoot))
    Dim hiIdx As Long   : hiIdx = Application.WorksheetFunction.Min(nBoot, CLng((1 - alpha) * nBoot))
    Dim f2Lo  As Double : f2Lo  = sorted(loIdx)
    Dim f2Hi  As Double : f2Hi  = sorted(hiIdx)
    Dim f2Mn  As Double : f2Mn  = ArrMean(f2Boots, nBoot)
    Dim f2Med As Double : f2Med = sorted(CLng(nBoot / 2))
    Dim f2SD  As Double : f2SD  = ArrSD(f2Boots, nBoot, f2Mn)
    Dim isSim As Boolean: isSim = (f2Lo > 50)

    '─── 写回结果 / Write results ─────────────────────────────
    WriteResults wsRes, f1Orig, f2Orig, rsdWarnRef, rsdWarnTst, nPtsUsed, nTP, _
                 nBoot, ci, sMode, aRule, _
                 f2Lo, f2Hi, f2Mn, f2Med, f2SD, _
                 sorted(1), sorted(nBoot), _
                 refMO, tstMO, mask0, isSim

    '─── 提示框 / Result message ──────────────────────────────
    Dim msg As String
    If isSim Then
        msg = "✔  相似 SIMILAR" & vbCrLf & vbCrLf & _
              "f2* (下置信限 Lower CI) = " & Format(f2Lo, "0.00") & " > 50" & vbCrLf & _
              "溶出曲线相似性已确认！Dissolution profiles confirmed similar."
        MsgBox msg, vbInformation, "Bootstrap f2 结果"
    Else
        msg = "✘  不相似 NOT SIMILAR" & vbCrLf & vbCrLf & _
              "f2* (下置信限 Lower CI) = " & Format(f2Lo, "0.00") & " ≤ 50" & vbCrLf & _
              "不能确认相似。建议将处方 f2 提升至约 " & _
              Format(f2Orig + (f2Orig - f2Lo) + 2, "0") & " 以上后重新验证。" & vbCrLf & _
              "Profiles NOT confirmed similar. Consider reformulating."
        MsgBox msg, vbCritical, "Bootstrap f2 结果"
    End If
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "发生错误 / Error: " & Err.Description & vbCrLf & _
           "(Error No. " & Err.Number & ")", vbCritical
End Sub

'==============================================================
' 数据读取辅助 / Data helpers
'==============================================================
Private Function ReadLong(ws As Worksheet, r As Integer, c As Integer, def As Long) As Long
    On Error Resume Next
    Dim v As Long : v = CLng(ws.Cells(r, c).Value)
    If Err.Number <> 0 Or v = 0 Then v = def
    On Error GoTo 0
    ReadLong = v
End Function

Private Function ReadDbl(ws As Worksheet, r As Integer, c As Integer, def As Double) As Double
    On Error Resume Next
    Dim v As Double : v = CDbl(ws.Cells(r, c).Value)
    If Err.Number <> 0 Or v = 0 Then v = def
    On Error GoTo 0
    ReadDbl = v
End Function

Private Function CountTimePoints(ws As Worksheet) As Integer
    Dim n As Integer : n = 0
    Dim r As Integer
    For r = 5 To 200
        If IsEmpty(ws.Cells(r, 2)) Or ws.Cells(r, 2).Value = "" Then Exit For
        n = n + 1
    Next r
    CountTimePoints = n
End Function

Private Function CountBatches(ws As Worksheet) As Integer
    Dim n As Integer : n = 0
    Dim c As Integer
    For c = 3 To 500
        If IsEmpty(ws.Cells(5, c)) Or ws.Cells(5, c).Value = "" Then Exit For
        n = n + 1
    Next c
    CountBatches = n
End Function

Private Sub LoadData(ws As Worksheet, d() As Double, nTP As Integer, nB As Integer)
    Dim ti As Integer, bi As Integer
    For ti = 1 To nTP
        For bi = 1 To nB
            On Error Resume Next
            d(ti, bi) = CDbl(ws.Cells(4 + ti, 2 + bi).Value)
            If Err.Number <> 0 Then d(ti, bi) = 0 : Err.Clear
            On Error GoTo 0
        Next bi
    Next ti
End Sub

Private Sub ColMeans(d() As Double, nTP As Integer, nB As Integer, m() As Double)
    Dim ti As Integer, bi As Integer, s As Double
    For ti = 1 To nTP
        s = 0
        For bi = 1 To nB : s = s + d(ti, bi) : Next bi
        m(ti) = s / nB
    Next ti
End Sub

Private Function CountMask(mask() As Boolean, nTP As Integer) As Integer
    Dim n As Integer : n = 0
    Dim ti As Integer
    For ti = 1 To nTP : If mask(ti) Then n = n + 1 : End If : Next ti
    CountMask = n
End Function

Private Function RSDCheck(d() As Double, nTP As Integer, nB As Integer) As String
    Dim result As String : result = ""
    Dim ti As Integer, bi As Integer
    For ti = 1 To nTP
        Dim s As Double, sq As Double : s = 0 : sq = 0
        For bi = 1 To nB : s = s + d(ti, bi) : Next bi
        Dim m As Double : m = s / nB
        For bi = 1 To nB : sq = sq + (d(ti, bi) - m) ^ 2 : Next bi
        Dim rv As Double
        If m > 0 Then rv = Sqr(sq / (nB - 1)) / m * 100 Else rv = 0
        Dim maxRSD As Double : maxRSD = IIf(ti = 1, 20, 10)
        If rv > maxRSD Then
            result = result & "t" & ti & "(RSD=" & Format(rv, "0.1") & "%) "
        End If
    Next ti
    If result = "" Then result = "OK (无超标 / No violation)"
    RSDCheck = result
End Function

'==============================================================
' 85% 截断规则 / Auto-rule
'==============================================================
Private Sub ApplyRule(refM() As Double, tstM() As Double, _
                      nTP As Integer, rule As String, mask() As Boolean)
    Dim ti As Integer
    For ti = 1 To nTP : mask(ti) = True : Next ti
    Dim r As String : r = LCase(Trim(rule))
    If InStr(r, "no") > 0 Then Exit Sub  ' no_auto

    Dim cutAt As Integer : cutAt = 0
    For ti = 1 To nTP
        Dim cond As Boolean
        If InStr(r, "2") > 0 Then
            cond = (refM(ti) > 85 And tstM(ti) > 85)
        Else
            cond = (refM(ti) > 85 Or tstM(ti) > 85)
        End If
        If cond Then cutAt = ti : Exit For
    Next ti

    If cutAt > 1 Then
        For ti = cutAt To nTP : mask(ti) = False : Next ti
    End If
End Sub

'==============================================================
' f1 / f2 计算 / Calculation
'==============================================================
Private Function CalcF1(refM() As Double, tstM() As Double, _
                        mask() As Boolean, nTP As Integer) As Double
    Dim num As Double, den As Double
    Dim ti As Integer
    For ti = 1 To nTP
        If mask(ti) Then
            num = num + Abs(refM(ti) - tstM(ti))
            den = den + refM(ti)
        End If
    Next ti
    If den > 0 Then CalcF1 = num / den * 100 Else CalcF1 = 0
End Function

Private Function CalcF2(refM() As Double, tstM() As Double, _
                        mask() As Boolean, nTP As Integer) As Double
    Dim sq As Double, n As Integer
    Dim ti As Integer
    For ti = 1 To nTP
        If mask(ti) Then
            sq = sq + (refM(ti) - tstM(ti)) ^ 2
            n = n + 1
        End If
    Next ti
    If n < 1 Then CalcF2 = 0 : Exit Function
    CalcF2 = 50 * Log(100 / Sqr(1 + sq / n)) / Log(10)
End Function

'==============================================================
' Bootstrap 采样 / Sampling
'==============================================================
Private Sub BootWhole(d() As Double, nTP As Integer, nB As Integer, m() As Double)
    ' 整体向量采样 / Whole-vector resampling
    Dim ti As Integer, b As Integer, idx As Integer
    For ti = 1 To nTP : m(ti) = 0 : Next ti
    For b = 1 To nB
        idx = Int(Rnd() * nB) + 1
        For ti = 1 To nTP : m(ti) = m(ti) + d(ti, idx) : Next ti
    Next b
    For ti = 1 To nTP : m(ti) = m(ti) / nB : Next ti
End Sub

Private Sub BootIndiv(d() As Double, nTP As Integer, nB As Integer, m() As Double)
    ' 单点采样 / Individual-value resampling
    Dim ti As Integer, b As Integer, idx As Integer
    For ti = 1 To nTP
        Dim s As Double : s = 0
        For b = 1 To nB
            idx = Int(Rnd() * nB) + 1
            s = s + d(ti, idx)
        Next b
        m(ti) = s / nB
    Next ti
End Sub

'==============================================================
' 统计辅助 / Statistics helpers
'==============================================================
Private Function ArrMean(a() As Double, n As Long) As Double
    Dim s As Double, i As Long
    For i = 1 To n : s = s + a(i) : Next i
    ArrMean = s / n
End Function

Private Function ArrSD(a() As Double, n As Long, mn As Double) As Double
    Dim sq As Double, i As Long
    For i = 1 To n : sq = sq + (a(i) - mn) ^ 2 : Next i
    If n > 1 Then ArrSD = Sqr(sq / (n - 1)) Else ArrSD = 0
End Function

' 快速排序 / QuickSort
Private Sub QuickSort(a() As Double, lo As Long, hi As Long)
    If lo >= hi Then Exit Sub
    Dim pivot As Double : pivot = a((lo + hi) \ 2)
    Dim i As Long : i = lo
    Dim j As Long : j = hi
    Dim tmp As Double
    Do
        Do While a(i) < pivot : i = i + 1 : Loop
        Do While a(j) > pivot : j = j - 1 : Loop
        If i <= j Then
            tmp = a(i) : a(i) = a(j) : a(j) = tmp
            i = i + 1 : j = j - 1
        End If
    Loop While i <= j
    If i Mod 2000 = 0 Then DoEvents
    QuickSort a, lo, j
    QuickSort a, i, hi
End Sub

'==============================================================
' 写回结果 / Write results to 计算结果 sheet
'==============================================================
Private Sub WriteResults(ws As Worksheet, _
    f1o As Double, f2o As Double, rsdR As String, rsdT As String, _
    nPts As Integer, nTP As Integer, _
    nBoot As Long, ci As Double, sMode As String, aRule As String, _
    f2Lo As Double, f2Hi As Double, f2Mn As Double, _
    f2Med As Double, f2SD As Double, f2Min As Double, f2Max As Double, _
    refM() As Double, tstM() As Double, mask() As Boolean, isSim As Boolean)

    Const GRN_BG As Long = &HDAEFE2  ' RGB(226,239,218)
    Const RED_BG As Long = &HCEC7FF  ' RGB(255,199,206)  -- VBA BGR
    Const GRN_FG As Long = &H235637
    Const RED_FG As Long = &H60009C

    ' A. Original
    SetCell ws, 4, 3, Round(f1o, 2),   "0.00"
    SetCell ws, 5, 3, Round(f2o, 2),   "0.00"
    SetCell ws, 6, 3, "Ref: " & rsdR & " | Test: " & rsdT, "@"
    SetCell ws, 7, 3, CStr(nPts) & " / " & CStr(nTP), "@"

    ' B. Bootstrap
    SetCell ws, 10, 3, CStr(nBoot), "@"
    SetCell ws, 11, 3, CStr(ci) & "%", "@"
    SetCell ws, 12, 3, sMode, "@"
    SetCell ws, 13, 3, Round(f2Mn, 2), "0.00"

    ' f2* key cell
    With ws.Cells(14, 3)
        .Value = Round(f2Lo, 2)
        .NumberFormat = "0.00"
        .Font.Bold = True
        .Font.Size = 13
        If isSim Then
            .Interior.Color = GRN_BG
            .Font.Color = GRN_FG
        Else
            .Interior.Color = RED_BG
            .Font.Color = RED_FG
        End If
    End With

    SetCell ws, 15, 3, Round(f2Hi, 2), "0.00"

    ' C. Decision
    Dim decTxt As String
    If isSim Then
        decTxt = "✔  相似 SIMILAR  (f2* = " & Format(f2Lo, "0.00") & " > 50)"
    Else
        decTxt = "✘  不相似 NOT SIMILAR  (f2* = " & Format(f2Lo, "0.00") & _
                 " ≤ 50)  →  建议 f2 提升至约 " & _
                 Format(f2o + (f2o - f2Lo) + 2, "0") & " 以上"
    End If
    With ws.Cells(18, 3)
        .Value = decTxt
        .NumberFormat = "@"
        .Font.Bold = True
        .Font.Size = 11
        If isSim Then
            .Interior.Color = GRN_BG : .Font.Color = GRN_FG
        Else
            .Interior.Color = RED_BG : .Font.Color = RED_FG
        End If
    End With

    ' D. Distribution
    SetCell ws, 21, 3, Round(f2Min, 2), "0.00"
    SetCell ws, 22, 3, Round(f2Med, 2), "0.00"
    SetCell ws, 23, 3, Round(f2Max, 2), "0.00"
    SetCell ws, 24, 3, Round(f2SD, 2),  "0.00"

    ' E. Time-point table
    Dim ti As Integer
    For ti = 1 To nTP
        Dim row As Integer : row = 27 + ti
        ws.Cells(row, 2).Value = ti  ' placeholder; ideally pull time values
        ws.Cells(row, 3).Value = Round(refM(ti), 2)
        ws.Cells(row, 4).Value = Round(tstM(ti), 2)
        ws.Cells(row, 5).Value = Round(Abs(refM(ti) - tstM(ti)), 2)
    Next ti
End Sub

Private Sub SetCell(ws As Worksheet, r As Integer, c As Integer, _
                    val As Variant, fmt As String)
    ws.Cells(r, c).Value = val
    ws.Cells(r, c).NumberFormat = fmt
End Sub
