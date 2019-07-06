Attribute VB_Name = "StockSystem2018"
Public PassedNumber As Integer
Public PassedProportionCount As Integer
Public PassedZDF As Double '涨跌幅阈值
Public PassedZT As Double '涨停阈值
Public PassedPer As Double '卖额均值大于买额均值的比例
Public RowCount As Integer

'从SOURCE表复制代码和简称到DATA表
Public Sub CopyHead()
    Application.ScreenUpdating = False
    
    Dim arr
    Dim i
    
    i = Worksheets("SOURCE").Cells(Rows.Count, 2).End(xlUp).row - 1
    arr = Worksheets("SOURCE").Range("A2:B" & i + 1)
    Worksheets("DATA").Range("A2:B" & i + 1).Resize(i) = arr
    
    Erase arr
        
    Application.ScreenUpdating = True
End Sub

'计算委比
Public Sub Calculate()
    Application.ScreenUpdating = False
    
    Dim arr, brr, crr, Proportion
    Dim i As Integer
    Dim x As Integer
    i = Worksheets("SOURCE").Cells(Rows.Count, 2).End(xlUp).row - 1
    ReDim Proportion(1 To i, 1 To 1)
    
    '将源数据装入数组
    arr = Worksheets("SOURCE").Range("H2:V" & i + 1)
   
    '循环计算，委比
    For x = 1 To i
        '排除出现0的情况后计算委比
        If arr(x, 1) <> 0 Or arr(x, 2) <> 0 Or arr(x, 3) <> 0 Or arr(x, 4) <> 0 Or arr(x, 5) <> 0 Or arr(x, 11) <> 0 Or arr(x, 12) <> 0 Or arr(x, 13) <> 0 Or arr(x, 14) <> 0 Or arr(x, 15) <> 0 Then
            Proportion(x, 1) = (arr(x, 1) + arr(x, 2) + arr(x, 3) + arr(x, 4) + arr(x, 5) - arr(x, 11) - arr(x, 12) - arr(x, 13) - arr(x, 14) - arr(x, 15)) / (arr(x, 1) + arr(x, 2) + arr(x, 3) + arr(x, 4) + arr(x, 5) + arr(x, 11) + arr(x, 12) + arr(x, 13) + arr(x, 14) + arr(x, 15)) * 100
        Else
            Proportion(x, 1) = 0
        End If
    Next x
    
    '从数组输出到表中
    Worksheets("DATA").Range("C2").Resize(i) = Proportion
 
    '清空数组
    Erase arr, Proportion
    
    Application.ScreenUpdating = True
End Sub

'首次复制委比前值
Public Sub CopyFrontValue()
    Application.ScreenUpdating = False
    
    Dim arr
    Dim i
    With Worksheets("DATA")
        i = .Cells(Rows.Count, 2).End(xlUp).row - 1
    
        arr = .Range("C2:C" & i + 1)
    
        .Range("D2").Resize(i) = arr
    
        Erase arr
    End With
        
    Application.ScreenUpdating = True

End Sub

'复制委比后值
Public Sub CopyBackValue()
    Application.ScreenUpdating = False
    
    Dim arr
    Dim i
    With Worksheets("DATA")
        i = .Cells(Rows.Count, 2).End(xlUp).row - 1
    
        arr = .Range("C2:C" & i + 1)
    
        .Range("E2").Resize(i) = arr
    
        Erase arr
    End With
        
    Application.ScreenUpdating = True
End Sub


'计算委比差值
Public Sub Difference()
    Application.ScreenUpdating = False
    Dim arr, Difference
    Dim i As Integer
    Dim x As Integer

    With Worksheets("DATA")
        i = .Cells(Rows.Count, 2).End(xlUp).row - 1
        ReDim Difference(1 To i, 1 To 1)
    
        '将源数据装入数组
        arr = .Range("D2:E" & i + 1)
    
        '循环计算委比差值
        For x = 1 To i
            Difference(x, 1) = arr(x, 2) - arr(x, 1)
        Next x
    
        '从数组输出到表中
        .Range("F2").Resize(i) = Difference
    
        Erase arr, Difference
    End With
    
    Application.ScreenUpdating = True
    
End Sub

'进行委比计数
Public Sub ProportionCount()
    Application.ScreenUpdating = False
    
    '定义参数值
    PassedNumber = Worksheets("PARAMETER").Range("B1").Value
    PassedProportionCount = Worksheets("PARAMETER").Range("B2").Value
    
    Dim arr, BuyCount
    Dim i As Integer
    Dim x As Integer
    
    With Worksheets("DATA")
        i = .Cells(Rows.Count, 2).End(xlUp).row - 1
    
        '将源数据装入数组
        arr = .Range("D2:G" & i + 1)
        BuyCount = .Range("G2:G" & i + 1)
    
        For x = 1 To i
            If arr(x, 1) < 0 And arr(x, 2) > 0 And arr(x, 3) >= PassedNumber Then
                BuyCount(x, 1) = BuyCount(x, 1) + 1
            End If
        Next x
    
        '从数组输出到表中
        .Range("G2").Resize(i) = BuyCount
        
        Erase arr, BuyCount
    End With
    
    Application.ScreenUpdating = True
    
End Sub

'计数完毕将后值剪切到前值函数

Public Sub CutBackValue()
    Application.ScreenUpdating = False
    
    With Worksheets("DATA")
        .Range("E2:E" & .Cells(Rows.Count, 2).End(xlUp).row).Cut .Range("D2")
    End With
    
    Application.ScreenUpdating = True
    
End Sub

'计算委买金额和委卖金额现值
Public Sub MoneyCalc()
    Application.ScreenUpdating = False
    
    Dim arr, BuyMoney, SellMoney
    Dim i As Integer
    Dim x As Integer
    i = Worksheets("SOURCE").Cells(Rows.Count, 2).End(xlUp).row - 1
    ReDim SellMoney(1 To i, 1 To 1), BuyMoney(1 To i, 1 To 1)
    
    '将源数据装入数组
    arr = Worksheets("SOURCE").Range("C2:V" & i + 1)
    
    '循环计算委卖金额，委买金额
    For x = 1 To i
        BuyMoney(x, 1) = (arr(x, 1) * arr(x, 6) + arr(x, 2) * arr(x, 7) + arr(x, 3) * arr(x, 8) + arr(x, 4) * arr(x, 9) + arr(x, 5) * arr(x, 10)) * 100
        SellMoney(x, 1) = (arr(x, 11) * arr(x, 16) + arr(x, 12) * arr(x, 17) + arr(x, 13) * arr(x, 18) + arr(x, 14) * arr(x, 19) + arr(x, 15) * arr(x, 20)) * 100
    Next x
    
    '从数组输出到表中
    With Worksheets("DATA")
        .Range("I2").Resize(i) = BuyMoney
        .Range("N2").Resize(i) = SellMoney
    End With
    
    '清空数组
    Erase arr, BuyMoney, SellMoney
       
    Application.ScreenUpdating = True
End Sub

'复制记录买额总和、卖额总和、卖额均值、买额均值
Public Sub CopyBuyAndSellMoney()
    Application.ScreenUpdating = False
    
    Dim crr, drr
    Dim AllBuyMoney, AllSellMoney, AllBuyCount, AllSellCount, BuyAvg, SellAvg
    Dim i As Integer
    Dim x As Integer
    Dim y As Integer
    
    i = Worksheets("SOURCE").Cells(Rows.Count, 2).End(xlUp).row - 1
    
    ReDim AllBuyMoney(1 To i, 1 To 1), AllSellMoney(1 To i, 1 To 1), AllBuyCount(1 To i, 1 To 1), AllSellCount(1 To i, 1 To 1)
    ReDim BuyAvg(1 To i, 1 To 1), SellAvg(1 To i, 1 To 1)
    arr = Worksheets("DATA").Range("I2:I" & i + 1)
    brr = Worksheets("DATA").Range("N2:N" & i + 1)
    
    '处理买额总和及买额均值，次数
    crr = Worksheets("DATA").Range("I2:J" & i + 1)
    AllBuyCount = Worksheets("DATA").Range("L2:L" & i + 1)
    BuyAvg = Worksheets("DATA").Range("M2:M" & i + 1)
    '循环计算委买金额
 If (TimeValue(Now()) >= TimeValue("09:30:00") And TimeValue(Now()) <= TimeValue("11:30:00")) Or (TimeValue(Now()) >= TimeValue("13:00:00") And TimeValue(Now()) <= TimeValue("15:00:00")) Then
        For x = 1 To i
            AllBuyMoney(x, 1) = crr(x, 1) + crr(x, 2)
            AllBuyCount(x, 1) = AllBuyCount(x, 1) + 1
            BuyAvg(x, 1) = AllBuyMoney(x, 1) / AllBuyCount(x, 1)
        Next x
        
        '填入数组
        Worksheets("DATA").Range("K2").Resize(i) = AllBuyMoney
        Worksheets("DATA").Range("L2").Resize(i) = AllBuyCount
        Worksheets("DATA").Range("M2").Resize(i) = BuyAvg
        '将买额总值赋值给买额总值前值
        Worksheets("DATA").Range("J2").Resize(i) = AllBuyMoney
    End If
      
    '处理卖额总和及卖额均值，次数
    drr = Worksheets("DATA").Range("N2:O" & i + 1)
    AllSellCount = Worksheets("DATA").Range("Q2:Q" & i + 1)
    SellAvg = Worksheets("DATA").Range("R2:R" & i + 1)
    '循环计算委卖金额
     If (TimeValue(Now()) >= TimeValue("09:30:00") And TimeValue(Now()) <= TimeValue("11:30:00")) Or (TimeValue(Now()) >= TimeValue("13:00:00") And TimeValue(Now()) <= TimeValue("15:00:00")) Then
        For y = 1 To i
            AllSellMoney(y, 1) = drr(y, 1) + drr(y, 2)
            AllSellCount(y, 1) = AllSellCount(y, 1) + 1
            SellAvg(y, 1) = AllSellMoney(y, 1) / AllSellCount(y, 1)
        Next y
        
        '填入数组
        Worksheets("DATA").Range("P2").Resize(i) = AllSellMoney
        Worksheets("DATA").Range("Q2").Resize(i) = AllSellCount
        Worksheets("DATA").Range("R2").Resize(i) = SellAvg
        '将卖额总值赋值给买额总值前值
        Worksheets("DATA").Range("O2").Resize(i) = AllSellMoney
    End If
    
    Erase crr, drr, AllBuyMoney, AllSellMoney, AllBuyCount, AllSellCount, BuyAvg, SellAvg

    Application.ScreenUpdating = True
End Sub

'计算（买额总和+卖额总和）/总股本*MA10
Public Sub SpeedCalc()
    Application.ScreenUpdating = False
    
    Dim BuySum, SellSum, ZGB, MAArr, Result
    Dim i As Integer
    Dim x As Integer
    
    i = Worksheets("DATA").Cells(Rows.Count, 2).End(xlUp).row - 1
    
    ReDim BuySum(1 To i, 1 To 1), SellSum(1 To i, 1 To 1), ZGB(1 To i, 1 To 1), MAArr(1 To i, 1 To 1), Result(1 To i, 1 To 1)
    
    BuySum = Worksheets("DATA").Range("K2:K" & i + 1)
    SellSum = Worksheets("DATA").Range("P2:P" & i + 1)
    ZGB = Worksheets("SOURCE").Range("X2:X" & i + 1)
    MAArr = Worksheets("DATA").Range("H2:H" & i + 1)
    
    '循环计算（买额总和+卖额总和）/总股本*MA10
    For x = 1 To i
        Result(x, 1) = (BuySum(x, 1) + SellSum(x, 1)) / ZGB(x, 1) * MAArr(x, 1)
    Next x
    
    '填入数组
    Worksheets("DATA").Range("S2").Resize(i) = Result
    
    Erase BuySum, SellSum, ZGB, MAArr, Result
    
    Application.ScreenUpdating = True
End Sub

'结果输出
Public Sub ResultOutput()
    Application.ScreenUpdating = False
    
    '定义参数值
    PassedProportionCount = Worksheets("PARAMETER").Range("B2").Value
    PassedZDF = Worksheets("PARAMETER").Range("B3").Value
    PassedZT = Worksheets("PARAMETER").Range("B4").Value
    PassedPer = Worksheets("PARAMETER").Range("B5").Value
    
    Dim i As Integer
    Dim k As Integer
    
    For i = 2 To Worksheets("DATA").Cells(Rows.Count, 2).End(xlUp).row
        If Worksheets("DATA").Range("G" & i) >= PassedProportionCount And Worksheets("DATA").Range("R" & i) >= Worksheets("DATA").Range("M" & i) * PassedPer And Worksheets("SOURCE").Range("W" & i) > PassedZDF And Worksheets("SOURCE").Range("W" & i) <= PassedZT Then
            k = Worksheets("OUTPUT").Cells(Rows.Count, 7).End(xlUp).row + 1
            With Worksheets("OUTPUT")
                .Range("A" & k) = Worksheets("SOURCE").Range("A" & i)
                .Range("B" & k) = Worksheets("SOURCE").Range("B" & i)
                .Range("C" & k) = Worksheets("DATA").Range("G" & i)
                .Range("D" & k) = Worksheets("DATA").Range("M" & i)
                .Range("E" & k) = Worksheets("DATA").Range("R" & i)
                .Range("F" & k) = Format(Now, "yyyy/m/d h:mm")
                .Range("G" & k) = Worksheets("DATA").Range("S" & i)
            End With
        End If
    Next i
    
    Application.ScreenUpdating = True
    
End Sub

'删除结果重复项(暂时不用，备用）
Public Sub ClearBuyRepetition()
    Application.ScreenUpdating = False
    
    Dim d As Object
    Dim d1 As Object
    Dim R%, i%
    Dim arr, xm, brr, aa
    Dim rng As Range
    Set d = CreateObject("scripting.dictionary")
    Set d1 = CreateObject("scripting.dictionary")
    With Worksheets("OUTPUT")
    R = .Cells(.Rows.Count, 1).End(xlUp).row
    arr = .Range("A1:G" & R)
    For i = 2 To UBound(arr)
      xm = arr(i, 1) & "+" & arr(i, 2)
      If Not d.Exists(xm) Then
        d(xm) = Array(i, arr(i, 3), arr(i, 4), arr(i, 5), arr(i, 6), arr(i, 7))
      Else
        brr = d(xm)
        If brr(1) = arr(i, 3) Then
          If brr(4) > arr(i, 6) Then
            d1(brr(0)) = ""
          Else
            d1(i) = ""
          End If
        Else
          If brr(1) > arr(i, 3) Then
            d1(i) = ""
          Else
            d1(brr(0)) = ""
          End If
        End If
      End If
    Next
    For Each aa In d1.Keys
      If rng Is Nothing Then
        Set rng = .Rows(aa)
      Else
        Set rng = Union(rng, .Rows(aa))
      End If
    Next
    If Not rng Is Nothing Then
      rng.Delete
    End If
  End With
  
  Application.ScreenUpdating = True
End Sub

'按排序标准项排序结果
Public Sub ResultSort()
    Application.ScreenUpdating = False
    
    Dim i As Integer
    
    i = Worksheets("OUTPUT").Cells(Rows.Count, 7).End(xlUp).row - 1
    
    Worksheets("OUTPUT").Sort.SortFields.Clear
    Worksheets("OUTPUT").Sort.SortFields.Add Key:=Range("G2:G" & i + 1), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With Worksheets("OUTPUT").Sort
        .SetRange Range("A1:G" & i + 1)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Application.ScreenUpdating = True
End Sub

'清除OUTPUT页面选股结果
Public Sub ClearOutput()
    Application.ScreenUpdating = False
    
    Dim i As Integer
    
    i = Worksheets("OUTPUT").Cells(Rows.Count, 7).End(xlUp).row - 1
    '防止只有第一行时反向删除第一行
    If i > 1 Then
        Worksheets("OUTPUT").Range("A2:G" & i + 1).ClearContents
    End If
    Application.ScreenUpdating = True
End Sub

'按钮-写入历史记录表
Public Sub ButtonCopyToHistory()
    Application.ScreenUpdating = False

    Dim i, k
    i = Worksheets("OUTPUT").Cells(Rows.Count, 7).End(xlUp).row - 2
    k = Worksheets("HISTORY").Cells(Rows.Count, 7).End(xlUp).row + 1
    
    Worksheets("OUTPUT").Range("A2:G" & i + 2).Copy Worksheets("HISTORY").Range("A" & k)
           
    Application.ScreenUpdating = True
End Sub

'按钮-初始化委比及买卖额区域计数

Public Sub ButtonCountInit()
    Application.ScreenUpdating = False
    
    With Worksheets("DATA")
        .Range("G2:G" & .Cells(Rows.Count, 1).End(xlUp).row).Value = 0
        .Range("I2:S" & .Cells(Rows.Count, 1).End(xlUp).row).Value = 0
    End With
    
    CreateObject("sapi.spvoice").Speak "数据初始化完毕"
    
    Application.ScreenUpdating = True
End Sub

'按钮-隐藏表
Public Sub ButtonSheetHidden()
    Worksheets("SOURCE").Visible = xlSheetVeryHidden
    Worksheets("DATA").Visible = xlSheetVeryHidden
    Worksheets("PARAMETER").Visible = xlSheetVeryHidden
End Sub

'按钮-显示表
Public Sub ButtonSheetVisible()
    Application.ScreenUpdating = False
    
    Dim PassWord As String
    PassWord = Application.InputBox("请输入密码")
    If PassWord = "112233" Then
        Worksheets("DATA").Visible = xlSheetVisible
        Worksheets("PARAMETER").Visible = xlSheetVisible
        Worksheets("SOURCE").Visible = xlSheetVisible
    Else
        MsgBox ("密码错误！")
    End If
    
    Application.ScreenUpdating = True
End Sub

'语音提示
Public Sub Prompt()
    Dim i As Integer
    Dim speech
    speech = "注意，有新股票"
    
    i = Worksheets("OUTPUT").Cells(Rows.Count, 7).End(xlUp).row
    
    If i > RowCount Then
        
        Application.speech.Speak speech
        
    End If
End Sub

'欢迎语音
Public Sub Welcome()
    Dim welcomespeech
    welcomespeech = "欢迎使用自动选股系统"
    
    Application.speech.Speak welcomespeech
End Sub

'一分钟循环，负责处理委比计数
Public Sub LoopModule1()

    Call CopyBackValue
    Call Difference
    Call ProportionCount
    Call CutBackValue
        
    '按照指定时间循环
    Application.OnTime Now + TimeValue("00:01:00"), "LoopModule1"
End Sub

'六秒循环部分
Public Sub LoopModule2()

    Call Calculate
    Call MoneyCalc
        
    '按照指定时间循环
    Application.OnTime Now + TimeValue("00:00:06"), "LoopModule2"
End Sub

'三分钟循环部分
Public Sub LoopModule3()

    RowCount = Worksheets("OUTPUT").Cells(Rows.Count, 7).End(xlUp).row
    Call CopyBuyAndSellMoney
    Call SpeedCalc
    Call ClearOutput
    Call ResultOutput
    Call ResultSort
    Call Prompt
        
    '按照指定时间循环
    Application.OnTime Now + TimeValue("00:03:00"), "LoopModule3"
End Sub

'主要功能循环
Public Sub MainFunc()
    Call CopyFrontValue
    Call LoopModule1
    Call LoopModule2
    Call LoopModule3
End Sub

'主函数
Public Sub Main()
    Dim latespeech
    Dim earlyspeech
    
    latespeech = "软件开启太晚，需要手动初始化数据并开启选股"
    earlyspeech = "选股将九点三十分自动开启，请等待"
    
    'Call ButtonSheetHidden
    Call CopyHead
    '有效期认证
    If Environ("COMPUTERNAME") = "DELL-WOO" Or Environ("COMPUTERNAME") = "LENOVO-WOO" Then
        Call Welcome
        
         '启动时间识别是否需要手动开启循环
        If Worksheets("SOURCE").Range("AB2") > "09:30:00" Then
            Application.speech.Speak latespeech
        Else
            Application.speech.Speak earlyspeech
            '定时启动循环数据
            Application.OnTime TimeValue("9:30:00"), "MainFunc"
        End If
          
        '定时初始化计数
        Application.OnTime TimeValue("9:29:50"), "ButtonCountInit"
    Else
        MsgBox ("软件出现故障")
    End If

End Sub

'主功能手动开启开关
Public Sub ButtonMainFuncOn()
    Dim onspeech
    
    onspeech = "选股功能开启成功"
    
    Call MainFunc
    
    Application.speech.Speak onspeech
    
End Sub


'复制委比后值
Public Sub SS()
    Application.ScreenUpdating = False
    
    Dim arr, brr, crr
    Dim i, x
    With Worksheets("HISTORY")
        i = .Cells(Rows.Count, 2).End(xlUp).row - 1
    
        arr = .Range("D2:E" & i + 1)
        brr = .Range("H2:H" & i + 1)
        crr = .Range("I2:I" & i + 1)
    
        For x = 2 To i
            If arr(x, 2) >= arr(x, 1) * 1.5 And crr(x, 1) > -0.1 Then
                brr(x, 1) = 1
            End If
        Next x
        
        '填入数组
        Worksheets("HISTORY").Range("H2").Resize(i) = brr
    
        Erase arr, brr
    End With
        
    Application.ScreenUpdating = True
End Sub

