Attribute VB_Name = "Module1"
Sub Add_data_to_Rows_From_The_Right()
prepare
    Dim i As Long, j As Long
    Dim dev_num As String
    Dim rast_dog_dict, obr_dict
    Set rast_dog_dict = CreateObject("Scripting.Dictionary")
    Set obr_dict = CreateObject("Scripting.Dictionary")
    With Worksheets("Расторгнутые договора")
        For j = 1 To 11719
            dev_num = CStr(.Cells(j, 3).Value)
            rast_dog_dict(CStr(.Cells(j, 3).Value)) = j
        Next j
    End With
    With Worksheets("Обращения")
        For j = 1 To 36741
            dev_num = CStr(.Cells(j, 30).Value)
            If rast_dog_dict.exists(dev_num) Then 'если расторг - посчитаем, сколько было обращений
                If obr_dict.exists(dev_num) Then
                    obr_dict(dev_num) = obr_dict(dev_num) + 1
                Else
                    obr_dict(dev_num) = 1
                End If
            End If
        Next j
    End With
    
    'Выгрузим на отдельный лист для рисования графика
    With Worksheets("Analysis")
        .Range("A1").Resize(rast_dog_dict.Count).Value = Application.Transpose(rast_dog_dict.keys)
        
        For i = 1 To rast_dog_dict.Count
             dev_num = CStr(.Cells(i, 1).Value)
             If obr_dict.exists(dev_num) Then
                    .Cells(i, 2).Value = obr_dict(dev_num)
            Else
'                .Rows(i).Delete
'                i = i - 1
'                deleted_rows = deleted_rows + 1
'                If i > rast_dog_dict.Count - deleted_rows Then Exit For
            End If

       Next i
       'Посчитаем среднее время пользования с заявками и без
       Dim sum_without As Double, sumwith As Double, n As Long
       n = 0
       sum_without = 0
       sumwith = 0
       For i = 2 To rast_dog_dict.Count
            .Cells(i, 5) = CDate(.Cells(i, 4)) - CDate(.Cells(i, 3))
            
            If .Cells(i, 2) <> "" Then
                n = n + 1
                sumwith = sumwith + .Cells(i, 5).Value
            Else
                sum_without = sum_without + .Cells(i, 5).Value
            End If
       Next i
       .Cells(1, 7) = sum_without / (rast_dog_dict.Count - n)
       .Cells(2, 7) = sumwith / (n)
       
       ' C реднее время пользования с заявками и без по видам услуг
       Dim d_n_with, d_n_without, d_sums_with, d_sums_without
        Set d_n_with = CreateObject("Scripting.Dictionary")
        Set d_n_without = CreateObject("Scripting.Dictionary")
        Set d_sums_with = CreateObject("Scripting.Dictionary")
        Set d_sums_without = CreateObject("Scripting.Dictionary")
        With Worksheets("Расторгнутые договора")
        For i = 2 To 11719
            d_n_with(.Cells(i, 25).Value) = 0
            d_n_without(.Cells(i, 25).Value) = 0
            d_sums_with(.Cells(i, 25).Value) = 0
            d_sums_without(.Cells(i, 25).Value) = 0
        Next i
        End With
        Dim key
        With Worksheets("Analysis")
        
'        n = 0
'       sum_without = 0
'       sumwith = 0
       
        For i = 2 To rast_dog_dict.Count
            .Cells(i, 5) = CDate(.Cells(i, 4)) - CDate(.Cells(i, 3))
            key = Worksheets("Расторгнутые договора").Cells(i, 25).Value
            If .Cells(i, 2) <> "" Then
            '''''
'                n = n + 1
'                sumwith = sumwith + .Cells(i, 5).Value
            '''''
                d_n_with(key) = d_n_with(key) + 1
                d_sums_with(key) = d_sums_with(key) + .Cells(i, 5).Value
            Else
            '''''
'                sum_without = sum_without + .Cells(i, 5).Value
            '''''
                d_n_without(key) = d_n_without(key) + 1
                d_sums_without(key) = d_sums_without(key) + .Cells(i, 5).Value
            End If
       Next i
''''''       .Cells(1, 7) = sum_without / (rast_dog_dict.Count - n)
''''''       .Cells(2, 7) = sumwith / (n)
       i = 8
       For Each key In d_n_with
            .Cells(i, 6).Value = key
            If d_n_with(key) <> 0 Then
                .Cells(i, 7).Value = CStr(d_sums_with(key) / d_n_with(key))
            Else
                .Cells(i, 7).Value = "0"
            End If
            i = i + 1
       Next
       i = 8
       For Each key In d_n_without
        If d_n_without(key) <> 0 Then
            .Cells(i, 8).Value = CStr(d_sums_without(key) / d_n_without(key))
        Else
            .Cells(i, 8).Value = "0"
            
        End If
        i = i + 1
       Next
       
       'ПОсмотрим теперь на количество заявок по тарифным планам и среднее время пользования ими
       Dim d_tarify_av_time
       Dim d_tarify_n
       Dim d_tarify_n_zayav
       Dim d_tarify_D
       Set d_tarify_av_time = CreateObject("Scripting.Dictionary")
       Set d_tarify_n = CreateObject("Scripting.Dictionary")
       Set d_tarify_n_zayav = CreateObject("Scripting.Dictionary")
       Set d_tarify_D = CreateObject("Scripting.Dictionary")
       
       
       For i = 2 To 11719
            d_tarify_av_time(Worksheets("Расторгнутые договора").Cells(i, 13).Value) = 0
            d_tarify_n(Worksheets("Расторгнутые договора").Cells(i, 13).Value) = 0
            d_tarify_D(Worksheets("Расторгнутые договора").Cells(i, 13).Value) = 0
        Next i
        
        For i = 2 To 11719
            key = Worksheets("Расторгнутые договора").Cells(i, 13).Value
            If .Cells(i, 2) <> "" Then d_tarify_n_zayav(key) = d_tarify_n_zayav(key) + 1
            d_tarify_n(key) = d_tarify_n(key) + 1
            d_tarify_av_time(key) = d_tarify_av_time(key) + .Cells(i, 5).Value
            
            
       Next i
       i = 21
       Dim Zscale
       Zscale = 10
       
       For Each key In d_tarify_av_time
            .Cells(i, 6).Value = key
            If d_tarify_n(key) <> 0 Then
                 d_tarify_av_time(key) = d_tarify_av_time(key) / d_tarify_n(key)
                 
                 .Cells(i, 7).Value = d_tarify_av_time(key)
            Else
                .Cells(i, 7).Value = "0"
            End If
            .Cells(i, 8).Value = CLng(d_tarify_n_zayav(key)) * Zscale
            i = i + 1
       Next
       
       'Посчитаем дисперсию, среднее линейное отклонение
       For i = 2 To 11719
            key = Worksheets("Расторгнутые договора").Cells(i, 13).Value
            If d_tarify_n(key) <> 0 Then
                d_tarify_D(key) = d_tarify_D(key) + (.Cells(i, 5).Value - d_tarify_av_time(key)) * (.Cells(i, 5).Value - d_tarify_av_time(key))
            End If
       Next i
       i = 21
       For Each key In d_tarify_av_time
            
            If d_tarify_n(key) <> 0 Then
                d_tarify_D(key) = d_tarify_D(key) / d_tarify_n(key) 'Дисперсия по тарифу
                .Cells(i, 9).Value = d_tarify_D(key)
            End If
            i = i + 1
        Next key
        
        
        
        
    End With
    End With
    ended

End Sub
Sub count_Types()
    Dim d
    Set d = CreateObject("Scripting.Dictionary")
    With Worksheets("Analysis")
    For i = 2 To 11719
        d(.Cells(i, 8).Value) = Empty
    Next i
    End With
    
End Sub
Sub prepare()
     Application.ScreenUpdating = False
     Application.Calculation = xlCalculationManual
      Application.EnableEvents = False
End Sub
Sub ended()
Application.ScreenUpdating = True
     Application.Calculation = xlCalculationAutomatic
      Application.EnableEvents = True
End Sub
