Option Compare Database
Option Explicit

Private Sub sub_Calculate()
    Dim var_Стоимость_листа_белой_бумаги_65г As Double
    Dim var_Стоимость_листа_белой_бумаги_80г As Double
    Dim var_Стоимость_вывода_1стр_оригинал_макета As Double
    Dim var_Стоимость_вывода_всего_оригинал_макета_на_принтере As Double
    Dim var_Стоимость_листа_цветной_бумаги As Double
    Dim var_Листов_белой_бумаги_на_тираж As Double
    Dim var_Листов_цветной_бумаги_на_тираж As Double
    Dim var_Количество_страниц As Double
    Dim var_Количество_тетрадей As Double
    Dim var_Количество_брошюр As Double
    Dim var_Пачек_бумаги As Double
    Dim var_Затраты_на_белую_бумагу As Double
    Dim var_Затраты_на_цветную_бумагу As Double
    Dim var_Затраты_на_бумагу_ИТОГО As Double
    Dim var_Затраты_на_бумагу_ИТОГО_с_учетом_брака_и_транспортных_расходов As Double
    Dim var_Стоимость_кадра_МП As Double
    Dim var_Затраты_на_МП As Double
    Dim var_Затраты_на_краску As Double
    Dim var_Стоимость_1_скобы As Double
    Dim var_Скоб_на_тираж As Double
    Dim var_Затраты_на_скобы_на_тираж As Double
    Dim var_Затраты_на_расходные_материалы_ИТОГО As Double
    Dim var_Затраты_на_тираж As Double
    Dim var_Затраты_на_книгу As Double
    Dim var_Цена_работ_на_тираж As Double
    Dim var_Сумма_налога As Double
    Dim var_Сумма As Double
    Dim var_Итоговые_затраты
    Dim var_Развитие_материально_технической_базы_сумма As Double
    Dim var_Цена_на_тираж As Double
    Dim var_Цена_на_тираж_НДС As Double
    Dim var_Цена_книги As Double
    Dim var_Цена_книги_с_учетом_рассылки As Double
    Dim A As Double
    Dim B As Double
    Dim C As Double
    Dim D As Double
    Dim E As Double
    Dim F As Double
    Dim G As Double
    Dim H As Double
    Dim I As Double
    Dim J As Double
    Dim K As Double
    Dim L As Double
    Dim M As Double
    Dim N As Double
    Dim O As Double
    Dim P As Double
    Dim count As Double
    Dim ДПЦ As Double, ПЦ As Double, ППЦ As Double
'------------------------------------------------------------------------------------------------------------------------------------------------
    
    If Число_страниц > 3 Then
        While ((Число_страниц / 4) - Int(Число_страниц / 4))
            Число_страниц = Число_страниц + 1
        Wend
    Else
        Число_страниц = 4
    End If
    ЧислоЛистовМакета = Число_страниц / 4
    var_Количество_тетрадей = ЧислоЛистовМакета
    var_Количество_страниц = Число_страниц
    
    var_Количество_тетрадей = var_Количество_тетрадей
    var_Стоимость_листа_белой_бумаги_65г = ЦенаБумаги / Число_листов_в_пачке
    var_Стоимость_листа_белой_бумаги_80г = Цена_пачки_белой_бумаги_80г / Листов_в_пачке_80г
    var_Стоимость_вывода_1стр_оригинал_макета = Вывод_1_листа_макета_на_HP
    var_Стоимость_вывода_всего_оригинал_макета_на_принтере = var_Стоимость_вывода_1стр_оригинал_макета * var_Количество_страниц
    var_Стоимость_листа_цветной_бумаги = ЦенаОбложки / ЦветнойЛист
    var_Листов_белой_бумаги_на_тираж = var_Количество_тетрадей * Тираж
    var_Листов_цветной_бумаги_на_тираж = Тираж
    
    If var_Количество_тетрадей <= 32 Then
        var_Количество_брошюр = 1
        Клей = 0
    Else
        var_Количество_брошюр = Int(var_Количество_тетрадей / 8 + 0.9999999)
    End If
    ЧислоБрошюр = var_Количество_брошюр
    var_Пачек_бумаги = Int(var_Количество_тетрадей * Тираж / Число_листов_в_пачке + 0.9999999)
    пачек_белой_бумаги = var_Пачек_бумаги
    var_Затраты_на_белую_бумагу = var_Пачек_бумаги * ЦенаБумаги + var_Стоимость_вывода_всего_оригинал_макета_на_принтере + var_Стоимость_листа_белой_бумаги_80г * Число_страниц
    var_Затраты_на_цветную_бумагу = var_Стоимость_листа_цветной_бумаги * Тираж
    var_Затраты_на_бумагу_ИТОГО = var_Затраты_на_белую_бумагу + var_Затраты_на_цветную_бумагу
    var_Затраты_на_бумагу_ИТОГО_с_учетом_брака_и_транспортных_расходов = var_Затраты_на_бумагу_ИТОГО * (1 + (Брак + ТранспортныеРасходы) / 100)
    var_Стоимость_кадра_МП = ЦенаРулонаПленки / ЧислоМастерПленки
    var_Затраты_на_МП = var_Количество_тетрадей * var_Стоимость_кадра_МП * 2
    var_Затраты_на_краску = ЦенаТубы_Краски / КраскоЛисты * 2 * (var_Листов_белой_бумаги_на_тираж + var_Листов_цветной_бумаги_на_тираж)
    var_Скоб_на_тираж = var_Количество_брошюр * Тираж * 2
    var_Стоимость_1_скобы = Стоимость_пачки_скоб / Скоб_в_упаковке
    var_Затраты_на_скобы_на_тираж = var_Стоимость_1_скобы * Тираж * Скоб_На_Брошюру
    var_Затраты_на_расходные_материалы_ИТОГО = var_Затраты_на_МП + var_Затраты_на_краску + var_Затраты_на_скобы_на_тираж + Клей
    var_Затраты_на_тираж = var_Затраты_на_бумагу_ИТОГО_с_учетом_брака_и_транспортных_расходов + var_Затраты_на_расходные_материалы_ИТОГО
    var_Затраты_на_книгу = var_Затраты_на_тираж / Тираж
    
    A = Рецензирование * var_Количество_тетрадей
    B = Корректура * var_Количество_тетрадей
    'C = (A + B) * Набор / 100
    'D = (A + B + C) * Правка_на_ПК__в___ / 100
    'E = (A + B + C + D) * КоэффицентСложности / 100
    'F = (A + B + C + D + E) * Сканирование / 100
    'G = (A + B + C + D + E + F) * Тиражирование / 100
    'H = (A + B + C + D + E + F + G) * Макетирование / 100
    'I = (A + B + C + D + E + F + G + H) * Сортировка / 100
    'J = (A + B + C + D + E + F + G + H + I) * ЧислоСторонДляРезки * Резка / 100
    'K = (A + B + C + D + E + F + G + H + I + J) * ПереплетБрошюровка / 100
    'L = (A + B + C + D + E + F + G + H + I + J + K) * Фальцовка / 100
    'M = (A + B + C + D + E + F + G + H + I + J + K + L) * ПереплетТермоклеевой / 100
    'Допечатный цикл
    C = (A + B) * Набор / 100
    D = (A + B + C) * Правка_на_ПК__в___ / 100
    E = (A + B + C + D) * КоэффицентСложности / 100
    F = (A + B + C + D + E) * Сканирование / 100
    ДПЦ = A + B + C + D + E + F
    'Печатный цикл
    G = (ДПЦ) * Тиражирование / 100
    H = (ДПЦ) * Макетирование / 100
    ПЦ = G + H
    'Послепечатный цикл
    I = (ДПЦ) * Сортировка / 100
    J = (ДПЦ) * ЧислоСторонДляРезки * Резка / 100
    K = (ДПЦ) * ПереплетБрошюровка / 100
    L = (ДПЦ) * Фальцовка / 100
    M = (ДПЦ) * ПереплетТермоклеевой / 100
    ППЦ = I + J + K + L + M
    'C = (A + B) * Набор / 100
    'D = (A + B) * Правка_на_ПК__в___ / 100
    'E = (A + B) * КоэффицентСложности / 100
    'F = (A + B) * Сканирование / 100
    'G = (A + B) * Тиражирование / 100
    'H = (A + B) * Макетирование / 100
    'I = (A + B) * Сортировка / 100
    'J = (A + B) * ЧислоСторонДляРезки * Резка / 100
    'K = (A + B) * ПереплетБрошюровка / 100
    'L = (A + B) * Фальцовка / 100
    'M = (A + B) * ПереплетТермоклеевой / 100

    'var_Цена_работ_на_тираж = A + B + C + D + E + F + G + H + I + J + K + L + M
    var_Цена_работ_на_тираж = ДПЦ + ПЦ + ППЦ
    
    var_Сумма_налога = var_Цена_работ_на_тираж * НалогНаЗП / 100
    var_Сумма = var_Цена_работ_на_тираж + var_Сумма_налога
    var_Итоговые_затраты = var_Затраты_на_тираж + var_Сумма + ISBN
    var_Развитие_материально_технической_базы_сумма = var_Итоговые_затраты * ФондРазвитияМатБазы_процент / 100
    var_Цена_на_тираж = var_Развитие_материально_технической_базы_сумма + var_Итоговые_затраты
    var_Цена_на_тираж_НДС = var_Цена_на_тираж + (var_Цена_на_тираж * НДС / 100)
    If Тираж = БесплатнаяРассылка Then
        var_Цена_книги = var_Цена_на_тираж_НДС / Тираж
        var_Цена_книги_с_учетом_рассылки = var_Цена_книги
    Else
        var_Цена_книги = var_Цена_на_тираж_НДС / Тираж
        var_Цена_книги_с_учетом_рассылки = var_Цена_на_тираж_НДС / (Тираж - БесплатнаяРассылка)
    End If
'************************************************************************************************************************************************
    Расходы_на_тираж = var_Затраты_на_тираж
    Расходы_на_книгу = var_Затраты_на_книгу
    Цена_работы_на_тираж = var_Цена_работ_на_тираж
    СуммаНалога = var_Сумма_налога
    Сумма = var_Сумма
    Итоговые_Затраты = var_Итоговые_затраты
    ФондРазвитияМатБазы_сумма = var_Развитие_материально_технической_базы_сумма
    Цена_на_тираж_НДС = var_Цена_на_тираж_НДС
    ЦЕНА_ЗА_ТИРАЖ_без_НДС = var_Цена_на_тираж
    Цена_книги = var_Цена_книги
    ЦЕНА_КНИГИ_С_УЧЕТОМ_БР = var_Цена_книги_с_учетом_рассылки
    Число_страниц = var_Количество_страниц
'************************************************************************************************************************************************
    ТестовыйВывод.Value = "Стоимость_листа_белой_бумаги_65г = " & var_Стоимость_листа_белой_бумаги_65г & Chr(13) & Chr(10) & _
                            "Стоимость_листа_белой_бумаги_80г = " & var_Стоимость_листа_белой_бумаги_80г & Chr(13) & Chr(10) & _
                            "Стоимость_вывода_1стр_оригинал_макета = " & Вывод_1_листа_макета_на_HP & Chr(13) & Chr(10) & _
                            "Стоимость_вывода_всего_оригинал_макета_на_принтере = " & var_Стоимость_вывода_всего_оригинал_макета_на_принтере & Chr(13) & Chr(10) & _
                            "var_Стоимость_листа_цветной_бумаги = " & var_Стоимость_листа_цветной_бумаги & Chr(13) & Chr(10) & _
                            "var_Листов_белой_бумаги_на_тираж = " & var_Листов_белой_бумаги_на_тираж & Chr(13) & Chr(10) & _
                            "var_Листов_цветной_бумаги_на_тираж = " & var_Листов_цветной_бумаги_на_тираж & Chr(13) & Chr(10) & _
                            "var_Количество_брошюр = " & var_Количество_брошюр & Chr(13) & Chr(10) & _
                            "var_Пачек_бумаги = " & var_Пачек_бумаги & Chr(13) & Chr(10) & _
                            "var_Затраты_на_белую_бумагу = " & var_Затраты_на_белую_бумагу & Chr(13) & Chr(10) & _
                            "var_Затраты_на_цветную_бумагу = " & var_Затраты_на_цветную_бумагу & Chr(13) & Chr(10) & _
                            "var_Затраты_на_бумагу_ИТОГО = " & var_Затраты_на_бумагу_ИТОГО & Chr(13) & Chr(10) & _
                            "var_Затраты_на_бумагу_ИТОГО_с_учетом_брака_и_транспортных_расходов = " & var_Затраты_на_бумагу_ИТОГО_с_учетом_брака_и_транспортных_расходов & CrLf & _
                            "var_Стоимость_кадра_МП = " & var_Стоимость_кадра_МП & CrLf & _
                            "var_Затраты_на_МП = " & var_Затраты_на_МП & CrLf & _
                            "var_Затраты_на_краску = " & var_Затраты_на_краску & CrLf & _
                            "var_Скоб_на_тираж = " & var_Скоб_на_тираж & CrLf & _
                            "var_Затраты_на_скобы_на_тираж = " & var_Затраты_на_скобы_на_тираж & CrLf & _
                            "var_Затраты_на_расходные_материалы_ИТОГО = " & var_Затраты_на_расходные_материалы_ИТОГО & CrLf & _
                            "var_Затраты_на_тираж = " & var_Затраты_на_тираж & CrLf & _
                            "var_Затраты_на_книгу = " & var_Затраты_на_книгу & CrLf & _
                            "var_Цена_работ_на_тираж = " & var_Цена_работ_на_тираж & CrLf
End Sub

Private Sub btn_ShowReport_Click()
    Me.Refresh
    DoCmd.OpenReport "smeta_report", acViewPreview
End Sub

Private Sub Form_Open(Cancel As Integer)
    DoCmd.Maximize
End Sub

Private Sub ISBN_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub БесплатнаяРассылка_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub Брак_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub Вывод_1_листа_макета_на_HP_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub Клей_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub Корректура_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub КоэффицентСложности_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub КраскоЛисты_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub Листов_в_пачке_80г_Exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub Макетирование_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub Набор_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub НалогНаЗП_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub НДС_Exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub ПереплетБрошюровка_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub ПереплетТермоклеевой_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub Правка_на_ПК__в____exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub Резка_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub Рецензирование_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub Сканирование_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub Скрепка_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub Скоб_в_упаковке_Exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub Скоб_На_Брошюру_Exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub Сортировка_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub Стоимость_пачки_скоб_Exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub Тираж_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub Тиражирование_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub ТорговыеАгенты_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub ТранспортныеРасходы_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub Фальцовка_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub ФондРазвитияМатБазы_exit(Cancel As Integer)
    sub_Calculate
End Sub


Private Sub ФондРазвитияМатБазы_процент_Exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub ЦветнойЛист_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub Цена_пачки_белой_бумаги_80г_Exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub ЦенаБумаги_Exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub ЦенаОбложки_Exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub ЦенаРулонаПленки_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub ЦенаТубы_Краски_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub Число_листов_в_пачке_Exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub Число_страниц_Exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub ЧислоЛистовМакета_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub ЧислоМастерПленки_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub ЧислоСторон_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Sub ЧислоСторонДляРезки_exit(Cancel As Integer)
    sub_Calculate
End Sub

Private Function CrLf()
    CrLf = (Chr(13) & Chr(10))
End Function
