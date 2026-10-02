Attribute VB_Name = "Module1"
Option Explicit

'=============================================================================
' КОНФИГУРАТОР ШКАФОВ
'
' 1. Создать шкафы P410 и сразу зарезервировать верхнюю зону.
' 2. Разместить в каждом шкафу P410 до 4 KDI и до 4 KDO
'    вместе с барьерами. Не поместившееся реле остаётся свободным.
' 3. Пересчитать остаток реле и назначить его модулям DIO.
'    Для DIO64: максимум 8 реле в любой пропорции 4 KDI + 4 KDO.
'    Для DIO32: максимум 4 реле в любой пропорции 4 KDI + 4 KDO.
'    Все группы, кроме последней, должны быть полными.
'    Последняя группа может быть неполной без предупреждения.
' 4. Проверить, помещается ли КАЖДАЯ ЦЕЛАЯ группа DIO в существующий
'    шкаф. Если нет — создать новый шкаф.
'
' Число реек КАЖДОГО шкафа задаётся ячейкой D18.
'=============================================================================

Private Const TYP_NONE As Long = 0
Private Const TYP_DI As Long = 1
Private Const TYP_DO As Long = 2
Private Const TYP_DIO32 As Long = 32
Private Const TYP_DIO64 As Long = 64

Private Const FAM_TC As Long = 1
Private Const FAM_AI As Long = 2
Private Const FAM_DIO As Long = 3

Private Const ROWS_DIO As String = "33"
Private Const ROWS_KDI As String = "34,35"
Private Const ROWS_KDO As String = "36,37"

Private Const FAM_TC_MOD As String = "28,29"
Private Const FAM_TC_BAR As String = "49,50"
Private Const FAM_AI_MOD As String = "30,31,32"
Private Const FAM_AI_BAR As String = "51,52"
Private Const FAM_DIO_REL As String = "34,35,36,37"
Private Const FAM_DIO_BAR As String = "53,54,55"

Private Const ROWS_BAR_KDI As String = "53,54,55"
Private Const ROWS_BAR_KDO As String = "53"

Private Const ROWS_SEPARATE As String = _
    "34,35,36,37,47,48,49,50,51,52,53,54,55"

Private Const RELAY_CH_DEFAULT As Long = 8
Private Const BAR_CH_DEFAULT As Long = 1

Private Const ADDR_EXI_DI As String = "V10"
Private Const ADDR_NAMUR_DI As String = "X10"
Private Const ADDR_EXI_DO As String = "AA10"
Private Const ADDR_FREQ_DI As String = "AC10"

' Service: E — наименование, F — каналы, H — ширина.
Private Const SVC_COL_NAME As Long = 5
Private Const SVC_COL_CH As Long = 6
Private Const SVC_COL_WIDTH As Long = 8
' Service: D — категория ("module AI" / "module AO").
Private Const SVC_COL_CAT As Long = 4
Private Const SVC_CAT_AI As String = "MODULE AI"
Private Const SVC_CAT_AO As String = "MODULE AO"

' Ёмкость аналогового модуля, если его нет в Service.
Private Const ANALOG_CH_DEFAULT As Long = 8

' Строки листа "1" с барьерами для AO (через запятую).
' Пусто — тип определяется по имени (AO / ВЫХ), иначе считается AI.


Private Const DIR_AI As Long = 1
Private Const DIR_AO As Long = 2
Private Const DIR_AI_AO As Long = 3

' Единственная строка общих барьеров AI/AO на листе "1".
Private Const ROWS_BAR_AI_AO As String = "52"

Private Const RELAYS_PER_DIO64 As Long = 8
Private Const RELAYS_PER_DIO32 As Long = 4

Private Const DI_PER_CTRL As Long = 4
Private Const DO_PER_CTRL As Long = 4

Private Const SWITCH_RESERVE_ALL_RAILS As Boolean = True
Private Const SWITCH_RESERVE_BACK_RAILS As Boolean = False
Private Const SWITCH_RAILS As Long = 1

Private Const COL_NAME As String = "C"
Private Const COL_QTY As String = "D"
Private Const SH_MOD_HDR_ROW As Long = 9

Private Const CTRL_ZONE_MM As Double = 224
Private Const CTRL_BLOCK_PX As Double = 52
Private Const CTRL_OFFSET_PX As Double = 56

Private Const PLAN_EPS As Double = 0.0001
Private Const PLAN_NODE_LIMIT As Long = 300000
Private Const PLAN_SEARCH_BLOCK_LIMIT As Long = 120

' Рейки и шкафы.
Private rCab() As Long
Private rKind() As String
Private rSide() As String
Private rSideIdx() As Long
Private rCap() As Double
Private rUsed() As Double
Private rPhys() As Double
Private rCtrl() As Boolean
Private rSwitch() As Boolean
Private cabCtrl() As Long

Private railCount As Long
Private cabCount As Long

Private gRailsTotal As Long
Private gRailLen As Double
Private gTwoSided As Boolean
Private gFreeFrac As Double
Private gCtrlRemain As Long
Private gCtrlPerBlock As Long
Private gSeparateMode As Boolean
Private gLastDio As Long
Private gWarn As String

' Экземпляры модулей.
Private mName() As String
Private mRow() As Long
Private mLen() As Double
Private mColor() As Long
Private mSep() As Boolean
Private mRail() As Long
Private mPos() As Double
Private mPlaced() As Boolean
Private mGrpType() As Long
Private mFam() As Long
Private mIsBar() As Boolean
Private mCh() As Long
Private mBarType() As Long
Private mBundle() As Long
Private mAnalogOwner() As Long
Private mCapAI() As Long      ' каналы AI модуля
Private mCapAO() As Long      ' каналы AO модуля
Private mAnDir() As Long      ' DIR_AI / DIR_AO для аналогового барьера
Private gTotalMods As Long

' Связки «реле + его барьеры».
Private bHead() As Long
Private bKind() As String
Private bPlaced() As Boolean
Private bCab() As Long
Private bDioOwner() As Long
Private bCtrlCab() As Long
Private gBundleCount As Long

Private Function ЭтоНоль(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function

    If IsNumeric(v) Then
        ЭтоНоль = (CDbl(v) = 0)
    Else
        ЭтоНоль = (Trim$(CStr(v)) = "0")
    End If
End Function

Private Function ЭтоНольИлиПрочерк(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function

    If Trim$(CStr(v)) = "-" Then
        ЭтоНольИлиПрочерк = True
    Else
        ЭтоНольИлиПрочерк = ЭтоНоль(v)
    End If
End Function

'=============================================================================
' ГЛАВНАЯ ПРОЦЕДУРА
'=============================================================================

Public Sub СоздатьКонфигуратор()
    Dim ws As Worksheet
    Dim wsS As Worksheet
    Dim wl As Worksheet

    Dim defRows As Variant
    Dim defLensFallback As Variant
    Dim defLens() As Double

    Dim v As Variant
    Dim fv As Double
    Dim ctrlCount As Long
    Dim total As Long
    Dim allocN As Long

    Dim capAI As Long
    Dim capAO As Long
    Dim anDir As Long

    Dim t As Long
    Dim r As Long
    Dim q As Long
    Dim k As Long
    Dim idx As Long
    Dim c As Long
    Dim b As Long
    Dim fam As Long

    Dim rawNm As String
    Dim nm As String
    Dim col As Long
    Dim typ As Long
    Dim isBar As Boolean
    Dim ch As Long
    Dim barT As Long
    Dim isSep As Boolean

    Dim legName() As String
    Dim legLen() As Double
    Dim legQty() As Long
    Dim legColor() As Long
    Dim legN As Long

    Dim nDI As Long
    Dim nDO As Long
    Dim nDIO As Long
    Dim relayChDI As Long
    Dim relayChDO As Long
    Dim relayCapDI As Double
    Dim relayCapDO As Double
    Dim sigDI As Double
    Dim sigDO As Double
    Dim needKDI As Long
    Dim needKDO As Long

    Dim maxLen As Double
    Dim capFull As Double
    Dim capCtrl As Double

    Dim gotCDI As Long
    Dim gotCDO As Long

    Dim bl() As Long
    Dim nb As Long
    Dim relayLimit As Long

    Dim actualDI As Long
    Dim actualDO As Long

    Dim modCnt As Long
    Dim barCnt As Long
    Dim iMod As Long
    Dim takeBar As Long
    Dim items() As Long
    Dim n As Long

    Dim usedCnt As Long
    Dim msg As String
    Dim errText As String

    gWarn = vbNullString
    gLastDio = 0

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("1")
    Set wsS = ThisWorkbook.Worksheets("Service")
    Set wl = ThisWorkbook.Worksheets("1 расположение")
    On Error GoTo ErrH

    If ws Is Nothing Then
        MsgBox "ОШИБКА: не найден лист ""1"".", vbCritical
        Exit Sub
    End If

    If wsS Is Nothing Then
        MsgBox "ОШИБКА: не найден лист ""Service"".", vbCritical
        Exit Sub
    End If

    '---------------------------------------------------------------------
    ' Параметры шкафа. D18 — ЧИСЛО РЕЕК В КАЖДОМ ШКАФУ.
    '---------------------------------------------------------------------
    gRailsTotal = ИзвлечьЧисло(ws.Range("D18").Value2)

    v = ws.Range("B18").Value2

    If IsError(v) Then
        Err.Raise vbObjectError + 10, , _
            "Ошибка в B18: укажите длину рейки."
    End If

    If IsNumeric(v) And Not IsEmpty(v) Then
        gRailLen = CDbl(v)
    Else
        gRailLen = ИзвлечьЧисло(v)
    End If

    If gRailLen <= 0 Then
        Err.Raise vbObjectError + 11, , _
            "В B18 должна быть положительная длина рейки."
    End If

    Select Case gRailsTotal
        Case 3, 4
            gTwoSided = False

        Case 6, 8
            gTwoSided = True

        Case Else
            Err.Raise vbObjectError + 12, , _
                "Недопустимое значение D18. Нужно 3, 4, 6 или 8 реек."
    End Select

    fv = 0
    v = ws.Range("D19").Value2

    If Not IsError(v) Then
        If IsNumeric(v) Then fv = CDbl(v)
    End If

    If fv > 1 Then
        gFreeFrac = fv / 100#
    Else
        gFreeFrac = fv
    End If

    If gFreeFrac < 0 Then gFreeFrac = 0

    If gFreeFrac >= 0.9 Then
        gFreeFrac = 0.9
        gWarn = gWarn & vbCrLf & _
            "  • Резерв D19 ограничен значением 90%."
    End If

    gSeparateMode = Да(ws.Range("D13").Value2)

    ctrlCount = ЧитатьКол(ws, 26)

    gCtrlPerBlock = 1
    If Да(ws.Range("D12").Value2) Then gCtrlPerBlock = 2

    capFull = gRailLen * (1 - gFreeFrac)
    capCtrl = capFull - CTRL_ZONE_MM

    If ctrlCount > 0 And capCtrl < -PLAN_EPS Then
        Err.Raise vbObjectError + 13, , _
            "Полезная длина рейки меньше верхней зоны P410/" & _
            "коммутаторов (" & CTRL_ZONE_MM & " мм). " & _
            "Проверьте B18 и D19."
    End If

    '---------------------------------------------------------------------
    ' Чтение модулей.
    '---------------------------------------------------------------------
    defRows = Array(27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, _
                    47, 48, 49, 50, 51, 52, 53, 54, 55)

    defLensFallback = Array(35, 164, 164, 157, 157, 157, 157, _
                            110, 110, 100, 100, 23, 67, 13, 23, _
                            23, 13, 13, 13, 23)

    ReDim defLens(LBound(defRows) To UBound(defRows))

    total = 0

    For t = LBound(defRows) To UBound(defRows)
        r = CLng(defRows(t))
        q = ЧитатьКол(ws, r)
        total = total + q

        If q > 0 Then
            v = ws.Range(COL_NAME & r).Value2

            If IsError(v) Then
                Err.Raise vbObjectError + 14, , _
                    "Ошибка в наименовании модуля, строка " & r & "."
            End If

            defLens(t) = ШиринаИзService( _
                wsS, CStr(v), CDbl(defLensFallback(t)), r)
        Else
            defLens(t) = CDbl(defLensFallback(t))
        End If
    Next t

    If total = 0 And ctrlCount = 0 Then
        MsgBox "Нет модулей и контроллеров для размещения.", _
               vbExclamation
        Exit Sub
    End If

    allocN = total
    If allocN < 1 Then allocN = 1

    ReDim mName(1 To allocN)
    ReDim mRow(1 To allocN)
    ReDim mLen(1 To allocN)
    ReDim mColor(1 To allocN)
    ReDim mSep(1 To allocN)
    ReDim mRail(1 To allocN)
    ReDim mPos(1 To allocN)
    ReDim mPlaced(1 To allocN)
    ReDim mGrpType(1 To allocN)
    ReDim mFam(1 To allocN)
    ReDim mIsBar(1 To allocN)
    ReDim mCh(1 To allocN)
    ReDim mBarType(1 To allocN)
    ReDim mBundle(1 To allocN)
    ReDim mAnalogOwner(1 To allocN)

    ReDim mCapAI(1 To allocN)
    ReDim mCapAO(1 To allocN)
    ReDim mAnDir(1 To allocN)

    ReDim bHead(1 To allocN)
    ReDim bKind(1 To allocN)
    ReDim bPlaced(1 To allocN)
    ReDim bCab(1 To allocN)
    ReDim bDioOwner(1 To allocN)
    ReDim bCtrlCab(1 To allocN)

    gTotalMods = total
    gBundleCount = 0

    ReDim legName(1 To UBound(defRows) + 1)
    ReDim legLen(1 To UBound(defRows) + 1)
    ReDim legQty(1 To UBound(defRows) + 1)
    ReDim legColor(1 To UBound(defRows) + 1)

    idx = 0
    legN = 0

    For t = LBound(defRows) To UBound(defRows)
        r = CLng(defRows(t))
        q = ЧитатьКол(ws, r)

        If q > 0 Then
            rawNm = CStr(ws.Range(COL_NAME & r).Value2)
            nm = ОчиститьИмя(rawNm, r)
            col = Палитра(t + 1)

            isSep = gSeparateMode And ВСписке(r, ROWS_SEPARATE)
            typ = ОпределитьТип(r, nm)
            fam = Семейство(r)
            isBar = ЭтоБарьер(r)

            ch = 0
            barT = TYP_NONE

            If typ = TYP_DI Or typ = TYP_DO Then
                ch = КаналыИзService( _
                    wsS, rawNm, RELAY_CH_DEFAULT, r, "реле")

            ElseIf isBar And fam = FAM_DIO Then
                ch = КаналыИзService( _
                    wsS, rawNm, BAR_CH_DEFAULT, r, "барьер")

                barT = ТипБарьераDIO(r, nm)
            End If
            capAI = 0
            capAO = 0
            anDir = 0

            If fam = FAM_TC Or fam = FAM_AI Then
                If isBar Then
                    ch = КаналыИзService( _
                        wsS, rawNm, BAR_CH_DEFAULT, r, _
                        "аналоговый барьер")

                    anDir = НаправлениеАналогБарьера(r, nm)
                Else
                    ЁмкостьАналогМодуля wsS, rawNm, r, capAI, capAO
                End If
            End If

            legN = legN + 1
            legName(legN) = nm
            legLen(legN) = defLens(t)
            legQty(legN) = q
            legColor(legN) = col

            For k = 1 To q
                idx = idx + 1

                mName(idx) = nm
                mRow(idx) = r
                mLen(idx) = defLens(t)
                mColor(idx) = col
                mSep(idx) = isSep
                mGrpType(idx) = typ
                mFam(idx) = fam
                mIsBar(idx) = isBar
                mCh(idx) = ch
                mBarType(idx) = barT
                mCapAI(idx) = capAI
                mCapAO(idx) = capAO
                mAnDir(idx) = anDir
                
                If mLen(idx) > maxLen Then
                    maxLen = mLen(idx)
                End If
            Next k
        End If
    Next t

    If maxLen > capFull + PLAN_EPS Then
        Err.Raise vbObjectError + 15, , _
            "Модуль шириной " & Format(maxLen, "0") & _
            " мм не помещается на рейку. Полезная длина: " & _
            Format(capFull, "0") & " мм."
    End If

    If ctrlCount > 0 And maxLen > capCtrl + PLAN_EPS Then
        gWarn = gWarn & vbCrLf & _
            "  • Некоторые модули не помещаются на рейку с верхней " & _
            "зоной P410/коммутаторов; будут использованы другие рейки."
    End If

    '---------------------------------------------------------------------
    ' Диагностика сигналов.
    '---------------------------------------------------------------------
    For idx = 1 To total
        Select Case mGrpType(idx)
            Case TYP_DI
                nDI = nDI + 1
                relayCapDI = relayCapDI + mCh(idx)
                If relayChDI = 0 Then relayChDI = mCh(idx)

            Case TYP_DO
                nDO = nDO + 1
                relayCapDO = relayCapDO + mCh(idx)
                If relayChDO = 0 Then relayChDO = mCh(idx)

            Case TYP_DIO32, TYP_DIO64
                nDIO = nDIO + 1
        End Select
    Next idx

    If relayChDI <= 0 Then relayChDI = RELAY_CH_DEFAULT
    If relayChDO <= 0 Then relayChDO = RELAY_CH_DEFAULT

    sigDI = ЧитатьЧисло(ws, ADDR_EXI_DI) + _
            ЧитатьЧисло(ws, ADDR_NAMUR_DI) + _
            ЧитатьЧисло(ws, ADDR_FREQ_DI)

    sigDO = ЧитатьЧисло(ws, ADDR_EXI_DO)
    РаспределитьОбщиеБарьерыDIO sigDO

    If sigDI > 0 Then
        needKDI = -Int(-sigDI / relayChDI)

        If relayCapDI < sigDI Then
            gWarn = gWarn & vbCrLf & _
                "  • Сигналы DI: " & Format(sigDI, "0") & _
                "; требуется ориентировочно " & needKDI & _
                " KDI по " & relayChDI & _
                " каналов. Имеющихся каналов KDI: " & _
                Format(relayCapDI, "0") & "."
        End If
    End If

    If sigDO > 0 Then
        needKDO = -Int(-sigDO / relayChDO)

        If relayCapDO < sigDO Then
            gWarn = gWarn & vbCrLf & _
                "  • Сигналы DO: " & Format(sigDO, "0") & _
                "; требуется ориентировочно " & needKDO & _
                " KDO по " & relayChDO & _
                " каналов. Имеющихся каналов KDO: " & _
                Format(relayCapDO, "0") & "."
        End If
    End If

    ' D13=YES физически разносит реле/барьеры и основные модули.
    If gSeparateMode Then
        If nDIO > 0 Then
            Err.Raise vbObjectError + 16, , _
                "D13=YES отправляет реле и барьеры в отдельный шкаф, " & _
                "а DIO оставляет в основном. Для группы " & _
                "DIO + реле + барьеры установите D13=NO."
        End If

        If ctrlCount > 0 And (nDI > 0 Or nDO > 0) Then
            Err.Raise vbObjectError + 17, , _
                "При D13=YES реле размещаются отдельно от P410. " & _
                "Для подключения реле к контроллеру установите D13=NO."
        End If

        If ЕстьВСемействе(FAM_TC, False) And _
           ЕстьВСемействе(FAM_TC, True) Then

            Err.Raise vbObjectError + 18, , _
                "D13=YES разделяет модуль TC/RTD и его барьеры. " & _
                "Для совместного размещения установите D13=NO."
        End If

        If ЕстьВСемействе(FAM_AI, False) And _
           ЕстьВСемействе(FAM_AI, True) Then

            Err.Raise vbObjectError + 19, , _
                "D13=YES разделяет модуль AI/AO и его барьеры. " & _
                "Для совместного размещения установите D13=NO."
        End If
    End If

    ' Привязываем барьеры к реле с учётом числа каналов.
    СформироватьСвязки

    '---------------------------------------------------------------------
    ' ЭТАП 1. Создать шкафы P410.
    ' ДобавитьШкаф резервирует место сверху ДО размещения реле.
    '---------------------------------------------------------------------
    ReDim rCab(1 To 64)
    ReDim rKind(1 To 64)
    ReDim rSide(1 To 64)
    ReDim rSideIdx(1 To 64)
    ReDim rCap(1 To 64)
    ReDim rUsed(1 To 64)
    ReDim rPhys(1 To 64)
    ReDim rCtrl(1 To 64)
    ReDim rSwitch(1 To 64)
    ReDim cabCtrl(1 To 64)

    railCount = 0
    cabCount = 0
    gCtrlRemain = ctrlCount

    Do While gCtrlRemain > 0
        ДобавитьШкаф "M"
    Loop

    If cabCount = 0 Then
        For idx = 1 To total
            If Not mSep(idx) Then
                ДобавитьШкаф "M"
                Exit For
            End If
        Next idx
    End If

    '---------------------------------------------------------------------
    ' ЭТАП 1.1. Реле, подключаемые к P410.
    '
    ' Связка считается занятой контроллером ТОЛЬКО если реле и барьеры
    ' действительно разместились в данном шкафу.
    '---------------------------------------------------------------------
    If Not gSeparateMode Then
        For c = 1 To cabCount
            If cabCtrl(c) > 0 Then
                ЗакрепитьРелеЗаКонтроллером _
                    c, TYP_DI, DI_PER_CTRL, gotCDI

                ЗакрепитьРелеЗаКонтроллером _
                    c, TYP_DO, DO_PER_CTRL, gotCDO
            End If
        Next c
    End If

    '---------------------------------------------------------------------
    ' ЭТАП 2. Пересчитать СВОБОДНЫЕ реле после P410 и назначить их DIO.
    ' Последнему DIO разрешено получить меньше нормы каждого типа.
    '---------------------------------------------------------------------
    ReserveDioBundles

    ReDim bl(1 To 16)

    For idx = 1 To total
        If mGrpType(idx) = TYP_DIO32 Or _
           mGrpType(idx) = TYP_DIO64 Then

            If mPlaced(idx) Then
                Err.Raise vbObjectError + 20, , _
                    "DIO уже размещён до формирования группы."
            End If

If mGrpType(idx) = TYP_DIO64 Then
    relayLimit = RELAYS_PER_DIO64
Else
    relayLimit = RELAYS_PER_DIO32
End If

            nb = 0
            actualDI = 0
            actualDO = 0

            For b = 1 To gBundleCount
                If bDioOwner(b) = idx Then
                    If bPlaced(b) Then
                        Err.Raise vbObjectError + 21, , _
                            "Реле для " & mName(idx) & _
                            " уже размещено отдельно от DIO."
                    End If

                    nb = nb + 1

                    If nb > UBound(bl) Then
                        ReDim Preserve bl(1 To nb + 16)
                    End If

                    bl(nb) = b

                    Select Case ТипСвязки(b)
                        Case TYP_DI
                            actualDI = actualDI + 1

                        Case TYP_DO
                            actualDO = actualDO + 1

                        Case Else
                            Err.Raise vbObjectError + 22, , _
                                "Неизвестный тип реле в группе DIO."
                    End Select
                End If
            Next b

If actualDI + actualDO > relayLimit Then
    Err.Raise vbObjectError + 23, , _
        "Для " & mName(idx) & _
        " назначено больше допустимого числа реле: " & _
        (actualDI + actualDO) & " из " & relayLimit & "."
End If

If idx <> gLastDio Then
    If actualDI + actualDO <> relayLimit Then
        Err.Raise vbObjectError + 24, , _
            "Непоследняя группа " & mName(idx) & _
            " должна содержать " & relayLimit & _
            " реле KDI/KDO суммарно."
    End If
End If


            ' Для ПОСЛЕДНЕГО DIO меньшие actualDI/actualDO допустимы.
            ' Не добавляем по этому поводу ни предупреждений, ни справок.
            РазместитьГруппуDIO idx, bl, nb
        End If
    Next idx

    '---------------------------------------------------------------------
    ' ЭТАП 2.2. Аналоговые модули и их барьеры.
    '
    ' 49–50: TC/RTD -> температурные входы.
    ' 51:    RTD>AI -> только входы AI.
    ' 52:    общий пул AI/AO -> оставшиеся AI и AO.
    '
    ' Общие барьеры могут назначаться также выходам AO
    ' температурных модулей.
    '---------------------------------------------------------------------
    If Not gSeparateMode Then

        ' Сначала барьеры с ограниченной совместимостью.
        РаспределитьАналоговыеБарьеры FAM_TC, DIR_AI
        РаспределитьАналоговыеБарьеры FAM_AI, DIR_AI

        ' Затем единственный общий пул AI/AO.
        РаспределитьАналоговыеБарьеры FAM_AI, DIR_AI_AO

        ReDim items(1 To 32)

        For idx = 1 To total
            If Not mPlaced(idx) And Not mSep(idx) And _
               Not mIsBar(idx) Then

                If mFam(idx) = FAM_TC Or _
                   mFam(idx) = FAM_AI Then

                    n = 0
                    ДобавитьВСписок idx, items, n

                    ' Собираем ВСЕ барьеры владельца:
                    ' семейство барьера может отличаться от
                    ' семейства модуля, например AO у TC-модуля.
                    For k = 1 To total
                        If mAnalogOwner(k) = idx And _
                           Not mPlaced(k) Then

                            ДобавитьВСписок k, items, n
                        End If
                    Next k

                    If n > 1 Then
                        If n > 2 Then
                            СортироватьЧастьПоДлине items, 2, n
                        End If

                        РазместитьАналоговуюГруппу items, n, idx
                    End If

                    ' Модули без назначенных барьеров будут
                    ' размещены на этапе одиночных элементов.
                End If
            End If
        Next idx
    End If

    '---------------------------------------------------------------------
    ' ЭТАП 2.3. Связки реле, не назначенные ни P410, ни DIO.
    '---------------------------------------------------------------------
    For b = 1 To gBundleCount
        If Not bPlaced(b) Then
            If bDioOwner(b) <> 0 Then
                Err.Raise vbObjectError + 25, , _
                    "Связка реле DIO осталась неразмещённой. " & _
                    "Размещать её отдельно запрещено."
            End If

            РазместитьСвязку b
        End If
    Next b

    '---------------------------------------------------------------------
    ' ЭТАП 3. Остальные одиночные элементы.
    '---------------------------------------------------------------------
    For idx = 1 To total
        If Not mPlaced(idx) And mBundle(idx) = 0 And _
           Not mIsBar(idx) Then

            If mGrpType(idx) = TYP_DIO32 Or _
               mGrpType(idx) = TYP_DIO64 Then

                Err.Raise vbObjectError + 26, , _
                    "DIO не прошёл этап размещения группы."
            End If

            РазместитьОдиночный idx
        End If
    Next idx

    For idx = 1 To total
        If Not mPlaced(idx) And mBundle(idx) = 0 Then
            If mFam(idx) = FAM_DIO And mIsBar(idx) Then
                Err.Raise vbObjectError + 27, , _
                    "Барьер DIO остался без назначенного реле."
            End If

            РазместитьОдиночный idx
        End If
    Next idx

    ПроверитьРазмещение

    ' Рисуем только после успешного расчёта.
    If wl Is Nothing Then
        Set wl = ThisWorkbook.Worksheets.Add(after:=ws)
        wl.Name = "1 расположение"
    End If

    Application.ScreenUpdating = False

    НарисоватьРасположение wl, legName, legLen, legQty, _
                           legColor, legN, ctrlCount, usedCnt

    Application.ScreenUpdating = True

    msg = "Готово!" & vbCrLf & _
          "Шкафов задействовано: " & usedCnt

    If Len(gWarn) > 0 Then
        MsgBox msg & vbCrLf & vbCrLf & _
               "ВНИМАНИЕ:" & gWarn, vbExclamation
    Else
        MsgBox msg, vbInformation
    End If

    Exit Sub

ErrH:
    errText = Err.Description
    Application.ScreenUpdating = True

    MsgBox "Ошибка конфигуратора:" & vbCrLf & vbCrLf & _
           errText, vbCritical
End Sub

'=============================================================================
' ПРИВЯЗКА БАРЬЕРОВ К РЕЛЕ
'=============================================================================

Private Sub СформироватьСвязки()
    Dim i As Long

    gBundleCount = 0

    For i = 1 To gTotalMods
        mBundle(i) = 0
        bDioOwner(i) = 0
        bCtrlCab(i) = 0
    Next i

    ФормироватьСвязкиТипа TYP_DI, "KDI"
    ФормироватьСвязкиТипа TYP_DO, "KDO"
End Sub

Private Sub ФормироватьСвязкиТипа( _
    ByVal relType As Long, ByVal relName As String)

    Dim relB() As Long
    Dim leftCh() As Long
    Dim bars() As Long

    Dim nRel As Long
    Dim nBars As Long

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim barIdx As Long
    Dim tmp As Long
    Dim best As Long
    Dim afterCh As Long
    Dim bestAfter As Long

    ReDim relB(1 To gTotalMods + 1)
    ReDim bars(1 To gTotalMods + 1)

    For i = 1 To gTotalMods
        If mGrpType(i) = relType Then
            gBundleCount = gBundleCount + 1

            bHead(gBundleCount) = i
            bKind(gBundleCount) = IIf(mSep(i), "S", "M")
            bPlaced(gBundleCount) = False
            bCab(gBundleCount) = 0
            bDioOwner(gBundleCount) = 0
            bCtrlCab(gBundleCount) = 0

            mBundle(i) = gBundleCount

            nRel = nRel + 1
            relB(nRel) = gBundleCount
        End If
    Next i

    ReDim leftCh(1 To nRel + 1)

    For k = 1 To nRel
        leftCh(k) = mCh(bHead(relB(k)))
    Next k

    For i = 1 To gTotalMods
        If mIsBar(i) And mFam(i) = FAM_DIO Then
            If mBarType(i) = relType Then
                nBars = nBars + 1
                bars(nBars) = i
            End If
        End If
    Next i

    ' Сначала более многоканальные барьеры.
    For i = 2 To nBars
        tmp = bars(i)
        j = i - 1

        Do While j >= 1
            If mCh(bars(j)) >= mCh(tmp) Then Exit Do

            bars(j + 1) = bars(j)
            j = j - 1
        Loop

        bars(j + 1) = tmp
    Next i

    For i = 1 To nBars
        barIdx = bars(i)
        best = 0
        bestAfter = 2147483647

        For k = 1 To nRel
            If mSep(barIdx) = mSep(bHead(relB(k))) Then
                If leftCh(k) >= mCh(barIdx) Then
                    afterCh = leftCh(k) - mCh(barIdx)

                    If afterCh < bestAfter Then
                        bestAfter = afterCh
                        best = k
                    End If
                End If
            End If
        Next k

        If best = 0 Then
            Err.Raise vbObjectError + 30, , _
                "Барьер """ & mName(barIdx) & """ (" & _
                mCh(barIdx) & " кан.) не удалось назначить " & _
                "реле " & relName & " без превышения числа каналов." & _
                vbCrLf & "Добавьте реле или проверьте Service!F."
        End If

        mBundle(barIdx) = relB(best)
        leftCh(best) = leftCh(best) - mCh(barIdx)
    Next i
End Sub

'=============================================================================
' РЕЛЕ КОНТРОЛЛЕРА
'
' Не поместившаяся связка не помечается занятой контроллером.
' Проверяются следующие свободные связки, пока не размещено до want штук.
'=============================================================================

Private Sub ЗакрепитьРелеЗаКонтроллером( _
    ByVal cabId As Long, ByVal relType As Long, _
    ByVal want As Long, ByRef gotCnt As Long)

    Dim b As Long

    gotCnt = 0
    If want <= 0 Then Exit Sub

    For b = 1 To gBundleCount
        If gotCnt >= want Then Exit For

        If Not bPlaced(b) And _
           bDioOwner(b) = 0 And _
           bCtrlCab(b) = 0 And _
           bKind(b) = "M" And _
           ТипСвязки(b) = relType Then

            If РазместитьСвязкуВШкаф(b, cabId) Then
                bCtrlCab(b) = cabId
                gotCnt = gotCnt + 1
            End If
        End If
    Next b
End Sub

'=============================================================================
' НАЗНАЧЕНИЕ ОСТАТКА РЕЛЕ МОДУЛЯМ DIO
'=============================================================================

Private Function СвободноДляDIO(ByVal b As Long) As Boolean
    СвободноДляDIO = (Not bPlaced(b)) And _
                     bDioOwner(b) = 0 And _
                     bCtrlCab(b) = 0 And _
                     bKind(b) = "M"
End Function
Private Function ШкафСКонтроллеромИлиDIO( _
    ByVal cabId As Long) As Boolean

    Dim i As Long

    If cabCtrl(cabId) > 0 Then
        ШкафСКонтроллеромИлиDIO = True
        Exit Function
    End If

    For i = 1 To gTotalMods
        If mPlaced(i) Then
            If mGrpType(i) = TYP_DIO32 Or _
               mGrpType(i) = TYP_DIO64 Then

                If rCab(mRail(i)) = cabId Then
                    ШкафСКонтроллеромИлиDIO = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function FreeRelaysForDio(ByVal relType As Long) As Long
    Dim b As Long

    For b = 1 To gBundleCount
        If СвободноДляDIO(b) Then
            If ТипСвязки(b) = relType Then
                FreeRelaysForDio = FreeRelaysForDio + 1
            End If
        End If
    Next b
End Function

Private Sub AssignRelaysToDio( _
    ByVal dioIdx As Long, ByVal relType As Long, _
    ByVal qty As Long)

    Dim b As Long
    Dim assigned As Long

    If qty <= 0 Then Exit Sub

    For b = 1 To gBundleCount
        If СвободноДляDIO(b) Then
            If ТипСвязки(b) = relType Then
                bDioOwner(b) = dioIdx
                assigned = assigned + 1

                If assigned = qty Then Exit For
            End If
        End If
    Next b

    If assigned <> qty Then
        Err.Raise vbObjectError + 31, , _
            "Не удалось закрепить все реле за " & mName(dioIdx) & "."
    End If
End Sub

Private Sub ReserveDioBundles()
    Dim i As Long
    Dim relayLimit As Long
    Dim availableDI As Long
    Dim availableDO As Long
    Dim takeDI As Long
    Dim takeDO As Long

    gLastDio = 0

    For i = 1 To gTotalMods
        If mGrpType(i) = TYP_DIO32 Or _
           mGrpType(i) = TYP_DIO64 Then
            gLastDio = i
        End If
    Next i

    For i = 1 To gTotalMods
        If mGrpType(i) = TYP_DIO32 Or _
           mGrpType(i) = TYP_DIO64 Then

            If mGrpType(i) = TYP_DIO64 Then
                relayLimit = RELAYS_PER_DIO64
            Else
                relayLimit = RELAYS_PER_DIO32
            End If

            ' Пересчитываем остаток после фактического размещения
            ' реле P410 и после предыдущих групп DIO.
            availableDI = FreeRelaysForDio(TYP_DI)
            availableDO = FreeRelaysForDio(TYP_DO)

            If i <> gLastDio Then
                If availableDI + availableDO < relayLimit Then
                    Err.Raise vbObjectError + 32, , _
                        "Недостаточно реле для непоследней группы " & _
                        mName(i) & "." & vbCrLf & _
                        "Нужно " & relayLimit & _
                        " реле KDI/KDO суммарно; свободно KDI x" & _
                        availableDI & ", KDO x" & availableDO & "."
                End If
            End If

            ' Никаких отдельных квот 4+4 или 2+2.
            ' В группе допускается любая пропорция типов.
            takeDI = availableDI
            If takeDI > relayLimit Then takeDI = relayLimit

            takeDO = availableDO
            If takeDO > relayLimit - takeDI Then
                takeDO = relayLimit - takeDI
            End If

            AssignRelaysToDio i, TYP_DI, takeDI
            AssignRelaysToDio i, TYP_DO, takeDO
        End If
    Next i
End Sub

Private Sub СобратьСвязку( _
    ByVal bid As Long, ByRef items() As Long, ByRef n As Long)

    Dim i As Long

    If bid < 1 Or bid > gBundleCount Then
        Err.Raise vbObjectError + 33, , _
            "Неверный номер связки реле."
    End If

    If bHead(bid) < 1 Then
        Err.Raise vbObjectError + 34, , _
            "Обнаружена связка барьеров без реле."
    End If

    If bPlaced(bid) Then
        Err.Raise vbObjectError + 35, , _
            "Повторная попытка разместить связку реле."
    End If

    n = 0

    If mPlaced(bHead(bid)) Then
        Err.Raise vbObjectError + 36, , _
            "Реле из неразмещённой связки уже стоит в шкафу."
    End If

    ДобавитьВСписок bHead(bid), items, n

    For i = 1 To gTotalMods
        If i <> bHead(bid) And mBundle(i) = bid Then
            If mPlaced(i) Then
                Err.Raise vbObjectError + 37, , _
                    "Барьер уже размещён отдельно от своего реле."
            End If

            ДобавитьВСписок i, items, n
        End If
    Next i

    If n > 2 Then
        СортироватьЧастьПоДлине items, 2, n
    End If
End Sub

Private Function ТипСвязки(ByVal bid As Long) As Long
    If bHead(bid) > 0 Then
        ТипСвязки = mGrpType(bHead(bid))
    End If
End Function

Private Function ИмяСвязки(ByVal bid As Long) As String
    If bHead(bid) > 0 Then
        ИмяСвязки = mName(bHead(bid)) & " + барьеры"
    Else
        ИмяСвязки = "барьеры без реле"
    End If
End Function

'=============================================================================
' РАЗМЕЩЕНИЕ НЕДЕЛИМЫХ ГРУПП
'=============================================================================

Private Function РазместитьСвязкуВШкаф( _
    ByVal bid As Long, ByVal cabId As Long) As Boolean

    Dim items() As Long
    Dim blocks() As Long
    Dim n As Long
    Dim j As Long
    Dim railsUsed As Long
    Dim wasSplit As Boolean
    Dim hitLimit As Boolean

    If bPlaced(bid) Then
        РазместитьСвязкуВШкаф = (bCab(bid) = cabId)
        Exit Function
    End If

    If bDioOwner(bid) <> 0 Then
        Err.Raise vbObjectError + 40, , _
            "Реле, закреплённое за DIO, нельзя разместить отдельно."
    End If

    If ВидШкафа(cabId) <> bKind(bid) Then Exit Function

    ReDim items(1 To 16)
    СобратьСвязку bid, items, n

    ReDim blocks(1 To n)

    For j = 1 To n
        blocks(j) = 1
    Next j

    If TryPlaceGroupInCab( _
        cabId, bKind(bid), items, blocks, n, 1, _
        railsUsed, wasSplit, hitLimit) Then

        bPlaced(bid) = True
        bCab(bid) = cabId
        РазместитьСвязкуВШкаф = True
    End If
End Function

Private Sub РазместитьСвязку(ByVal bid As Long)
    Dim c As Long
    Dim items() As Long
    Dim n As Long

    If bPlaced(bid) Then Exit Sub

    If bDioOwner(bid) <> 0 Then
        Err.Raise vbObjectError + 41, , _
            "Связку DIO нельзя разместить отдельно."
    End If

    For c = 1 To cabCount
    If ВидШкафа(c) = bKind(bid) Then
        If bKind(bid) <> "M" Then
            If РазместитьСвязкуВШкаф(bid, c) Then Exit Sub

        ElseIf Not ШкафСКонтроллеромИлиDIO(c) Then
            If РазместитьСвязкуВШкаф(bid, c) Then Exit Sub
        End If
    End If
Next c
    ДобавитьШкаф bKind(bid)

    If РазместитьСвязкуВШкаф(bid, cabCount) Then Exit Sub

    ReDim items(1 To 16)
    СобратьСвязку bid, items, n

    Err.Raise vbObjectError + 42, , _
        "Связка """ & ИмяСвязки(bid) & """ (" & _
        Format(СуммаДлин(items, n), "0") & _
        " мм) не помещается целиком даже в пустой шкаф." & _
        vbCrLf & "Проверьте B18, D18 и D19."
End Sub

Private Sub РазместитьГруппуDIO( _
    ByVal modIdx As Long, ByRef bl() As Long, ByVal nb As Long)

    Dim items() As Long
    Dim blocks() As Long
    Dim tmp() As Long

    Dim n As Long
    Dim nBlk As Long
    Dim tn As Long
    Dim k As Long
    Dim j As Long
    Dim c As Long

    Dim railsUsed As Long
    Dim wasSplit As Boolean
    Dim hitLimit As Boolean

    ReDim items(1 To 32)
    ReDim blocks(1 To 32)
    ReDim tmp(1 To 16)

    ' Сначала СОБИРАЕМ ВСЮ группу; ничего пока не размещаем.
    n = 0
    nBlk = 1

    ДобавитьВСписокБлк modIdx, 1, items, blocks, n

    For k = 1 To nb
        If bDioOwner(bl(k)) <> modIdx Then
            Err.Raise vbObjectError + 43, , _
                "Реле назначено другому модулю DIO."
        End If

        СобратьСвязку bl(k), tmp, tn
        nBlk = nBlk + 1

        For j = 1 To tn
            ДобавитьВСписокБлк tmp(j), nBlk, _
                                 items, blocks, n
        Next j
    Next k

    ' Существующие шкафы проверяются первыми.
    ' Шкафы P410 созданы раньше остальных.
    For c = 1 To cabCount
        If ВидШкафа(c) = "M" Then
            If TryPlaceGroupInCab( _
                c, "M", items, blocks, n, nBlk, _
                railsUsed, wasSplit, hitLimit) Then

                ОтметитьСвязки bl, nb, c
                Exit Sub
            End If
        End If
    Next c

    ' Если вся группа не вошла в существующий шкаф — новый шкаф.
    ДобавитьШкаф "M"

    If TryPlaceGroupInCab( _
        cabCount, "M", items, blocks, n, nBlk, _
        railsUsed, wasSplit, hitLimit) Then

        ОтметитьСвязки bl, nb, cabCount
        Exit Sub
    End If

    If hitLimit Then
        Err.Raise vbObjectError + 44, , _
            "Для группы """ & mName(modIdx) & _
            " + реле + барьеры"" достигнут предел поиска укладки."
    Else
        Err.Raise vbObjectError + 45, , _
            "Группа """ & mName(modIdx) & _
            " + реле + барьеры"" (" & _
            Format(СуммаДлин(items, n), "0") & _
            " мм) не помещается целиком даже в пустой шкаф " & _
            "на " & gRailsTotal & " рейках." & vbCrLf & _
            "Проверьте B18, D18 и D19."
    End If
End Sub

Private Sub ОтметитьСвязки( _
    ByRef bl() As Long, ByVal nb As Long, ByVal cabId As Long)

    Dim k As Long

    For k = 1 To nb
        bPlaced(bl(k)) = True
        bCab(bl(k)) = cabId
    Next k
End Sub

Private Sub РазместитьАналоговуюГруппу( _
    ByRef items() As Long, ByVal n As Long, _
    ByVal headIdx As Long)

    Dim blocks() As Long
    Dim j As Long
    Dim c As Long
    Dim railsUsed As Long
    Dim wasSplit As Boolean
    Dim hitLimit As Boolean

    ReDim blocks(1 To n)

    For j = 1 To n
        blocks(j) = 1
    Next j

    For c = 1 To cabCount
        If ВидШкафа(c) = "M" Then
            If TryPlaceGroupInCab( _
                c, "M", items, blocks, n, 1, _
                railsUsed, wasSplit, hitLimit) Then

                Exit Sub
            End If
        End If
    Next c

    ДобавитьШкаф "M"

    If TryPlaceGroupInCab( _
        cabCount, "M", items, blocks, n, 1, _
        railsUsed, wasSplit, hitLimit) Then

        Exit Sub
    End If

    If hitLimit Then
        Err.Raise vbObjectError + 46, , _
            "Не удалось подобрать укладку для модуля """ & _
            mName(headIdx) & """ и его барьеров: " & _
            "достигнут предел поиска."
    Else
        Err.Raise vbObjectError + 47, , _
            "Модуль """ & mName(headIdx) & _
            """ и его барьеры (" & _
            Format(СуммаДлин(items, n), "0") & _
            " мм) не помещаются целиком даже в пустой шкаф." & _
            vbCrLf & "Проверьте B18, D18 и D19."
    End If
End Sub

'=============================================================================
' ПЛАНИРОВЩИК РЕЕК
'
' Сначала ищет укладку целых блоков. Если блоки не помещаются,
' разрешает распределить их элементы по разным рейкам ОДНОГО шкафа.
'
' План фиксируется только после успешной укладки ВСЕЙ группы.
'=============================================================================

Private Function TryPlaceGroupInCab( _
    ByVal cabId As Long, ByVal kind As String, _
    ByRef items() As Long, ByRef blocks() As Long, _
    ByVal n As Long, ByVal nBlk As Long, _
    ByRef railsUsed As Long, ByRef wasSplit As Boolean, _
    ByRef hitLimit As Boolean) As Boolean

    Dim plan() As Long
    Dim singleBlocks() As Long
    Dim localLimit As Boolean
    Dim j As Long

    railsUsed = 0
    wasSplit = False
    hitLimit = False

    If TryPlanBlocks( _
        cabId, kind, items, blocks, n, nBlk, _
        plan, localLimit) Then

        CommitPlan items, blocks, n, plan, railsUsed
        TryPlaceGroupInCab = True
        Exit Function
    End If

    If localLimit Then hitLimit = True

    If nBlk < n Then
        ReDim singleBlocks(1 To n)

        For j = 1 To n
            singleBlocks(j) = j
        Next j

        If TryPlanBlocks( _
            cabId, kind, items, singleBlocks, n, n, _
            plan, localLimit) Then

            CommitPlan items, singleBlocks, n, _
                       plan, railsUsed

            wasSplit = True
            TryPlaceGroupInCab = True
            Exit Function
        End If

        If localLimit Then hitLimit = True
    End If
End Function

Private Function TryPlanBlocks( _
    ByVal cabId As Long, ByVal kind As String, _
    ByRef items() As Long, ByRef blocks() As Long, _
    ByVal n As Long, ByVal nBlk As Long, _
    ByRef plan() As Long, _
    ByRef hitLimit As Boolean) As Boolean

    Dim blkLen() As Double
    Dim railIds() As Long
    Dim freeCap() As Double
    Dim ord() As Long

    Dim nRails As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim blk As Long
    Dim tmp As Long
    Dim best As Long
    Dim nodes As Long

    Dim totalLen As Double
    Dim totalFree As Double
    Dim maxFree As Double
    Dim waste As Double
    Dim bestWaste As Double
    Dim greedyOK As Boolean

    hitLimit = False

    If n <= 0 Or nBlk <= 0 Then
        Err.Raise vbObjectError + 50, , _
            "Попытка спланировать пустую группу."
    End If

    If railCount <= 0 Then Exit Function

    ReDim blkLen(1 To nBlk)
    ReDim railIds(1 To railCount)
    ReDim freeCap(1 To railCount)
    ReDim plan(1 To nBlk)

    For j = 1 To n
        If items(j) < 1 Or items(j) > gTotalMods Then
            Err.Raise vbObjectError + 51, , _
                "Неверный индекс модуля при планировании."
        End If

        If mPlaced(items(j)) Then
            Err.Raise vbObjectError + 52, , _
                "Группа содержит уже размещённый элемент: " & _
                mName(items(j)) & "."
        End If

        blk = blocks(j)

        If blk < 1 Or blk > nBlk Then
            Err.Raise vbObjectError + 53, , _
                "Неверный номер блока при планировании."
        End If

        blkLen(blk) = blkLen(blk) + mLen(items(j))
        totalLen = totalLen + mLen(items(j))
    Next j

    ' Свободное место на ВСЕХ рейках проверяемого шкафа.
    For i = 1 To railCount
        If rCab(i) = cabId And rKind(i) = kind Then
            nRails = nRails + 1
            railIds(nRails) = i

            freeCap(nRails) = rCap(i) - rUsed(i)
            If freeCap(nRails) < 0 Then freeCap(nRails) = 0

            totalFree = totalFree + freeCap(nRails)

            If freeCap(nRails) > maxFree Then
                maxFree = freeCap(nRails)
            End If
        End If
    Next i

    If nRails = 0 Then Exit Function
    If totalLen > totalFree + PLAN_EPS Then Exit Function

    ' Приоритет — вся группа на одной рейке.
    best = 0
    bestWaste = 1E+30

    For i = 1 To nRails
        waste = freeCap(i) - totalLen

        If waste >= -PLAN_EPS And waste < bestWaste Then
            best = i
            bestWaste = waste
        End If
    Next i

    If best <> 0 Then
        For blk = 1 To nBlk
            plan(blk) = railIds(best)
        Next blk

        TryPlanBlocks = True
        Exit Function
    End If

    For blk = 1 To nBlk
        If blkLen(blk) > maxFree + PLAN_EPS Then
            Exit Function
        End If
    Next blk

    ReDim ord(1 To nBlk)

    For blk = 1 To nBlk
        ord(blk) = blk
    Next blk

    ' Более длинные блоки — первыми.
    For j = 2 To nBlk
        tmp = ord(j)
        k = j - 1

        Do While k >= 1
            If blkLen(ord(k)) >= blkLen(tmp) Then Exit Do

            ord(k + 1) = ord(k)
            k = k - 1
        Loop

        ord(k + 1) = tmp
    Next j

    ' Быстрая попытка best-fit decreasing.
    greedyOK = True

    For j = 1 To nBlk
        blk = ord(j)
        best = 0
        bestWaste = 1E+30

        For i = 1 To nRails
            waste = freeCap(i) - blkLen(blk)

            If waste >= -PLAN_EPS And waste < bestWaste Then
                best = i
                bestWaste = waste
            End If
        Next i

        If best = 0 Then
            greedyOK = False
            Exit For
        End If

        plan(blk) = railIds(best)
        freeCap(best) = freeCap(best) - blkLen(blk)
    Next j

    If greedyOK Then
        TryPlanBlocks = True
        Exit Function
    End If

    ' Если быстрая попытка не нашла укладку — перебор комбинаций.
    If nBlk > PLAN_SEARCH_BLOCK_LIMIT Then
        hitLimit = True
        Exit Function
    End If

    For i = 1 To nRails
        freeCap(i) = rCap(railIds(i)) - rUsed(railIds(i))
        If freeCap(i) < 0 Then freeCap(i) = 0
    Next i

    For blk = 1 To nBlk
        plan(blk) = 0
    Next blk

    nodes = 0

    If SearchBlocks( _
        1, nBlk, ord, blkLen, railIds, nRails, _
        freeCap, plan, nodes, hitLimit) Then

        TryPlanBlocks = True
    End If
End Function

Private Function SearchBlocks( _
    ByVal level As Long, ByVal nBlk As Long, _
    ByRef ord() As Long, ByRef blkLen() As Double, _
    ByRef railIds() As Long, ByVal nRails As Long, _
    ByRef freeCap() As Double, ByRef plan() As Long, _
    ByRef nodes As Long, _
    ByRef hitLimit As Boolean) As Boolean

    Dim seenFree() As Double
    Dim seenCount As Long
    Dim duplicate As Boolean

    Dim blk As Long
    Dim i As Long
    Dim j As Long

    If hitLimit Then Exit Function

    nodes = nodes + 1

    If nodes > PLAN_NODE_LIMIT Then
        hitLimit = True
        Exit Function
    End If

    If level > nBlk Then
        SearchBlocks = True
        Exit Function
    End If

    blk = ord(level)
    ReDim seenFree(1 To nRails)

    For i = 1 To nRails
        If freeCap(i) + PLAN_EPS >= blkLen(blk) Then
            duplicate = False

            For j = 1 To seenCount
                If Abs(freeCap(i) - seenFree(j)) <= PLAN_EPS Then
                    duplicate = True
                    Exit For
                End If
            Next j

            If Not duplicate Then
                seenCount = seenCount + 1
                seenFree(seenCount) = freeCap(i)

                freeCap(i) = freeCap(i) - blkLen(blk)
                plan(blk) = railIds(i)

                If SearchBlocks( _
                    level + 1, nBlk, ord, blkLen, _
                    railIds, nRails, freeCap, plan, _
                    nodes, hitLimit) Then

                    SearchBlocks = True
                    Exit Function
                End If

                plan(blk) = 0
                freeCap(i) = freeCap(i) + blkLen(blk)

                If hitLimit Then Exit Function
            End If
        End If
    Next i
End Function

Private Sub CommitPlan( _
    ByRef items() As Long, ByRef blocks() As Long, _
    ByVal n As Long, ByRef plan() As Long, _
    ByRef railsUsed As Long)

    Dim seenRails() As Boolean
    Dim j As Long
    Dim ri As Long

    ReDim seenRails(1 To railCount)
    railsUsed = 0

    For j = 1 To n
        ri = plan(blocks(j))

        If ri < 1 Or ri > railCount Then
            Err.Raise vbObjectError + 54, , _
                "Планировщик вернул неверный номер рейки."
        End If

        PlaceModule items(j), ri

        If Not seenRails(ri) Then
            seenRails(ri) = True
            railsUsed = railsUsed + 1
        End If
    Next j
End Sub

'=============================================================================
' ПРОВЕРКИ, СПИСКИ, ОДИНОЧНЫЕ ЭЛЕМЕНТЫ
'=============================================================================

Private Sub ПроверитьРазмещение()
    Dim i As Long
    Dim b As Long
    Dim owner As Long
    Dim cabId As Long

    For i = 1 To gTotalMods
        If Not mPlaced(i) Then
            Err.Raise vbObjectError + 60, , _
                "Остался неразмещённый элемент: " & _
                mName(i) & " (индекс " & i & ")."
        End If

        If mRail(i) < 1 Or mRail(i) > railCount Then
            Err.Raise vbObjectError + 61, , _
                "У размещённого элемента нет корректной рейки."
        End If

        owner = mAnalogOwner(i)

        If owner > 0 Then
            If rCab(mRail(i)) <> rCab(mRail(owner)) Then
                Err.Raise vbObjectError + 62, , _
                    "Аналоговый модуль и его барьер оказались " & _
                    "в разных шкафах."
            End If
        End If
    Next i

    For b = 1 To gBundleCount
        If Not bPlaced(b) Or bCab(b) < 1 Then
            Err.Raise vbObjectError + 63, , _
                "Не размещена связка """ & ИмяСвязки(b) & """."
        End If

        cabId = bCab(b)
        If bCtrlCab(b) = 0 And bDioOwner(b) = 0 And bKind(b) = "M" Then
    If ШкафСКонтроллеромИлиDIO(cabId) Then
        Err.Raise vbObjectError + 76, , _
            "Реле без назначения попало в шкаф с P410 или DIO."
    End If
End If

        For i = 1 To gTotalMods
            If mBundle(i) = b Then
                If rCab(mRail(i)) <> cabId Then
                    Err.Raise vbObjectError + 64, , _
                        "Связка """ & ИмяСвязки(b) & _
                        """ разорвана между шкафами."
                End If
            End If
        Next i

        If bDioOwner(b) > 0 Then
            If rCab(mRail(bDioOwner(b))) <> cabId Then
                Err.Raise vbObjectError + 65, , _
                    "DIO и назначенное ему реле оказались " & _
                    "в разных шкафах."
            End If

            If bCtrlCab(b) > 0 Then
                Err.Raise vbObjectError + 66, , _
                    "Одно реле одновременно назначено P410 и DIO."
            End If
        End If

        If bCtrlCab(b) > 0 Then
            If bCtrlCab(b) <> cabId Then
                Err.Raise vbObjectError + 67, , _
                    "Реле контроллера размещено не в шкафу P410."
            End If

            If cabCtrl(cabId) <= 0 Then
                Err.Raise vbObjectError + 68, , _
                    "Реле назначено шкафу без P410."
            End If
        End If
    Next b
End Sub

Private Sub РазместитьОдиночный(ByVal idx As Long)
    Dim kind As String
    Dim ri As Long

    If mPlaced(idx) Then Exit Sub

    kind = IIf(mSep(idx), "S", "M")
    ri = НайтиРейку(kind, mLen(idx))

    If ri = 0 Then
        ДобавитьШкаф kind
        ri = НайтиРейку(kind, mLen(idx))
    End If

    If ri = 0 Then
        Err.Raise vbObjectError + 69, , _
            "Не найдено место для """ & mName(idx) & _
            """ даже в новом шкафу."
    End If

    PlaceModule idx, ri
End Sub

Private Sub PlaceModule(ByVal idx As Long, ByVal ri As Long)
    If idx < 1 Or idx > gTotalMods Then
        Err.Raise vbObjectError + 70, , _
            "Неверный индекс размещаемого элемента."
    End If

    If mPlaced(idx) Then
        Err.Raise vbObjectError + 71, , _
            "Повторное размещение элемента: " & _
            mName(idx) & "."
    End If

    If ri < 1 Or ri > railCount Then
        Err.Raise vbObjectError + 72, , _
            "Не найдена рейка для """ & mName(idx) & """."
    End If

    If rUsed(ri) + mLen(idx) > rCap(ri) + PLAN_EPS Then
        Err.Raise vbObjectError + 73, , _
            "Переполнение рейки шкафа " & rCab(ri) & _
            " элементом """ & mName(idx) & """." & vbCrLf & _
            "Свободно " & _
            Format(rCap(ri) - rUsed(ri), "0") & _
            " мм, требуется " & _
            Format(mLen(idx), "0") & " мм."
    End If

    mRail(idx) = ri
    mPos(idx) = rUsed(ri)
    rUsed(ri) = rUsed(ri) + mLen(idx)
    mPlaced(idx) = True
End Sub

Private Sub ДобавитьВСписок( _
    ByVal idx As Long, ByRef items() As Long, ByRef n As Long)

    n = n + 1

    If n > UBound(items) Then
        ReDim Preserve items(1 To n + 32)
    End If

    items(n) = idx
End Sub

Private Sub ДобавитьВСписокБлк( _
    ByVal idx As Long, ByVal blk As Long, _
    ByRef items() As Long, ByRef blocks() As Long, _
    ByRef n As Long)

    n = n + 1

    If n > UBound(items) Then
        ReDim Preserve items(1 To n + 32)
        ReDim Preserve blocks(1 To n + 32)
    End If

    items(n) = idx
    blocks(n) = blk
End Sub

Private Function ВзятьБарьеры( _
    ByVal fam As Long, ByVal reqCount As Long, _
    ByVal ownerIdx As Long, _
    ByRef items() As Long, ByRef n As Long) As Long

    Dim i As Long
    Dim got As Long

    If reqCount <= 0 Then Exit Function

    For i = 1 To gTotalMods
        If got >= reqCount Then Exit For

        If Not mPlaced(i) And Not mSep(i) And _
           mBundle(i) = 0 And mAnalogOwner(i) = 0 Then

            If mFam(i) = fam And mIsBar(i) Then
                mAnalogOwner(i) = ownerIdx
                ДобавитьВСписок i, items, n
                got = got + 1
            End If
        End If
    Next i

    If got <> reqCount Then
        Err.Raise vbObjectError + 74, , _
            "Не удалось назначить все аналоговые барьеры " & _
            "модулю """ & mName(ownerIdx) & """."
    End If

    ВзятьБарьеры = got
End Function

Private Function Доля( _
    ByVal total As Long, ByVal parts As Long, _
    ByVal i As Long) As Long

    If parts <= 0 Or total <= 0 Then Exit Function

    Доля = total \ parts

    If i <= total Mod parts Then
        Доля = Доля + 1
    End If
End Function

Private Function КолВСемействе( _
    ByVal fam As Long, ByVal barriers As Boolean) As Long

    Dim i As Long

    For i = 1 To gTotalMods
        If Not mPlaced(i) And Not mSep(i) Then
            If mFam(i) = fam And _
               mIsBar(i) = barriers Then

                КолВСемействе = КолВСемействе + 1
            End If
        End If
    Next i
End Function

Private Function ЕстьВСемействе( _
    ByVal fam As Long, ByVal barriers As Boolean) As Boolean

    Dim i As Long

    For i = 1 To gTotalMods
        If mFam(i) = fam And mIsBar(i) = barriers Then
            ЕстьВСемействе = True
            Exit Function
        End If
    Next i
End Function

Private Sub СортироватьЧастьПоДлине( _
    ByRef items() As Long, _
    ByVal firstIdx As Long, ByVal lastIdx As Long)

    Dim i As Long
    Dim j As Long
    Dim tmp As Long

    If firstIdx >= lastIdx Then Exit Sub

    For i = firstIdx + 1 To lastIdx
        tmp = items(i)
        j = i - 1

        Do While j >= firstIdx
            If mLen(items(j)) >= mLen(tmp) Then Exit Do

            items(j + 1) = items(j)
            j = j - 1
        Loop

        items(j + 1) = tmp
    Next i
End Sub

Private Function СуммаДлин( _
    ByRef items() As Long, ByVal n As Long) As Double

    Dim j As Long

    For j = 1 To n
        СуммаДлин = СуммаДлин + mLen(items(j))
    Next j
End Function

Private Function НайтиРейку( _
    ByVal kind As String, ByVal length As Double) As Long

    Dim i As Long

    For i = 1 To railCount
        If rKind(i) = kind Then
            If rUsed(i) + length <= rCap(i) + PLAN_EPS Then
                НайтиРейку = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function ВидШкафа(ByVal cabId As Long) As String
    Dim i As Long

    For i = 1 To railCount
        If rCab(i) = cabId Then
            ВидШкафа = rKind(i)
            Exit Function
        End If
    Next i
End Function

'=============================================================================
' СОЗДАНИЕ ШКАФОВ И РЕЕК
'=============================================================================

Private Sub РасширитьМассивыРеек(ByVal need As Long)
    Dim newSize As Long

    newSize = UBound(rCab)
    If newSize >= need Then Exit Sub

    Do While newSize < need
        newSize = newSize * 2
    Loop

    ReDim Preserve rCab(1 To newSize)
    ReDim Preserve rKind(1 To newSize)
    ReDim Preserve rSide(1 To newSize)
    ReDim Preserve rSideIdx(1 To newSize)
    ReDim Preserve rCap(1 To newSize)
    ReDim Preserve rUsed(1 To newSize)
    ReDim Preserve rPhys(1 To newSize)
    ReDim Preserve rCtrl(1 To newSize)
    ReDim Preserve rSwitch(1 To newSize)
    ReDim Preserve cabCtrl(1 To newSize)
End Sub

Private Sub ДобавитьШкаф(ByVal kind As String)
    Dim frontN As Long
    Dim backN As Long
    Dim kk As Long
    Dim firstRail As Long
    Dim nThis As Long
    Dim placedC As Long
    Dim doneSw As Long
    Dim i As Long

    РасширитьМассивыРеек railCount + gRailsTotal + 4

    cabCount = cabCount + 1
    cabCtrl(cabCount) = 0
    firstRail = railCount + 1

    ' ВАЖНО: frontN + backN = D18.
    ' Дополнительных реек для шкафа S не создаём.
    If gTwoSided Then
        frontN = gRailsTotal \ 2
        backN = gRailsTotal \ 2
    Else
        frontN = gRailsTotal
        backN = 0
    End If

    For kk = 1 To frontN
        railCount = railCount + 1

        rCab(railCount) = cabCount
        rKind(railCount) = kind
        rSide(railCount) = "F"
        rSideIdx(railCount) = kk

        rPhys(railCount) = gRailLen
        rCap(railCount) = gRailLen * (1 - gFreeFrac)
        rUsed(railCount) = 0
        rCtrl(railCount) = False
        rSwitch(railCount) = False
    Next kk

    For kk = 1 To backN
        railCount = railCount + 1

        rCab(railCount) = cabCount
        rKind(railCount) = kind
        rSide(railCount) = "B"
        rSideIdx(railCount) = kk

        rPhys(railCount) = gRailLen
        rCap(railCount) = gRailLen * (1 - gFreeFrac)
        rUsed(railCount) = 0
        rCtrl(railCount) = False
        rSwitch(railCount) = False
    Next kk

    If railCount - firstRail + 1 <> gRailsTotal Then
        Err.Raise vbObjectError + 80, , _
            "Число созданных реек не совпадает с D18."
    End If

    ' Верхняя зона резервируется немедленно при создании шкафа P410.
    If kind = "M" And gCtrlRemain > 0 Then
        nThis = gCtrlPerBlock
        If gCtrlRemain < nThis Then nThis = gCtrlRemain

        placedC = 0
        doneSw = 0

        For i = firstRail To railCount
            If rSide(i) = "F" Then
                If placedC < nThis Then
                    rCtrl(i) = True
                    rCap(i) = rCap(i) - CTRL_ZONE_MM
                    placedC = placedC + 1

                Else
                    If SWITCH_RESERVE_ALL_RAILS Or _
                       doneSw < SWITCH_RAILS Then

                        rSwitch(i) = True
                        rCap(i) = rCap(i) - CTRL_ZONE_MM
                        doneSw = doneSw + 1
                    End If
                End If

            ElseIf rSide(i) = "B" And _
                   SWITCH_RESERVE_BACK_RAILS Then

                rSwitch(i) = True
                rCap(i) = rCap(i) - CTRL_ZONE_MM
            End If

            If rCap(i) < 0 Then rCap(i) = 0
        Next i

        If placedC <> nThis Then
            Err.Raise vbObjectError + 81, , _
                "Не хватило фронтальных реек для P410."
        End If

        cabCtrl(cabCount) = nThis
        gCtrlRemain = gCtrlRemain - nThis
    End If
End Sub

'=============================================================================
' ЧТЕНИЕ ДАННЫХ И КЛАССИФИКАЦИЯ
'=============================================================================

Private Function Да(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function

    Да = (UCase$(Trim$(CStr(v))) = "YES")
End Function

Private Function ВСписке( _
    ByVal r As Long, ByVal lst As String) As Boolean

    ВСписке = (InStr( _
        1, "," & Replace(lst, " ", "") & ",", _
        "," & CStr(r) & ",") > 0)
End Function

Private Function Семейство(ByVal r As Long) As Long
    If ВСписке(r, FAM_TC_MOD) Or _
       ВСписке(r, FAM_TC_BAR) Then

        Семейство = FAM_TC

    ElseIf ВСписке(r, FAM_AI_MOD) Or _
           ВСписке(r, FAM_AI_BAR) Then

        Семейство = FAM_AI

    ElseIf ВСписке(r, FAM_DIO_REL) Or _
           ВСписке(r, FAM_DIO_BAR) Then

        Семейство = FAM_DIO
    End If
End Function

Private Function ЭтоБарьер(ByVal r As Long) As Boolean
    ЭтоБарьер = ВСписке(r, "47,48") Or _
                ВСписке(r, FAM_TC_BAR) Or _
                ВСписке(r, FAM_AI_BAR) Or _
                ВСписке(r, FAM_DIO_BAR)
End Function

Private Function ИмяСемейства(ByVal fam As Long) As String
    Select Case fam
        Case FAM_TC
            ИмяСемейства = "TC/RTD"

        Case FAM_AI
            ИмяСемейства = "AI/AO"

        Case FAM_DIO
            ИмяСемейства = "DIO"
    End Select
End Function

Private Function ОпределитьТип( _
    ByVal r As Long, ByVal nm As String) As Long

    If ВСписке(r, ROWS_KDI) Then
        ОпределитьТип = TYP_DI
        Exit Function
    End If

    If ВСписке(r, ROWS_KDO) Then
        ОпределитьТип = TYP_DO
        Exit Function
    End If

    If ВСписке(r, ROWS_DIO) Then
        If InStr(1, nm, "32", vbTextCompare) > 0 Then
            ОпределитьТип = TYP_DIO32

        ElseIf InStr(1, nm, "64", vbTextCompare) > 0 Then
            ОпределитьТип = TYP_DIO64

        Else
            ОпределитьТип = TYP_DIO64

            gWarn = gWarn & vbCrLf & _
                "  • Строка " & r & " (""" & nm & _
                """): не найдено 32/64; принято DIO64."
        End If
    End If
End Function

Private Function НормализоватьИмя(ByVal nm As String) As String
    Dim s As String
    Dim i As Long
    Dim delimiters As String

    s = UCase$(nm)
    delimiters = "-._/\(),:;+"

    For i = 1 To Len(delimiters)
        s = Replace(s, Mid$(delimiters, i, 1), " ")
    Next i

    Do While InStr(1, s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    НормализоватьИмя = " " & Trim$(s) & " "
End Function

Private Function ТипБарьераDIO( _
    ByVal r As Long, ByVal nm As String) As Long

    Dim s As String
    Dim hasDI As Boolean
    Dim hasDO As Boolean

    ' Строки 54–55 относятся к KDI.
    If r <> 53 Then
        ТипБарьераDIO = TYP_DI
        Exit Function
    End If

    s = НормализоватьИмя(nm)

    hasDI = (InStr(1, s, " KDI ") > 0 Or _
             InStr(1, s, " DI ") > 0 Or _
             InStr(1, s, " NAMUR") > 0 Or _
             InStr(1, s, " ЧАСТ") > 0 Or _
             InStr(1, s, " ВХ") > 0 Or _
             InStr(1, s, " FREQ") > 0)

    hasDO = (InStr(1, s, " KDO ") > 0 Or _
             InStr(1, s, " DO ") > 0 Or _
             InStr(1, s, " ВЫХ") > 0)

    If hasDI And Not hasDO Then
        ТипБарьераDIO = TYP_DI
    ElseIf hasDO And Not hasDI Then
        ТипБарьераDIO = TYP_DO
    Else
        ' Например, «ЛПА-340-100»: назначим экземпляры
        ' после чтения AA10.
        ТипБарьераDIO = TYP_NONE
    End If
End Function

Private Sub РаспределитьОбщиеБарьерыDIO(ByVal sigAA10 As Double)
    Dim i As Long
    Dim remainingDO As Double

    remainingDO = sigAA10

    ' Учитываем барьеры, явно обозначенные как KDO.
    For i = 1 To gTotalMods
        If mIsBar(i) And mFam(i) = FAM_DIO Then
            If mBarType(i) = TYP_DO Then
                remainingDO = remainingDO - mCh(i)
            End If
        End If
    Next i

    ' Универсальные барьеры строки 53: сначала для AA10,
    ' остальные — для KDI.
    For i = 1 To gTotalMods
        If mIsBar(i) And mFam(i) = FAM_DIO Then
            If mBarType(i) = TYP_NONE Then
                If mRow(i) <> 53 Then
                    Err.Raise vbObjectError + 75, , _
                        "Не определён тип барьера DIO в строке " & mRow(i) & "."
                End If

                If remainingDO > PLAN_EPS Then
                    mBarType(i) = TYP_DO
                    remainingDO = remainingDO - mCh(i)
                Else
                    mBarType(i) = TYP_DI
                End If
            End If
        End If
    Next i
End Sub


'=============================================================================
' АНАЛОГОВЫЕ БАРЬЕРЫ: ЁМКОСТЬ МОДУЛЕЙ И РАСПРЕДЕЛЕНИЕ
'=============================================================================

' Первая строка Service, где E = nm и D = категория.
' Допускается 0 каналов (модуль не поддерживает это направление).
Private Function КаналыПоКатегории( _
    ByVal wsS As Worksheet, ByVal nm As String, _
    ByVal cat As String, ByRef found As Boolean) As Long

    Dim lastRow As Long
    Dim r As Long
    Dim key As String
    Dim v As Variant
    Dim nameValue As Variant
    Dim catValue As Variant
    Dim d As Double

    found = False
    key = НормИмяService(nm)
    If Len(key) = 0 Then Exit Function

    lastRow = wsS.Cells(wsS.Rows.count, SVC_COL_NAME).End(xlUp).Row

    For r = 1 To lastRow
        nameValue = wsS.Cells(r, SVC_COL_NAME).Value2
        catValue = wsS.Cells(r, SVC_COL_CAT).Value2

        If Not IsError(nameValue) And Not IsError(catValue) Then
            If НормИмяService(CStr(nameValue)) = key And _
               НормКатегория(CStr(catValue)) = cat Then

                v = wsS.Cells(r, SVC_COL_CH).Value2

                If Not IsError(v) And Not IsEmpty(v) Then
                    If IsNumeric(v) Then
                        d = CDbl(v)

                        If d >= 0 And d = Fix(d) Then
                            КаналыПоКатегории = CLng(d)
                            found = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next r
End Function

Private Function НормКатегория(ByVal s As String) As String
    s = НормИмяService(s)

    Do While InStr(1, s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    НормКатегория = s
End Function

Private Sub ЁмкостьАналогМодуля( _
    ByVal wsS As Worksheet, ByVal nm As String, _
    ByVal rowIn1 As Long, _
    ByRef capAI As Long, ByRef capAO As Long)

    Dim fAI As Boolean
    Dim fAO As Boolean

    capAI = КаналыПоКатегории(wsS, nm, SVC_CAT_AI, fAI)
    capAO = КаналыПоКатегории(wsS, nm, SVC_CAT_AO, fAO)

    If Not fAI And Not fAO Then
        capAI = ANALOG_CH_DEFAULT
        capAO = 0

        gWarn = gWarn & vbCrLf & _
            "  • """ & nm & """ (строка " & rowIn1 & _
            "): нет в Service с категорией module AI/AO; " & _
            "принято AI x" & ANALOG_CH_DEFAULT & " кан."
    End If
End Sub

Private Function НаправлениеАналогБарьера( _
    ByVal r As Long, ByVal nm As String) As Long

    If ВСписке(r, ROWS_BAR_AI_AO) Then
        НаправлениеАналогБарьера = DIR_AI_AO
    Else
        ' TC/RTD и RTD>AI относятся к входным каналам.
        НаправлениеАналогБарьера = DIR_AI
    End If
End Function

' Назначает свободные барьеры семейства fam модулям этого семейства.
' Критерий: минимальная загрузка (used + ch) / cap после назначения.
' При равенстве — модуль с меньшим числом барьеров, затем больший.
Private Sub РаспределитьАналоговыеБарьеры( _
    ByVal fam As Long, ByVal dirType As Long)

    Dim mods() As Long
    Dim bars() As Long
    Dim capA() As Long
    Dim usedA() As Long
    Dim cntA() As Long

    Dim nMod As Long
    Dim nBar As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim tmp As Long
    Dim curCap As Long
    Dim ch As Long
    Dim best As Long

    Dim score As Double
    Dim bestScore As Double
    Dim totalCap As Double
    Dim totalNeed As Double
    Dim overflow As Double

    Dim better As Boolean
    Dim countOld As Boolean
    Dim groupName As String

    If gTotalMods <= 0 Then Exit Sub

    ReDim mods(1 To gTotalMods)
    ReDim bars(1 To gTotalMods)
    ReDim capA(1 To gTotalMods)
    ReDim usedA(1 To gTotalMods)
    ReDim cntA(1 To gTotalMods)

    If dirType = DIR_AI_AO Then
        groupName = "AI/AO"
    ElseIf fam = FAM_TC Then
        groupName = "TC/RTD"
    ElseIf dirType = DIR_AO Then
        groupName = "AO"
    Else
        groupName = "AI, включая RTD>AI"
    End If

    '-----------------------------------------------------------------
    ' Собираем свободные барьеры и совместимые с ними модули.
    '-----------------------------------------------------------------
    For i = 1 To gTotalMods
        If Not mPlaced(i) And Not mSep(i) Then

            If mIsBar(i) Then

                If mFam(i) = fam And _
                   mAnDir(i) = dirType And _
                   mAnalogOwner(i) = 0 Then

                    nBar = nBar + 1
                    bars(nBar) = i
                    totalNeed = totalNeed + mCh(i)
                End If

            Else
                curCap = 0

                If dirType = DIR_AI_AO Then

                    Select Case mFam(i)
                        Case FAM_AI
                            ' Общий барьер может обслуживать
                            ' входы AI и выходы AO.
                            curCap = mCapAI(i) + mCapAO(i)

                        Case FAM_TC
                            ' mCapAI здесь означает температурные
                            ' входы. Общему AI/AO-барьеру доступны
                            ' только выходы AO этого модуля.
                            curCap = mCapAO(i)
                    End Select

                ElseIf mFam(i) = fam Then

                    If dirType = DIR_AO Then
                        curCap = mCapAO(i)
                    Else
                        curCap = mCapAI(i)
                    End If
                End If

                ' Модули без подходящих каналов не включаем.
                If curCap > 0 Then
                    nMod = nMod + 1
                    mods(nMod) = i
                    capA(nMod) = curCap
                    totalCap = totalCap + curCap
                End If
            End If
        End If
    Next i

    If nBar = 0 Then Exit Sub

    If nMod = 0 Then
        gWarn = gWarn & vbCrLf & _
            "  • " & groupName & _
            ": у модулей нет подходящих каналов. " & _
            "Барьеры (" & nBar & _
            " шт.) будут размещены отдельно."

        Exit Sub
    End If

    '-----------------------------------------------------------------
    ' Учитываем ранее назначенные барьеры.
    '
    ' Для общего пула AI/AO уже назначенные RTD>AI занимают
    ' часть входов AI.
    '
    ' Температурные барьеры не занимают выходы AO
    ' температурного модуля.
    '-----------------------------------------------------------------
    For k = 1 To nMod
        For i = 1 To gTotalMods

            If mAnalogOwner(i) = mods(k) Then

                If dirType = DIR_AI_AO Then
                    countOld = (mFam(i) = FAM_AI)
                Else
                    countOld = (mFam(i) = fam And _
                                mAnDir(i) = dirType)
                End If

                If countOld Then
                    usedA(k) = usedA(k) + mCh(i)
                    cntA(k) = cntA(k) + 1
                    totalNeed = totalNeed + mCh(i)
                End If
            End If
        Next i
    Next k

    '-----------------------------------------------------------------
    ' Многоканальные барьеры назначаем первыми.
    '-----------------------------------------------------------------
    For i = 2 To nBar
        tmp = bars(i)
        j = i - 1

        Do While j >= 1
            If mCh(bars(j)) >= mCh(tmp) Then Exit Do

            bars(j + 1) = bars(j)
            j = j - 1
        Loop

        bars(j + 1) = tmp
    Next i

    '-----------------------------------------------------------------
    ' Назначение барьеров.
    ' Каждый физический барьер получает ровно одного владельца.
    '-----------------------------------------------------------------
    For i = 1 To nBar
        ch = mCh(bars(i))
        best = 0
        bestScore = 1E+30

        ' Сначала ищем назначение без превышения ёмкости.
        For k = 1 To nMod
            If capA(k) - usedA(k) >= ch Then

                score = (usedA(k) + ch) / capA(k)
                better = False

                If best = 0 Then
                    better = True

                ElseIf score < bestScore - PLAN_EPS Then
                    better = True

                ElseIf Abs(score - bestScore) <= PLAN_EPS Then

                    If cntA(k) < cntA(best) Then
                        better = True

                    ElseIf cntA(k) = cntA(best) And _
                           capA(k) > capA(best) Then

                        better = True
                    End If
                End If

                If better Then
                    best = k
                    bestScore = score
                End If
            End If
        Next k

        ' Как в исходном коде: если назначить без превышения
        ' не удалось, выполняем условное распределение.
        ' После назначения будет выдано предупреждение.
        If best = 0 Then
            bestScore = 1E+30

            For k = 1 To nMod
                score = (usedA(k) + ch) / capA(k)

                If score < bestScore Then
                    best = k
                    bestScore = score
                End If
            Next k
        End If

        mAnalogOwner(bars(i)) = mods(best)
        usedA(best) = usedA(best) + ch
        cntA(best) = cntA(best) + 1
    Next i

    ' Фактическое превышение по отдельным модулям.
    For k = 1 To nMod
        If usedA(k) > capA(k) Then
            overflow = overflow + usedA(k) - capA(k)
        End If
    Next k

    If overflow > 0 Then
        gWarn = gWarn & vbCrLf & _
            "  • " & groupName & _
            ": учтено барьеров на " & _
            Format(totalNeed, "0") & " кан.; " & _
            "ёмкость модулей — " & _
            Format(totalCap, "0") & " кан. " & _
            "Превышение по отдельным модулям: " & _
            Format(overflow, "0") & " кан. " & _
            "Размещение выполнено условно; " & _
            "проверьте количество модулей и распределение каналов."
    End If
End Sub

Private Function НормИмяService(ByVal nm As String) As String
    НормИмяService = UCase$( _
        Trim$(Replace(nm, ChrW$(160), " ")))
End Function

Private Function ШиринаИзService( _
    ByVal wsS As Worksheet, ByVal nm As String, _
    ByVal fallback As Double, ByVal rowIn1 As Long) As Double

    Dim lastRow As Long
    Dim r As Long
    Dim key As String
    Dim v As Variant
    Dim nameValue As Variant

    key = НормИмяService(nm)

    If Len(key) > 0 Then
        lastRow = wsS.Cells( _
            wsS.Rows.count, SVC_COL_NAME).End(xlUp).Row

        For r = 1 To lastRow
            nameValue = wsS.Cells(r, SVC_COL_NAME).Value2

            If Not IsError(nameValue) Then
                If НормИмяService(CStr(nameValue)) = key Then
                    v = wsS.Cells(r, SVC_COL_WIDTH).Value2

                    If Not IsError(v) Then
                        If IsNumeric(v) Then
                            If CDbl(v) > 0 Then
                                ШиринаИзService = CDbl(v)
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        Next r
    End If

    gWarn = gWarn & vbCrLf & _
        "  • """ & nm & """ (строка " & rowIn1 & _
        "): нет положительной ширины в Service!H; " & _
        "взято " & fallback & " мм."

    ШиринаИзService = fallback
End Function

Private Function КаналыИзService( _
    ByVal wsS As Worksheet, ByVal nm As String, _
    ByVal fallback As Long, ByVal rowIn1 As Long, _
    ByVal what As String) As Long

    Dim lastRow As Long
    Dim r As Long
    Dim key As String
    Dim v As Variant
    Dim nameValue As Variant
    Dim d As Double

    key = НормИмяService(nm)

    If Len(key) > 0 Then
        lastRow = wsS.Cells( _
            wsS.Rows.count, SVC_COL_NAME).End(xlUp).Row

        For r = 1 To lastRow
            nameValue = wsS.Cells(r, SVC_COL_NAME).Value2

            If Not IsError(nameValue) Then
                If НормИмяService(CStr(nameValue)) = key Then
                    v = wsS.Cells(r, SVC_COL_CH).Value2

                    If Not IsError(v) Then
                        If IsNumeric(v) Then
                            d = CDbl(v)

                            If d > 0 And d <= 2147483647# Then
                                If d = Fix(d) Then
                                    КаналыИзService = CLng(d)
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next r
    End If

    gWarn = gWarn & vbCrLf & _
        "  • """ & nm & """ (строка " & rowIn1 & _
        "): нет целого положительного числа каналов " & _
        "в Service!F; принято " & fallback & _
        " кан. на " & what & "."

    КаналыИзService = fallback
End Function

Private Function ЧитатьКол( _
    ByVal ws As Worksheet, ByVal r As Long) As Long

    Dim v As Variant

    v = ws.Range(COL_QTY & r).Value2

    If IsError(v) Then Exit Function

    If IsNumeric(v) Then
        ЧитатьКол = CLng(Int(CDbl(v)))

        If ЧитатьКол < 0 Then ЧитатьКол = 0
    End If
End Function

Private Function ЧитатьЧисло( _
    ByVal ws As Worksheet, ByVal addr As String) As Double

    Dim v As Variant

    v = ws.Range(addr).Value2

    If IsError(v) Or IsEmpty(v) Then Exit Function

    If IsNumeric(v) Then
        ЧитатьЧисло = CDbl(v)
    Else
        ЧитатьЧисло = ИзвлечьЧисло(v)
    End If

    If ЧитатьЧисло < 0 Then ЧитатьЧисло = 0
End Function

Private Function ОчиститьИмя( _
    ByVal s As String, ByVal r As Long) As String

    Dim p As Long
    Dim pref As String

    pref = "АРКС400."
    p = InStr(1, s, pref, vbTextCompare)

    If p > 0 Then
        ОчиститьИмя = Trim$(Mid$(s, p + Len(pref)))
    Else
        ОчиститьИмя = Trim$(s)
    End If

    If Len(ОчиститьИмя) = 0 Then
        ОчиститьИмя = "Mod" & r
    End If
End Function

Private Function ИзвлечьЧисло(ByVal v As Variant) As Long
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim digits As String

    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function

    s = CStr(v)

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)

        If ch >= "0" And ch <= "9" Then
            digits = digits & ch

        ElseIf Len(digits) > 0 Then
            Exit For
        End If
    Next i

    If Len(digits) > 0 Then
        ИзвлечьЧисло = CLng(digits)
    End If
End Function

'=============================================================================
' ОТРИСОВКА
'=============================================================================

Private Sub НарисоватьРасположение( _
    ByVal wl As Worksheet, _
    ByRef legName() As String, ByRef legLen() As Double, _
    ByRef legQty() As Long, ByRef legColor() As Long, _
    ByVal legN As Long, ByVal ctrlCount As Long, _
    ByRef usedCnt As Long)

    Const MM As Double = 0.25
    Const RAILW As Double = 78
    Const RAILGAP As Double = 6
    Const SIDEGAP As Double = 22
    Const CABGAP As Double = 55
    Const TOPY As Double = 120
    Const LEFTX As Double = 40

    Dim cabUsed() As Boolean
    Dim drawList() As Long
    Dim railX() As Double

    Dim si As Long
    Dim i As Long
    Dim c As Long
    Dim rr As Long
    Dim idx As Long
    Dim dN As Long

    Dim x As Double
    Dim maxH As Double
    Dim minX As Double
    Dim maxX As Double
    Dim hPx As Double
    Dim capPx As Double
    Dim modH As Double
    Dim yOff As Double
    Dim legX As Double
    Dim legY As Double

    Dim kindC As String
    Dim txtCol As Long
    Dim rb As Shape
    Dim msh As Shape

    If cabCount < 1 Or railCount < 1 Then
        Err.Raise vbObjectError + 90, , _
            "Нет шкафов для отрисовки."
    End If

    For si = wl.Shapes.count To 1 Step -1
        wl.Shapes(si).Delete
    Next si

    wl.Cells.Clear

    ReDim cabUsed(1 To cabCount)

    For i = 1 To railCount
        If rUsed(i) > 0 Then
            cabUsed(rCab(i)) = True
        End If
    Next i

    For i = 1 To cabCount
        If cabCtrl(i) > 0 Then
            cabUsed(i) = True
        End If
    Next i

    ReDim drawList(1 To railCount)
    dN = 0

    For i = 1 To railCount
        If cabUsed(rCab(i)) Then
            dN = dN + 1
            drawList(dN) = i
        End If
    Next i

    ReDim railX(1 To railCount)

    x = LEFTX

    For i = 1 To dN
        If i > 1 Then
            x = x + RAILW + RAILGAP

            If rCab(drawList(i)) <> rCab(drawList(i - 1)) Then
                x = x + CABGAP

            ElseIf rSide(drawList(i)) <> _
                   rSide(drawList(i - 1)) Then

                x = x + SIDEGAP
            End If
        End If

        railX(drawList(i)) = x
    Next i

    maxH = gRailLen * MM

    ДобТекст wl, LEFTX, 5, 500, 15, _
        "КОНФИГУРАТОР РАСПОЛОЖЕНИЯ МОДУЛЕЙ", _
        14, True, RGB(0, 0, 0), msoAlignLeft

    ' Контуры шкафов.
    For c = 1 To cabCount
        If cabUsed(c) Then
            minX = 1E+30
            maxX = -1E+30
            kindC = "M"

            For i = 1 To railCount
                If rCab(i) = c Then
                    If railX(i) < minX Then minX = railX(i)
                    If railX(i) > maxX Then maxX = railX(i)

                    kindC = rKind(i)
                End If
            Next i

            ДобРект wl, minX - 14, TOPY - 36, _
                (maxX + RAILW + 14) - (minX - 14), _
                maxH + 52, _
                RGB(248, 249, 251), _
                RGB(120, 130, 145), 1.5

            ДобТекст wl, minX - 14, TOPY - 56, _
                maxX - minX + RAILW + 28, 18, _
                "Шкаф " & c & " — " & _
                IIf(kindC = "S", _
                    "отдельный барьерный/релейный", _
                    IIf(gTwoSided, _
                        "двухсторонний", _
                        "односторонний")), _
                10, True, RGB(20, 40, 80), _
                msoAlignLeft
        End If
    Next c

    ' Рейки и зарезервированные зоны.
    For i = 1 To dN
        rr = drawList(i)
        hPx = rPhys(rr) * MM
        capPx = rPhys(rr) * (1 - gFreeFrac) * MM

        ДобРект wl, railX(rr), TOPY, RAILW, hPx, _
            RGB(232, 235, 240), _
            RGB(150, 160, 175), 0.75

        If hPx - capPx > 1 Then
            Set rb = ДобРект( _
                wl, railX(rr), TOPY + capPx, _
                RAILW, hPx - capPx, _
                RGB(255, 244, 204), _
                RGB(220, 180, 60), 0.5)

            rb.line.DashStyle = msoLineDash

            If hPx - capPx > 16 Then
                ДобТекст wl, railX(rr), _
                    TOPY + capPx, _
                    RAILW, hPx - capPx, "резерв", _
                    6, False, RGB(150, 110, 0), _
                    msoAlignCenter
            End If
        End If

        If rCtrl(rr) Then
            ДобРект wl, railX(rr), TOPY, _
                RAILW, CTRL_BLOCK_PX, _
                RGB(46, 125, 50), _
                RGB(27, 70, 30), 0.75

            ДобТекст wl, railX(rr), TOPY, _
                RAILW, CTRL_BLOCK_PX, "P410", _
                8, True, RGB(255, 255, 255), _
                msoAlignCenter

        ElseIf rSwitch(rr) Then
            ДобРект wl, railX(rr), TOPY, _
                RAILW, CTRL_BLOCK_PX, _
                RGB(21, 101, 192), _
                RGB(13, 60, 115), 0.75

            ДобТекст wl, railX(rr), TOPY, _
                RAILW, CTRL_BLOCK_PX, _
                "Коммутаторы и БП", _
                7.5, True, RGB(255, 255, 255), _
                msoAlignCenter
        End If

        ДобТекст wl, railX(rr), TOPY - 30, _
            RAILW, 28, _
            IIf(rSide(rr) = "F", _
                "Фронтальная", "Задняя") & _
            " · Р" & rSideIdx(rr) & vbLf & _
            Format(rUsed(rr), "0") & "/" & _
            Format(rCap(rr), "0") & " мм", _
            7, True, RGB(40, 40, 40), _
            msoAlignCenter
    Next i

    ' Модули и барьеры.
    For idx = 1 To gTotalMods
        rr = mRail(idx)

        If rr >= 1 And rr <= railCount Then
            yOff = 0

            If rCtrl(rr) Or rSwitch(rr) Then
                yOff = CTRL_OFFSET_PX
            End If

            modH = mLen(idx) * MM

            Set msh = ДобРект( _
                wl, railX(rr) + 1, _
                TOPY + yOff + mPos(idx) * MM, _
                RAILW - 2, modH, mColor(idx), _
                RGB(60, 60, 60), 0.25)

            If modH >= 13 Then
                txtCol = ТекстЦвет(mColor(idx))

                ДобТекст wl, railX(rr) + 1, _
                    TOPY + yOff + mPos(idx) * MM, _
                    RAILW - 2, modH, mName(idx), _
                    6, False, txtCol, msoAlignCenter
            End If
        End If
    Next idx

    ' Легенда.
    legX = LEFTX
    legY = TOPY + maxH + 50

    ДобТекст wl, legX, legY, 400, 18, _
        "ЛЕГЕНДА", 11, True, _
        RGB(0, 0, 0), msoAlignLeft

    legY = legY + 24

    For i = 1 To legN
        ДобРект wl, legX, legY, 16, 14, _
            legColor(i), RGB(60, 60, 60), 0.5

        ДобТекст wl, legX + 22, legY - 2, _
            460, 18, _
            legName(i) & "   (" & _
            Format(legLen(i), "0") & " мм x " & _
            legQty(i) & " шт)", _
            8.5, False, RGB(30, 30, 30), _
            msoAlignLeft

        legY = legY + 19
    Next i

    legY = legY + 8

    ДобРект wl, legX, legY, 16, 14, _
        RGB(255, 244, 204), _
        RGB(220, 180, 60), 0.5

    ДобТекст wl, legX + 22, legY - 2, _
        500, 18, _
        "Свободное место / резерв: " & _
        Format(gFreeFrac, "0%"), _
        8.5, True, RGB(120, 90, 0), _
        msoAlignLeft

    legY = legY + 22

    If ctrlCount > 0 Then
        ДобРект wl, legX, legY, 16, 14, _
            RGB(46, 125, 50), _
            RGB(27, 70, 30), 0.5

        ДобТекст wl, legX + 22, legY - 2, _
            600, 18, "Контроллер P410", _
            8.5, True, RGB(20, 70, 30), _
            msoAlignLeft

        legY = legY + 22

        ДобРект wl, legX, legY, 16, 14, _
            RGB(21, 101, 192), _
            RGB(13, 60, 115), 0.5

        ДобТекст wl, legX + 22, legY - 2, _
            600, 18, "Коммутаторы и БП", _
            8.5, True, RGB(15, 70, 130), _
            msoAlignLeft
    End If

    usedCnt = 0

    For c = 1 To cabCount
        If cabUsed(c) Then usedCnt = usedCnt + 1
    Next c

    wl.Activate
    wl.Range("A1").Select
End Sub

Private Function Палитра(ByVal i As Long) As Long
    Static colors(1 To 20) As Long
    Static initialized As Boolean

    If Not initialized Then
        colors(1) = RGB(31, 119, 180)
        colors(2) = RGB(255, 127, 14)
        colors(3) = RGB(44, 160, 44)
        colors(4) = RGB(214, 39, 40)
        colors(5) = RGB(148, 103, 189)
        colors(6) = RGB(140, 86, 75)
        colors(7) = RGB(227, 119, 194)
        colors(8) = RGB(127, 127, 127)
        colors(9) = RGB(188, 189, 34)
        colors(10) = RGB(23, 190, 207)
        colors(11) = RGB(174, 199, 232)
        colors(12) = RGB(255, 187, 120)
        colors(13) = RGB(152, 223, 138)
        colors(14) = RGB(255, 152, 150)
        colors(15) = RGB(197, 176, 213)
        colors(16) = RGB(196, 156, 148)
        colors(17) = RGB(247, 182, 210)
        colors(18) = RGB(199, 199, 199)
        colors(19) = RGB(219, 219, 141)
        colors(20) = RGB(158, 218, 229)

        initialized = True
    End If

    If i < 1 Then i = 1

    Палитра = colors(((i - 1) Mod 20) + 1)
End Function

Private Function ТекстЦвет(ByVal bg As Long) As Long
    Dim rr As Long
    Dim gg As Long
    Dim bb As Long
    Dim lum As Double

    rr = bg Mod 256
    gg = (bg \ 256) Mod 256
    bb = (bg \ 65536) Mod 256

    lum = 0.299 * rr + 0.587 * gg + 0.114 * bb

    If lum < 140 Then
        ТекстЦвет = RGB(255, 255, 255)
    Else
        ТекстЦвет = RGB(0, 0, 0)
    End If
End Function

Private Function ДобРект( _
    ByVal wl As Worksheet, _
    ByVal l As Double, ByVal t As Double, _
    ByVal w As Double, ByVal h As Double, _
    ByVal fillCol As Long, ByVal lineCol As Long, _
    Optional ByVal lineW As Single = 0.75) As Shape

    Dim s As Shape

    If w < 1 Then w = 1
    If h < 1 Then h = 1

    Set s = wl.Shapes.AddShape( _
        msoShapeRectangle, l, t, w, h)

    s.fill.ForeColor.RGB = fillCol
    s.fill.Visible = msoTrue

    If lineCol < 0 Then
        s.line.Visible = msoFalse
    Else
        s.line.Visible = msoTrue
        s.line.ForeColor.RGB = lineCol
        s.line.Weight = lineW
    End If

    On Error Resume Next
    s.Shadow.Visible = msoFalse
    On Error GoTo 0

    Set ДобРект = s
End Function

Private Sub ДобТекст( _
    ByVal wl As Worksheet, _
    ByVal l As Double, ByVal t As Double, _
    ByVal w As Double, ByVal h As Double, _
    ByVal txt As String, ByVal sz As Single, _
    ByVal isBold As Boolean, ByVal col As Long, _
    Optional ByVal align As Long = 1)

    Dim s As Shape

    If w < 1 Then w = 1
    If h < 1 Then h = 1

    Set s = wl.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, l, t, w, h)

    s.fill.Visible = msoFalse
    s.line.Visible = msoFalse

    With s.TextFrame2
        .MarginLeft = 1
        .MarginRight = 1
        .MarginTop = 0
        .MarginBottom = 0
        .WordWrap = msoTrue
        .VerticalAnchor = msoAnchorMiddle

        With .TextRange
            .Text = txt
            .Font.Size = sz
            .Font.bold = isBold
            .Font.fill.ForeColor.RGB = col
            .ParagraphFormat.Alignment = align
        End With
    End With
End Sub

