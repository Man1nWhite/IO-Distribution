Attribute VB_Name = "Module1"
Option Explicit

'  онфигуратор шкафов ј– —400: раскладка по DIN-рейкам и чертЄж.
' “очка входа Ч —оздать онфигуратор. ѕравила и диагностика Ч в README.md.

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


Private Const ROWS_SEPARATE As String = _
    "34,35,36,37,47,48,49,50,51,52,53,54,55"

Private Const RELAY_CH_DEFAULT As Long = 8
Private Const BAR_CH_DEFAULT As Long = 1

Private Const ADDR_EXI_DI As String = "V10"
Private Const ADDR_NAMUR_DI As String = "X10"
Private Const ADDR_EXI_DO As String = "AA10"
Private Const ADDR_FREQ_DI As String = "AC10"

' Service: E Ч наименование, F Ч каналы, H Ч ширина.
Private Const SVC_COL_NAME As Long = 5
Private Const SVC_COL_CH As Long = 6
Private Const SVC_COL_WIDTH As Long = 8
' Service: D Ч категори€ ("module AI" / "module AO").
Private Const SVC_COL_CAT As Long = 4
Private Const SVC_CAT_AI As String = "MODULE AI"
Private Const SVC_CAT_AO As String = "MODULE AO"

Private Const ANALOG_CH_DEFAULT As Long = 8

' Ќаправление аналогового барьера задаЄтс€ номером строки листа "1":
' строка 52 (см. ROWS_BAR_AI_AO) Ч общие AI/AO, остальные барьеры Ч входы AI.

Private Const DIR_AI As Long = 1
Private Const DIR_AI_AO As Long = 3

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

Private Const CTRL_ZONE_MM As Double = 224
Private Const CTRL_BLOCK_PX As Double = 52
Private Const CTRL_OFFSET_PX As Double = 56

Private Const PLAN_EPS As Double = 0.0001
Private Const PLAN_NODE_LIMIT As Long = 300000
Private Const PLAN_SEARCH_BLOCK_LIMIT As Long = 120

' –ейки и шкафы.
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

' Ёкземпл€ры модулей.
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
Private mCapAI() As Long      ' каналы AI модул€
Private mCapAO() As Long      ' каналы AO модул€
Private mAnDir() As Long      ' DIR_AI / DIR_AI_AO дл€ аналогового барьера
Private gTotalMods As Long

' —в€зки Ђреле + его барьерыї.
Private bHead() As Long
Private bKind() As String
Private bPlaced() As Boolean
Private bCab() As Long
Private bDioOwner() As Long
Private bCtrlCab() As Long
Private gBundleCount As Long

'=============================================================================
' √Ћј¬Ќјя ѕ–ќ÷≈ƒ”–ј
'=============================================================================

Public Sub —оздать онфигуратор()
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

    Dim leftCnt As Long
    Dim leftDI As Long
    Dim leftDO As Long

    Dim bl() As Long
    Dim nb As Long
    Dim relayLimit As Long

    Dim actualDI As Long
    Dim actualDO As Long

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
        MsgBox "ќЎ»Ѕ ј: не найден лист ""1"".", vbCritical
        Exit Sub
    End If

    If wsS Is Nothing Then
        MsgBox "ќЎ»Ѕ ј: не найден лист ""Service"".", vbCritical
        Exit Sub
    End If

    ' ѕараметры шкафа. D18 Ч „»—Ћќ –≈≈  ¬  ј∆ƒќћ Ў ј‘”.
    gRailsTotal = »звлечь„исло(ws.Range("D18").Value2)

    v = ws.Range("B18").Value2

    If IsError(v) Then
        Err.Raise vbObjectError + 10, , _
            "ќшибка в B18: укажите длину рейки."
    End If

    If IsNumeric(v) And Not IsEmpty(v) Then
        gRailLen = CDbl(v)
    Else
        ' –езультат Ч Double (не »звлечь„исло): дробна€ часть сохран€етс€
        gRailLen = „исло»з“екста(CStr(v))
    End If

    If gRailLen <= 0 Then
        Err.Raise vbObjectError + 11, , _
            "¬ B18 должна быть положительна€ длина рейки."
    End If

    Select Case gRailsTotal
        Case 3, 4
            gTwoSided = False

        Case 6, 8
            gTwoSided = True

        Case Else
            Err.Raise vbObjectError + 12, , _
                "Ќедопустимое значение D18. Ќужно 3, 4, 6 или 8 реек."
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
            "  Х –езерв D19 ограничен значением 90%."
    End If

    gSeparateMode = ƒа(ws.Range("D13").Value2)

    ctrlCount = „итать ол(ws, 26)

    gCtrlPerBlock = 1
    If ƒа(ws.Range("D12").Value2) Then gCtrlPerBlock = 2

    ' –езервированна€ пара P410 делит один комплект KDI/KDO: нечЄтное D26
    ' оставл€ет один шкаф с одним контроллером Ч об этом сообщаем заранее.
    If gCtrlPerBlock = 2 And ctrlCount > 0 And _
       (ctrlCount Mod 2) <> 0 Then

        gWarn = gWarn & vbCrLf & _
            "  Х ¬ D12 выбрано резервирование контроллеров (YES), " & _
            "но в D26 указано нечЄтное количество P410 (" & _
            ctrlCount & "); один шкаф будет с 1 контроллером."
    End If

    capFull = gRailLen * (1 - gFreeFrac)
    capCtrl = capFull - CTRL_ZONE_MM

    If ctrlCount > 0 And capCtrl < -PLAN_EPS Then
        Err.Raise vbObjectError + 13, , _
            "ѕолезна€ длина рейки меньше верхней зоны P410/" & _
            "коммутаторов (" & CTRL_ZONE_MM & " мм). " & _
            "ѕроверьте B18 и D19."
    End If

    defRows = Array(27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, _
                    47, 48, 49, 50, 51, 52, 53, 54, 55)

    ' ѕроверка текстовых количеств: одно предупреждение на строку.
    For t = LBound(defRows) To UBound(defRows)
        ѕроверитьячейку ол ws, CLng(defRows(t))
    Next t

    ѕроверитьячейку ол ws, 26

    defLensFallback = Array(35, 164, 164, 157, 157, 157, 157, _
                            110, 110, 100, 100, 23, 67, 13, 23, _
                            23, 13, 13, 13, 23)

    ReDim defLens(LBound(defRows) To UBound(defRows))

    total = 0

    For t = LBound(defRows) To UBound(defRows)
        r = CLng(defRows(t))
        q = „итать ол(ws, r)
        total = total + q

        If q > 0 Then
            v = ws.Range(COL_NAME & r).Value2

            If IsError(v) Then
                Err.Raise vbObjectError + 14, , _
                    "ќшибка в наименовании модул€, строка " & r & "."
            End If

            defLens(t) = Ўирина»зService( _
                wsS, CStr(v), CDbl(defLensFallback(t)), r)
        Else
            defLens(t) = CDbl(defLensFallback(t))
        End If
    Next t

    If total = 0 And ctrlCount = 0 Then
        MsgBox "Ќет модулей и контроллеров дл€ размещени€.", _
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
        q = „итать ол(ws, r)

        If q > 0 Then
            rawNm = CStr(ws.Range(COL_NAME & r).Value2)
            nm = ќчистить»м€(rawNm, r)
            col = ѕалитра(t + 1)

            isSep = gSeparateMode And ¬—писке(r, ROWS_SEPARATE)
            typ = ќпределить“ип(r, nm)
            fam = —емейство(r)
            isBar = ЁтоЅарьер(r)

            ch = 0
            barT = TYP_NONE

            If typ = TYP_DI Or typ = TYP_DO Then
                ch =  аналы»зService( _
                    wsS, rawNm, RELAY_CH_DEFAULT, r, "реле")

            ElseIf isBar And fam = FAM_DIO Then
                ch =  аналы»зService( _
                    wsS, rawNm, BAR_CH_DEFAULT, r, "барьер")

                barT = “ипЅарьераDIO(r, nm)
            End If
            capAI = 0
            capAO = 0
            anDir = 0

            If fam = FAM_TC Or fam = FAM_AI Then
                If isBar Then
                    ch =  аналы»зService( _
                        wsS, rawNm, BAR_CH_DEFAULT, r, _
                        "аналоговый барьер")

                    anDir = ЌаправлениејналогЅарьера(r)
                Else
                    ®мкостьјналогћодул€ wsS, rawNm, r, capAI, capAO
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
            "ћодуль шириной " & Format(maxLen, "0") & _
            " мм не помещаетс€ на рейку. ѕолезна€ длина: " & _
            Format(capFull, "0") & " мм."
    End If

    If ctrlCount > 0 And maxLen > capCtrl + PLAN_EPS Then
        gWarn = gWarn & vbCrLf & _
            "  Х Ќекоторые модули не помещаютс€ на рейку с верхней " & _
            "зоной P410/коммутаторов; будут использованы другие рейки."
    End If

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

    sigDI = „итать„исло(ws, ADDR_EXI_DI) + _
            „итать„исло(ws, ADDR_NAMUR_DI) + _
            „итать„исло(ws, ADDR_FREQ_DI)

    sigDO = „итать„исло(ws, ADDR_EXI_DO)
    –аспределитьќбщиеЅарьерыDIO sigDO

    If sigDI > 0 Then
        needKDI = -Int(-sigDI / relayChDI)

        If relayCapDI < sigDI Then
            gWarn = gWarn & vbCrLf & _
                "  Х —игналы DI: " & Format(sigDI, "0") & _
                "; требуетс€ ориентировочно " & needKDI & _
                " KDI по " & relayChDI & _
                " каналов. »меющихс€ каналов KDI: " & _
                Format(relayCapDI, "0") & "."
        End If
    End If

    If sigDO > 0 Then
        needKDO = -Int(-sigDO / relayChDO)

        If relayCapDO < sigDO Then
            gWarn = gWarn & vbCrLf & _
                "  Х —игналы DO: " & Format(sigDO, "0") & _
                "; требуетс€ ориентировочно " & needKDO & _
                " KDO по " & relayChDO & _
                " каналов. »меющихс€ каналов KDO: " & _
                Format(relayCapDO, "0") & "."
        End If
    End If

    ' D13=YES физически разносит реле/барьеры и основные модули.
    If gSeparateMode Then
        If nDIO > 0 Then
            Err.Raise vbObjectError + 16, , _
                "D13=YES отправл€ет реле и барьеры в отдельный шкаф, " & _
                "а DIO оставл€ет в основном. ƒл€ группы " & _
                "DIO + реле + барьеры установите D13=NO."
        End If

        If ctrlCount > 0 And (nDI > 0 Or nDO > 0) Then
            Err.Raise vbObjectError + 17, , _
                "ѕри D13=YES реле размещаютс€ отдельно от P410. " & _
                "ƒл€ подключени€ реле к контроллеру установите D13=NO."
        End If

        If ≈сть¬—емействе(FAM_TC, False) And _
           ≈сть¬—емействе(FAM_TC, True) Then

            Err.Raise vbObjectError + 18, , _
                "D13=YES раздел€ет модуль TC/RTD и его барьеры. " & _
                "ƒл€ совместного размещени€ установите D13=NO."
        End If

        If ≈сть¬—емействе(FAM_AI, False) And _
           ≈сть¬—емействе(FAM_AI, True) Then

            Err.Raise vbObjectError + 19, , _
                "D13=YES раздел€ет модуль AI/AO и его барьеры. " & _
                "ƒл€ совместного размещени€ установите D13=NO."
        End If
    End If

    —формировать—в€зки

    ' Ё“јѕ 1. Ўкафы P410: ƒобавитьЎкаф резервирует место сверху ƒќ реле.
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
        ƒобавитьЎкаф "M"
    Loop

    If cabCount = 0 Then
        For idx = 1 To total
            If Not mSep(idx) Then
                ƒобавитьЎкаф "M"
                Exit For
            End If
        Next idx
    End If

    ' Ё“јѕ 1.1. –еле, подключаемые к P410: св€зка считаетс€ зан€той
    ' контроллером “ќЋ№ ќ если реле и барьеры разместились в данном шкафу.
    If Not gSeparateMode Then
        For c = 1 To cabCount
            If cabCtrl(c) > 0 Then
                «акрепить–еле«а онтроллером _
                    c, TYP_DI, DI_PER_CTRL, gotCDI

                «акрепить–еле«а онтроллером _
                    c, TYP_DO, DO_PER_CTRL, gotCDO
            End If
        Next c
    End If

    ' Ё“јѕ 2. —¬ќЅќƒЌџ≈ реле после P410 назначаютс€ модул€м DIO;
    ' последнему DIO разрешено получить меньше нормы каждого типа.
    ReserveDioBundles

    ReDim bl(1 To 16)

    For idx = 1 To total
        If mGrpType(idx) = TYP_DIO32 Or _
           mGrpType(idx) = TYP_DIO64 Then

            If mPlaced(idx) Then
                Err.Raise vbObjectError + 20, , _
                    "DIO уже размещЄн до формировани€ группы."
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
                            "–еле дл€ " & mName(idx) & _
                            " уже размещено отдельно от DIO."
                    End If

                    nb = nb + 1

                    If nb > UBound(bl) Then
                        ReDim Preserve bl(1 To nb + 16)
                    End If

                    bl(nb) = b

                    Select Case “ип—в€зки(b)
                        Case TYP_DI
                            actualDI = actualDI + 1

                        Case TYP_DO
                            actualDO = actualDO + 1

                        Case Else
                            Err.Raise vbObjectError + 22, , _
                                "Ќеизвестный тип реле в группе DIO."
                    End Select
                End If
            Next b

            If actualDI + actualDO > relayLimit Then
                Err.Raise vbObjectError + 23, , _
                    "ƒл€ " & mName(idx) & _
                    " назначено больше допустимого числа реле: " & _
                    (actualDI + actualDO) & " из " & relayLimit & "."
            End If

            If idx <> gLastDio Then
                If actualDI + actualDO <> relayLimit Then
                    Err.Raise vbObjectError + 24, , _
                        "Ќепоследн€€ группа " & mName(idx) & _
                        " должна содержать " & relayLimit & _
                        " реле KDI/KDO суммарно."
                End If
            End If


            ' ƒл€ ѕќ—Ћ≈ƒЌ≈√ќ DIO меньше нормы допустимо Ч без предупреждений.
            –азместить√руппуDIO idx, bl, nb
        End If
    Next idx

    ' Ё“јѕ 2.2. јналоговые модули и их барьеры: 49-50 -> температурные входы,
    ' 51 -> только входы AI, 52 -> общий пул AI/AO (он же Ч выходы AO
    ' температурных модулей). —начала барьеры с ограниченной совместимостью,
    ' затем единственный общий пул.
    If Not gSeparateMode Then

        –аспределитьјналоговыеЅарьеры FAM_TC, DIR_AI
        –аспределитьјналоговыеЅарьеры FAM_AI, DIR_AI
        –аспределитьјналоговыеЅарьеры FAM_AI, DIR_AI_AO

        ReDim items(1 To 32)

        For idx = 1 To total
            If Not mPlaced(idx) And Not mSep(idx) And _
               Not mIsBar(idx) Then

                If mFam(idx) = FAM_TC Or _
                   mFam(idx) = FAM_AI Then

                    n = 0
                    ƒобавить¬—писок idx, items, n

                    ' —обираем ¬—≈ барьеры владельца: семейство барьера может
                    ' отличатьс€ от семейства модул€, например AO у TC-модул€.
                    For k = 1 To total
                        If mAnalogOwner(k) = idx And _
                           Not mPlaced(k) Then

                            ƒобавить¬—писок k, items, n
                        End If
                    Next k

                    If n > 1 Then
                        If n > 2 Then
                            —ортировать„астьѕоƒлине items, 2, n
                        End If

                        –азместитьјналоговую√руппу items, n, idx
                    End If

                    ' ћодули без назначенных барьеров уйдут на этап одиночных.
                End If
            End If
        Next idx
    End If

    ' Ё“јѕ 2.3. ќстаток реле: при D13=NO Ч ошибка 28,
    ' при D13=YES Ч размещение св€зок в отдельном шкафу S.
    leftCnt = 0
    leftDI = 0
    leftDO = 0

    For b = 1 To gBundleCount
        If Not bPlaced(b) And bDioOwner(b) = 0 And _
           bKind(b) = "M" Then

            leftCnt = leftCnt + 1

            If “ип—в€зки(b) = TYP_DI Then
                leftDI = leftDI + 1
            ElseIf “ип—в€зки(b) = TYP_DO Then
                leftDO = leftDO + 1
            End If
        End If
    Next b

    If leftCnt > 0 Then
        Err.Raise vbObjectError + 28, , _
            "ќстались реле, не назначенные ни P410, ни модул€м DIO: " & _
            "KDI x" & leftDI & ", KDO x" & leftDO & "." & vbCrLf & _
            " аждое реле должно обслуживать либо контроллер P410 " & _
            "(до " & DI_PER_CTRL & " KDI и " & DO_PER_CTRL & _
            " KDO на шкаф), либо группу DIO (до " & RELAYS_PER_DIO64 & _
            " реле дл€ DIO64 / " & RELAYS_PER_DIO32 & " дл€ DIO32)." & _
            vbCrLf & _
            "ѕроверьте количество реле (строки 34-37) и число модулей " & _
            "DIO; также убедитесь, что св€зки реле с барьерами " & _
            "помещаютс€ в шкафы P410 (B18, D18, D19, ширины в Service!H)." & _
            vbCrLf & _
            "≈сли нужен отдельный релейный шкаф (без P410 и DIO), " & _
            "установите D13=YES."
    End If

    For b = 1 To gBundleCount
        If Not bPlaced(b) Then
            If bDioOwner(b) <> 0 Then
                Err.Raise vbObjectError + 25, , _
                    "—в€зка реле DIO осталась неразмещЄнной. " & _
                    "–азмещать еЄ отдельно запрещено."
            End If

            –азместить—в€зку b
        End If
    Next b

    ' Ё“јѕ 3. ќстальные одиночные элементы.
    For idx = 1 To total
        If Not mPlaced(idx) And mBundle(idx) = 0 And _
           Not mIsBar(idx) Then

            If mGrpType(idx) = TYP_DIO32 Or _
               mGrpType(idx) = TYP_DIO64 Then

                Err.Raise vbObjectError + 26, , _
                    "DIO не прошЄл этап размещени€ группы."
            End If

            –азместитьќдиночный idx
        End If
    Next idx

    For idx = 1 To total
        If Not mPlaced(idx) And mBundle(idx) = 0 Then
            If mFam(idx) = FAM_DIO And mIsBar(idx) Then
                Err.Raise vbObjectError + 27, , _
                    "Ѕарьер DIO осталс€ без назначенного реле."
            End If

            –азместитьќдиночный idx
        End If
    Next idx

    ѕроверить–азмещение

    ' –исуем только после успешного расчЄта.
    If wl Is Nothing Then
        Set wl = ThisWorkbook.Worksheets.Add(after:=ws)
        wl.Name = "1 расположение"
    End If

    Application.ScreenUpdating = False

    Ќарисовать–асположение wl, legName, legLen, legQty, _
                           legColor, legN, ctrlCount, usedCnt

    Application.ScreenUpdating = True

    msg = "√отово!" & vbCrLf & _
          "Ўкафов задействовано: " & usedCnt

    If Len(gWarn) > 0 Then
        MsgBox msg & vbCrLf & vbCrLf & _
               "¬Ќ»ћјЌ»≈:" & gWarn, vbExclamation
    Else
        MsgBox msg, vbInformation
    End If

    Exit Sub

ErrH:
    errText = Err.Description
    Application.ScreenUpdating = True

    MsgBox "ќшибка конфигуратора:" & vbCrLf & vbCrLf & _
           errText, vbCritical
End Sub

'=============================================================================
' ѕ–»¬я« ј Ѕј–№≈–ќ¬   –≈Ћ≈
'=============================================================================

Private Sub —формировать—в€зки()
    Dim i As Long

    gBundleCount = 0

    For i = 1 To gTotalMods
        mBundle(i) = 0
        bDioOwner(i) = 0
        bCtrlCab(i) = 0
    Next i

    ‘ормировать—в€зки“ипа TYP_DI, "KDI"
    ‘ормировать—в€зки“ипа TYP_DO, "KDO"
End Sub

Private Sub ‘ормировать—в€зки“ипа( _
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

    ' —начала более многоканальные барьеры.
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
                "Ѕарьер """ & mName(barIdx) & """ (" & _
                mCh(barIdx) & " кан.) не удалось назначить " & _
                "реле " & relName & " без превышени€ числа каналов." & _
                vbCrLf & "ƒобавьте реле или проверьте Service!F."
        End If

        mBundle(barIdx) = relB(best)
        leftCh(best) = leftCh(best) - mCh(barIdx)
    Next i
End Sub

'=============================================================================
' –≈Ћ≈  ќЌ“–ќЋЋ≈–ј. Ќе поместивша€с€ св€зка не помечаетс€ зан€той
' контроллером; перебираютс€ следующие свободные св€зки, пока не размещено
' до want штук.
'=============================================================================

Private Sub «акрепить–еле«а онтроллером( _
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
           “ип—в€зки(b) = relType Then

            ' —начала св€зка целиком Ч на наименее загруженную рейку шкафа:
            ' так до 4 KDI и 4 KDO расход€тс€ по разным рейкам, а не набиваютс€ на первую.
            ' Ўтатный планировщик (с дроблением элементов) Ч только если
            ' св€зка не помещаетс€ на одну рейку целиком.
            If –азместить—в€зкуЌа—вободной–ейке(b, cabId) Then
                bCtrlCab(b) = cabId
                gotCnt = gotCnt + 1

            ElseIf –азместить—в€зку¬Ўкаф(b, cabId) Then
                bCtrlCab(b) = cabId
                gotCnt = gotCnt + 1
            End If
        End If
    Next b
End Sub

' ѕробует уложить св€зку реле контроллера ÷≈Ћ» ќћ на одну рейку шкафа P410:
' выбираетс€ наименее загруженна€ рейка (максимум свободного места), куда
' св€зка [–еле + его барьеры] влезает целиком. ≈сли такой рейки нет, ничего не
' размещаетс€: вызывающий код примен€ет обычный планировщик, вплоть
' до дроблени€ элементов св€зки по рейкам.
Private Function –азместить—в€зкуЌа—вободной–ейке( _
    ByVal bid As Long, ByVal cabId As Long) As Boolean

    Dim items() As Long
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim bestRail As Long
    Dim bestFree As Double
    Dim curFree As Double
    Dim totalLen As Double

    If bPlaced(bid) Then Exit Function
    If bDioOwner(bid) <> 0 Then Exit Function
    If ¬идЎкафа(cabId) <> bKind(bid) Then Exit Function

    ReDim items(1 To 16)
    —обрать—в€зку bid, items, n
    totalLen = —уммаƒлин(items, n)

    bestFree = -1

    For i = 1 To railCount
        If rCab(i) = cabId And rKind(i) = bKind(bid) Then
            curFree = rCap(i) - rUsed(i)

            If curFree < 0 Then curFree = 0

            If curFree + PLAN_EPS >= totalLen Then
                If curFree > bestFree Then
                    bestRail = i
                    bestFree = curFree
                End If
            End If
        End If
    Next i

    If bestRail = 0 Then Exit Function

    For j = 1 To n
        PlaceModule items(j), bestRail
    Next j

    bPlaced(bid) = True
    bCab(bid) = cabId
    –азместить—в€зкуЌа—вободной–ейке = True
End Function

'=============================================================================
' Ќј«Ќј„≈Ќ»≈ ќ—“ј“ ј –≈Ћ≈ ћќƒ”Ћяћ DIO
'=============================================================================

Private Function —вободноƒл€DIO(ByVal b As Long) As Boolean
    —вободноƒл€DIO = (Not bPlaced(b)) And _
                     bDioOwner(b) = 0 And _
                     bCtrlCab(b) = 0 And _
                     bKind(b) = "M"
End Function
Private Function Ўкаф— онтроллером»лиDIO( _
    ByVal cabId As Long) As Boolean

    Dim i As Long

    If cabCtrl(cabId) > 0 Then
        Ўкаф— онтроллером»лиDIO = True
        Exit Function
    End If

    For i = 1 To gTotalMods
        If mPlaced(i) Then
            If mGrpType(i) = TYP_DIO32 Or _
               mGrpType(i) = TYP_DIO64 Then

                If rCab(mRail(i)) = cabId Then
                    Ўкаф— онтроллером»лиDIO = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function FreeRelaysForDio(ByVal relType As Long) As Long
    Dim b As Long

    For b = 1 To gBundleCount
        If —вободноƒл€DIO(b) Then
            If “ип—в€зки(b) = relType Then
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
        If —вободноƒл€DIO(b) Then
            If “ип—в€зки(b) = relType Then
                bDioOwner(b) = dioIdx
                assigned = assigned + 1

                If assigned = qty Then Exit For
            End If
        End If
    Next b

    If assigned <> qty Then
        Err.Raise vbObjectError + 31, , _
            "Ќе удалось закрепить все реле за " & mName(dioIdx) & "."
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

            ' ќстаток Ч после фактического размещени€ реле P410 и прошлых DIO.
            availableDI = FreeRelaysForDio(TYP_DI)
            availableDO = FreeRelaysForDio(TYP_DO)

            If i <> gLastDio Then
                If availableDI + availableDO < relayLimit Then
                    Err.Raise vbObjectError + 32, , _
                        "Ќедостаточно реле дл€ непоследней группы " & _
                        mName(i) & "." & vbCrLf & _
                        "Ќужно " & relayLimit & _
                        " реле KDI/KDO суммарно; свободно KDI x" & _
                        availableDI & ", KDO x" & availableDO & "."
                End If
            End If

            '  вот 4+4 или 2+2 нет: в группе допустима люба€ пропорци€ типов.
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

Private Sub —обрать—в€зку( _
    ByVal bid As Long, ByRef items() As Long, ByRef n As Long)

    Dim i As Long

    If bid < 1 Or bid > gBundleCount Then
        Err.Raise vbObjectError + 33, , _
            "Ќеверный номер св€зки реле."
    End If

    If bHead(bid) < 1 Then
        Err.Raise vbObjectError + 34, , _
            "ќбнаружена св€зка барьеров без реле."
    End If

    If bPlaced(bid) Then
        Err.Raise vbObjectError + 35, , _
            "ѕовторна€ попытка разместить св€зку реле."
    End If

    n = 0

    If mPlaced(bHead(bid)) Then
        Err.Raise vbObjectError + 36, , _
            "–еле из неразмещЄнной св€зки уже стоит в шкафу."
    End If

    ƒобавить¬—писок bHead(bid), items, n

    For i = 1 To gTotalMods
        If i <> bHead(bid) And mBundle(i) = bid Then
            If mPlaced(i) Then
                Err.Raise vbObjectError + 37, , _
                    "Ѕарьер уже размещЄн отдельно от своего реле."
            End If

            ƒобавить¬—писок i, items, n
        End If
    Next i

    If n > 2 Then
        —ортировать„астьѕоƒлине items, 2, n
    End If
End Sub

Private Function “ип—в€зки(ByVal bid As Long) As Long
    If bHead(bid) > 0 Then
        “ип—в€зки = mGrpType(bHead(bid))
    End If
End Function

Private Function »м€—в€зки(ByVal bid As Long) As String
    If bHead(bid) > 0 Then
        »м€—в€зки = mName(bHead(bid)) & " + барьеры"
    Else
        »м€—в€зки = "барьеры без реле"
    End If
End Function

'=============================================================================
' –ј«ћ≈ў≈Ќ»≈ Ќ≈ƒ≈Ћ»ћџ’ √–”ѕѕ
'=============================================================================

Private Function –азместить—в€зку¬Ўкаф( _
    ByVal bid As Long, ByVal cabId As Long) As Boolean

    Dim items() As Long
    Dim blocks() As Long
    Dim n As Long
    Dim j As Long
    Dim railsUsed As Long
    Dim wasSplit As Boolean
    Dim hitLimit As Boolean

    If bPlaced(bid) Then
        –азместить—в€зку¬Ўкаф = (bCab(bid) = cabId)
        Exit Function
    End If

    If bDioOwner(bid) <> 0 Then
        Err.Raise vbObjectError + 40, , _
            "–еле, закреплЄнное за DIO, нельз€ разместить отдельно."
    End If

    If ¬идЎкафа(cabId) <> bKind(bid) Then Exit Function

    ReDim items(1 To 16)
    —обрать—в€зку bid, items, n

    ReDim blocks(1 To n)

    For j = 1 To n
        blocks(j) = 1
    Next j

    If TryPlaceGroupInCab( _
        cabId, bKind(bid), items, blocks, n, 1, _
        railsUsed, wasSplit, hitLimit) Then

        bPlaced(bid) = True
        bCab(bid) = cabId
        –азместить—в€зку¬Ўкаф = True
    End If
End Function

Private Sub –азместить—в€зку(ByVal bid As Long)
    Dim c As Long
    Dim items() As Long
    Dim n As Long

    If bPlaced(bid) Then Exit Sub

    If bDioOwner(bid) <> 0 Then
        Err.Raise vbObjectError + 41, , _
            "—в€зку DIO нельз€ разместить отдельно."
    End If

    For c = 1 To cabCount
        If ¬идЎкафа(c) = bKind(bid) Then
            If bKind(bid) <> "M" Then
                If –азместить—в€зку¬Ўкаф(bid, c) Then Exit Sub

            ElseIf Not Ўкаф— онтроллером»лиDIO(c) Then
                If –азместить—в€зку¬Ўкаф(bid, c) Then Exit Sub
            End If
        End If
    Next c
    ƒобавитьЎкаф bKind(bid)

    If –азместить—в€зку¬Ўкаф(bid, cabCount) Then Exit Sub

    ReDim items(1 To 16)
    —обрать—в€зку bid, items, n

    Err.Raise vbObjectError + 42, , _
        "—в€зка """ & »м€—в€зки(bid) & """ (" & _
        Format(—уммаƒлин(items, n), "0") & _
        " мм) не помещаетс€ целиком даже в пустой шкаф." & _
        vbCrLf & "ѕроверьте B18, D18 и D19."
End Sub

Private Sub –азместить√руппуDIO( _
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

    ' —начала —ќЅ»–ј≈ћ ¬—ё группу; ничего пока не размещаем.
    n = 0
    nBlk = 1

    ƒобавить¬—писокЅлк modIdx, 1, items, blocks, n

    For k = 1 To nb
        If bDioOwner(bl(k)) <> modIdx Then
            Err.Raise vbObjectError + 43, , _
                "–еле назначено другому модулю DIO."
        End If

        —обрать—в€зку bl(k), tmp, tn
        nBlk = nBlk + 1

        For j = 1 To tn
            ƒобавить¬—писокЅлк tmp(j), nBlk, _
                                 items, blocks, n
        Next j
    Next k

    ' —уществующие шкафы провер€ютс€ первыми: шкафы P410 созданы раньше.
    For c = 1 To cabCount
        If ¬идЎкафа(c) = "M" Then
            If TryPlaceGroupInCab( _
                c, "M", items, blocks, n, nBlk, _
                railsUsed, wasSplit, hitLimit) Then

                ќтметить—в€зки bl, nb, c
                Exit Sub
            End If
        End If
    Next c

    ƒобавитьЎкаф "M"

    If TryPlaceGroupInCab( _
        cabCount, "M", items, blocks, n, nBlk, _
        railsUsed, wasSplit, hitLimit) Then

        ќтметить—в€зки bl, nb, cabCount
        Exit Sub
    End If

    If hitLimit Then
        Err.Raise vbObjectError + 44, , _
            "ƒл€ группы """ & mName(modIdx) & _
            " + реле + барьеры"" достигнут предел поиска укладки."
    Else
        Err.Raise vbObjectError + 45, , _
            "√руппа """ & mName(modIdx) & _
            " + реле + барьеры"" (" & _
            Format(—уммаƒлин(items, n), "0") & _
            " мм) не помещаетс€ целиком даже в пустой шкаф " & _
            "на " & gRailsTotal & " рейках." & vbCrLf & _
            "ѕроверьте B18, D18 и D19."
    End If
End Sub

Private Sub ќтметить—в€зки( _
    ByRef bl() As Long, ByVal nb As Long, ByVal cabId As Long)

    Dim k As Long

    For k = 1 To nb
        bPlaced(bl(k)) = True
        bCab(bl(k)) = cabId
    Next k
End Sub

Private Sub –азместитьјналоговую√руппу( _
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
        If ¬идЎкафа(c) = "M" Then
            If TryPlaceGroupInCab( _
                c, "M", items, blocks, n, 1, _
                railsUsed, wasSplit, hitLimit) Then

                Exit Sub
            End If
        End If
    Next c

    ƒобавитьЎкаф "M"

    If TryPlaceGroupInCab( _
        cabCount, "M", items, blocks, n, 1, _
        railsUsed, wasSplit, hitLimit) Then

        Exit Sub
    End If

    If hitLimit Then
        Err.Raise vbObjectError + 46, , _
            "Ќе удалось подобрать укладку дл€ модул€ """ & _
            mName(headIdx) & """ и его барьеров: " & _
            "достигнут предел поиска."
    Else
        Err.Raise vbObjectError + 47, , _
            "ћодуль """ & mName(headIdx) & _
            """ и его барьеры (" & _
            Format(—уммаƒлин(items, n), "0") & _
            " мм) не помещаютс€ целиком даже в пустой шкаф." & _
            vbCrLf & "ѕроверьте B18, D18 и D19."
    End If
End Sub

'=============================================================================
' ѕЋјЌ»–ќ¬ў»  –≈≈ . —начала ищет укладку целых блоков; если блоки не
' помещаютс€, разрешает распределить их элементы по разным рейкам ќƒЌќ√ќ
' шкафа. ѕлан фиксируетс€ только после успешной укладки ¬—≈… группы.
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
        ' —в€зки приоритетны: дробитс€ по одному самый широкий составной
        ' блок, а остальные остаютс€ неразрывными [–еле + его барьеры].
        If TryPlacePartialSplit( _
            cabId, kind, items, blocks, n, nBlk, _
            railsUsed, hitLimit) Then

            wasSplit = True
            TryPlaceGroupInCab = True
            Exit Function
        End If

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
            "ѕопытка спланировать пустую группу."
    End If

    If railCount <= 0 Then Exit Function

    ReDim blkLen(1 To nBlk)
    ReDim railIds(1 To railCount)
    ReDim freeCap(1 To railCount)
    ReDim plan(1 To nBlk)

    For j = 1 To n
        If items(j) < 1 Or items(j) > gTotalMods Then
            Err.Raise vbObjectError + 51, , _
                "Ќеверный индекс модул€ при планировании."
        End If

        If mPlaced(items(j)) Then
            Err.Raise vbObjectError + 52, , _
                "√руппа содержит уже размещЄнный элемент: " & _
                mName(items(j)) & "."
        End If

        blk = blocks(j)

        If blk < 1 Or blk > nBlk Then
            Err.Raise vbObjectError + 53, , _
                "Ќеверный номер блока при планировании."
        End If

        blkLen(blk) = blkLen(blk) + mLen(items(j))
        totalLen = totalLen + mLen(items(j))
    Next j

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

    ' ѕриоритет Ч вс€ группа на одной рейке.
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

    ' Ѕыстра€ попытка best-fit decreasing.
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

    ' ≈сли быстра€ попытка не нашла укладку Ч перебор комбинаций.
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
    Dim startRail As Long

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

    ' Ѕлоки одинаковой ширины идут в ord подр€д: перестановки равноценны,
    ' поэтому рейки перебираютс€ начина€ с рейки предыдущего такого блока.
    startRail = 1

    If level > 1 Then
        If Abs(blkLen(blk) - blkLen(ord(level - 1))) <= PLAN_EPS Then
            For i = 1 To nRails
                If railIds(i) = plan(ord(level - 1)) Then
                    startRail = i
                    Exit For
                End If
            Next i
        End If
    End If

    For i = startRail To nRails
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

' ћ€гкое дробление группы: св€зки остаютс€ ÷≈Ћџћ» настолько, насколько это
' возможно. ‘аза 1 Ч дробитс€ ровно ќƒ»Ќ составной блок, по очереди от самого
' широкого (остальные св€зки [–еле + его барьеры] целы). ‘аза 2 Ч к самому
' широкому блоку добавл€етс€ следующий по ширине, и так далее. ≈сли ни один
' вариант не уложилс€, вызывающий код переходит к штатному дроблению на
' отдельные элементы.
Private Function TryPlacePartialSplit( _
    ByVal cabId As Long, ByVal kind As String, _
    ByRef items() As Long, ByRef blocks() As Long, _
    ByVal n As Long, ByVal nBlk As Long, _
    ByRef railsUsed As Long, _
    ByRef hitLimit As Boolean) As Boolean

    Dim part() As Long
    Dim blkMap() As Long
    Dim blkCnt() As Long
    Dim blkLen() As Double
    Dim done() As Boolean
    Dim tried() As Boolean
    Dim plan() As Long

    Dim blk As Long
    Dim best As Long
    Dim nextId As Long
    Dim partN As Long
    Dim phase As Long
    Dim nTry As Long
    Dim j As Long

    Dim localLimit As Boolean

    If nBlk < 2 Then Exit Function

    ReDim blkCnt(1 To nBlk)
    ReDim blkLen(1 To nBlk)
    ReDim done(1 To nBlk)
    ReDim tried(1 To nBlk)
    ReDim blkMap(1 To nBlk)

    For j = 1 To n
        blk = blocks(j)

        If blk < 1 Or blk > nBlk Then Exit Function

        blkCnt(blk) = blkCnt(blk) + 1
        blkLen(blk) = blkLen(blk) + mLen(items(j))
    Next j

    For phase = 1 To 2
        For blk = 1 To nBlk
            done(blk) = False
            tried(blk) = False
        Next blk

        For nTry = 1 To nBlk
            best = 0

            For blk = 1 To nBlk
                If blkCnt(blk) > 1 And Not tried(blk) Then
                    If best = 0 Then
                        best = blk
                    ElseIf blkLen(blk) > blkLen(best) + PLAN_EPS Then
                        best = blk
                    End If
                End If
            Next blk

            If best = 0 Then Exit For

            tried(best) = True

            ' ‘аза 1 дробит по одному блоку, фаза 2 накапливает их.
            If phase = 1 Then
                For blk = 1 To nBlk
                    done(blk) = False
                Next blk
            End If

            done(best) = True

            ' ¬ариант блоков: раздробленные Ч по элементам, остальные целиком.
            nextId = 0

            For blk = 1 To nBlk
                blkMap(blk) = 0

                If blkCnt(blk) > 0 And Not done(blk) Then
                    nextId = nextId + 1
                    blkMap(blk) = nextId
                End If
            Next blk

            ReDim part(1 To n)
            partN = nextId

            For j = 1 To n
                If done(blocks(j)) Then
                    partN = partN + 1
                    part(j) = partN
                Else
                    part(j) = blkMap(blocks(j))
                End If
            Next j

            ' ÷елых блоков не осталось Ч это штатный фоллбэк вызывающего.
            If partN >= n Then Exit For

            If TryPlanBlocks( _
                cabId, kind, items, part, n, partN, _
                plan, localLimit) Then

                CommitPlan items, part, n, plan, railsUsed
                TryPlacePartialSplit = True
                Exit Function
            End If

            ' ѕредел поиска уже достигнут: следующие варианты только дороже.
            If localLimit Then
                hitLimit = True
                Exit Function
            End If
        Next nTry
    Next phase
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
                "ѕланировщик вернул неверный номер рейки."
        End If

        PlaceModule items(j), ri

        If Not seenRails(ri) Then
            seenRails(ri) = True
            railsUsed = railsUsed + 1
        End If
    Next j
End Sub

'=============================================================================
' ѕ–ќ¬≈– », —ѕ»— », ќƒ»Ќќ„Ќџ≈ ЁЋ≈ћ≈Ќ“џ
'=============================================================================

Private Sub ѕроверить–азмещение()
    Dim i As Long
    Dim b As Long
    Dim owner As Long
    Dim cabId As Long
    Dim sumLen() As Double
    Dim cntRail() As Long
    Dim tol As Double

    For i = 1 To gTotalMods
        If Not mPlaced(i) Then
            Err.Raise vbObjectError + 60, , _
                "ќсталс€ неразмещЄнный элемент: " & _
                mName(i) & " (индекс " & i & ")."
        End If

        If mRail(i) < 1 Or mRail(i) > railCount Then
            Err.Raise vbObjectError + 61, , _
                "” размещЄнного элемента нет корректной рейки."
        End If

        owner = mAnalogOwner(i)

        If owner > 0 Then
            If rCab(mRail(i)) <> rCab(mRail(owner)) Then
                Err.Raise vbObjectError + 62, , _
                    "јналоговый модуль и его барьер оказались " & _
                    "в разных шкафах."
            End If
        End If
    Next i

    For b = 1 To gBundleCount
        If Not bPlaced(b) Or bCab(b) < 1 Then
            Err.Raise vbObjectError + 63, , _
                "Ќе размещена св€зка """ & »м€—в€зки(b) & """."
        End If

        cabId = bCab(b)
        If bCtrlCab(b) = 0 And bDioOwner(b) = 0 And bKind(b) = "M" Then
            If Ўкаф— онтроллером»лиDIO(cabId) Then
                Err.Raise vbObjectError + 76, , _
                    "–еле без назначени€ попало в шкаф с P410 или DIO."
            End If
        End If

        For i = 1 To gTotalMods
            If mBundle(i) = b Then
                If rCab(mRail(i)) <> cabId Then
                    Err.Raise vbObjectError + 64, , _
                        "—в€зка """ & »м€—в€зки(b) & _
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
                    "ќдно реле одновременно назначено P410 и DIO."
            End If
        End If

        If bCtrlCab(b) > 0 Then
            If bCtrlCab(b) <> cabId Then
                Err.Raise vbObjectError + 67, , _
                    "–еле контроллера размещено не в шкафу P410."
            End If

            If cabCtrl(cabId) <= 0 Then
                Err.Raise vbObjectError + 68, , _
                    "–еле назначено шкафу без P410."
            End If
        End If
    Next b

    ' »нвариант рейки: суммарна€ ширина размещЄнных элементов не может
    ' превышать полезную длину рейки. —умма считаетс€ независимо по
    ' элементам, а не по накопленному учЄту, поэтому проверка ловит и
    ' расхождение rUsed с суммой ширин. ƒопуск PLAN_EPS на элемент Ч
    ' PlaceModule допускает такое превышение Ємкости дл€ последнего элемента.
    If railCount >= 1 Then
        ReDim sumLen(1 To railCount)
        ReDim cntRail(1 To railCount)

        For i = 1 To gTotalMods
            sumLen(mRail(i)) = sumLen(mRail(i)) + mLen(i)
            cntRail(mRail(i)) = cntRail(mRail(i)) + 1
        Next i

        For i = 1 To railCount
            tol = PLAN_EPS * (cntRail(i) + 1)

            If sumLen(i) > rCap(i) + tol Then
                Err.Raise vbObjectError + 77, , _
                    "–ейка " & i & " шкафа " & rCab(i) & _
                    " переполнена: " & Format(sumLen(i), "0") & _
                    " мм при Ємкости " & Format(rCap(i), "0") & " мм."
            End If

            If Abs(sumLen(i) - rUsed(i)) > tol Then
                Err.Raise vbObjectError + 78, , _
                    "”чЄт зан€тости рейки " & i & " разошЄлс€ с суммой " & _
                    "ширины элементов: " & Format(rUsed(i), "0") & _
                    " мм против " & Format(sumLen(i), "0") & " мм."
            End If
        Next i
    End If
End Sub

Private Sub –азместитьќдиночный(ByVal idx As Long)
    Dim kind As String
    Dim ri As Long

    If mPlaced(idx) Then Exit Sub

    kind = IIf(mSep(idx), "S", "M")
    ri = Ќайти–ейку(kind, mLen(idx))

    If ri = 0 Then
        ƒобавитьЎкаф kind
        ri = Ќайти–ейку(kind, mLen(idx))
    End If

    If ri = 0 Then
        Err.Raise vbObjectError + 69, , _
            "Ќе найдено место дл€ """ & mName(idx) & _
            """ даже в новом шкафу."
    End If

    PlaceModule idx, ri
End Sub

Private Sub PlaceModule(ByVal idx As Long, ByVal ri As Long)
    If idx < 1 Or idx > gTotalMods Then
        Err.Raise vbObjectError + 70, , _
            "Ќеверный индекс размещаемого элемента."
    End If

    If mPlaced(idx) Then
        Err.Raise vbObjectError + 71, , _
            "ѕовторное размещение элемента: " & _
            mName(idx) & "."
    End If

    If ri < 1 Or ri > railCount Then
        Err.Raise vbObjectError + 72, , _
            "Ќе найдена рейка дл€ """ & mName(idx) & """."
    End If

    If rUsed(ri) + mLen(idx) > rCap(ri) + PLAN_EPS Then
        Err.Raise vbObjectError + 73, , _
            "ѕереполнение рейки шкафа " & rCab(ri) & _
            " элементом """ & mName(idx) & """." & vbCrLf & _
            "—вободно " & _
            Format(rCap(ri) - rUsed(ri), "0") & _
            " мм, требуетс€ " & _
            Format(mLen(idx), "0") & " мм."
    End If

    mRail(idx) = ri
    mPos(idx) = rUsed(ri)
    rUsed(ri) = rUsed(ri) + mLen(idx)
    mPlaced(idx) = True
End Sub

Private Sub ƒобавить¬—писок( _
    ByVal idx As Long, ByRef items() As Long, ByRef n As Long)

    n = n + 1

    If n > UBound(items) Then
        ReDim Preserve items(1 To n + 32)
    End If

    items(n) = idx
End Sub

Private Sub ƒобавить¬—писокЅлк( _
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

Private Function ≈сть¬—емействе( _
    ByVal fam As Long, ByVal barriers As Boolean) As Boolean

    Dim i As Long

    For i = 1 To gTotalMods
        If mFam(i) = fam And mIsBar(i) = barriers Then
            ≈сть¬—емействе = True
            Exit Function
        End If
    Next i
End Function

Private Sub —ортировать„астьѕоƒлине( _
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

Private Function —уммаƒлин( _
    ByRef items() As Long, ByVal n As Long) As Double

    Dim j As Long

    For j = 1 To n
        —уммаƒлин = —уммаƒлин + mLen(items(j))
    Next j
End Function

Private Function Ќайти–ейку( _
    ByVal kind As String, ByVal length As Double) As Long

    Dim i As Long

    For i = 1 To railCount
        If rKind(i) = kind Then
            If rUsed(i) + length <= rCap(i) + PLAN_EPS Then
                Ќайти–ейку = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function ¬идЎкафа(ByVal cabId As Long) As String
    Dim i As Long

    For i = 1 To railCount
        If rCab(i) = cabId Then
            ¬идЎкафа = rKind(i)
            Exit Function
        End If
    Next i
End Function

'=============================================================================
' —ќ«ƒјЌ»≈ Ў ј‘ќ¬ » –≈≈ 
'=============================================================================

Private Sub –асширитьћассивы–еек(ByVal need As Long)
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

Private Sub ƒобавитьЎкаф(ByVal kind As String)
    Dim frontN As Long
    Dim backN As Long
    Dim kk As Long
    Dim firstRail As Long
    Dim nThis As Long
    Dim placedC As Long
    Dim doneSw As Long
    Dim i As Long

    –асширитьћассивы–еек railCount + gRailsTotal + 4

    cabCount = cabCount + 1
    cabCtrl(cabCount) = 0
    firstRail = railCount + 1

    ' ¬ј∆Ќќ: frontN + backN = D18.
    ' ƒополнительных реек дл€ шкафа S не создаЄм.
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
            "„исло созданных реек не совпадает с D18."
    End If

    ' ¬ерхн€€ зона резервируетс€ немедленно при создании шкафа P410.
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
                "Ќе хватило фронтальных реек дл€ P410."
        End If

        cabCtrl(cabCount) = nThis
        gCtrlRemain = gCtrlRemain - nThis
    End If
End Sub

'=============================================================================
' „“≈Ќ»≈ ƒјЌЌџ’ »  Ћј——»‘» ј÷»я
'=============================================================================

Private Function ƒа(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function

    ƒа = (UCase$(Trim$(CStr(v))) = "YES")
End Function

Private Function ¬—писке( _
    ByVal r As Long, ByVal lst As String) As Boolean

    ¬—писке = (InStr( _
        1, "," & Replace(lst, " ", "") & ",", _
        "," & CStr(r) & ",") > 0)
End Function

Private Function —емейство(ByVal r As Long) As Long
    If ¬—писке(r, FAM_TC_MOD) Or _
       ¬—писке(r, FAM_TC_BAR) Then

        —емейство = FAM_TC

    ElseIf ¬—писке(r, FAM_AI_MOD) Or _
           ¬—писке(r, FAM_AI_BAR) Then

        —емейство = FAM_AI

    ElseIf ¬—писке(r, FAM_DIO_REL) Or _
           ¬—писке(r, FAM_DIO_BAR) Then

        —емейство = FAM_DIO
    End If
End Function

Private Function ЁтоЅарьер(ByVal r As Long) As Boolean
    ЁтоЅарьер = ¬—писке(r, "47,48") Or _
                ¬—писке(r, FAM_TC_BAR) Or _
                ¬—писке(r, FAM_AI_BAR) Or _
                ¬—писке(r, FAM_DIO_BAR)
End Function

Private Function ќпределить“ип( _
    ByVal r As Long, ByVal nm As String) As Long

    If ¬—писке(r, ROWS_KDI) Then
        ќпределить“ип = TYP_DI
        Exit Function
    End If

    If ¬—писке(r, ROWS_KDO) Then
        ќпределить“ип = TYP_DO
        Exit Function
    End If

    If ¬—писке(r, ROWS_DIO) Then
        If InStr(1, nm, "32", vbTextCompare) > 0 Then
            ќпределить“ип = TYP_DIO32

        ElseIf InStr(1, nm, "64", vbTextCompare) > 0 Then
            ќпределить“ип = TYP_DIO64

        Else
            ќпределить“ип = TYP_DIO64

            gWarn = gWarn & vbCrLf & _
                "  Х —трока " & r & " (""" & nm & _
                """): не найдено 32/64; прин€то DIO64."
        End If
    End If
End Function

Private Function Ќормализовать»м€(ByVal nm As String) As String
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

    Ќормализовать»м€ = " " & Trim$(s) & " "
End Function

Private Function “ипЅарьераDIO( _
    ByVal r As Long, ByVal nm As String) As Long

    Dim s As String
    Dim hasDI As Boolean
    Dim hasDO As Boolean

    ' —троки 54Ц55 относ€тс€ к KDI.
    If r <> 53 Then
        “ипЅарьераDIO = TYP_DI
        Exit Function
    End If

    s = Ќормализовать»м€(nm)

    hasDI = (InStr(1, s, " KDI ") > 0 Or _
             InStr(1, s, " DI ") > 0 Or _
             InStr(1, s, " NAMUR") > 0 Or _
             InStr(1, s, " „ј—“") > 0 Or _
             InStr(1, s, " ¬’") > 0 Or _
             InStr(1, s, " FREQ") > 0)

    hasDO = (InStr(1, s, " KDO ") > 0 Or _
             InStr(1, s, " DO ") > 0 Or _
             InStr(1, s, " ¬џ’") > 0)

    If hasDI And Not hasDO Then
        “ипЅарьераDIO = TYP_DI
    ElseIf hasDO And Not hasDI Then
        “ипЅарьераDIO = TYP_DO
    Else
        ' Ќапример, ЂЋѕј-340-100ї: назначим экземпл€ры после чтени€ AA10.
        “ипЅарьераDIO = TYP_NONE
    End If
End Function

Private Sub –аспределитьќбщиеЅарьерыDIO(ByVal sigAA10 As Double)
    Dim i As Long
    Dim remainingDO As Double

    remainingDO = sigAA10

    ' ”читываем барьеры, €вно обозначенные как KDO.
    For i = 1 To gTotalMods
        If mIsBar(i) And mFam(i) = FAM_DIO Then
            If mBarType(i) = TYP_DO Then
                remainingDO = remainingDO - mCh(i)
            End If
        End If
    Next i

    ' ”ниверсальные барьеры строки 53: сначала под DO (AA10), остальные Ч под KDI.
    For i = 1 To gTotalMods
        If mIsBar(i) And mFam(i) = FAM_DIO Then
            If mBarType(i) = TYP_NONE Then
                If mRow(i) <> 53 Then
                    Err.Raise vbObjectError + 75, , _
                        "Ќе определЄн тип барьера DIO в строке " & mRow(i) & "."
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
' јЌјЋќ√ќ¬џ≈ Ѕј–№≈–џ: ®ћ ќ—“№ ћќƒ”Ћ≈… » –ј—ѕ–≈ƒ≈Ћ≈Ќ»≈
'=============================================================================

' ѕерва€ строка Service, где E = nm и D = категори€.
' ƒопускаетс€ 0 каналов (модуль не поддерживает это направление).
Private Function  аналыѕо атегории( _
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
    key = Ќорм»м€Service(nm)
    If Len(key) = 0 Then Exit Function

    lastRow = wsS.Cells(wsS.Rows.count, SVC_COL_NAME).End(xlUp).Row

    For r = 1 To lastRow
        nameValue = wsS.Cells(r, SVC_COL_NAME).Value2
        catValue = wsS.Cells(r, SVC_COL_CAT).Value2

        If Not IsError(nameValue) And Not IsError(catValue) Then
            If Ќорм»м€Service(CStr(nameValue)) = key And _
               Ќорм атегори€(CStr(catValue)) = cat Then

                v = wsS.Cells(r, SVC_COL_CH).Value2

                If Not IsError(v) And Not IsEmpty(v) Then
                    If IsNumeric(v) Then
                        d = CDbl(v)

                        If d >= 0 And d = Fix(d) Then
                             аналыѕо атегории = CLng(d)
                            found = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next r
End Function

Private Function Ќорм атегори€(ByVal s As String) As String
    s = Ќорм»м€Service(s)

    Do While InStr(1, s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    Ќорм атегори€ = s
End Function

Private Sub ®мкостьјналогћодул€( _
    ByVal wsS As Worksheet, ByVal nm As String, _
    ByVal rowIn1 As Long, _
    ByRef capAI As Long, ByRef capAO As Long)

    Dim fAI As Boolean
    Dim fAO As Boolean

    capAI =  аналыѕо атегории(wsS, nm, SVC_CAT_AI, fAI)
    capAO =  аналыѕо атегории(wsS, nm, SVC_CAT_AO, fAO)

    If Not fAI And Not fAO Then
        capAI = ANALOG_CH_DEFAULT
        capAO = 0

        gWarn = gWarn & vbCrLf & _
            "  Х """ & nm & """ (строка " & rowIn1 & _
            "): нет в Service с категорией module AI/AO; " & _
            "прин€то AI x" & ANALOG_CH_DEFAULT & " кан."
    End If
End Sub

' Ќаправление определ€етс€ только номером строки (им€ не используетс€).
Private Function ЌаправлениејналогЅарьера( _
    ByVal r As Long) As Long

    If ¬—писке(r, ROWS_BAR_AI_AO) Then
        ЌаправлениејналогЅарьера = DIR_AI_AO
    Else
        ' TC/RTD и RTD>AI относ€тс€ к входным каналам.
        ЌаправлениејналогЅарьера = DIR_AI
    End If
End Function

' Ќазначает свободные барьеры семейства fam модул€м этого семейства.
'  ритерий: минимальна€ загрузка (used + ch) / cap после назначени€.
' ѕри равенстве Ч модуль с меньшим числом барьеров, затем больший.
Private Sub –аспределитьјналоговыеЅарьеры( _
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
    Dim oldAI As Long
    Dim oldCnt As Long
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
    Else
        groupName = "AI, включа€ RTD>AI"
    End If

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
                            curCap = mCapAI(i) + mCapAO(i)

                        Case FAM_TC
                            ' mCapAI здесь означает температурные
                            ' входы. ќбщему AI/AO-барьеру доступны
                            ' только выходы AO этого модул€.
                            curCap = mCapAO(i)
                    End Select

                ElseIf mFam(i) = fam Then

                    curCap = mCapAI(i)
                End If

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
            "  Х " & groupName & _
            ": у модулей нет подход€щих каналов. " & _
            "Ѕарьеры (" & nBar & _
            " шт.) будут размещены отдельно."

        Exit Sub
    End If

    ' ”читываем ранее назначенные барьеры: в общем пуле AI/AO часть входов AI
    ' уже зан€та RTD>AI, а температурные барьеры выходы AO температурного
    ' модул€ не занимают. —трока 51 подключаетс€ только ко входам AI,
    ' поэтому еЄ вклад в общий пул ограничен mCapAI: выходы AO остаютс€ свободными,
    ' а превышение прохода 51 не дублируетс€ в overflow прохода 52.
    For k = 1 To nMod
        oldAI = 0
        oldCnt = 0

        For i = 1 To gTotalMods

            If mAnalogOwner(i) = mods(k) Then

                If dirType = DIR_AI_AO Then
                    countOld = (mFam(i) = FAM_AI)
                Else
                    countOld = (mFam(i) = fam And _
                                mAnDir(i) = dirType)
                End If

                If countOld Then
                    oldAI = oldAI + mCh(i)
                    oldCnt = oldCnt + 1
                End If
            End If
        Next i

        If dirType = DIR_AI_AO Then
            ' —вободный остаток дл€ строки 52 равен
            ' Max(0, mCapAI - usedAI) + mCapAO.
            If oldAI > mCapAI(mods(k)) Then
                oldAI = mCapAI(mods(k))
            End If
        End If

        usedA(k) = oldAI
        cntA(k) = oldCnt
        totalNeed = totalNeed + oldAI
    Next k

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

    ' Ќазначение барьеров: каждый физический барьер Ч ровно одному владельцу.
    For i = 1 To nBar
        ch = mCh(bars(i))
        best = 0
        bestScore = 1E+30

        ' —начала ищем назначение без превышени€ Ємкости.
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

        ' ≈сли назначить без превышени€ не удалось Ч условное распределение,
        ' после назначени€ будет выдано предупреждение.
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

    For k = 1 To nMod
        If usedA(k) > capA(k) Then
            overflow = overflow + usedA(k) - capA(k)
        End If
    Next k

    If overflow > 0 Then
        gWarn = gWarn & vbCrLf & _
            "  Х " & groupName & _
            ": учтено барьеров на " & _
            Format(totalNeed, "0") & " кан.; " & _
            "Ємкость модулей Ч " & _
            Format(totalCap, "0") & " кан. " & _
            "ѕревышение по отдельным модул€м: " & _
            Format(overflow, "0") & " кан. " & _
            "–азмещение выполнено условно; " & _
            "проверьте количество модулей и распределение каналов."
    End If
End Sub

Private Function Ќорм»м€Service(ByVal nm As String) As String
    Ќорм»м€Service = UCase$( _
        Trim$(Replace(nm, ChrW$(160), " ")))
End Function

Private Function Ўирина»зService( _
    ByVal wsS As Worksheet, ByVal nm As String, _
    ByVal fallback As Double, ByVal rowIn1 As Long) As Double

    Dim lastRow As Long
    Dim r As Long
    Dim key As String
    Dim v As Variant
    Dim nameValue As Variant

    key = Ќорм»м€Service(nm)

    If Len(key) > 0 Then
        lastRow = wsS.Cells( _
            wsS.Rows.count, SVC_COL_NAME).End(xlUp).Row

        For r = 1 To lastRow
            nameValue = wsS.Cells(r, SVC_COL_NAME).Value2

            If Not IsError(nameValue) Then
                If Ќорм»м€Service(CStr(nameValue)) = key Then
                    v = wsS.Cells(r, SVC_COL_WIDTH).Value2

                    If Not IsError(v) Then
                        If IsNumeric(v) Then
                            If CDbl(v) > 0 Then
                                Ўирина»зService = CDbl(v)
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        Next r
    End If

    gWarn = gWarn & vbCrLf & _
        "  Х """ & nm & """ (строка " & rowIn1 & _
        "): нет положительной ширины в Service!H; " & _
        "вз€то " & fallback & " мм."

    Ўирина»зService = fallback
End Function

Private Function  аналы»зService( _
    ByVal wsS As Worksheet, ByVal nm As String, _
    ByVal fallback As Long, ByVal rowIn1 As Long, _
    ByVal what As String) As Long

    Dim lastRow As Long
    Dim r As Long
    Dim key As String
    Dim v As Variant
    Dim nameValue As Variant
    Dim d As Double

    key = Ќорм»м€Service(nm)

    If Len(key) > 0 Then
        lastRow = wsS.Cells( _
            wsS.Rows.count, SVC_COL_NAME).End(xlUp).Row

        For r = 1 To lastRow
            nameValue = wsS.Cells(r, SVC_COL_NAME).Value2

            If Not IsError(nameValue) Then
                If Ќорм»м€Service(CStr(nameValue)) = key Then
                    v = wsS.Cells(r, SVC_COL_CH).Value2

                    If Not IsError(v) Then
                        If IsNumeric(v) Then
                            d = CDbl(v)

                            If d > 0 And d <= 2147483647# Then
                                If d = Fix(d) Then
                                     аналы»зService = CLng(d)
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
        "  Х """ & nm & """ (строка " & rowIn1 & _
        "): нет целого положительного числа каналов " & _
        "в Service!F; прин€то " & fallback & _
        " кан. на " & what & "."

     аналы»зService = fallback
End Function

Private Function „итать ол( _
    ByVal ws As Worksheet, ByVal r As Long) As Long

    Dim v As Variant
    Dim d As Double

    v = ws.Range(COL_QTY & r).Value2

    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function

    ' Ќечисловой текст разбираетс€ „исло»з“екста;
    ' сам факт разбора отмечает ѕроверитьячейку ол.
    If IsNumeric(v) Then
        d = Int(CDbl(v))
    Else
        d = Int(„исло»з“екста(CStr(v)))
    End If

    If d < 0 Then d = 0
    If d > 2147483647# Then d = 2147483647#

    „итать ол = CLng(d)
End Function

' —толбец D об€зан содержать число. ≈сли там текст, количество берЄтс€ из
' первой группы цифр Ч и об€зательно попадает в предупреждение, чтобы тиха€
' подмена (или тихий пропуск строки) не осталась незаметной. „исло, пуста€
' €чейка, пробелы и ошибка Ч молча; на строку одно сообщение.
Private Sub ѕроверитьячейку ол(ByVal ws As Worksheet, ByVal r As Long)
    Dim v As Variant
    Dim t As String
    Dim q As Long

    v = ws.Range(COL_QTY & r).Value2

    If IsError(v) Then Exit Sub
    If IsEmpty(v) Then Exit Sub
    If IsNumeric(v) Then Exit Sub

    t = Trim$(CStr(v))
    If Len(t) = 0 Then Exit Sub

    q = „итать ол(ws, r)

    If q > 0 Then
        gWarn = gWarn & vbCrLf & _
            "  Х —трока " & r & ", столбец " & COL_QTY & _
            ": количество записано текстом (""" & t & """); " & _
            "прин€то " & q & "."
    Else
        gWarn = gWarn & vbCrLf & _
            "  Х —трока " & r & ", столбец " & COL_QTY & _
            ": """ & t & """ Ч не число; строка не участвует в расчЄте."
    End If
End Sub

Private Function „итать„исло( _
    ByVal ws As Worksheet, ByVal addr As String) As Double

    Dim v As Variant

    v = ws.Range(addr).Value2

    If IsError(v) Or IsEmpty(v) Then Exit Function

    If IsNumeric(v) Then
        „итать„исло = CDbl(v)
    Else
        „итать„исло = „исло»з“екста(CStr(v))
    End If

    If „итать„исло < 0 Then „итать„исло = 0
End Function

Private Function ќчистить»м€( _
    ByVal s As String, ByVal r As Long) As String

    Dim p As Long
    Dim pref As String

    pref = "ј– —400."
    p = InStr(1, s, pref, vbTextCompare)

    If p > 0 Then
        ќчистить»м€ = Trim$(Mid$(s, p + Len(pref)))
    Else
        ќчистить»м€ = Trim$(s)
    End If

    If Len(ќчистить»м€) = 0 Then
        ќчистить»м€ = "Mod" & r
    End If
End Function

' „исло из текста. ѕравила (локаль не участвует: цифры накапливаютс€
' арифметически). ¬екторы сверки Ч tools/parse-spec.py:
'   - нормализаци€: NBSP U+00A0 и тонкие пробелы U+2009/U+200A/U+202F ->
'     обычный пробел, минус U+2212 -> Ђ-ї;
'   - Ђ-ї/Ђ+ї в начале строки или после пробела Ч знак; дефис внутри
'     наименовани€ (ЂLPA-340-100ї) знаком не считаетс€;
'   - одиночный пробел между цифрами Ч разделитель групп разр€дов, если
'     дальше ровно 3 цифры (Ђ1 500ї, Ђ1 500 000ї; перва€ группа Ч 1-3);
'   - Ђ,ї и Ђ.ї Ч всегда дес€тичный разделитель (Ђ1,500ї = 1,5);
'   - текст после числа игнорируетс€ (Ђ2000 ммї = 2000, Ђ5 шт 160ї = 5);
'   - если цифр нет Ч 0. –анние Exit Function возвращают 0 Ч значение
'     Double по умолчанию, €вного присваивани€ нет.
Private Function „исло»з“екста(ByVal s As String) As Double
    Dim n As Long
    Dim i As Long
    Dim c As String
    Dim d As Double
    Dim scale As Double
    Dim intDigits As Long
    Dim groups As Long
    Dim decSeen As Boolean
    Dim neg As Boolean
    Dim signOK As Boolean

    s = Replace(s, ChrW$(160), " ")
    s = Replace(s, ChrW$(8239), " ")
    s = Replace(s, ChrW$(8201), " ")
    s = Replace(s, ChrW$(8202), " ")
    s = Replace(s, ChrW$(8722), "-")

    n = Len(s)
    If n = 0 Then Exit Function

    i = 1

    Do While i <= n
        c = Mid$(s, i, 1)

        If c >= "0" And c <= "9" Then
            Exit Do
        End If

        If c = "-" Or c = "+" Then
            If i = 1 Then
                signOK = True
            Else
                signOK = (Mid$(s, i - 1, 1) = " ")
            End If

            If signOK And i < n Then
                If Mid$(s, i + 1, 1) >= "0" Then
                    If Mid$(s, i + 1, 1) <= "9" Then
                        neg = (c = "-")
                        i = i + 1
                        Exit Do
                    End If
                End If
            End If
        End If

        i = i + 1
    Loop

    If i > n Then Exit Function

    ' ÷ела€ часть и группы разр€дов через пробел.
    Do While i <= n
        c = Mid$(s, i, 1)

        If c >= "0" And c <= "9" Then
            d = d * 10# + CDbl(AscW(c) - 48)
            intDigits = intDigits + 1
            i = i + 1

        ElseIf c = " " Then
            If Not ≈сть√руппа–азр€дов(s, i, n, intDigits, groups) Then
                Exit Do
            End If

            i = i + 1
            d = d * 1000#
            d = d + CDbl(AscW(Mid$(s, i, 1)) - 48) * 100#
            d = d + CDbl(AscW(Mid$(s, i + 1, 1)) - 48) * 10#
            d = d + CDbl(AscW(Mid$(s, i + 2, 1)) - 48)
            intDigits = 0
            groups = groups + 1
            i = i + 3

        ElseIf c = "," Or c = "." Then
            decSeen = True
            i = i + 1
            Exit Do

        Else
            Exit Do
        End If
    Loop

    If decSeen Then
        scale = 10#

        Do While i <= n
            c = Mid$(s, i, 1)

            If c < "0" Or c > "9" Then Exit Do

            d = d + CDbl(AscW(c) - 48) / scale
            scale = scale * 10#
            i = i + 1
        Loop
    End If

    If neg Then d = -d

    „исло»з“екста = d
End Function

' ѕробел на позиции i раздел€ет группы разр€дов, если дальше идут ровно
' 3 цифры, не продолженные цифрой (перва€ группа Ч 1-3). –анние выходы
' возвращают False Ч значение Boolean по умолчанию.
Private Function ≈сть√руппа–азр€дов( _
    ByVal s As String, ByVal i As Long, ByVal n As Long, _
    ByVal intDigits As Long, ByVal groups As Long) As Boolean

    Dim d1 As String
    Dim d2 As String
    Dim d3 As String

    If i + 3 > n Then Exit Function

    d1 = Mid$(s, i + 1, 1)
    d2 = Mid$(s, i + 2, 1)
    d3 = Mid$(s, i + 3, 1)

    If d1 < "0" Or d1 > "9" Then Exit Function
    If d2 < "0" Or d2 > "9" Then Exit Function
    If d3 < "0" Or d3 > "9" Then Exit Function

    If i + 4 <= n Then
        If Mid$(s, i + 4, 1) >= "0" Then
            If Mid$(s, i + 4, 1) <= "9" Then Exit Function
        End If
    End If

    If groups = 0 Then
        ≈сть√руппа–азр€дов = (intDigits >= 1 And intDigits <= 3)
    Else
        ≈сть√руппа–азр€дов = True
    End If
End Function

' ÷елочисленный доступ к чтению €чейки: сначала IsNumeric и CDbl (ветка
' локально-зависима), иначе „исло»з“екста; дробна€ часть отбрасываетс€
' (Fix, без округлени€); выход за диапазон Long
' ограничиваетс€; ошибка или пуста€ €чейка Ч 0 без присваивани€.
Private Function »звлечь„исло(ByVal v As Variant) As Long
    Dim d As Double

    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function

    If IsNumeric(v) Then
        d = CDbl(v)
    Else
        d = „исло»з“екста(CStr(v))
    End If

    d = Fix(d)

    If d > 2147483647# Then
        d = 2147483647#
    ElseIf d < -2147483647# Then
        d = -2147483647#
    End If

    »звлечь„исло = CLng(d)
End Function

'=============================================================================
' ќ“–»—ќ¬ ј
'=============================================================================

Private Sub Ќарисовать–асположение( _
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

    ' ћасштаб отрисовки об€зан быть согласован с высотой верхней зоны:
    ' CTRL_ZONE_MM мм при MM px/мм дают ровно CTRL_OFFSET_PX пикселей, иначе
    ' полоса Ђрезервї и стартова€ позици€ модулей разойдутс€.
    If Abs(CTRL_OFFSET_PX - CTRL_ZONE_MM * MM) > 0.001 Then
        Err.Raise vbObjectError + 82, , _
            "—лужебна€ ошибка: масштаб отрисовки не согласован " & _
            "(CTRL_ZONE_MM = " & CTRL_ZONE_MM & " мм, MM = " & MM & _
            ", CTRL_OFFSET_PX = " & CTRL_OFFSET_PX & " px)."
    End If


    If cabCount < 1 Or railCount < 1 Then
        Err.Raise vbObjectError + 90, , _
            "Ќет шкафов дл€ отрисовки."
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

    ƒоб“екст wl, LEFTX, 5, 500, 15, _
        " ќЌ‘»√”–ј“ќ– –ј—ѕќЋќ∆≈Ќ»я ћќƒ”Ћ≈…", _
        14, True, RGB(0, 0, 0), msoAlignLeft

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

            ƒоб–ект wl, minX - 14, TOPY - 36, _
                (maxX + RAILW + 14) - (minX - 14), _
                maxH + 52, _
                RGB(248, 249, 251), _
                RGB(120, 130, 145), 1.5

            ƒоб“екст wl, minX - 14, TOPY - 56, _
                maxX - minX + RAILW + 28, 18, _
                "Ўкаф " & c & " Ч " & _
                IIf(kindC = "S", _
                    "отдельный барьерный/релейный", _
                    IIf(gTwoSided, _
                        "двухсторонний", _
                        "односторонний")), _
                10, True, RGB(20, 40, 80), _
                msoAlignLeft
        End If
    Next c

    For i = 1 To dN
        rr = drawList(i)
        hPx = rPhys(rr) * MM
        capPx = rPhys(rr) * (1 - gFreeFrac) * MM

        ƒоб–ект wl, railX(rr), TOPY, RAILW, hPx, _
            RGB(232, 235, 240), _
            RGB(150, 160, 175), 0.75

        If hPx - capPx > 1 Then
            Set rb = ƒоб–ект( _
                wl, railX(rr), TOPY + capPx, _
                RAILW, hPx - capPx, _
                RGB(255, 244, 204), _
                RGB(220, 180, 60), 0.5)

            rb.line.DashStyle = msoLineDash

            If hPx - capPx > 16 Then
                ƒоб“екст wl, railX(rr), _
                    TOPY + capPx, _
                    RAILW, hPx - capPx, "резерв", _
                    6, False, RGB(150, 110, 0), _
                    msoAlignCenter
            End If
        End If

        If rCtrl(rr) Then
            ƒоб–ект wl, railX(rr), TOPY, _
                RAILW, CTRL_BLOCK_PX, _
                RGB(46, 125, 50), _
                RGB(27, 70, 30), 0.75

            ƒоб“екст wl, railX(rr), TOPY, _
                RAILW, CTRL_BLOCK_PX, "P410", _
                8, True, RGB(255, 255, 255), _
                msoAlignCenter

        ElseIf rSwitch(rr) Then
            ƒоб–ект wl, railX(rr), TOPY, _
                RAILW, CTRL_BLOCK_PX, _
                RGB(21, 101, 192), _
                RGB(13, 60, 115), 0.75

            ƒоб“екст wl, railX(rr), TOPY, _
                RAILW, CTRL_BLOCK_PX, _
                " оммутаторы и Ѕѕ", _
                7.5, True, RGB(255, 255, 255), _
                msoAlignCenter
        End If

        ƒоб“екст wl, railX(rr), TOPY - 30, _
            RAILW, 28, _
            IIf(rSide(rr) = "F", _
                "‘ронтальна€", "«адн€€") & _
            " Ј –" & rSideIdx(rr) & vbLf & _
            Format(rUsed(rr), "0") & "/" & _
            Format(rCap(rr), "0") & " мм", _
            7, True, RGB(40, 40, 40), _
            msoAlignCenter
    Next i

    For idx = 1 To gTotalMods
        rr = mRail(idx)

        If rr >= 1 And rr <= railCount Then
            yOff = 0

            If rCtrl(rr) Or rSwitch(rr) Then
                yOff = CTRL_OFFSET_PX
            End If

            modH = mLen(idx) * MM

            Set msh = ƒоб–ект( _
                wl, railX(rr) + 1, _
                TOPY + yOff + mPos(idx) * MM, _
                RAILW - 2, modH, mColor(idx), _
                RGB(60, 60, 60), 0.25)

            If modH >= 13 Then
                txtCol = “екст÷вет(mColor(idx))

                ƒоб“екст wl, railX(rr) + 1, _
                    TOPY + yOff + mPos(idx) * MM, _
                    RAILW - 2, modH, mName(idx), _
                    6, False, txtCol, msoAlignCenter
            End If
        End If
    Next idx

    legX = LEFTX
    legY = TOPY + maxH + 50

    ƒоб“екст wl, legX, legY, 400, 18, _
        "Ћ≈√≈Ќƒј", 11, True, _
        RGB(0, 0, 0), msoAlignLeft

    legY = legY + 24

    For i = 1 To legN
        ƒоб–ект wl, legX, legY, 16, 14, _
            legColor(i), RGB(60, 60, 60), 0.5

        ƒоб“екст wl, legX + 22, legY - 2, _
            460, 18, _
            legName(i) & "   (" & _
            Format(legLen(i), "0") & " мм x " & _
            legQty(i) & " шт)", _
            8.5, False, RGB(30, 30, 30), _
            msoAlignLeft

        legY = legY + 19
    Next i

    legY = legY + 8

    ƒоб–ект wl, legX, legY, 16, 14, _
        RGB(255, 244, 204), _
        RGB(220, 180, 60), 0.5

    ƒоб“екст wl, legX + 22, legY - 2, _
        500, 18, _
        "—вободное место / резерв: " & _
        Format(gFreeFrac, "0%"), _
        8.5, True, RGB(120, 90, 0), _
        msoAlignLeft

    legY = legY + 22

    If ctrlCount > 0 Then
        ƒоб–ект wl, legX, legY, 16, 14, _
            RGB(46, 125, 50), _
            RGB(27, 70, 30), 0.5

        ƒоб“екст wl, legX + 22, legY - 2, _
            600, 18, " онтроллер P410", _
            8.5, True, RGB(20, 70, 30), _
            msoAlignLeft

        legY = legY + 22

        ƒоб–ект wl, legX, legY, 16, 14, _
            RGB(21, 101, 192), _
            RGB(13, 60, 115), 0.5

        ƒоб“екст wl, legX + 22, legY - 2, _
            600, 18, " оммутаторы и Ѕѕ", _
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

Private Function ѕалитра(ByVal i As Long) As Long
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

    ѕалитра = colors(((i - 1) Mod 20) + 1)
End Function

Private Function “екст÷вет(ByVal bg As Long) As Long
    Dim rr As Long
    Dim gg As Long
    Dim bb As Long
    Dim lum As Double

    rr = bg Mod 256
    gg = (bg \ 256) Mod 256
    bb = (bg \ 65536) Mod 256

    lum = 0.299 * rr + 0.587 * gg + 0.114 * bb

    If lum < 140 Then
        “екст÷вет = RGB(255, 255, 255)
    Else
        “екст÷вет = RGB(0, 0, 0)
    End If
End Function

Private Function ƒоб–ект( _
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

    Set ƒоб–ект = s
End Function

Private Sub ƒоб“екст( _
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

