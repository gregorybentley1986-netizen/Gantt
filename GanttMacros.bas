Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
    Private Declare PtrSafe Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hdc As LongPtr, ByVal lpsz As LongPtr, ByVal cbString As Long, ByRef lpSize As API_SIZE) As Long
    Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
    Private Declare PtrSafe Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectW" (ByRef lpLogFont As LOGFONTW) As LongPtr
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
#Else
    ' 32-bit VBA: Declare without PtrSafe. Editor shows red in 64-bit Office — this branch is not compiled then.
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
    Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, ByRef lpSize As API_SIZE) As Long
    Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
    Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (ByRef lpLogFont As LOGFONTA) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
#End If

Private Type API_SIZE
    cx As Long
    cy As Long
End Type

#If VBA7 Then
Private Type LOGFONTW
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(0 To 31) As Integer
End Type
#Else
Private Type LOGFONTA
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(0 To 31) As Byte
End Type
#End If

Private Const FW_NORMAL As Long = 400
Private Const FW_BOLD As Long = 700
Private Const LOGPIXELSY As Long = 90

Private Type timelineData
    PeriodStarts() As Date
    PeriodEnds() As Date
    PeriodLabels() As String
    PeriodWidths() As Double
    CumulativeWidths() As Double
    count As Long
    TotalWidth As Double
End Type

Private Type fontSettingsType
    name As String
    Size As Double
    Color As Long
    bold As Boolean
    italic As Boolean
End Type

Private Const SHEET_GANTT As String = "Gantt"
Private Const SHEET_DATA As String = "Данные"
Private Const SHEET_SETTINGS As String = "Настройки"
Private Const SHEET_REF As String = "Справочники"

Private Const TABLE_DATA As String = "ОснДанные"
Private Const TABLE_SETTINGS As String = "ОснНастройки"
Private Const TABLE_COLORS As String = "ЦветНастройки"
Private Const TABLE_FONTS As String = "НастройкаШрифтов"
Private Const TABLE_LINE_TYPES As String = "ТипыЛиний"
Private Const TABLE_EVENT_TYPES As String = "ТипыСобытий"
Private Const TABLE_HEADER As String = "Заголовок"

Private Const BUTTON_NAME As String = "btnCreateGantt"
Private Const BG_SHAPE_NAME As String = "ganttBackground"
Private Const GAP_1MM_PT As Double = 2.83465 ' 1 мм в пунктах

Public Sub SetupGanttSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_GANTT)

    EnsureGanttButton ws
End Sub

Public Sub CreateGanttDiagram()
    Dim wsGantt As Worksheet
    Dim wsData As Worksheet
    Dim wsSettings As Worksheet
    Dim wsRef As Worksheet
    Dim dataTable As ListObject
    Dim settingsTable As ListObject
    Dim colorsTable As ListObject
    Dim fontsTable As ListObject
    Dim lineTypesTable As ListObject
    Dim eventTypesTable As ListObject
    Dim startCell As Range
    Dim timelineStart As Date
    Dim timelineEnd As Date
    Dim periodName As String
    Dim dayWidth As Double
    Dim periodWidth As Double
    Dim rowHeight As Double
    Dim headerHeight As Double
    Dim periodHeaderHeight As Double
    Dim totalHeaderHeight As Double
    Dim timelineLengthMm As Double
    Dim timelineLengthPoints As Double
    Dim scaleFactor As Double
    Dim sidebarMode As String
    Dim sidebarGap As Double
    Dim sidebarWidth As Double
    Dim sectionColWidth As Double
    Dim taskColWidth As Double
    Dim timelineLeft As Double
    Dim diagramLeft As Double
    Dim rowIndex As Long
    Dim shapeTop As Double
    Dim shapeLeft As Double
    Dim barHeight As Double
    Dim barHeightNominal As Double
    Dim stackHeight As Double
    Dim startDate As Date
    Dim endDate As Date
    Dim taskName As String
    Dim sectionName As String
    Dim lineType As String
    Dim barColor As Long
    Dim barTransparency As Double
    Dim lineColor As Long
    Dim lineTransparency As Double
    Dim eventTransparency As Double
    Dim eventItems As Collection
    Dim ev As Variant
    Dim textOffsets() As Double
    Dim labelTiers() As Long
    Dim secondTierHeightForTask() As Double
    Dim evIdx As Long
    Dim lblOffset As Double
    Dim barShape As Shape
    Dim bgColor As Long
    Dim yearHeaderHeightMm As Double
    Dim yearHeaderHeightPoints As Double
    Dim yearHeaderColor As Long
    Dim periodHeaderColor As Long
    Dim timelineWidth As Double
    Dim diagramHeight As Double
    Dim legendHeight As Double
    Dim legendPadding As Double
    Dim diagramWidth As Double
    Dim timelineData As timelineData
    Dim i As Long
    Dim j As Long
    Dim g As Long
    Dim groupTop As Double
    Dim rowCount As Long
    Dim taskCount As Long
    Dim displayRowCount As Long
    Dim rotateAllPeriods As Boolean
    Dim sectionNames() As String
    Dim subsectionNames() As String
    Dim taskNames() As String
    Dim taskRowMap() As Long
    Dim groupLengths() As Long
    Dim groupNames() As String
    Dim groupTypes() As Long
    Dim groupRowIndices() As Long
    Dim taskRowIndices() As Long
    Dim taskStackIndices() As Long
    Dim taskStackCounts() As Long
    Dim sidebarTaskNames() As String
    Dim sidebarTaskRowIndices() As Long
    Dim groupCount As Long
    Dim sidebarColor As Long
    Dim sectionRowColor As Long
    Dim subsectionRowColor As Long
    Dim barLabel As String
    Dim hasEventLabel As Boolean
    Dim sidebarEnabled As Boolean
    Dim sidebarBoth As Boolean
    Dim vertDividerThickness As Double
    Dim horizDividerThickness As Double
    Dim vertDividerColor As Long
    Dim yearDividerColor As Long
    Dim horizDividerColor As Long
    Dim todayStripeMode As String
    Dim todayLabelMode As String
    Dim todayStripeThickness As Double
    Dim todayStripeColor As Long
    Dim pastEventsOverlayMode As String
    Dim pastEventsColor As Long
    Dim pastEventsTransparency As Double
    Dim bgTransparencyColor As Double
    Dim yearHeaderTransparency As Double
    Dim periodHeaderTransparency As Double
    Dim sidebarTransparency As Double
    Dim sectionTransparencyColor As Double
    Dim subsectionTransparencyColor As Double
    Dim vertDividerTransparency As Double
    Dim yearDividerTransparency As Double
    Dim horizDividerTransparency As Double
    Dim todayStripeTransparency As Double
    Dim sectionTransparencyFinal As Double
    Dim subsectionTransparencyFinal As Double
    Dim sectionTransparency As Double
    Dim taskMinHeight As Double
    Dim sectionHeightMm As Double
    Dim sectionHeightPoints As Double
    Dim subsectionHeightMm As Double
    Dim subsectionHeightPoints As Double
    Dim sidebarTextHeight As Double
    Dim maxStackHeight As Double
    Dim tmpBarHeight As Double
    Dim tmpStackHeight As Double
    Dim barGapMm As Double
    Dim barGapPoints As Double
    Dim barOffsetDownPoints As Double
    Dim sidebarWidthMm As Double
    Dim sidebarWidthPoints As Double
    Dim trackMinHeight As Double
    Dim trackPadMm As Double
    Dim trackPadPoints As Double
    Dim baseTrackHeight As Double
    Dim rowHeights() As Double
    Dim isSectionRow() As Boolean
    Dim rowTops() As Double
    Dim totalTasksHeight As Double
    Dim sidebarTextIndentPoints As Double
    Dim sidebarSectionTextIndentPoints As Double
    Dim sidebarSubsectionTextIndentPoints As Double
    Dim sidebarTextPaddingMm As Double
    Dim sidebarTextPaddingPoints As Double
    Dim yearDividerThickness As Double
    Dim usedLineTypes() As String
    Dim usedEventTypes() As String
    Dim usedLineTypeCount As Long
    Dim usedEventTypeCount As Long
    Dim maxLegendEventSize As Double
    Dim legendOffsetLeft As Double
    Dim legendMode As String
    Dim sectionRenderMode As String
    Dim barDisplayMode As String
    Dim taskBarShapeType As MsoAutoShapeType
    Dim drawSectionInTimeline As Boolean
    Dim availSidebar As Double
    Dim availTimeline As Double
    Dim availWidth As Double
    Dim linesCount As Long
    Dim showLegend As Boolean
    Dim taskBorderMode As String
    Dim eventBorderMode As String
    Dim eventLabelGapTopMm As Double
    Dim eventLabelGapTopPoints As Double
    Dim eventLabelGapBottomMm As Double
    Dim eventLabelGapBottomPoints As Double
    Dim eventLabelWidthMm As Double
    Dim eventLabelWidthPoints As Double
    Dim eventLabelMaxLines As Long
    Dim eventLabelReservePoints As Double
    Dim eventLabelReserveForTask() As Double
    Dim needGapAbove As Boolean
    Dim eventScanAvailW As Double
    Dim eventScanLineCount As Long
    Dim eventScanLblH As Double
    Dim eventScanGap As Double
    Dim showTaskBorder As Boolean
    Dim showEventBorder As Boolean
    Dim eventIdx As Long
    Dim eventName As String
    
    Dim eventLabelText As String
    Dim eventDateValue As Variant
    Dim eventDescValue As Variant
    Dim eventDate As Date
    Dim eventShapeType As MsoAutoShapeType
    Dim eventColor As Long
    Dim eventHeight As Double
    Dim eventWidth As Double
    Dim eventLeft As Double
    Dim eventTop As Double
    Dim eventColumnIndex As Long
    Dim eventDateColumnIndex As Long
    Dim eventDescColumnIndex As Long
    Dim eventColumnIndices(1 To 10) As Long
    
    Dim eventDescColumnIndices(1 To 10) As Long
    Dim eventDateColumnIndices(1 To 10) As Long
    Dim eventFound As Boolean
    Dim dataRowIndex As Long

    Application.ScreenUpdating = False

    Set wsGantt = GetWorksheetByName(SHEET_GANTT)
    Set wsData = GetWorksheetByNameOrTable(SHEET_DATA, TABLE_DATA)
    Set wsSettings = GetWorksheetByNameOrTable(SHEET_SETTINGS, TABLE_SETTINGS)
    Set wsRef = GetWorksheetByNameOrTable(SHEET_REF, TABLE_LINE_TYPES)

    If wsGantt Is Nothing Or wsData Is Nothing Or wsSettings Is Nothing Or wsRef Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "Не найдены листы (Gantt/Данные/Настройки/Справочники) или таблицы.", vbExclamation
        Exit Sub
    End If

    Set dataTable = wsData.ListObjects(TABLE_DATA)
    Set settingsTable = wsSettings.ListObjects(TABLE_SETTINGS)
    Set colorsTable = wsSettings.ListObjects(TABLE_COLORS)
    On Error Resume Next
    Set fontsTable = wsSettings.ListObjects(TABLE_FONTS)
    On Error GoTo 0
    Set lineTypesTable = wsRef.ListObjects(TABLE_LINE_TYPES)
    On Error Resume Next
    Set eventTypesTable = wsRef.ListObjects(TABLE_EVENT_TYPES)
    On Error GoTo 0

    Dim fontYears As fontSettingsType
    Dim fontPeriods As fontSettingsType
    Dim fontSections As fontSettingsType
    Dim fontTasks As fontSettingsType
    Dim fontEventDesc As fontSettingsType
    Dim fontLegend As fontSettingsType
    fontYears = GetFontSettingsFromTable(fontsTable, "Годы", "Calibri", 25, RGB(0, 0, 0), False, True)
    fontPeriods = GetFontSettingsFromTable(fontsTable, "Периоды", "Arial", 20, RGB(0, 0, 0), False, False)
    fontSections = GetFontSettingsFromTable(fontsTable, "Разделы", "Century Gothic", 15, RGB(0, 0, 0), False, False)
    fontTasks = GetFontSettingsFromTable(fontsTable, "Задачи", "Century Gothic", 12, RGB(0, 0, 0), True, False)
    fontEventDesc = GetFontSettingsFromTable(fontsTable, "Описание событий", "Arial", 8, RGB(0, 0, 0), False, False)
    fontLegend = GetFontSettingsFromTable(fontsTable, "Легенда", "Times New Roman", 12, RGB(0, 0, 0), True, True)
    Dim fontHeader As fontSettingsType
    fontHeader = GetFontSettingsFromTable(fontsTable, "Заголовок", "Arial", 18, RGB(0, 0, 0), True, False)
    Dim fontSubsections As fontSettingsType
    fontSubsections = GetFontSettingsFromTable(fontsTable, "Подразделы", "Times New Roman", 15, RGB(0, 0, 0), True, True)
    Dim fontTodayLabel As fontSettingsType
    fontTodayLabel = GetFontSettingsFromTable(fontsTable, "Надпись Сегодня", "Calibri", 9, RGB(0, 0, 0), False, False)

    Set startCell = wsGantt.Range("E1")

    timelineStart = GetSettingValue(settingsTable, "Дата начала таймлайна", Date)
    timelineEnd = GetSettingValue(settingsTable, "Дата окончания таймлайна", Date)
    periodName = CStr(GetSettingValue(settingsTable, "Период", "Дни"))

    dayWidth = GetSettingValue(settingsTable, "Ширина дня", 18)
    periodWidth = GetSettingValue(settingsTable, "Ширина периода", 60)
    rowHeight = GetSettingValue(settingsTable, "Высота строки", 5)
    headerHeight = GetSettingValue(settingsTable, "Высота шапки", rowHeight)
    yearHeaderHeightMm = GetSettingValue(settingsTable, "Высота строки с годами, мм", 6)
    If IsNumeric(yearHeaderHeightMm) And CDbl(yearHeaderHeightMm) > 0 Then
        yearHeaderHeightPoints = CDbl(yearHeaderHeightMm) * 2.83465
    Else
        yearHeaderHeightPoints = 6 * 2.83465
    End If
    timelineLengthMm = GetSettingValue(settingsTable, "Длина таймлайна, мм", 0)
    taskMinHeight = GetSettingValue(settingsTable, "Высота бара задачи, мм", 0)
    If taskMinHeight > 0 Then
        taskMinHeight = taskMinHeight * 2.83465
    End If
    barHeightNominal = taskMinHeight
    If barHeightNominal <= 0 Then
        barHeightNominal = rowHeight * 0.6
    End If
    sidebarMode = CStr(GetSettingValue(settingsTable, "Сайдбар", "Отключен"))
    sidebarGap = 0
    vertDividerThickness = GetSettingValue(settingsTable, "Толщина вертикальных разделителей, пт", 0.75)
    yearDividerThickness = GetSettingValue(settingsTable, "Толщина вертикальных разделителей между годами, пт", _
                                           vertDividerThickness)
    horizDividerThickness = GetSettingValue(settingsTable, "Толщина горизонтальных разделителей, пт", 0.75)
    todayStripeMode = Trim$(CStr(GetSettingValue(settingsTable, "Полоса Сегодня", "Выключить")))
    todayLabelMode = Trim$(CStr(GetSettingValue(settingsTable, "Надпись Сегодня", "Выключить")))
    todayStripeThickness = GetSettingValue(settingsTable, "Толщина полосы Сегодня, пт", 1.5)
    pastEventsOverlayMode = CStr(GetSettingValue(settingsTable, "Включить закрашивание прошедших событий", "Нет"))
    sectionTransparency = GetSettingValue(settingsTable, "Прозрачность разделов", 0)
    legendMode = CStr(GetSettingValue(settingsTable, "Легенда", "Включить"))
    taskBorderMode = CStr(GetSettingValue(settingsTable, "Отображение рамки в задачах", "Включить"))
    eventBorderMode = CStr(GetSettingValue(settingsTable, "Отображение рамки в фигурах", "Включить"))
    eventLabelGapTopMm = GetSettingValue(settingsTable, "Верхний отступ от текста описания события, мм", 0)
    If eventLabelGapTopMm >= 0 Then
        eventLabelGapTopPoints = eventLabelGapTopMm * 2.83465
    Else
        eventLabelGapTopPoints = 0
    End If
    eventLabelGapBottomMm = GetSettingValue(settingsTable, "Нижний отступ от текста описания события, мм", 1)
    If eventLabelGapBottomMm >= 0 Then
        eventLabelGapBottomPoints = eventLabelGapBottomMm * 2.83465
    Else
        eventLabelGapBottomPoints = 2.83465
    End If
    eventLabelWidthMm = CDbl(GetSettingValue(settingsTable, "Ширина описания события, мм", 40))
    If eventLabelWidthMm > 0 Then
        eventLabelWidthPoints = eventLabelWidthMm * 2.83465
    Else
        eventLabelWidthPoints = 40 * 2.83465
    End If
    eventLabelMaxLines = CLng(GetSettingValue(settingsTable, "Количество отображаемых строк в событии", 3))
    If eventLabelMaxLines < 1 Then eventLabelMaxLines = 1
    sidebarTextIndentPoints = GetSettingValue(settingsTable, "Отступ текста задачи от левого края в сайдбаре, мм", 0)
    If sidebarTextIndentPoints > 0 Then
        sidebarTextIndentPoints = sidebarTextIndentPoints * 2.83465
    Else
        sidebarTextIndentPoints = 4
    End If
    sidebarTextPaddingMm = GetSettingValue(settingsTable, "Вертикальные отступы от текста в задачах, мм", 1.5)
    If sidebarTextPaddingMm >= 0 Then
        sidebarTextPaddingPoints = sidebarTextPaddingMm * 2.83465
    Else
        sidebarTextPaddingPoints = 0
    End If
    sidebarWidthMm = GetSettingValue(settingsTable, "Ширина сайдбара, мм", 0)
    If sidebarWidthMm > 0 Then
        sidebarWidthPoints = sidebarWidthMm * 2.83465
    End If
    trackMinHeight = GetSettingValue(settingsTable, "Высота трека задачи (min), мм", 0)
    If trackMinHeight > 0 Then
        trackMinHeight = trackMinHeight * 2.83465
    End If
    trackPadMm = GetSettingValue(settingsTable, "Минимальный отступ бара от границ трека, мм", 0)
    If trackPadMm > 0 Then
        trackPadPoints = trackPadMm * 2.83465
    Else
        trackPadPoints = 0
    End If

    sectionRenderMode = CStr(GetSettingValue(settingsTable, "Режим отображения раздела", "Сайдбар+таймлайн"))
    barDisplayMode = CStr(GetSettingValue(settingsTable, "Вид отображения баров", "Уменьшение высоты"))
    If LCase$(Trim$(CStr(GetSettingValue(settingsTable, "Режим отображения бара задачи", "Прямоугольник со скругленными углами")))) = "стрелка вправо" Then
        taskBarShapeType = msoShapePentagon
    Else
        taskBarShapeType = msoShapeRoundedRectangle
    End If
    drawSectionInTimeline = (LCase$(sectionRenderMode) = "сайдбар+таймлайн")

    sectionHeightMm = GetSettingValue(settingsTable, "Высота раздела, мм", 0)
    If sectionHeightMm > 0 Then
        sectionHeightPoints = sectionHeightMm * 2.83465
    Else
        sectionHeightPoints = 0
    End If
    subsectionHeightMm = GetSettingValue(settingsTable, "Высота подраздела, мм", 0)
    If subsectionHeightMm > 0 Then
        subsectionHeightPoints = subsectionHeightMm * 2.83465
    Else
        subsectionHeightPoints = sectionHeightPoints
    End If
    sidebarSectionTextIndentPoints = GetSettingValue(settingsTable, "Отступ текста раздела от левого края в сайдбаре, мм", 0)
    If sidebarSectionTextIndentPoints > 0 Then
        sidebarSectionTextIndentPoints = sidebarSectionTextIndentPoints * 2.83465
    Else
        sidebarSectionTextIndentPoints = sidebarTextIndentPoints
    End If
    sidebarSubsectionTextIndentPoints = GetSettingValue(settingsTable, "Отступ текста подраздела от левого края в сайдбаре, мм", 0)
    If sidebarSubsectionTextIndentPoints > 0 Then
        sidebarSubsectionTextIndentPoints = sidebarSubsectionTextIndentPoints * 2.83465
    Else
        sidebarSubsectionTextIndentPoints = sidebarSectionTextIndentPoints
    End If
    showLegend = LCase$(legendMode) = "включить"
    showTaskBorder = LCase$(taskBorderMode) <> "выключить"
    showEventBorder = LCase$(eventBorderMode) <> "выключить"
    eventLabelReservePoints = eventLabelMaxLines * EstimateTextHeight(8) + (eventLabelMaxLines - 1) * 1 + 2 + eventLabelGapTopPoints + eventLabelGapBottomPoints

    ClearGanttShapes wsGantt
    Application.CutCopyMode = False
    DoEvents
    EnsureGanttButton wsGantt

    If dataTable.DataBodyRange Is Nothing Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    Dim subsectionColIdx As Long
    subsectionColIdx = GetTableColumnIndex(dataTable, "Подраздел")
    rowCount = dataTable.ListRows.count
    If rowCount > 0 Then
        ReDim sectionNames(1 To rowCount)
        ReDim subsectionNames(1 To rowCount)
        ReDim taskNames(1 To rowCount)
        ReDim taskRowMap(1 To rowCount)
    End If

    taskCount = 0
    ' Collect events to draw after all bars (so events are always on top of bars)
    Set eventItems = New Collection
    Dim barItems As Collection
    Set barItems = New Collection

    For i = 1 To rowCount
        taskName = Trim$(CStr(dataTable.ListColumns("Задача").DataBodyRange.Cells(i, 1).value))
        If Len(taskName) > 0 Then
            taskCount = taskCount + 1
            sectionNames(taskCount) = Trim$(CStr(dataTable.ListColumns("Раздел").DataBodyRange.Cells(i, 1).value))
            If subsectionColIdx > 0 Then
                subsectionNames(taskCount) = Trim$(CStr(dataTable.ListColumns("Подраздел").DataBodyRange.Cells(i, 1).value))
            Else
                subsectionNames(taskCount) = ""
            End If
            taskNames(taskCount) = taskName
            taskRowMap(taskCount) = i
        End If
    Next i

    If taskCount > 0 Then
        ReDim Preserve sectionNames(1 To taskCount)
        ReDim Preserve subsectionNames(1 To taskCount)
        ReDim Preserve taskNames(1 To taskCount)
        ReDim Preserve taskRowMap(1 To taskCount)
    Else
        Erase sectionNames
        Erase subsectionNames
        Erase taskNames
        Erase taskRowMap
    End If

    rowCount = taskCount

    For eventIdx = 1 To 10
        eventColumnIndices(eventIdx) = GetTableColumnIndex(dataTable, "Событие " & eventIdx)
        eventDescColumnIndices(eventIdx) = GetTableColumnIndex(dataTable, "Описание события " & eventIdx)
        eventDateColumnIndices(eventIdx) = GetTableColumnIndex(dataTable, "Дата " & eventIdx)
    Next eventIdx

    sidebarEnabled = LCase$(sidebarMode) <> "отключен"
    sidebarBoth = LCase$(sidebarMode) = "с обеих сторон"

    If rowCount > 0 Then
        If sidebarEnabled Then
            BuildSidebarLayoutWithDuplicates sectionNames, subsectionNames, subsectionColIdx, taskNames, groupRowIndices, groupLengths, groupNames, groupTypes, _
                groupCount, taskRowIndices, taskStackIndices, taskStackCounts, sidebarTaskNames, _
                sidebarTaskRowIndices, displayRowCount
        Else
            BuildSidebarLayout sectionNames, subsectionNames, subsectionColIdx, taskNames, groupRowIndices, groupLengths, groupNames, groupTypes, groupCount, _
                taskRowIndices, displayRowCount
            ' When sidebar disabled: no stacking, each task gets own row
            ReDim taskStackIndices(1 To rowCount)
            ReDim taskStackCounts(1 To rowCount)
            For i = 1 To rowCount
                taskStackIndices(i) = 1
                taskStackCounts(i) = 1
            Next i
        End If
    End If

    If sidebarEnabled Then
        If rowCount > 0 Then
            ComputeSidebarWidths sectionNames, subsectionNames, groupNames, groupTypes, groupCount, fontSections, fontSubsections, _
                sidebarTaskNames, CLng(fontTasks.Size), 6, sectionColWidth, taskColWidth
        End If
        If sidebarWidthPoints > 0 Then
            taskColWidth = sidebarWidthPoints
        End If
        sectionColWidth = 0
        sidebarWidth = taskColWidth
    Else
        sidebarWidth = 0
        sectionColWidth = 0
        taskColWidth = 0
        If displayRowCount = 0 Then
            displayRowCount = rowCount
        End If
    End If

    barGapMm = GetSettingValue(settingsTable, "Зазор между барами в задаче, мм", 0)
    If barGapMm > 0 Then
        barGapPoints = barGapMm * 2.83465
    Else
        barGapPoints = 0
    End If
    Dim barOffsetDownMm As Double
    barOffsetDownMm = GetSettingValue(settingsTable, "Сдвиг бара вниз, мм", 0)
    If barOffsetDownMm > 0 Then
        barOffsetDownPoints = barOffsetDownMm * 2.83465
    Else
        barOffsetDownPoints = 0
    End If
    baseTrackHeight = Application.WorksheetFunction.Max(rowHeight, barHeightNominal + 2 * trackPadPoints)
    If trackMinHeight > 0 And trackMinHeight > baseTrackHeight Then
        baseTrackHeight = trackMinHeight
    End If

    If displayRowCount > 0 Then
        ReDim rowHeights(1 To displayRowCount)
        ReDim isSectionRow(1 To displayRowCount)
        If groupCount > 0 Then
            For i = 1 To groupCount
                If groupRowIndices(i) >= 1 And groupRowIndices(i) <= displayRowCount Then
                    isSectionRow(groupRowIndices(i)) = True
                End If
            Next i
        End If

        
        For i = 1 To displayRowCount
            rowHeights(i) = baseTrackHeight
        Next i
    End If


    '--- Apply section/subsection row height from settings (like tasks)
' Settings: "Высота раздела, мм", "Высота подраздела, мм"
If groupCount > 0 And displayRowCount > 0 Then
    For i = 1 To groupCount
        If groupRowIndices(i) >= 1 And groupRowIndices(i) <= displayRowCount Then
            If i <= UBound(groupTypes) And groupTypes(i) = 2 Then
                If subsectionHeightPoints > 0 Then
                    rowHeights(groupRowIndices(i)) = subsectionHeightPoints
                Else
                    rowHeights(groupRowIndices(i)) = baseTrackHeight
                End If
            Else
                If sectionHeightPoints > 0 Then
                    rowHeights(groupRowIndices(i)) = sectionHeightPoints
                Else
                    rowHeights(groupRowIndices(i)) = baseTrackHeight
                End If
            End If
        End If
    Next i
End If

If sidebarEnabled And rowCount > 0 Then
        For i = LBound(sidebarTaskNames) To UBound(sidebarTaskNames)
            rowIndex = sidebarTaskRowIndices(i)
            If rowIndex < 1 Or rowIndex > displayRowCount Then GoTo NextSidebarRow
            If isSectionRow(rowIndex) Then GoTo NextSidebarRow

            sidebarTextHeight = GetSidebarRowTextHeight(sidebarTaskNames(i), taskColWidth, fontTasks, sidebarTextIndentPoints, sidebarTextPaddingPoints)
            If sidebarTextHeight > rowHeights(rowIndex) Then
                rowHeights(rowIndex) = sidebarTextHeight
            End If

NextSidebarRow:
        Next i
End If


    maxStackHeight = barHeightNominal
    ' === Stack row height depends on bar display mode ===
    Select Case LCase$(barDisplayMode)
        Case "одинаковая высота"
            For i = 1 To rowCount
                If taskStackCounts(i) > 1 Then
                    maxStackHeight = (barHeightNominal * taskStackCounts(i)) + (barGapPoints * (taskStackCounts(i) - 1))
                    If maxStackHeight + 2 * trackPadPoints > rowHeights(taskRowIndices(i)) Then
                        rowHeights(taskRowIndices(i)) = maxStackHeight + 2 * trackPadPoints
                    End If
                End If
            Next i

        Case Else
            ' "Уменьшение высоты" and "В одну линию"
            For i = 1 To rowCount
                If taskStackCounts(i) > 1 Then
                    maxStackHeight = barHeightNominal

                    ' If user specified a gap, the stack may exceed barHeightNominal when minimum bar height is reached.
                    ' In that case we expand the row height to avoid overlaps/clipping.
                    If (LCase$(barDisplayMode) = "уменьшение высоты") And (barGapPoints > 0) Then
                        tmpBarHeight = (barHeightNominal - (barGapPoints * (taskStackCounts(i) - 1))) / taskStackCounts(i)
                        If tmpBarHeight < 1 Then tmpBarHeight = 1
                        tmpStackHeight = (tmpBarHeight * taskStackCounts(i)) + (barGapPoints * (taskStackCounts(i) - 1))
                        If tmpStackHeight > maxStackHeight Then maxStackHeight = tmpStackHeight
                    End If
                lineTransparency = GetLineTypeTransparency(lineTypesTable, lineType, 0)

                    If maxStackHeight + 2 * trackPadPoints > rowHeights(taskRowIndices(i)) Then
                        rowHeights(taskRowIndices(i)) = maxStackHeight + 2 * trackPadPoints
                    End If
                End If
            Next i
    End Select

    ReDim eventLabelReserveForTask(1 To Application.WorksheetFunction.Max(rowCount, 1))
    For i = 1 To UBound(eventLabelReserveForTask)
        eventLabelReserveForTask(i) = 0
    Next i
    If displayRowCount > 0 And rowCount > 0 Then
    For i = 1 To rowCount
        If sidebarEnabled Then
            rowIndex = taskRowIndices(i)
        ElseIf groupCount > 0 Then
            rowIndex = taskRowIndices(i)
        Else
            rowIndex = i
        End If
        If rowIndex < 1 Or rowIndex > displayRowCount Then GoTo NextEventScan
        If isSectionRow(rowIndex) Then GoTo NextEventScan
        dataRowIndex = taskRowMap(i)
        For eventIdx = 1 To 10
            eventColumnIndex = eventColumnIndices(eventIdx)
            eventDateColumnIndex = eventDateColumnIndices(eventIdx)
            eventDescColumnIndex = eventDescColumnIndices(eventIdx)
            If eventColumnIndex > 0 And eventDateColumnIndex > 0 And eventDescColumnIndex > 0 Then
                eventName = Trim$(CStr(dataTable.ListColumns(eventColumnIndex).DataBodyRange.Cells(dataRowIndex, 1).value))
                eventDateValue = dataTable.ListColumns(eventDateColumnIndex).DataBodyRange.Cells(dataRowIndex, 1).value
                eventDescValue = dataTable.ListColumns(eventDescColumnIndex).DataBodyRange.Cells(dataRowIndex, 1).value
                eventLabelText = Trim$(CStr(eventDescValue))
                If Len(eventName) > 0 And IsDate(eventDateValue) And Len(eventLabelText) > 0 Then
                    eventDate = CDate(eventDateValue)
                    If eventDate >= timelineStart And eventDate <= timelineEnd Then
                        eventScanAvailW = eventLabelWidthPoints
                        eventScanLineCount = EstimateWrappedLinesForFont(eventLabelText, eventScanAvailW, fontEventDesc)
                        If eventScanLineCount < 1 Then eventScanLineCount = 1
                        If eventScanLineCount > eventLabelMaxLines Then eventScanLineCount = eventLabelMaxLines
                        eventScanLblH = eventScanLineCount * EstimateTextHeight(CLng(fontEventDesc.Size)) + (eventScanLineCount - 1) * 1 + 2
                        eventScanGap = eventLabelGapTopPoints + eventLabelGapBottomPoints
                        eventLabelReserveForTask(i) = Application.WorksheetFunction.Max(eventLabelReserveForTask(i), _
                            eventScanLblH + eventScanGap)
                    End If
                End If
            End If
        Next eventIdx
NextEventScan:
    Next i
    End If

    timelineData = BuildTimeline(timelineStart, timelineEnd, periodName, dayWidth, periodWidth)
    timelineWidth = timelineData.TotalWidth

    If timelineData.count = 0 Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    If timelineLengthMm > 0 Then
        timelineLengthPoints = timelineLengthMm * 2.83465
        If timelineWidth > 0 Then
            scaleFactor = timelineLengthPoints / timelineWidth
            ScaleTimeline timelineData, scaleFactor
            timelineWidth = timelineData.TotalWidth
        End If
    End If

    NormalizeTimelinePeriodWidths timelineData
    timelineWidth = timelineData.TotalWidth

    ' Если в строке три и более событий с подписями и среднее перекрывает соседей — добавляем высоту под второй ярус (1 мм от верха первого уровня до низа второго)
    ' Сохраняем исходные значения резерва первого уровня для правильного расчета tierOffsetPoints
    Dim firstTierReserveForTask() As Double
    Dim maxFirstTierLabelHeightForTask() As Double ' Объявляем здесь, чтобы была доступна в цикле добавления резерва
    If rowCount > 0 Then
        ReDim secondTierHeightForTask(1 To rowCount)
        ReDim firstTierReserveForTask(1 To rowCount)
        ReDim maxFirstTierLabelHeightForTask(1 To rowCount)
        Dim r As Long
        For r = 1 To rowCount
            secondTierHeightForTask(r) = 0
            firstTierReserveForTask(r) = eventLabelReserveForTask(r) ' Сохраняем исходное значение
            maxFirstTierLabelHeightForTask(r) = 0 ' Инициализируем
        Next r
    End If
    AddSecondTierReserveForOverlappingMiddle dataTable, taskRowMap, taskRowIndices, taskStackIndices, rowCount, displayRowCount, _
        isSectionRow, eventColumnIndices, eventDateColumnIndices, eventDescColumnIndices, timelineData, _
        timelineStart, timelineEnd, rowHeights, eventLabelReserveForTask, eventLabelWidthPoints, _
        eventLabelReservePoints, eventLabelMaxLines, fontEventDesc, secondTierHeightForTask
    
    ' НОВАЯ ЛОГИКА: Высота трека складывается снизу вверх для каждого бара в стеке
    ' Выполняем ПОСЛЕ AddSecondTierReserveForOverlappingMiddle
    If displayRowCount > 0 And rowCount > 0 Then
        ' Проходим по строкам отображения
        Dim rowIdx As Long
        For rowIdx = 1 To displayRowCount
            If isSectionRow(rowIdx) Then GoTo NextRowHeightNew
            
            ' Находим все задачи в этой строке и определяем количество баров в стеке
            Dim stackCountForRow As Long
            stackCountForRow = 0
            Dim taskIdxInRow As Long
            For taskIdxInRow = 1 To rowCount
                Dim taskRowIdx As Long
                If sidebarEnabled Then
                    taskRowIdx = taskRowIndices(taskIdxInRow)
                ElseIf groupCount > 0 Then
                    taskRowIdx = taskRowIndices(taskIdxInRow)
                Else
                    taskRowIdx = taskIdxInRow
                End If
                If taskRowIdx = rowIdx Then
                    If taskStackIndices(taskIdxInRow) > stackCountForRow Then
                        stackCountForRow = taskStackIndices(taskIdxInRow)
                    End If
                End If
            Next taskIdxInRow
            If stackCountForRow = 0 Then stackCountForRow = 1
            
            ' Вычисляем высоту трека для этой строки
            Dim trackHeightForRow As Double
            trackHeightForRow = 0
            
            ' Проходим по всем барам в стеке снизу вверх (от нижнего к верхнему)
            Dim stackPos As Long
            For stackPos = stackCountForRow To 1 Step -1
                ' Находим задачу на этой позиции стека
                Dim taskIdxAtPos As Long
                taskIdxAtPos = FindTaskAtStackPos(rowCount, taskRowIndices, taskStackIndices, rowIdx, stackPos)
                If taskIdxAtPos <= 0 Then GoTo NextStackPos
                
                ' Вычисляем высоту бара на позиции stackPos
                Dim barH As Double
                If LCase$(barDisplayMode) = "уменьшение высоты" And stackCountForRow > 1 Then
                    barH = (barHeightNominal - (barGapPoints * (stackCountForRow - 1))) / stackCountForRow
                    If barH < 1 Then barH = 1
                Else
                    barH = barHeightNominal
                End If
                
                ' Определяем высоту первого яруса для этого бара
                Dim firstTierHeight As Double
                firstTierHeight = 0
                If taskIdxAtPos <= UBound(maxFirstTierLabelHeightForTask) And maxFirstTierLabelHeightForTask(taskIdxAtPos) > 0 Then
                    firstTierHeight = maxFirstTierLabelHeightForTask(taskIdxAtPos)
                ElseIf taskIdxAtPos <= UBound(firstTierReserveForTask) And firstTierReserveForTask(taskIdxAtPos) > 0 Then
                    ' Fallback: используем резерв первого яруса минус отступы
                    firstTierHeight = firstTierReserveForTask(taskIdxAtPos) - eventLabelGapTopPoints - eventLabelGapBottomPoints
                    If firstTierHeight < 0 Then firstTierHeight = 0
                End If
                
                ' Определяем высоту второго яруса для этого бара
                Dim secondTierHeight As Double
                secondTierHeight = 0
                If taskIdxAtPos <= UBound(secondTierHeightForTask) Then
                    secondTierHeight = secondTierHeightForTask(taskIdxAtPos)
                End If
                
                ' Добавляем элементы снизу вверх:
                ' 1. Для нижнего бара: зазор между баром и нижней границей трека
                If stackPos = stackCountForRow Then
                    trackHeightForRow = trackHeightForRow + trackPadPoints
                End If
                
                ' 2. Высота бара
                trackHeightForRow = trackHeightForRow + barH
                
                ' 3. Если есть первый ярус: зазор 1мм между баром и первым ярусом + высота первого яруса
                If firstTierHeight > 0 Then
                    trackHeightForRow = trackHeightForRow + GAP_1MM_PT + firstTierHeight
                End If
                
                ' 4. Если есть второй ярус: зазор 1мм между первым и вторым ярусом + высота второго яруса
                If secondTierHeight > 0 Then
                    trackHeightForRow = trackHeightForRow + GAP_1MM_PT + secondTierHeight
                End If
                
                ' 5. Зазор между барами: добавляется после каждого бара (кроме верхнего)
                ' Зазор должен быть сразу после текущего бара и его ярусов, перед следующим баром выше.
                ' Если у текущего бара есть описание (ярусы), используем не меньше trackPadPoints (настройка "Минимальный отступ бара от границ трека, мм").
                If stackPos > 1 Then
                    Dim gapBetweenBars As Double
                    gapBetweenBars = barGapPoints
                    If (firstTierHeight > 0 Or secondTierHeight > 0) And trackPadPoints > gapBetweenBars Then
                        gapBetweenBars = trackPadPoints
                    End If
                    trackHeightForRow = trackHeightForRow + gapBetweenBars
                End If
                
                ' 6. Для верхнего бара: всегда отступ от верхней границы трека — настройка "Минимальный отступ бара от границ трека, мм"
                If stackPos = 1 Then
                    trackHeightForRow = trackHeightForRow + trackPadPoints
                End If
NextStackPos:
            Next stackPos
            
            ' Устанавливаем высоту строки как максимум из базовой высоты и вычисленной высоты трека
            If trackHeightForRow > rowHeights(rowIdx) Then
                rowHeights(rowIdx) = trackHeightForRow
            End If
NextRowHeightNew:
        Next rowIdx
    End If

    diagramLeft = startCell.Left
    timelineLeft = diagramLeft
    If LCase$(sidebarMode) = "слева" Or sidebarBoth Then
        timelineLeft = diagramLeft + sidebarWidth + sidebarGap
    End If

    rotateAllPeriods = ShouldRotateAllPeriodLabels(timelineData, CLng(fontPeriods.Size))
    periodHeaderHeight = GetPeriodHeaderHeight(timelineData, headerHeight, CLng(fontPeriods.Size), rotateAllPeriods)
    ' Extra vertical padding to avoid month label clipping
    periodHeaderHeight = periodHeaderHeight + (2 * 2.83465)

    ' Заголовок диаграммы: текст из таблицы Заголовок
    Dim headerTitleText As String
    Dim headerTitleColor As Long
    Dim headerTitleTransparency As Double
    Dim titleHeight As Double
    Dim baseTitleHeight As Double
    Dim availTitleWidth As Double
    Dim titleLines As Long
    Dim neededTitleHeight As Double
    Const TITLE_PADDING_PT As Double = 8

    diagramWidth = timelineWidth
    If sidebarEnabled Then
        If sidebarBoth Then
            diagramWidth = timelineWidth + 2 * sidebarWidth + sidebarGap * 2
        Else
            diagramWidth = timelineWidth + sidebarWidth + sidebarGap
        End If
    End If

    headerTitleText = GetHeaderTextFromTable(wsData)
    titleHeight = 0
    If Len(Trim$(headerTitleText)) > 0 And sidebarEnabled Then
        baseTitleHeight = yearHeaderHeightPoints + periodHeaderHeight
        availTitleWidth = Application.WorksheetFunction.Max(1, sidebarWidth - TITLE_PADDING_PT * 2)
        titleLines = EstimateWrappedLinesForFont(headerTitleText, availTitleWidth, fontHeader)
        If titleLines < 1 Then titleLines = 1
        neededTitleHeight = titleLines * EstimateTextHeight(CLng(fontHeader.Size)) + (titleLines - 1) * 1 + TITLE_PADDING_PT
        titleHeight = Application.WorksheetFunction.Max(baseTitleHeight, neededTitleHeight)
    End If

    totalHeaderHeight = Application.WorksheetFunction.Max(titleHeight, yearHeaderHeightPoints + periodHeaderHeight)
    If titleHeight > 0 And titleHeight > yearHeaderHeightPoints + periodHeaderHeight Then
        periodHeaderHeight = titleHeight - yearHeaderHeightPoints
    End If

    If displayRowCount > 0 Then
        ComputeRowTops startCell.Top + totalHeaderHeight, rowHeights, rowTops, totalTasksHeight
    End If

    bgColor = GetColorFromTable(colorsTable, "Фон диаграммы", bgTransparencyColor)
    If bgColor = -1 Then
        bgColor = RGB(242, 242, 242)
    End If

    yearHeaderColor = GetColorFromTable(colorsTable, "Фон шапки (годы)", yearHeaderTransparency)
    If yearHeaderColor = -1 Then
        yearHeaderColor = RGB(217, 217, 217)
    End If

    periodHeaderColor = GetColorFromTable(colorsTable, "Фон шапки (периоды)", periodHeaderTransparency)
    If periodHeaderColor = -1 Then
        periodHeaderColor = RGB(198, 224, 180)
    End If

    headerTitleColor = GetColorFromTable(colorsTable, "Фон заголовка", headerTitleTransparency)
    If headerTitleColor = -1 Then
        headerTitleColor = RGB(217, 217, 217)
    End If

    diagramWidth = timelineWidth
    If sidebarEnabled Then
        If sidebarBoth Then
            diagramWidth = timelineWidth + 2 * sidebarWidth + sidebarGap * 2
        Else
            diagramWidth = timelineWidth + sidebarWidth + sidebarGap
        End If
    End If
    If displayRowCount = 0 Then
        displayRowCount = rowCount
    End If
    legendPadding = 6
    legendOffsetLeft = 10 * 2.83465
    If showLegend Then
        CollectUsedLegendItems dataTable, eventTypesTable, taskRowMap, rowCount, timelineStart, timelineEnd, _
            eventColumnIndices, eventDateColumnIndices, barHeightNominal, usedLineTypes, usedLineTypeCount, _
            usedEventTypes, usedEventTypeCount, maxLegendEventSize
        legendHeight = GetLegendHeightFromCounts(usedLineTypeCount, usedEventTypeCount, legendPadding, _
            maxLegendEventSize)
    Else
        legendHeight = 0
    End If
    Dim todayLabelReserveHeight As Double
    Dim todayLabelShapeWidth As Double
    Dim todayLabelGap As Double
    Dim todayStripeEnabled As Boolean
    todayLabelReserveHeight = 0
    todayLabelGap = 2 ' Отступ фигуры "Сегодня" от таймлайна
    todayStripeEnabled = (StrComp(todayStripeMode, "Полоса", vbTextCompare) = 0 Or StrComp(todayStripeMode, "Полоса с точкой", vbTextCompare) = 0)
    ' Надпись показывается только если включена полоса "Сегодня"
    If todayStripeEnabled And (StrComp(todayLabelMode, "Горизонтально", vbTextCompare) = 0 Or StrComp(todayLabelMode, "Вертикально", vbTextCompare) = 0) Then
        If StrComp(todayLabelMode, "Вертикально", vbTextCompare) = 0 Then
            todayLabelShapeWidth = MeasureTextWidthWithFont("Сегодня", fontTodayLabel.name, CLng(fontTodayLabel.Size), fontTodayLabel.bold, fontTodayLabel.italic) + 4
            ' Высота фона: высота фигуры надписи (после поворота -90° = ширина бокса lblW+4) + четыре отступа от таймлайна
            todayLabelReserveHeight = todayLabelShapeWidth + 4 + (4 * todayLabelGap)
        Else
            ' Высота фона: высота фигуры + два отступа от таймлайна
            todayLabelReserveHeight = EstimateTextHeight(CLng(fontTodayLabel.Size)) + 2 + 4 + (2 * todayLabelGap)
        End If
    End If
    diagramHeight = totalHeaderHeight + totalTasksHeight + legendHeight + todayLabelReserveHeight

    bgTransparencyColor = ResolveTransparency(bgTransparencyColor, 0)

    DrawDiagramBackground wsGantt, diagramLeft, startCell.Top, diagramWidth, diagramHeight, bgColor, _
        bgTransparencyColor

    If titleHeight > 0 And Len(Trim$(headerTitleText)) > 0 Then
        headerTitleTransparency = ResolveTransparency(headerTitleTransparency, 0)
        If LCase$(sidebarMode) = "слева" Or sidebarBoth Then
            DrawDiagramTitle wsGantt, diagramLeft, startCell.Top, sidebarWidth, titleHeight, headerTitleText, _
                fontHeader, headerTitleColor, headerTitleTransparency
        End If
        If LCase$(sidebarMode) = "справа" Or sidebarBoth Then
            DrawDiagramTitle wsGantt, timelineLeft + timelineWidth + sidebarGap, startCell.Top, sidebarWidth, titleHeight, headerTitleText, _
                fontHeader, headerTitleColor, headerTitleTransparency
        End If
    End If

    DrawHeaderShapes wsGantt, timelineLeft, startCell.Top, timelineData, yearHeaderHeightPoints, periodHeaderHeight, _
        yearHeaderColor, yearHeaderTransparency, periodHeaderColor, periodHeaderTransparency, rotateAllPeriods, _
        fontYears, fontPeriods

    sidebarColor = GetColorFromTable(colorsTable, "Цвет сайдбара", sidebarTransparency)
    If sidebarColor = -1 Then
        sidebarColor = RGB(242, 242, 242)
    End If

    sectionRowColor = GetColorFromTable(colorsTable, "Цвет раздела", sectionTransparencyColor)
    If sectionRowColor = -1 Then
        sectionRowColor = RGB(217, 217, 217)
    End If

    subsectionRowColor = GetColorFromTable(colorsTable, "Цвет подраздела", subsectionTransparencyColor)
    If subsectionRowColor = -1 Then
        subsectionRowColor = RGB(232, 232, 232)
    End If

    vertDividerColor = GetColorFromTable(colorsTable, "Цвет вертикальных разделителей", vertDividerTransparency)
    If vertDividerColor = -1 Then
        vertDividerColor = RGB(191, 191, 191)
    End If

    yearDividerColor = GetColorFromTable(colorsTable, "Цвет вертикальных разделителей между годами", yearDividerTransparency)
    If yearDividerColor = -1 Then
        yearDividerColor = vertDividerColor
    End If

    horizDividerColor = GetColorFromTable(colorsTable, "Цвет горизонтальных разделителей", horizDividerTransparency)
    If horizDividerColor = -1 Then
        horizDividerColor = RGB(191, 191, 191)
    End If

    todayStripeColor = GetColorFromTable(colorsTable, "Цвет полосы Сегодня", todayStripeTransparency)
    If todayStripeColor = -1 Then
        todayStripeColor = RGB(255, 0, 0)
    End If
    pastEventsColor = GetColorFromTable(colorsTable, "Закрашивание прошедших событий", pastEventsTransparency)
    If pastEventsColor = -1 Then
        pastEventsColor = RGB(180, 180, 180)
    End If
    pastEventsTransparency = ResolveTransparency(pastEventsTransparency, 50)

    yearHeaderTransparency = ResolveTransparency(yearHeaderTransparency, 0)
    periodHeaderTransparency = ResolveTransparency(periodHeaderTransparency, 0)
    sidebarTransparency = ResolveTransparency(sidebarTransparency, 0)
    vertDividerTransparency = ResolveTransparency(vertDividerTransparency, 0)
    yearDividerTransparency = ResolveTransparency(yearDividerTransparency, vertDividerTransparency)
    horizDividerTransparency = ResolveTransparency(horizDividerTransparency, 0)
    todayStripeTransparency = ResolveTransparency(todayStripeTransparency, 0)
    sectionTransparencyFinal = ResolveTransparency(sectionTransparencyColor, sectionTransparency)
    subsectionTransparencyFinal = ResolveTransparency(subsectionTransparencyColor, sectionTransparency)

    If LCase$(sidebarMode) = "справа" Then
        DrawSidebar wsGantt, timelineLeft + timelineWidth + sidebarGap, startCell.Top, totalHeaderHeight, _
            totalTasksHeight, titleHeight, displayRowCount, sectionColWidth, taskColWidth, groupRowIndices, groupNames, groupTypes, groupCount, _
            sidebarTaskNames, sidebarTaskRowIndices, sidebarColor, sidebarTransparency, sectionRowColor, sectionTransparencyFinal, _
            subsectionRowColor, subsectionTransparencyFinal, drawSectionInTimeline, rowTops, rowHeights, sidebarTextIndentPoints, sidebarSectionTextIndentPoints, sidebarSubsectionTextIndentPoints, sidebarTextPaddingPoints, fontSections, fontSubsections, fontTasks
    ElseIf LCase$(sidebarMode) = "слева" Then
        DrawSidebar wsGantt, diagramLeft, startCell.Top, totalHeaderHeight, _
            totalTasksHeight, titleHeight, displayRowCount, sectionColWidth, taskColWidth, groupRowIndices, groupNames, groupTypes, groupCount, _
            sidebarTaskNames, sidebarTaskRowIndices, sidebarColor, sidebarTransparency, sectionRowColor, sectionTransparencyFinal, _
            subsectionRowColor, subsectionTransparencyFinal, drawSectionInTimeline, rowTops, rowHeights, sidebarTextIndentPoints, sidebarSectionTextIndentPoints, sidebarSubsectionTextIndentPoints, sidebarTextPaddingPoints, fontSections, fontSubsections, fontTasks
    ElseIf sidebarBoth Then
        DrawSidebar wsGantt, diagramLeft, startCell.Top, totalHeaderHeight, _
            totalTasksHeight, titleHeight, displayRowCount, sectionColWidth, taskColWidth, groupRowIndices, groupNames, groupTypes, groupCount, _
            sidebarTaskNames, sidebarTaskRowIndices, sidebarColor, sidebarTransparency, sectionRowColor, sectionTransparencyFinal, _
            subsectionRowColor, subsectionTransparencyFinal, drawSectionInTimeline, rowTops, rowHeights, sidebarTextIndentPoints, sidebarSectionTextIndentPoints, sidebarSubsectionTextIndentPoints, sidebarTextPaddingPoints, fontSections, fontSubsections, fontTasks
        DrawSidebar wsGantt, timelineLeft + timelineWidth + sidebarGap, startCell.Top, totalHeaderHeight, _
            totalTasksHeight, titleHeight, displayRowCount, sectionColWidth, taskColWidth, groupRowIndices, groupNames, groupTypes, groupCount, _
            sidebarTaskNames, sidebarTaskRowIndices, sidebarColor, sidebarTransparency, sectionRowColor, sectionTransparencyFinal, _
            subsectionRowColor, subsectionTransparencyFinal, drawSectionInTimeline, rowTops, rowHeights, sidebarTextIndentPoints, sidebarSectionTextIndentPoints, sidebarSubsectionTextIndentPoints, sidebarTextPaddingPoints, fontSections, fontSubsections, fontTasks
    End If

    DrawTimelineDividers wsGantt, diagramLeft, startCell.Top, diagramWidth, yearHeaderHeightPoints, totalHeaderHeight, _
        totalTasksHeight, displayRowCount, rowTops, rowHeights, timelineLeft, timelineWidth, timelineData, _
        sidebarMode, sidebarWidth, sidebarGap, vertDividerThickness, yearDividerThickness, horizDividerThickness, _
        vertDividerColor, vertDividerTransparency, yearDividerColor, yearDividerTransparency, horizDividerColor, horizDividerTransparency

    If drawSectionInTimeline Then
        DrawSectionRows wsGantt, diagramLeft, startCell.Top + totalHeaderHeight, _
            diagramWidth, groupRowIndices, groupLengths, groupNames, groupTypes, groupCount, _
            sectionRowColor, sectionTransparencyFinal, subsectionRowColor, subsectionTransparencyFinal, _
            sidebarSectionTextIndentPoints, sidebarSubsectionTextIndentPoints, rowTops, rowHeights, fontSections, fontSubsections
    End If

    For i = 1 To rowCount
        sectionName = sectionNames(i)
        taskName = taskNames(i)
        dataRowIndex = taskRowMap(i)
        startDate = dataTable.ListColumns("Дата начала").DataBodyRange.Cells(dataRowIndex, 1).value
        endDate = dataTable.ListColumns("Дата окончания").DataBodyRange.Cells(dataRowIndex, 1).value
        lineType = CStr(dataTable.ListColumns("Тип линии").DataBodyRange.Cells(dataRowIndex, 1).value)

        If endDate < startDate Then
            GoTo NextRow
        End If

        If sidebarEnabled Then
            rowIndex = taskRowIndices(i)
        ElseIf groupCount > 0 Then
            rowIndex = taskRowIndices(i)
        Else
            rowIndex = i
        End If
        If sidebarEnabled Then
            barLabel = ""
        Else
            barLabel = TruncateTextWithEllipsis(taskName, _
                GetTimelineOffset(timelineData, endDate + 1) - GetTimelineOffset(timelineData, startDate), 9, 4)
        End If
        barHeight = barHeightNominal
        shapeTop = rowTops(rowIndex) + trackPadPoints

        Select Case LCase$(barDisplayMode)
            Case "уменьшение высоты"
                If taskStackCounts(i) > 1 Then
                    barHeight = (barHeightNominal - (barGapPoints * (taskStackCounts(i) - 1))) / taskStackCounts(i)
                    If barHeight < 1 Then barHeight = 1
                    stackHeight = (barHeight * taskStackCounts(i)) + (barGapPoints * (taskStackCounts(i) - 1))
                    shapeTop = rowTops(rowIndex) + trackPadPoints + ComputeBarTopWithEventReserve(i, rowIndex, taskStackIndices(i), barHeight, barGapPoints, trackPadPoints, _
                        rowCount, taskRowIndices, taskStackIndices, eventLabelReserveForTask, isSectionRow, rowTops, rowHeights, barHeightNominal, barDisplayMode, _
                        maxFirstTierLabelHeightForTask, firstTierReserveForTask, secondTierHeightForTask, eventLabelGapTopPoints, eventLabelGapBottomPoints)
                Else
                    barHeight = barHeightNominal
                    shapeTop = rowTops(rowIndex) + trackPadPoints + ComputeBarTopWithEventReserve(i, rowIndex, 1, barHeight, barGapPoints, trackPadPoints, _
                        rowCount, taskRowIndices, taskStackIndices, eventLabelReserveForTask, isSectionRow, rowTops, rowHeights, barHeightNominal, barDisplayMode, _
                        maxFirstTierLabelHeightForTask, firstTierReserveForTask, secondTierHeightForTask, eventLabelGapTopPoints, eventLabelGapBottomPoints)
                End If

            Case "одинаковая высота"
                barHeight = barHeightNominal
                If taskStackCounts(i) > 1 Then
                    stackHeight = (barHeight * taskStackCounts(i)) + (barGapPoints * (taskStackCounts(i) - 1))
                    shapeTop = rowTops(rowIndex) + trackPadPoints + ComputeBarTopWithEventReserve(i, rowIndex, taskStackIndices(i), barHeight, barGapPoints, trackPadPoints, _
                        rowCount, taskRowIndices, taskStackIndices, eventLabelReserveForTask, isSectionRow, rowTops, rowHeights, barHeightNominal, barDisplayMode, _
                        maxFirstTierLabelHeightForTask, firstTierReserveForTask, secondTierHeightForTask, eventLabelGapTopPoints, eventLabelGapBottomPoints)
                Else
                    shapeTop = rowTops(rowIndex) + trackPadPoints + ComputeBarTopWithEventReserve(i, rowIndex, 1, barHeight, barGapPoints, trackPadPoints, _
                        rowCount, taskRowIndices, taskStackIndices, eventLabelReserveForTask, isSectionRow, rowTops, rowHeights, barHeightNominal, barDisplayMode, _
                        maxFirstTierLabelHeightForTask, firstTierReserveForTask, secondTierHeightForTask, eventLabelGapTopPoints, eventLabelGapBottomPoints)
                End If

            Case "в одну линию"
                barHeight = barHeightNominal
                shapeTop = rowTops(rowIndex) + trackPadPoints + ComputeBarTopWithEventReserve(i, rowIndex, 1, barHeight, barGapPoints, trackPadPoints, _
                    rowCount, taskRowIndices, taskStackIndices, eventLabelReserveForTask, isSectionRow, rowTops, rowHeights, barHeightNominal, barDisplayMode, _
                    maxFirstTierLabelHeightForTask, firstTierReserveForTask, secondTierHeightForTask, eventLabelGapTopPoints, eventLabelGapBottomPoints)

            Case Else
                barHeight = barHeightNominal
                shapeTop = rowTops(rowIndex) + trackPadPoints + ComputeBarTopWithEventReserve(i, rowIndex, taskStackIndices(i), barHeight, barGapPoints, trackPadPoints, _
                    rowCount, taskRowIndices, taskStackIndices, eventLabelReserveForTask, isSectionRow, rowTops, rowHeights, barHeightNominal, barDisplayMode, _
                    maxFirstTierLabelHeightForTask, firstTierReserveForTask, secondTierHeightForTask, eventLabelGapTopPoints, eventLabelGapBottomPoints)
        End Select
        If taskStackCounts(i) = 1 Then
            hasEventLabel = False
            If i >= LBound(eventLabelReserveForTask) And i <= UBound(eventLabelReserveForTask) Then
                hasEventLabel = (eventLabelReserveForTask(i) > 0)
            End If
            If Not hasEventLabel Then
                shapeTop = rowTops(rowIndex) + (rowHeights(rowIndex) - barHeight) / 2
            End If
        End If
        ' Применяем сдвиг бара вниз из настроек
        shapeTop = shapeTop + barOffsetDownPoints
        shapeLeft = timelineLeft + GetTimelineOffset(timelineData, startDate)
        Dim barWidth As Double
        barWidth = GetTimelineOffset(timelineData, endDate + 1) - GetTimelineOffset(timelineData, startDate)

        barColor = GetLineTypeColor(lineTypesTable, lineType, RGB(91, 155, 213))
        barTransparency = GetLineTypeTransparency(lineTypesTable, lineType, 0)
        lineColor = barColor

        Set barShape = DrawTaskBar(wsGantt, shapeLeft, shapeTop, barWidth, _
            barHeight, barLabel, sectionName, barColor, lineColor, showTaskBorder, fontTasks, taskBarShapeType)
        barShape.Fill.transparency = ClampTransparency(barTransparency) / 100
        barItems.Add Array(shapeLeft, shapeTop, barWidth, barHeight)

        For eventIdx = 1 To 10
            eventColumnIndex = eventColumnIndices(eventIdx)
            eventDateColumnIndex = eventDateColumnIndices(eventIdx)
            eventDescColumnIndex = eventDescColumnIndices(eventIdx)
            eventLabelText = vbNullString
            If eventColumnIndex > 0 And eventDateColumnIndex > 0 Then
                eventName = Trim$(CStr(dataTable.ListColumns(eventColumnIndex).DataBodyRange.Cells(dataRowIndex, 1).value))
                eventDateValue = dataTable.ListColumns(eventDateColumnIndex).DataBodyRange.Cells(dataRowIndex, 1).value
                eventDescValue = vbNullString
                If eventDescColumnIndex > 0 Then
                    eventDescValue = dataTable.ListColumns(eventDescColumnIndex).DataBodyRange.Cells(dataRowIndex, 1).value
                End If
                eventLabelText = Trim$(CStr(eventDescValue))
                If Len(eventName) > 0 And IsDate(eventDateValue) Then
                    eventDate = CDate(eventDateValue)
                    If eventDate >= timelineStart And eventDate <= timelineEnd Then
                        eventFound = GetEventTypeInfo(eventTypesTable, eventName, eventShapeType, eventColor, eventHeight)
                        eventTransparency = GetEventTypeTransparency(eventTypesTable, eventName, 0)
                        If Not eventFound Then
                            eventShapeType = msoShapeOval
                            eventColor = barColor
                            eventHeight = barHeightNominal
                            eventTransparency = 0
                        End If
                        If eventHeight <= 0 Then
                            eventHeight = barHeightNominal
                        End If
                        eventWidth = eventHeight
                        eventLeft = timelineLeft + GetTimelineOffset(timelineData, eventDate + 0.5) - eventWidth / 2
                        ' Позиционируем событие относительно нижнего края бара (снизу вверх)
                        eventTop = shapeTop + barHeight - eventHeight
                        eventItems.Add Array(eventShapeType, eventLeft, eventTop, eventWidth, eventHeight, eventColor, showEventBorder, eventTransparency, eventLabelText, GetEventShapeName(eventTypesTable, eventName), i)
                    End If
                End If
            End If
        Next eventIdx
NextRow:
    Next i

    Dim overlayShapes As Collection
    Set overlayShapes = New Collection
    If LCase$(pastEventsOverlayMode) = "да" Then
        DrawPastEventsOverlays wsGantt, timelineLeft, timelineData, pastEventsColor, pastEventsTransparency, _
            overlayShapes, displayRowCount, rowTops, rowHeights, isSectionRow
    End If
    If Not eventItems Is Nothing Then
        ComputeEventLabelOffsets eventItems, eventLabelWidthPoints, eventLabelMaxLines, textOffsets, labelTiers, _
            timelineLeft, timelineWidth, diagramLeft, diagramWidth, sidebarMode, sidebarEnabled, fontEventDesc
        
        ' Вычисляем фактическую высоту самого высокого описания первого уровня для каждой задачи
        ' maxFirstTierLabelHeightForTask уже объявлена выше, просто заполняем значения
        evIdx = 0
        Dim taskIndexForEvPreCalc As Long
        For Each ev In eventItems
            If IsArray(ev) Then
                evIdx = evIdx + 1
                If UBound(ev) >= 9 And UBound(ev) >= 10 Then
                    taskIndexForEvPreCalc = CLng(ev(10))
                    If taskIndexForEvPreCalc >= 1 And taskIndexForEvPreCalc <= rowCount Then
                        If evIdx <= UBound(labelTiers) And labelTiers(evIdx) = 0 Then
                            ' Это событие первого яруса - вычисляем его фактическую высоту
                            Dim labelTextForHeight As String
                            labelTextForHeight = CStr(ev(8))
                            If Len(Trim$(labelTextForHeight)) > 0 Then
                                Dim lblWForHeight As Double, availWForHeight As Double
                                Dim lineCountForHeight As Long, actualLinesForHeight As Long
                                Dim lblHForHeight As Double
                                lblWForHeight = Application.WorksheetFunction.Min(eventLabelWidthPoints, _
                                    Application.WorksheetFunction.Max(eventLabelWidthPoints * 0.5, _
                                    MeasureTextWidthWithFont(labelTextForHeight, fontEventDesc.name, CLng(fontEventDesc.Size), fontEventDesc.bold, fontEventDesc.italic) + 8))
                                availWForHeight = Application.WorksheetFunction.Max(1, lblWForHeight)
                                lineCountForHeight = EstimateWrappedLinesForFont(labelTextForHeight, availWForHeight, fontEventDesc)
                                If lineCountForHeight < 1 Then lineCountForHeight = 1
                                actualLinesForHeight = Application.WorksheetFunction.Min(lineCountForHeight, eventLabelMaxLines)
                                lblHForHeight = actualLinesForHeight * EstimateTextHeight(CLng(fontEventDesc.Size)) + (actualLinesForHeight - 1) * 1 + 2
                                If lblHForHeight > maxFirstTierLabelHeightForTask(taskIndexForEvPreCalc) Then
                                    maxFirstTierLabelHeightForTask(taskIndexForEvPreCalc) = lblHForHeight
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next ev
        
        evIdx = 0
        Dim taskIndexForEv As Long
        Dim tierOffsetPoints As Double
        For Each ev In eventItems
            If IsArray(ev) Then
                evIdx = evIdx + 1
                If UBound(ev) >= 8 Then
                    lblOffset = 0
                    If evIdx <= UBound(textOffsets) Then lblOffset = textOffsets(evIdx)
                    taskIndexForEv = 0
                    If UBound(ev) >= 10 Then taskIndexForEv = CLng(ev(10))
                    tierOffsetPoints = 0
                    If evIdx <= UBound(labelTiers) And labelTiers(evIdx) = 1 And taskIndexForEv >= 1 Then
                        If rowCount > 0 And taskIndexForEv <= UBound(secondTierHeightForTask) And secondTierHeightForTask(taskIndexForEv) > 0 Then
                        ' Отступ 1 мм: от верхней границы самого высокого описания 1-го уровня до нижней границы описания 2-го уровня.
                        ' Используем фактическую высоту самого высокого описания первого уровня (не резерв, а реальную высоту текста)
                            If taskIndexForEv <= UBound(maxFirstTierLabelHeightForTask) And maxFirstTierLabelHeightForTask(taskIndexForEv) > 0 Then
                                ' tierOffsetPoints = фактическая_высота_текста_первого_уровня + GAP_1MM_PT
                                tierOffsetPoints = maxFirstTierLabelHeightForTask(taskIndexForEv) + GAP_1MM_PT
                            ElseIf taskIndexForEv <= UBound(firstTierReserveForTask) Then
                                ' Fallback: используем исходный резерв минус отступы
                                tierOffsetPoints = firstTierReserveForTask(taskIndexForEv) - eventLabelGapBottomPoints - eventLabelGapTopPoints + GAP_1MM_PT
                            Else
                                ' Fallback: используем текущее значение минус второй ярус
                                tierOffsetPoints = eventLabelReserveForTask(taskIndexForEv) - secondTierHeightForTask(taskIndexForEv) - eventLabelGapBottomPoints - eventLabelGapTopPoints
                            End If
                            If tierOffsetPoints < 0 Then tierOffsetPoints = 0
                        End If
                    End If
                    DrawEventShape wsGantt, ev(0), ev(1), ev(2), ev(3), ev(4), ev(5), ev(6), fontEventDesc, ev(7), ev(8), _
                        eventLabelGapBottomPoints, eventLabelWidthPoints, eventLabelMaxLines, lblOffset, _
                        IIf(UBound(ev) >= 9, CStr(ev(9)), ""), _
                        IIf(tierOffsetPoints > 0, 1, 0), tierOffsetPoints
                End If
            End If
        Next ev
        ' Overlay stays above bars (drawn after them) and below events (events drawn after overlay)
    End If

    If StrComp(todayStripeMode, "Полоса", vbTextCompare) = 0 Or StrComp(todayStripeMode, "Полоса с точкой", vbTextCompare) = 0 Then
        DrawTodayStripe wsGantt, timelineLeft, startCell.Top + totalHeaderHeight, _
            totalTasksHeight, displayRowCount, timelineData, todayStripeThickness, todayStripeColor, _
            todayStripeTransparency, todayStripeMode
    End If
    ' Надпись показывается только если включена полоса "Сегодня"
    If todayStripeEnabled And (StrComp(todayLabelMode, "Горизонтально", vbTextCompare) = 0 Or StrComp(todayLabelMode, "Вертикально", vbTextCompare) = 0) Then
        DrawTodayLabel wsGantt, timelineLeft, startCell.Top + totalHeaderHeight, totalTasksHeight, _
            timelineData, todayLabelMode, fontTodayLabel
    End If

    If legendHeight > 0 Then
        DrawLegend wsGantt, diagramLeft + legendOffsetLeft, startCell.Top + totalHeaderHeight + totalTasksHeight + todayLabelReserveHeight, _
            diagramWidth - legendOffsetLeft, legendHeight, lineTypesTable, eventTypesTable, legendPadding, _
            barHeightNominal, _
            usedLineTypes, usedLineTypeCount, usedEventTypes, usedEventTypeCount, showEventBorder, fontLegend
    End If

    Application.ScreenUpdating = True
End Sub

Private Sub DrawDiagramTitle(ByVal ws As Worksheet, ByVal leftPos As Double, ByVal topPos As Double, _
                             ByVal width As Double, ByVal height As Double, ByVal titleText As String, _
                             ByRef fontHeader As fontSettingsType, ByVal fillColor As Long, ByVal fillTransparency As Double)
    Dim shp As Shape

    If Len(Trim$(titleText)) = 0 Then Exit Sub
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, leftPos, topPos, width, height)
    shp.Fill.ForeColor.RGB = fillColor
    shp.Fill.transparency = ClampTransparency(fillTransparency) / 100
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.text = titleText
    ApplyFontToTextRange shp.TextFrame2.TextRange, fontHeader
    shp.TextFrame2.WordWrap = msoCTrue
    shp.TextFrame2.AutoSize = msoAutoSizeNone
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.TextRange.ParagraphFormat.SpaceBefore = 0
    shp.TextFrame2.TextRange.ParagraphFormat.SpaceAfter = 0
    shp.TextFrame2.MarginTop = 4
    shp.TextFrame2.MarginBottom = 4
    shp.TextFrame2.MarginLeft = 8
    shp.TextFrame2.MarginRight = 8
    shp.ZOrder msoBringToFront
End Sub

Private Sub DrawDiagramBackground(ByVal ws As Worksheet, ByVal leftPos As Double, ByVal topPos As Double, _
                                  ByVal width As Double, ByVal height As Double, ByVal fillColor As Long, _
                                  ByVal backgroundTransparency As Double)
    Dim shp As Shape

    Set shp = ws.Shapes.AddShape(msoShapeRectangle, leftPos, topPos, width, height)
    shp.name = BG_SHAPE_NAME
    shp.Fill.ForeColor.RGB = fillColor
    shp.Fill.transparency = ClampTransparency(backgroundTransparency) / 100
    shp.Line.Visible = msoFalse
End Sub

Private Sub DrawHeaderShapes(ByVal ws As Worksheet, ByVal startLeft As Double, ByVal startTop As Double, _
                             ByRef timelineData As timelineData, ByVal yearHeaderHeight As Double, _
                             ByVal periodHeaderHeight As Double, ByVal yearColor As Long, _
                             ByVal yearTransparency As Double, ByVal periodColor As Long, _
                             ByVal periodTransparency As Double, ByVal rotateAllPeriods As Boolean, _
                             ByRef fontYears As fontSettingsType, ByRef fontPeriods As fontSettingsType)
    Dim i As Long
    Dim periodLeft As Double
    Dim yearStartIndex As Long
    Dim yearWidth As Double
    Dim currentYear As Long
    Dim periodShape As Shape
    Dim yearShape As Shape

    periodLeft = startLeft
    currentYear = Year(timelineData.PeriodStarts(1))
    yearStartIndex = 1
    yearWidth = 0
    For i = 1 To timelineData.count
        Set periodShape = ws.Shapes.AddShape(msoShapeRectangle, periodLeft, startTop + yearHeaderHeight, _
                                             timelineData.PeriodWidths(i), periodHeaderHeight)
        periodShape.Fill.ForeColor.RGB = periodColor
        periodShape.Fill.transparency = ClampTransparency(periodTransparency) / 100
        periodShape.Line.Visible = msoFalse
        periodShape.TextFrame2.TextRange.text = timelineData.PeriodLabels(i)
        ApplyFontToTextRange periodShape.TextFrame2.TextRange, fontPeriods
        periodShape.TextFrame2.WordWrap = msoFalse
        periodShape.TextFrame2.TextRange.ParagraphFormat.WordWrap = msoFalse
        periodShape.TextFrame2.VerticalAnchor = msoAnchorMiddle
        periodShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        If rotateAllPeriods Then
            FitRotatedPeriodShape periodShape, periodLeft, startTop + yearHeaderHeight, _
                timelineData.PeriodWidths(i), periodHeaderHeight, timelineData.PeriodLabels(i), fontPeriods.Size
        Else
            FitHorizontalPeriodShape periodShape, periodLeft, startTop + yearHeaderHeight, _
                timelineData.PeriodWidths(i), periodHeaderHeight
        End If

        yearWidth = yearWidth + timelineData.PeriodWidths(i)

        If i = timelineData.count Then
            Set yearShape = ws.Shapes.AddShape(msoShapeRectangle, startLeft + timelineData.CumulativeWidths(yearStartIndex), _
                                               startTop, yearWidth, yearHeaderHeight)
            yearShape.Fill.ForeColor.RGB = yearColor
            yearShape.Fill.transparency = ClampTransparency(yearTransparency) / 100
            yearShape.Line.Visible = msoFalse
            yearShape.TextFrame2.TextRange.text = CStr(currentYear)
            ApplyFontToTextRange yearShape.TextFrame2.TextRange, fontYears
            yearShape.TextFrame2.VerticalAnchor = msoAnchorMiddle
            yearShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            yearShape.TextFrame2.MarginTop = 2
            yearShape.TextFrame2.MarginBottom = 2
        ElseIf Year(timelineData.PeriodStarts(i + 1)) <> currentYear Then
            Set yearShape = ws.Shapes.AddShape(msoShapeRectangle, startLeft + timelineData.CumulativeWidths(yearStartIndex), _
                                               startTop, yearWidth, yearHeaderHeight)
            yearShape.Fill.ForeColor.RGB = yearColor
            yearShape.Fill.transparency = ClampTransparency(yearTransparency) / 100
            yearShape.Line.Visible = msoFalse
            yearShape.TextFrame2.TextRange.text = CStr(currentYear)
            ApplyFontToTextRange yearShape.TextFrame2.TextRange, fontYears
            yearShape.TextFrame2.VerticalAnchor = msoAnchorMiddle
            yearShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            yearShape.TextFrame2.MarginTop = 2
            yearShape.TextFrame2.MarginBottom = 2

            currentYear = Year(timelineData.PeriodStarts(i + 1))
            yearStartIndex = i + 1
            yearWidth = 0
        End If

        periodLeft = periodLeft + timelineData.PeriodWidths(i)
    Next i
End Sub

Private Sub FitRotatedPeriodShape(ByVal periodShape As Shape, ByVal targetLeft As Double, _
                                  ByVal targetTop As Double, ByVal targetWidth As Double, _
                                  ByVal targetHeight As Double, ByVal labelText As String, _
                                  ByVal fontSize As Long)
    ' Rotate the whole shape (incl. text) and then re-center it inside the target cell.
    ' Important: after rotation Excel may reset/ignore vertical anchoring for some fonts/strings,
    ' so we re-apply VerticalAnchor and margins here to keep month names vertically centered.
    periodShape.Rotation = -90
    periodShape.width = targetHeight
    periodShape.height = targetWidth
    periodShape.Left = targetLeft + (targetWidth - periodShape.width) / 2
    periodShape.Top = targetTop + (targetHeight - periodShape.height) / 2

    ' Re-apply text formatting after rotation to avoid "sticking to top" for long labels
    With periodShape.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .MarginTop = 0
        .MarginBottom = 0
        .MarginLeft = 0
        .MarginRight = 0
        .WordWrap = msoFalse
        .TextRange.ParagraphFormat.WordWrap = msoFalse
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
End Sub

Private Sub FitHorizontalPeriodShape(ByVal periodShape As Shape, ByVal targetLeft As Double, _
                                     ByVal targetTop As Double, ByVal targetWidth As Double, _
                                     ByVal targetHeight As Double)
    periodShape.Rotation = 0
    periodShape.width = targetWidth
    periodShape.height = targetHeight
    periodShape.Left = targetLeft
    periodShape.Top = targetTop
    With periodShape.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .MarginTop = 0
        .MarginBottom = 0
        .MarginLeft = 0
        .MarginRight = 0
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
End Sub

Private Sub DrawSectionRows(ByVal ws As Worksheet, ByVal leftPos As Double, ByVal topPos As Double, _
                            ByVal timelineWidth As Double, ByRef groupRowIndices() As Long, _
                            ByRef groupLengths() As Long, ByRef groupNames() As String, ByRef groupTypes() As Long, _
                            ByVal groupCount As Long, ByVal sectionColor As Long, _
                            ByVal sectionTransparency As Double, ByVal subsectionColor As Long, _
                            ByVal subsectionTransparency As Double, ByVal sectionTextIndent As Double, _
                            ByVal subsectionTextIndent As Double, ByRef rowTops() As Double, _
                            ByRef rowHeights() As Double, ByRef fontSections As fontSettingsType, ByRef fontSubsections As fontSettingsType)
    Dim g As Long
    Dim shp As Shape
    Dim rowTop As Double

    If groupCount <= 0 Then
        Exit Sub
    End If

    For g = 1 To groupCount
        If groupLengths(g) <= 0 Or Len(Trim$(groupNames(g))) = 0 Then
            GoTo NextGroup
        End If
        rowTop = rowTops(groupRowIndices(g))
        Set shp = ws.Shapes.AddShape(msoShapeRectangle, leftPos, rowTop, timelineWidth, rowHeights(groupRowIndices(g)))
        If g <= UBound(groupTypes) And groupTypes(g) = 2 Then
            shp.Fill.ForeColor.RGB = subsectionColor
            shp.Fill.transparency = ClampTransparency(subsectionTransparency) / 100
        Else
            shp.Fill.ForeColor.RGB = sectionColor
            shp.Fill.transparency = ClampTransparency(sectionTransparency) / 100
        End If
        shp.Line.Visible = msoFalse
        shp.TextFrame2.TextRange.text = groupNames(g)
        If g <= UBound(groupTypes) And groupTypes(g) = 2 Then
            ApplyFontToTextRange shp.TextFrame2.TextRange, fontSubsections
        Else
            ApplyFontToTextRange shp.TextFrame2.TextRange, fontSections
        End If
        shp.TextFrame2.WordWrap = msoTrue
        shp.TextFrame2.TextRange.ParagraphFormat.WordWrap = msoTrue
        shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        If g <= UBound(groupTypes) And groupTypes(g) = 2 Then
            shp.TextFrame2.MarginLeft = subsectionTextIndent
        Else
            shp.TextFrame2.MarginLeft = sectionTextIndent
        End If
        shp.TextFrame2.MarginRight = 0
        shp.TextFrame2.MarginTop = 0
        shp.TextFrame2.MarginBottom = 0
NextGroup:
    Next g
End Sub

Private Sub DrawSidebar(ByVal ws As Worksheet, ByVal leftPos As Double, ByVal topPos As Double, _
                        ByVal headerHeight As Double, ByVal tasksHeight As Double, ByVal titleHeight As Double, _
                        ByVal displayRowCount As Long, ByVal sectionColWidth As Double, _
                        ByVal taskColWidth As Double, ByRef groupRowIndices() As Long, _
                        ByRef groupNames() As String, ByRef groupTypes() As Long, ByVal groupCount As Long, _
                        ByRef taskNames() As String, ByRef taskRowIndices() As Long, _
                        ByVal sidebarColor As Long, ByVal sidebarTransparency As Double, _
                        ByVal sectionColor As Long, ByVal sectionTransparency As Double, _
                        ByVal subsectionColor As Long, ByVal subsectionTransparency As Double, _
                        ByVal sectionsInTimeline As Boolean, _
                        ByRef rowTops() As Double, ByRef rowHeights() As Double, _
                        ByVal textIndent As Double, ByVal sectionTextIndent As Double, ByVal subsectionTextIndent As Double, ByVal textPaddingPoints As Double, _
                        ByRef fontSections As fontSettingsType, ByRef fontSubsections As fontSettingsType, ByRef fontTasks As fontSettingsType)
    Dim i As Long
    Dim g As Long
    Dim groupTop As Double
    Dim shp As Shape
    Dim taskTop As Double
    Dim totalHeight As Double

    If displayRowCount <= 0 Then
        Exit Sub
    End If

    totalHeight = headerHeight + tasksHeight
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, leftPos, topPos + titleHeight, taskColWidth, totalHeight - titleHeight)
    shp.Fill.ForeColor.RGB = sidebarColor
    shp.Fill.transparency = ClampTransparency(sidebarTransparency) / 100
    shp.Line.Visible = msoFalse

    If groupCount > 0 And Not sectionsInTimeline Then
        For g = 1 To groupCount
            If groupRowIndices(g) > 0 Then
                groupTop = rowTops(groupRowIndices(g))
                Set shp = ws.Shapes.AddShape(msoShapeRectangle, leftPos, groupTop, taskColWidth, rowHeights(groupRowIndices(g)))
                If g <= UBound(groupTypes) And groupTypes(g) = 2 Then
                    shp.Fill.ForeColor.RGB = subsectionColor
                    shp.Fill.transparency = ClampTransparency(subsectionTransparency) / 100
                Else
                    shp.Fill.ForeColor.RGB = sectionColor
                    shp.Fill.transparency = ClampTransparency(sectionTransparency) / 100
                End If
                shp.Line.Visible = msoFalse
                shp.TextFrame2.TextRange.text = groupNames(g)
                If g <= UBound(groupTypes) And groupTypes(g) = 2 Then
                    ApplyFontToTextRange shp.TextFrame2.TextRange, fontSubsections
                Else
                    ApplyFontToTextRange shp.TextFrame2.TextRange, fontSections
                End If
                shp.TextFrame2.WordWrap = msoTrue
                shp.TextFrame2.TextRange.ParagraphFormat.WordWrap = msoTrue
                shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
                shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
                If g <= UBound(groupTypes) And groupTypes(g) = 2 Then
                    shp.TextFrame2.MarginLeft = subsectionTextIndent
                Else
                    shp.TextFrame2.MarginLeft = sectionTextIndent
                End If
                shp.TextFrame2.MarginRight = 0
                shp.TextFrame2.MarginTop = 0
                shp.TextFrame2.MarginBottom = 0
            End If
        Next g
    End If

    For i = LBound(taskNames) To UBound(taskNames)
        taskTop = rowTops(taskRowIndices(i))
        Set shp = ws.Shapes.AddShape(msoShapeRectangle, leftPos, _
                                     taskTop, taskColWidth, rowHeights(taskRowIndices(i)))
        shp.Fill.ForeColor.RGB = sidebarColor
        shp.Fill.transparency = ClampTransparency(sidebarTransparency) / 100
        shp.Line.Visible = msoFalse
        shp.TextFrame2.TextRange.text = taskNames(i)
        ApplyFontToTextRange shp.TextFrame2.TextRange, fontTasks
        shp.TextFrame2.WordWrap = msoTrue
        shp.TextFrame2.TextRange.ParagraphFormat.WordWrap = msoTrue
        shp.TextFrame2.TextRange.ParagraphFormat.SpaceBefore = 0
        shp.TextFrame2.TextRange.ParagraphFormat.SpaceAfter = 0
        shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        shp.TextFrame2.MarginLeft = textIndent
        shp.TextFrame2.MarginRight = 0
        shp.TextFrame2.MarginTop = textPaddingPoints / 2
        shp.TextFrame2.MarginBottom = textPaddingPoints / 2
    Next i
End Sub

Private Sub BuildSidebarLayout(ByRef sectionNames() As String, ByRef subsectionNames() As String, ByVal subsectionColIdx As Long, _
                               ByRef taskNames() As String, ByRef groupRowIndices() As Long, ByRef groupLengths() As Long, _
                               ByRef groupNames() As String, ByRef groupTypes() As Long, ByRef groupCount As Long, _
                               ByRef taskRowIndices() As Long, ByRef displayRowCount As Long)
    Dim i As Long
    Dim lb As Long
    Dim sectionCount As Long
    Dim sectionList() As String
    Dim sectionIndexPerRow() As Long
    Dim sectionIndex As Long
    Dim hasSections As Boolean
    Dim useSubsections As Boolean
    Dim subsectionList() As String
    Dim subsectionIndexPerRow() As Long
    Dim subsectionIndex As Long
    Dim secSubKey As String
    Dim subListCnt As Long

    groupCount = 0
    displayRowCount = 0
    On Error Resume Next
    lb = LBound(sectionNames)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ReDim taskRowIndices(lb To UBound(taskNames))
    ReDim sectionIndexPerRow(lb To UBound(sectionNames))
    ReDim subsectionIndexPerRow(lb To UBound(subsectionNames))

    hasSections = False
    For i = lb To UBound(sectionNames)
        If Len(Trim$(sectionNames(i))) > 0 Or (subsectionColIdx > 0 And Len(Trim$(subsectionNames(i))) > 0) Then
            hasSections = True
            Exit For
        End If
    Next i

    useSubsections = (subsectionColIdx > 0)

    If Not hasSections Then
        For i = lb To UBound(taskNames)
            displayRowCount = displayRowCount + 1
            taskRowIndices(i) = displayRowCount
        Next i
        Exit Sub
    End If

    If useSubsections Then
        sectionCount = 0
        For i = lb To UBound(sectionNames)
            If Len(Trim$(sectionNames(i))) = 0 Then
                sectionIndexPerRow(i) = 0
            Else
                sectionIndex = FindKeyIndex(sectionList, sectionCount, sectionNames(i))
                If sectionIndex = 0 Then
                    sectionCount = sectionCount + 1
                    ReDim Preserve sectionList(1 To sectionCount)
                    sectionList(sectionCount) = sectionNames(i)
                    sectionIndex = sectionCount
                End If
                sectionIndexPerRow(i) = sectionIndex
            End If
        Next i

        Dim j As Long
        For sectionIndex = 1 To sectionCount
            subListCnt = 0
            For i = lb To UBound(sectionNames)
                If sectionIndexPerRow(i) = sectionIndex Then
                    secSubKey = subsectionNames(i)
                    If Len(Trim$(secSubKey)) = 0 Then
                        subsectionIndexPerRow(i) = 0
                    Else
                        subsectionIndex = FindKeyIndex(subsectionList, subListCnt, secSubKey)
                        If subsectionIndex = 0 Then
                            subListCnt = subListCnt + 1
                            ReDim Preserve subsectionList(1 To subListCnt)
                            subsectionList(subListCnt) = secSubKey
                            subsectionIndex = subListCnt
                        End If
                        subsectionIndexPerRow(i) = subsectionIndex
                    End If
                End If
            Next i

            groupCount = groupCount + 1
            ReDim Preserve groupRowIndices(1 To groupCount)
            ReDim Preserve groupLengths(1 To groupCount)
            ReDim Preserve groupNames(1 To groupCount)
            ReDim Preserve groupTypes(1 To groupCount)
            displayRowCount = displayRowCount + 1
            groupRowIndices(groupCount) = displayRowCount
            groupLengths(groupCount) = 1
            groupNames(groupCount) = sectionList(sectionIndex)
            groupTypes(groupCount) = 1

            For i = lb To UBound(sectionNames)
                If sectionIndexPerRow(i) = sectionIndex And subsectionIndexPerRow(i) = 0 Then
                    displayRowCount = displayRowCount + 1
                    taskRowIndices(i) = displayRowCount
                    groupLengths(groupCount) = groupLengths(groupCount) + 1
                End If
            Next i

            For subsectionIndex = 1 To subListCnt
                groupCount = groupCount + 1
                ReDim Preserve groupRowIndices(1 To groupCount)
                ReDim Preserve groupLengths(1 To groupCount)
                ReDim Preserve groupNames(1 To groupCount)
                ReDim Preserve groupTypes(1 To groupCount)
                displayRowCount = displayRowCount + 1
                groupRowIndices(groupCount) = displayRowCount
                groupLengths(groupCount) = 0
                groupNames(groupCount) = subsectionList(subsectionIndex)
                groupTypes(groupCount) = 2

                For i = lb To UBound(sectionNames)
                    If sectionIndexPerRow(i) = sectionIndex And subsectionIndexPerRow(i) = subsectionIndex Then
                        displayRowCount = displayRowCount + 1
                        taskRowIndices(i) = displayRowCount
                        groupLengths(groupCount) = groupLengths(groupCount) + 1
                    End If
                Next i
            Next subsectionIndex
        Next sectionIndex

        If sectionCount = 0 Then
            subListCnt = 0
            For i = lb To UBound(sectionNames)
                secSubKey = subsectionNames(i)
                If Len(Trim$(secSubKey)) > 0 Then
                    subsectionIndex = FindKeyIndex(subsectionList, subListCnt, secSubKey)
                    If subsectionIndex = 0 Then
                        subListCnt = subListCnt + 1
                        ReDim Preserve subsectionList(1 To subListCnt)
                        subsectionList(subListCnt) = secSubKey
                        subsectionIndex = subListCnt
                    End If
                    subsectionIndexPerRow(i) = subsectionIndex
                Else
                    subsectionIndexPerRow(i) = 0
                End If
            Next i

            For subsectionIndex = 1 To subListCnt
                groupCount = groupCount + 1
                ReDim Preserve groupRowIndices(1 To groupCount)
                ReDim Preserve groupLengths(1 To groupCount)
                ReDim Preserve groupNames(1 To groupCount)
                ReDim Preserve groupTypes(1 To groupCount)
                displayRowCount = displayRowCount + 1
                groupRowIndices(groupCount) = displayRowCount
                groupLengths(groupCount) = 0
                groupNames(groupCount) = subsectionList(subsectionIndex)
                groupTypes(groupCount) = 2

                For i = lb To UBound(sectionNames)
                    If subsectionIndexPerRow(i) = subsectionIndex Then
                        displayRowCount = displayRowCount + 1
                        taskRowIndices(i) = displayRowCount
                        groupLengths(groupCount) = groupLengths(groupCount) + 1
                    End If
                Next i
            Next subsectionIndex
        End If

        For i = lb To UBound(sectionNames)
            If sectionIndexPerRow(i) = 0 And subsectionIndexPerRow(i) = 0 Then
                displayRowCount = displayRowCount + 1
                taskRowIndices(i) = displayRowCount
            End If
        Next i
    Else
        For i = lb To UBound(sectionNames)
            If Len(Trim$(sectionNames(i))) = 0 Then
                sectionIndexPerRow(i) = 0
            Else
                sectionIndex = FindKeyIndex(sectionList, sectionCount, sectionNames(i))
                If sectionIndex = 0 Then
                    sectionCount = sectionCount + 1
                    ReDim Preserve sectionList(1 To sectionCount)
                    sectionList(sectionCount) = sectionNames(i)
                    sectionIndex = sectionCount
                End If
                sectionIndexPerRow(i) = sectionIndex
            End If
        Next i

        For sectionIndex = 1 To sectionCount
            groupCount = groupCount + 1
            ReDim Preserve groupRowIndices(1 To groupCount)
            ReDim Preserve groupLengths(1 To groupCount)
            ReDim Preserve groupNames(1 To groupCount)
            ReDim Preserve groupTypes(1 To groupCount)
            displayRowCount = displayRowCount + 1
            groupRowIndices(groupCount) = displayRowCount
            groupLengths(groupCount) = 0
            groupNames(groupCount) = sectionList(sectionIndex)
            groupTypes(groupCount) = 1

            For i = lb To UBound(sectionNames)
                If sectionIndexPerRow(i) = sectionIndex Then
                    displayRowCount = displayRowCount + 1
                    taskRowIndices(i) = displayRowCount
                    groupLengths(groupCount) = groupLengths(groupCount) + 1
                End If
            Next i
        Next sectionIndex
    End If

    If Not useSubsections Then
        For i = lb To UBound(sectionNames)
            If sectionIndexPerRow(i) = 0 Then
                displayRowCount = displayRowCount + 1
                taskRowIndices(i) = displayRowCount
            End If
        Next i
    End If
End Sub

Private Sub BuildSidebarLayoutWithDuplicates(ByRef sectionNames() As String, ByRef subsectionNames() As String, ByVal subsectionColIdx As Long, _
                                             ByRef taskNames() As String, ByRef groupRowIndices() As Long, ByRef groupLengths() As Long, _
                                             ByRef groupNames() As String, ByRef groupTypes() As Long, ByRef groupCount As Long, _
                                             ByRef taskRowIndices() As Long, ByRef taskStackIndices() As Long, _
                                             ByRef taskStackCounts() As Long, ByRef sidebarTaskNames() As String, _
                                             ByRef sidebarTaskRowIndices() As Long, ByRef displayRowCount As Long)
    Dim i As Long
    Dim lb As Long
    Dim key As String
    Dim keyIndex As Long
    Dim keyCount As Long
    Dim keyNames() As String
    Dim keyCounts() As Long
    Dim keyIndexPerRow() As Long
    Dim sectionCount As Long
    Dim sectionList() As String
    Dim sectionIndexPerRow() As Long
    Dim sectionIndex As Long
    Dim hasSections As Boolean
    Dim useSubsections As Boolean
    Dim secSubKey As String
    Dim subsectionList() As String
    Dim subListCnt As Long
    Dim subsectionIndexPerRow() As Long
    Dim subsectionIndex As Long

    groupCount = 0
    displayRowCount = 0
    On Error Resume Next
    lb = LBound(sectionNames)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    useSubsections = (subsectionColIdx > 0)
    ReDim taskRowIndices(lb To UBound(taskNames))
    ReDim taskStackIndices(lb To UBound(taskNames))
    ReDim taskStackCounts(lb To UBound(taskNames))
    ReDim keyIndexPerRow(lb To UBound(taskNames))
    ReDim sectionIndexPerRow(lb To UBound(sectionNames))
    ReDim subsectionIndexPerRow(lb To UBound(subsectionNames))

    hasSections = False
    For i = lb To UBound(sectionNames)
        If Len(Trim$(sectionNames(i))) > 0 Or (useSubsections And Len(Trim$(subsectionNames(i))) > 0) Then
            hasSections = True
            Exit For
        End If
    Next i

    If Not hasSections Then
        For i = lb To UBound(taskNames)
            key = taskNames(i)
            keyIndex = FindKeyIndex(keyNames, keyCount, key)
            If keyIndex = 0 Then
                keyCount = keyCount + 1
                ReDim Preserve keyNames(1 To keyCount)
                ReDim Preserve keyCounts(1 To keyCount)
                ReDim Preserve sidebarTaskNames(1 To keyCount)
                ReDim Preserve sidebarTaskRowIndices(1 To keyCount)
                keyNames(keyCount) = key
                keyCounts(keyCount) = 1
                displayRowCount = displayRowCount + 1
                sidebarTaskNames(keyCount) = taskNames(i)
                sidebarTaskRowIndices(keyCount) = displayRowCount
                taskRowIndices(i) = displayRowCount
                taskStackIndices(i) = 1
                keyIndexPerRow(i) = keyCount
            Else
                keyCounts(keyIndex) = keyCounts(keyIndex) + 1
                taskRowIndices(i) = sidebarTaskRowIndices(keyIndex)
                taskStackIndices(i) = keyCounts(keyIndex)
                keyIndexPerRow(i) = keyIndex
            End If
        Next i

        For i = lb To UBound(taskNames)
            taskStackCounts(i) = keyCounts(keyIndexPerRow(i))
        Next i
        Exit Sub
    End If

    If useSubsections Then
        sectionCount = 0
        For i = lb To UBound(sectionNames)
            If Len(Trim$(sectionNames(i))) = 0 Then
                sectionIndexPerRow(i) = 0
            Else
                sectionIndex = FindKeyIndex(sectionList, sectionCount, sectionNames(i))
                If sectionIndex = 0 Then
                    sectionCount = sectionCount + 1
                    ReDim Preserve sectionList(1 To sectionCount)
                    sectionList(sectionCount) = sectionNames(i)
                    sectionIndex = sectionCount
                End If
                sectionIndexPerRow(i) = sectionIndex
            End If
        Next i
    Else
        For i = lb To UBound(sectionNames)
            If Len(Trim$(sectionNames(i))) = 0 Then
                sectionIndexPerRow(i) = 0
            Else
                sectionIndex = FindKeyIndex(sectionList, sectionCount, sectionNames(i))
                If sectionIndex = 0 Then
                    sectionCount = sectionCount + 1
                    ReDim Preserve sectionList(1 To sectionCount)
                    sectionList(sectionCount) = sectionNames(i)
                    sectionIndex = sectionCount
                End If
                sectionIndexPerRow(i) = sectionIndex
            End If
        Next i
    End If

    If useSubsections Then
        subListCnt = 0
        For sectionIndex = 1 To sectionCount
            subListCnt = 0
            For i = lb To UBound(sectionNames)
                If sectionIndexPerRow(i) = sectionIndex Then
                    secSubKey = subsectionNames(i)
                    If Len(Trim$(secSubKey)) = 0 Then
                        subsectionIndexPerRow(i) = 0
                    Else
                        subsectionIndex = FindKeyIndex(subsectionList, subListCnt, secSubKey)
                        If subsectionIndex = 0 Then
                            subListCnt = subListCnt + 1
                            ReDim Preserve subsectionList(1 To subListCnt)
                            subsectionList(subListCnt) = secSubKey
                            subsectionIndex = subListCnt
                        End If
                        subsectionIndexPerRow(i) = subsectionIndex
                    End If
                End If
            Next i

            groupCount = groupCount + 1
            ReDim Preserve groupRowIndices(1 To groupCount)
            ReDim Preserve groupLengths(1 To groupCount)
            ReDim Preserve groupNames(1 To groupCount)
            ReDim Preserve groupTypes(1 To groupCount)
            displayRowCount = displayRowCount + 1
            groupRowIndices(groupCount) = displayRowCount
            groupLengths(groupCount) = 1
            groupNames(groupCount) = sectionList(sectionIndex)
            groupTypes(groupCount) = 1

            For i = lb To UBound(sectionNames)
                If sectionIndexPerRow(i) = sectionIndex And subsectionIndexPerRow(i) = 0 Then
                    key = sectionNames(i) & "|" & "|" & taskNames(i)
                    keyIndex = FindKeyIndex(keyNames, keyCount, key)
                    If keyIndex = 0 Then
                        keyCount = keyCount + 1
                        ReDim Preserve keyNames(1 To keyCount)
                        ReDim Preserve keyCounts(1 To keyCount)
                        ReDim Preserve sidebarTaskNames(1 To keyCount)
                        ReDim Preserve sidebarTaskRowIndices(1 To keyCount)
                        keyNames(keyCount) = key
                        keyCounts(keyCount) = 1
                        displayRowCount = displayRowCount + 1
                        sidebarTaskNames(keyCount) = taskNames(i)
                        sidebarTaskRowIndices(keyCount) = displayRowCount
                        taskRowIndices(i) = displayRowCount
                        taskStackIndices(i) = 1
                        keyIndexPerRow(i) = keyCount
                    Else
                        keyCounts(keyIndex) = keyCounts(keyIndex) + 1
                        taskRowIndices(i) = sidebarTaskRowIndices(keyIndex)
                        taskStackIndices(i) = keyCounts(keyIndex)
                        keyIndexPerRow(i) = keyIndex
                    End If
                    groupLengths(groupCount) = groupLengths(groupCount) + 1
                End If
            Next i

            For subsectionIndex = 1 To subListCnt
                groupCount = groupCount + 1
                ReDim Preserve groupRowIndices(1 To groupCount)
                ReDim Preserve groupLengths(1 To groupCount)
                ReDim Preserve groupNames(1 To groupCount)
                ReDim Preserve groupTypes(1 To groupCount)
                displayRowCount = displayRowCount + 1
                groupRowIndices(groupCount) = displayRowCount
                groupLengths(groupCount) = 0
                groupNames(groupCount) = subsectionList(subsectionIndex)
                groupTypes(groupCount) = 2

                For i = lb To UBound(sectionNames)
                    If sectionIndexPerRow(i) = sectionIndex And subsectionIndexPerRow(i) = subsectionIndex Then
                        key = sectionNames(i) & "|" & subsectionNames(i) & "|" & taskNames(i)
                        keyIndex = FindKeyIndex(keyNames, keyCount, key)
                        If keyIndex = 0 Then
                            keyCount = keyCount + 1
                            ReDim Preserve keyNames(1 To keyCount)
                            ReDim Preserve keyCounts(1 To keyCount)
                            ReDim Preserve sidebarTaskNames(1 To keyCount)
                            ReDim Preserve sidebarTaskRowIndices(1 To keyCount)
                            keyNames(keyCount) = key
                            keyCounts(keyCount) = 1
                            displayRowCount = displayRowCount + 1
                            sidebarTaskNames(keyCount) = taskNames(i)
                            sidebarTaskRowIndices(keyCount) = displayRowCount
                            taskRowIndices(i) = displayRowCount
                            taskStackIndices(i) = 1
                            keyIndexPerRow(i) = keyCount
                        Else
                            keyCounts(keyIndex) = keyCounts(keyIndex) + 1
                            taskRowIndices(i) = sidebarTaskRowIndices(keyIndex)
                            taskStackIndices(i) = keyCounts(keyIndex)
                            keyIndexPerRow(i) = keyIndex
                        End If
                        groupLengths(groupCount) = groupLengths(groupCount) + 1
                    End If
                Next i
            Next subsectionIndex
        Next sectionIndex

        If sectionCount = 0 Then
            subListCnt = 0
            For i = lb To UBound(sectionNames)
                secSubKey = subsectionNames(i)
                If Len(Trim$(secSubKey)) > 0 Then
                    subsectionIndex = FindKeyIndex(subsectionList, subListCnt, secSubKey)
                    If subsectionIndex = 0 Then
                        subListCnt = subListCnt + 1
                        ReDim Preserve subsectionList(1 To subListCnt)
                        subsectionList(subListCnt) = secSubKey
                        subsectionIndex = subListCnt
                    End If
                    subsectionIndexPerRow(i) = subsectionIndex
                Else
                    subsectionIndexPerRow(i) = 0
                End If
            Next i

            For subsectionIndex = 1 To subListCnt
                groupCount = groupCount + 1
                ReDim Preserve groupRowIndices(1 To groupCount)
                ReDim Preserve groupLengths(1 To groupCount)
                ReDim Preserve groupNames(1 To groupCount)
                ReDim Preserve groupTypes(1 To groupCount)
                displayRowCount = displayRowCount + 1
                groupRowIndices(groupCount) = displayRowCount
                groupLengths(groupCount) = 0
                groupNames(groupCount) = subsectionList(subsectionIndex)
                groupTypes(groupCount) = 2

                For i = lb To UBound(sectionNames)
                    If subsectionIndexPerRow(i) = subsectionIndex Then
                        key = "|" & subsectionNames(i) & "|" & taskNames(i)
                        keyIndex = FindKeyIndex(keyNames, keyCount, key)
                        If keyIndex = 0 Then
                            keyCount = keyCount + 1
                            ReDim Preserve keyNames(1 To keyCount)
                            ReDim Preserve keyCounts(1 To keyCount)
                            ReDim Preserve sidebarTaskNames(1 To keyCount)
                            ReDim Preserve sidebarTaskRowIndices(1 To keyCount)
                            keyNames(keyCount) = key
                            keyCounts(keyCount) = 1
                            displayRowCount = displayRowCount + 1
                            sidebarTaskNames(keyCount) = taskNames(i)
                            sidebarTaskRowIndices(keyCount) = displayRowCount
                            taskRowIndices(i) = displayRowCount
                            taskStackIndices(i) = 1
                            keyIndexPerRow(i) = keyCount
                        Else
                            keyCounts(keyIndex) = keyCounts(keyIndex) + 1
                            taskRowIndices(i) = sidebarTaskRowIndices(keyIndex)
                            taskStackIndices(i) = keyCounts(keyIndex)
                            keyIndexPerRow(i) = keyIndex
                        End If
                        groupLengths(groupCount) = groupLengths(groupCount) + 1
                    End If
                Next i
            Next subsectionIndex
        End If

        For i = lb To UBound(sectionNames)
            If sectionIndexPerRow(i) = 0 And subsectionIndexPerRow(i) = 0 Then
                key = "__EMPTY__|" & taskNames(i)
                keyIndex = FindKeyIndex(keyNames, keyCount, key)
                If keyIndex = 0 Then
                    keyCount = keyCount + 1
                    ReDim Preserve keyNames(1 To keyCount)
                    ReDim Preserve keyCounts(1 To keyCount)
                    ReDim Preserve sidebarTaskNames(1 To keyCount)
                    ReDim Preserve sidebarTaskRowIndices(1 To keyCount)
                    keyNames(keyCount) = key
                    keyCounts(keyCount) = 1
                    displayRowCount = displayRowCount + 1
                    sidebarTaskNames(keyCount) = taskNames(i)
                    sidebarTaskRowIndices(keyCount) = displayRowCount
                    taskRowIndices(i) = displayRowCount
                    taskStackIndices(i) = 1
                    keyIndexPerRow(i) = keyCount
                Else
                    keyCounts(keyIndex) = keyCounts(keyIndex) + 1
                    taskRowIndices(i) = sidebarTaskRowIndices(keyIndex)
                    taskStackIndices(i) = keyCounts(keyIndex)
                    keyIndexPerRow(i) = keyIndex
                End If
            End If
        Next i
    Else
        For sectionIndex = 1 To sectionCount
            groupCount = groupCount + 1
            ReDim Preserve groupRowIndices(1 To groupCount)
            ReDim Preserve groupLengths(1 To groupCount)
            ReDim Preserve groupNames(1 To groupCount)
            ReDim Preserve groupTypes(1 To groupCount)
            displayRowCount = displayRowCount + 1
            groupRowIndices(groupCount) = displayRowCount
            groupLengths(groupCount) = 0
            groupNames(groupCount) = sectionList(sectionIndex)
            groupTypes(groupCount) = 1

            For i = lb To UBound(sectionNames)
                If sectionIndexPerRow(i) = sectionIndex Then
                    key = sectionNames(i) & "|" & taskNames(i)
                    keyIndex = FindKeyIndex(keyNames, keyCount, key)
                    If keyIndex = 0 Then
                        keyCount = keyCount + 1
                        ReDim Preserve keyNames(1 To keyCount)
                        ReDim Preserve keyCounts(1 To keyCount)
                        ReDim Preserve sidebarTaskNames(1 To keyCount)
                        ReDim Preserve sidebarTaskRowIndices(1 To keyCount)
                        keyNames(keyCount) = key
                        keyCounts(keyCount) = 1
                        displayRowCount = displayRowCount + 1
                        sidebarTaskNames(keyCount) = taskNames(i)
                        sidebarTaskRowIndices(keyCount) = displayRowCount
                        taskRowIndices(i) = displayRowCount
                        taskStackIndices(i) = 1
                        keyIndexPerRow(i) = keyCount
                    Else
                        keyCounts(keyIndex) = keyCounts(keyIndex) + 1
                        taskRowIndices(i) = sidebarTaskRowIndices(keyIndex)
                        taskStackIndices(i) = keyCounts(keyIndex)
                        keyIndexPerRow(i) = keyIndex
                    End If
                    groupLengths(groupCount) = groupLengths(groupCount) + 1
                End If
            Next i
        Next sectionIndex
    End If

    If Not useSubsections Then
        For i = lb To UBound(sectionNames)
            If sectionIndexPerRow(i) = 0 Then
                key = "__EMPTY__|" & taskNames(i)
                keyIndex = FindKeyIndex(keyNames, keyCount, key)
                If keyIndex = 0 Then
                    keyCount = keyCount + 1
                    ReDim Preserve keyNames(1 To keyCount)
                    ReDim Preserve keyCounts(1 To keyCount)
                    ReDim Preserve sidebarTaskNames(1 To keyCount)
                    ReDim Preserve sidebarTaskRowIndices(1 To keyCount)
                    keyNames(keyCount) = key
                    keyCounts(keyCount) = 1
                    displayRowCount = displayRowCount + 1
                    sidebarTaskNames(keyCount) = taskNames(i)
                    sidebarTaskRowIndices(keyCount) = displayRowCount
                    taskRowIndices(i) = displayRowCount
                    taskStackIndices(i) = 1
                    keyIndexPerRow(i) = keyCount
                Else
                    keyCounts(keyIndex) = keyCounts(keyIndex) + 1
                    taskRowIndices(i) = sidebarTaskRowIndices(keyIndex)
                    taskStackIndices(i) = keyCounts(keyIndex)
                    keyIndexPerRow(i) = keyIndex
                End If
            End If
        Next i
    End If

    For i = lb To UBound(taskNames)
        taskStackCounts(i) = keyCounts(keyIndexPerRow(i))
    Next i
End Sub

Private Function FindKeyIndex(ByRef keyNames() As String, ByVal keyCount As Long, ByVal targetKey As String) As Long
    Dim i As Long

    For i = 1 To keyCount
        If keyNames(i) = targetKey Then
            FindKeyIndex = i
            Exit Function
        End If
    Next i

    FindKeyIndex = 0
End Function

'' Расчёт количества строк переноса с учётом фактического шрифта (имя, размер, bold, italic).
Private Function EstimateWrappedLinesForFont(ByVal labelText As String, ByVal availableWidth As Double, _
                                             ByRef fontSettings As fontSettingsType) As Long
    Dim paragraphs() As String
    Dim p As Long
    Dim totalLines As Long
    Dim normalized As String

    If Len(Trim$(labelText)) = 0 Then
        EstimateWrappedLinesForFont = 1
        Exit Function
    End If
    availableWidth = Application.WorksheetFunction.Max(1, availableWidth)
    normalized = Replace(labelText, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)
    paragraphs = Split(normalized, vbLf)
    totalLines = 0
    For p = LBound(paragraphs) To UBound(paragraphs)
        If Len(Trim$(paragraphs(p))) > 0 Then
            totalLines = totalLines + EstimateWrappedLinesSingleForFont(Trim$(paragraphs(p)), availableWidth, fontSettings)
        End If
    Next p
    If totalLines > 0 Then
        EstimateWrappedLinesForFont = totalLines
        Exit Function
    End If
    EstimateWrappedLinesForFont = EstimateWrappedLinesSingleForFont(Trim$(Replace(Replace(labelText, vbCrLf, " "), vbCr, " ")), availableWidth, fontSettings)
    If EstimateWrappedLinesForFont < 1 Then EstimateWrappedLinesForFont = 1
End Function

Private Function EstimateWrappedLinesSingleForFont(ByVal labelText As String, ByVal availableWidth As Double, _
                                                   ByRef fontSettings As fontSettingsType) As Long
    Dim words() As String
    Dim currentLine As String
    Dim i As Long
    Dim lineCount As Long
    Dim candidate As String
    Dim wordWid As Double
    Dim linesForWord As Long

    If Len(Trim$(labelText)) = 0 Then
        EstimateWrappedLinesSingleForFont = 1
        Exit Function
    End If
    availableWidth = Application.WorksheetFunction.Max(1, availableWidth)
    words = Split(labelText, " ")
    lineCount = 1
    currentLine = ""
    For i = LBound(words) To UBound(words)
        If currentLine = "" Then
            candidate = words(i)
        Else
            candidate = currentLine & " " & words(i)
        End If
        If MeasureTextWidthWithFont(candidate, fontSettings.name, CLng(fontSettings.Size), fontSettings.bold, fontSettings.italic) <= availableWidth Then
            currentLine = candidate
        Else
            If Len(currentLine) > 0 Then lineCount = lineCount + 1
            wordWid = MeasureTextWidthWithFont(words(i), fontSettings.name, CLng(fontSettings.Size), fontSettings.bold, fontSettings.italic)
            If wordWid > availableWidth Then
                linesForWord = Application.WorksheetFunction.RoundUp(wordWid / availableWidth, 0)
                lineCount = lineCount + linesForWord - 1
                currentLine = ""
            Else
                currentLine = words(i)
            End If
        End If
    Next i
    EstimateWrappedLinesSingleForFont = lineCount
End Function

Private Function EstimateWrappedLines(ByVal labelText As String, ByVal availableWidth As Double, _
                                      ByVal fontSize As Long) As Long
    Dim words() As String
    Dim currentLine As String
    Dim i As Long
    Dim lineCount As Long
    Dim candidate As String
    Dim wordWid As Double
    Dim linesForWord As Long
    Dim paragraphs() As String
    Dim p As Long
    Dim totalLines As Long
    Dim normalized As String

    If Len(Trim$(labelText)) = 0 Then
        EstimateWrappedLines = 1
        Exit Function
    End If

    availableWidth = Application.WorksheetFunction.Max(1, availableWidth)

    normalized = Replace(labelText, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)
    paragraphs = Split(normalized, vbLf)
    totalLines = 0
    For p = LBound(paragraphs) To UBound(paragraphs)
        If Len(Trim$(paragraphs(p))) > 0 Then
            totalLines = totalLines + EstimateWrappedLinesSingle(Trim$(paragraphs(p)), availableWidth, fontSize)
        End If
    Next p
    If totalLines > 0 Then
        EstimateWrappedLines = totalLines
        Exit Function
    End If

    EstimateWrappedLines = EstimateWrappedLinesSingle(Trim$(Replace(Replace(labelText, vbCrLf, " "), vbCr, " ")), availableWidth, fontSize)
    If EstimateWrappedLines < 1 Then EstimateWrappedLines = 1
End Function

Private Function EstimateWrappedLinesSingle(ByVal labelText As String, ByVal availableWidth As Double, _
                                            ByVal fontSize As Long) As Long
    Dim words() As String
    Dim currentLine As String
    Dim i As Long
    Dim lineCount As Long
    Dim candidate As String
    Dim wordWid As Double
    Dim linesForWord As Long

    If Len(Trim$(labelText)) = 0 Then
        EstimateWrappedLinesSingle = 1
        Exit Function
    End If

    availableWidth = Application.WorksheetFunction.Max(1, availableWidth)
    words = Split(labelText, " ")
    lineCount = 1
    currentLine = ""

    For i = LBound(words) To UBound(words)
        If currentLine = "" Then
            candidate = words(i)
        Else
            candidate = currentLine & " " & words(i)
        End If

        If EstimateTextWidth(candidate, fontSize) <= availableWidth Then
            currentLine = candidate
        Else
            If Len(currentLine) > 0 Then lineCount = lineCount + 1
            wordWid = EstimateTextWidth(words(i), fontSize)
            If wordWid > availableWidth Then
                linesForWord = Application.WorksheetFunction.RoundUp(wordWid / availableWidth, 0)
                lineCount = lineCount + linesForWord - 1
                currentLine = ""
            Else
                currentLine = words(i)
            End If
        End If
    Next i

    EstimateWrappedLinesSingle = lineCount
End Function

Private Function TruncateEventLabelToMaxLines(ByVal labelText As String, ByVal availableWidth As Double, _
                                              ByVal fontSize As Long, ByVal maxLines As Long) As String
    Dim words() As String
    Dim currentLine As String
    Dim i As Long
    Dim linesInResult As Long
    Dim candidate As String
    Dim result As String
    Dim j As Long
    Dim effectiveWidth As Double
    Dim ellipsisWidth As Double

    If Len(Trim$(labelText)) = 0 Or maxLines < 1 Then
        TruncateEventLabelToMaxLines = labelText
        Exit Function
    End If

    If EstimateWrappedLines(labelText, availableWidth, fontSize) <= maxLines And _
       EstimateTextWidth(labelText, fontSize) <= maxLines * availableWidth Then
        TruncateEventLabelToMaxLines = labelText
        Exit Function
    End If

    ellipsisWidth = EstimateTextWidth("...", fontSize)

    If InStr(labelText, " ") = 0 And EstimateTextWidth(labelText, fontSize) > maxLines * availableWidth Then
        For j = Len(labelText) To 1 Step -1
            If EstimateTextWidth(Left$(labelText, j) & "...", fontSize) <= maxLines * availableWidth Then
                TruncateEventLabelToMaxLines = Left$(labelText, j) & "..."
                Exit Function
            End If
        Next j
        TruncateEventLabelToMaxLines = "..."
        Exit Function
    End If

    words = Split(labelText, " ")
    linesInResult = 0
    currentLine = ""
    result = ""

    For i = LBound(words) To UBound(words)
        If currentLine = "" Then
            candidate = words(i)
        Else
            candidate = currentLine & " " & words(i)
        End If
        If linesInResult = maxLines - 1 Then
            effectiveWidth = availableWidth - ellipsisWidth
        Else
            effectiveWidth = availableWidth
        End If

        If EstimateTextWidth(candidate, fontSize) <= effectiveWidth Then
            currentLine = candidate
        Else
            If Len(currentLine) > 0 Then
                linesInResult = linesInResult + 1
                If linesInResult > maxLines Then
                    TruncateEventLabelToMaxLines = result & "..."
                    Exit Function
                End If
                If Len(result) > 0 Then result = result & " "
                result = result & currentLine
            End If
            currentLine = words(i)
        End If
    Next i

    If Len(currentLine) > 0 Then
        linesInResult = linesInResult + 1
        If linesInResult > maxLines Then
            TruncateEventLabelToMaxLines = result & "..."
            Exit Function
        End If
        If Len(result) > 0 Then result = result & " "
        result = result & currentLine
    End If

    TruncateEventLabelToMaxLines = result
End Function

Private Function TruncateTextWithEllipsis(ByVal labelText As String, ByVal availableWidth As Double, _
                                          ByVal fontSize As Long, ByVal padding As Double) As String
    Dim truncated As String
    Dim ellipsis As String
    Dim i As Long
    Dim maxWidth As Double

    ellipsis = "..."
    maxWidth = Application.WorksheetFunction.Max(1, availableWidth - padding * 2)

    If EstimateTextWidth(labelText, fontSize) <= maxWidth Then
        TruncateTextWithEllipsis = labelText
        Exit Function
    End If

    If EstimateTextWidth(ellipsis, fontSize) > maxWidth Then
        TruncateTextWithEllipsis = ""
        Exit Function
    End If

    truncated = labelText
    For i = Len(labelText) To 1 Step -1
        truncated = Left$(labelText, i) & ellipsis
        If EstimateTextWidth(truncated, fontSize) <= maxWidth Then
            TruncateTextWithEllipsis = truncated
            Exit Function
        End If
    Next i

    TruncateTextWithEllipsis = ellipsis
End Function

Private Sub ComputeRowTops(ByVal baseTop As Double, ByRef rowHeights() As Double, _
                           ByRef rowTops() As Double, ByRef totalHeight As Double)
    Dim i As Long
    Dim currentTop As Double

    If UBound(rowHeights) < LBound(rowHeights) Then
        totalHeight = 0
        Exit Sub
    End If

    ReDim rowTops(LBound(rowHeights) To UBound(rowHeights))
    currentTop = baseTop
    For i = LBound(rowHeights) To UBound(rowHeights)
        rowTops(i) = currentTop
        currentTop = currentTop + rowHeights(i)
    Next i
    totalHeight = currentTop - baseTop
End Sub

Private Function GetSidebarRowTextHeight(ByVal taskName As String, ByVal columnWidth As Double, _
                                         ByRef fontTasks As fontSettingsType, ByVal leftIndentPoints As Double, _
                                         ByVal verticalPaddingPoints As Double) As Double
    Dim availableWidth As Double
    Dim lines As Long
    Dim fontSize As Long
    Const LINE_GAP_PT As Double = 1
    Const LINE_HEIGHT_FACTOR As Double = 1.25
    Const TEXT_AREA_BUFFER_PT As Double = 4

    fontSize = CLng(fontTasks.Size)
    ' Ширина текста = колонка минус отступ слева (настройка "Отступ текста задачи от левого края в сайдбаре")
    ' минус запас на внутренние отступы TextFrame и погрешность
    availableWidth = Application.WorksheetFunction.Max(1, columnWidth - leftIndentPoints - TEXT_AREA_BUFFER_PT)
    lines = EstimateWrappedLinesForFont(taskName, availableWidth, fontTasks)
    If lines < 1 Then lines = 1
    GetSidebarRowTextHeight = lines * (fontSize * LINE_HEIGHT_FACTOR) + (lines - 1) * LINE_GAP_PT + verticalPaddingPoints
End Function

Private Sub ComputeSidebarWidths(ByRef sectionNames() As String, ByRef subsectionNames() As String, _
                                 ByRef groupNames() As String, ByRef groupTypes() As Long, ByVal groupCount As Long, _
                                 ByRef fontSections As fontSettingsType, ByRef fontSubsections As fontSettingsType, _
                                 ByRef taskNames() As String, ByVal taskFontSize As Long, ByVal padding As Double, _
                                 ByRef sectionColWidth As Double, ByRef taskColWidth As Double)
    Dim i As Long
    Dim widthCandidate As Double
    Dim lb As Long
    Dim grpFontSize As Long

    sectionColWidth = 0
    taskColWidth = 0

    On Error Resume Next
    lb = LBound(sectionNames)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    For i = 1 To groupCount
        If i <= UBound(groupTypes) And groupTypes(i) = 2 Then
            grpFontSize = CLng(fontSubsections.Size)
        Else
            grpFontSize = CLng(fontSections.Size)
        End If
        widthCandidate = EstimateTextWidth(groupNames(i), grpFontSize) + padding * 2
        If widthCandidate > sectionColWidth Then
            sectionColWidth = widthCandidate
        End If
    Next i
    If groupCount <= 0 Then
        For i = lb To UBound(sectionNames)
            widthCandidate = EstimateTextWidth(sectionNames(i), CLng(fontSections.Size)) + padding * 2
            If widthCandidate > sectionColWidth Then sectionColWidth = widthCandidate
        Next i
    End If

    For i = lb To UBound(taskNames)
        widthCandidate = EstimateTextWidth(taskNames(i), taskFontSize) + padding * 2
        If widthCandidate > taskColWidth Then
            taskColWidth = widthCandidate
        End If
    Next i
End Sub

' Returns vertical offset from (rowTop + trackPadPoints) to the top of the bar.
' НОВАЯ ЛОГИКА: позиционирование снизу вверх с учетом новой структуры зазоров
' Высота трека складывается снизу вверх:
' 1. trackPadPoints (зазор между баром и нижней границей трека) - только для нижнего бара
' 2. Высота бара
' 3. GAP_1MM_PT (зазор между баром и первым ярусом) + высота первого яруса (если есть)
' 4. GAP_1MM_PT (зазор между первым и вторым ярусом) + высота второго яруса (если есть)
' 5. GAP_1MM_PT (зазор между вторым ярусом и верхней границей трека) - только для верхнего бара
' 6. Если есть бар сверху: зазор 1мм до верхнего бара
' 7. Если есть бар снизу: зазор из настройки от верхнего элемента нижнего бара
Private Function ComputeBarTopWithEventReserve(ByVal taskIndex As Long, ByVal rowIndex As Long, ByVal stackPos As Long, ByVal barHeight As Double, _
    ByVal barGapPoints As Double, ByVal trackPadPoints As Double, ByVal rowCount As Long, ByRef taskRowIndices() As Long, _
    ByRef taskStackIndices() As Long, ByRef eventLabelReserveForTask() As Double, _
    ByRef isSectionRow() As Boolean, ByRef rowTops() As Double, ByRef rowHeights() As Double, _
    ByVal barHeightNominal As Double, ByVal barDisplayMode As String, _
    ByRef maxFirstTierLabelHeightForTask() As Double, ByRef firstTierReserveForTask() As Double, _
    ByRef secondTierHeightForTask() As Double, ByVal eventLabelGapTopPoints As Double, _
    ByVal eventLabelGapBottomPoints As Double) As Double
    Dim p As Long
    Dim taskIdxAtPos As Long
    Dim bottomOfTrack As Double
    Dim currentBottom As Double
    Dim stackCount As Long
    
    ' Находим нижнюю границу трека
    bottomOfTrack = rowTops(rowIndex) + rowHeights(rowIndex) - trackPadPoints
    
    ' Находим количество баров в стеке для этой строки
    stackCount = 0
    Dim taskIdx As Long
    For taskIdx = 1 To rowCount
        Dim taskRowIdx As Long
        If UBound(taskRowIndices) >= taskIdx Then
            taskRowIdx = taskRowIndices(taskIdx)
            If taskRowIdx = rowIndex Then
                If UBound(taskStackIndices) >= taskIdx Then
                    If taskStackIndices(taskIdx) > stackCount Then
                        stackCount = taskStackIndices(taskIdx)
                    End If
                End If
            End If
        End If
    Next taskIdx
    If stackCount = 0 Then stackCount = 1
    
    ' Вычисляем позицию снизу вверх: от последнего бара к текущему
    currentBottom = bottomOfTrack
    For p = stackCount To stackPos + 1 Step -1
        ' Находим задачу на позиции p
        taskIdxAtPos = FindTaskAtStackPos(rowCount, taskRowIndices, taskStackIndices, rowIndex, p)
        If taskIdxAtPos > 0 Then
            ' Вычисляем высоту бара на позиции p в зависимости от режима отображения
            Dim barH As Double
            If LCase$(barDisplayMode) = "уменьшение высоты" And stackCount > 1 Then
                ' В режиме "уменьшение высоты" зазор уже учтен в высоте бара
                barH = (barHeightNominal - (barGapPoints * (stackCount - 1))) / stackCount
                If barH < 1 Then barH = 1
            Else
                barH = barHeightNominal
            End If
            
            ' Определяем высоту первого яруса для этого бара
            Dim firstTierHeight As Double
            firstTierHeight = 0
            If taskIdxAtPos <= UBound(maxFirstTierLabelHeightForTask) And maxFirstTierLabelHeightForTask(taskIdxAtPos) > 0 Then
                firstTierHeight = maxFirstTierLabelHeightForTask(taskIdxAtPos)
            ElseIf taskIdxAtPos <= UBound(firstTierReserveForTask) And firstTierReserveForTask(taskIdxAtPos) > 0 Then
                firstTierHeight = firstTierReserveForTask(taskIdxAtPos) - eventLabelGapTopPoints - eventLabelGapBottomPoints
                If firstTierHeight < 0 Then firstTierHeight = 0
            End If
            
            ' Определяем высоту второго яруса для этого бара
            Dim secondTierHeight As Double
            secondTierHeight = 0
            If taskIdxAtPos <= UBound(secondTierHeightForTask) Then
                secondTierHeight = secondTierHeightForTask(taskIdxAtPos)
            End If
            
            ' Сдвигаемся вверх снизу вверх согласно новой логике:
            ' 1. Для нижнего бара: trackPadPoints уже учтен в bottomOfTrack
            ' 2. Высота бара
            currentBottom = currentBottom - barH
            
            ' 3. Если есть первый ярус: зазор 1мм + высота первого яруса
            If firstTierHeight > 0 Then
                currentBottom = currentBottom - GAP_1MM_PT - firstTierHeight
            End If
            
            ' 4. Если есть второй ярус: зазор 1мм + высота второго яруса
            If secondTierHeight > 0 Then
                currentBottom = currentBottom - GAP_1MM_PT - secondTierHeight
            End If
            
            ' 5. Зазор между барами: добавляется после каждого бара (кроме верхнего).
            ' Если у текущего бара есть описание (ярусы), используем не меньше trackPadPoints (настройка "Минимальный отступ бара от границ трека, мм").
            If p > 1 Then
                Dim gapBetweenBars As Double
                gapBetweenBars = barGapPoints
                If (firstTierHeight > 0 Or secondTierHeight > 0) And trackPadPoints > gapBetweenBars Then
                    gapBetweenBars = trackPadPoints
                End If
                currentBottom = currentBottom - gapBetweenBars
            End If
            
            ' 6. Для верхнего бара: всегда отступ от верхней границы трека — настройка "Минимальный отступ бара от границ трека, мм"
            If p = 1 Then
                currentBottom = currentBottom - trackPadPoints
            End If
        End If
    Next p
    
    ' Для текущего бара: позиция верха = текущая нижняя позиция - высота бара
    Dim barTop As Double
    barTop = currentBottom - barHeight
    
    ' Возвращаем смещение от (rowTop + trackPadPoints)
    ComputeBarTopWithEventReserve = barTop - (rowTops(rowIndex) + trackPadPoints)
End Function

Private Function FindTaskAtStackPos(ByVal rowCount As Long, ByRef taskRowIndices() As Long, _
    ByRef taskStackIndices() As Long, ByVal rowIndex As Long, ByVal stackPos As Long) As Long
    Dim j As Long
    FindTaskAtStackPos = 0
    For j = 1 To rowCount
        If taskRowIndices(j) = rowIndex And taskStackIndices(j) = stackPos Then
            FindTaskAtStackPos = j
            Exit Function
        End If
    Next j
End Function

Private Function DrawTaskBar(ByVal ws As Worksheet, ByVal leftPos As Double, ByVal topPos As Double, _
                             ByVal width As Double, ByVal height As Double, _
                             ByVal taskName As String, ByVal sectionName As String, _
                             ByVal fillColor As Long, ByVal lineColor As Long, _
                             ByVal showBorder As Boolean, ByRef fontTasks As fontSettingsType, _
                             ByVal barShapeType As MsoAutoShapeType) As Shape
    Dim shp As Shape

    Set shp = ws.Shapes.AddShape(barShapeType, leftPos, topPos, width, height)
    If barShapeType = msoShapePentagon Then
        shp.Rotation = 0
    End If
    shp.Fill.ForeColor.RGB = fillColor
    If showBorder Then
        shp.Line.ForeColor.RGB = DarkenColor(fillColor, 0.2)
        shp.Line.Weight = 0.75
        shp.Line.Visible = msoTrue
    Else
        shp.Line.Visible = msoFalse
    End If
    shp.TextFrame2.TextRange.text = taskName
    ApplyFontToTextRange shp.TextFrame2.TextRange, fontTasks
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.MarginTop = 0
    shp.TextFrame2.MarginBottom = 0
    shp.AlternativeText = sectionName

    Set DrawTaskBar = shp
End Function

' Для строки задачи: если между двумя соседними событиями с подписями есть третье и его подпись перекрывает обе —
' добавляет резерв высоты под второй ярус (подпись среднего события выше) и увеличивает высоту строки.
Private Sub AddSecondTierReserveForOverlappingMiddle(ByRef dataTable As ListObject, ByRef taskRowMap() As Long, _
    ByRef taskRowIndices() As Long, ByRef taskStackIndices() As Long, ByVal rowCount As Long, ByVal displayRowCount As Long, _
    ByRef isSectionRow() As Boolean, ByRef eventColumnIndices() As Long, ByRef eventDateColumnIndices() As Long, _
    ByRef eventDescColumnIndices() As Long, ByRef timelineData As timelineData, ByVal timelineStart As Date, _
    ByVal timelineEnd As Date, ByRef rowHeights() As Double, ByRef eventLabelReserveForTask() As Double, _
    ByVal eventLabelWidthPoints As Double, ByVal oneTierReservePoints As Double, ByVal eventLabelMaxLines As Long, _
    ByRef fontEventDesc As fontSettingsType, ByRef secondTierHeightForTask() As Double)
    Const MIN_GAP As Double = 2
    Dim i As Long, rowIndex As Long, dataRowIndex As Long
    Dim eventIdx As Long, evCount As Long, k As Long, t As Long
    Dim eventName As String, eventDateValue As Variant, eventDate As Date, eventLabelText As String
    Dim evCenters() As Double, lblWidths() As Double, evLabelTexts() As String
    Dim overlapLeft As Boolean, overlapRight As Boolean
    Dim tmpCenter As Double, tmpWid As Double, tmpText As String
    Dim maxLblW As Double

    If rowCount < 1 Then Exit Sub
    maxLblW = IIf(eventLabelWidthPoints > 0, eventLabelWidthPoints, 40 * 2.83465)
    ReDim evCenters(1 To 20)
    ReDim lblWidths(1 To 20)
    ReDim evLabelTexts(1 To 20)

    For i = 1 To rowCount
        rowIndex = taskRowIndices(i)
        If rowIndex < 1 Or rowIndex > displayRowCount Then GoTo NextTaskRow
        If isSectionRow(rowIndex) Then GoTo NextTaskRow
        dataRowIndex = taskRowMap(i)
        evCount = 0
        For eventIdx = 1 To 10
            If eventColumnIndices(eventIdx) <= 0 Or eventDateColumnIndices(eventIdx) <= 0 Or eventDescColumnIndices(eventIdx) <= 0 Then GoTo NextEvCol
            eventName = Trim$(CStr(dataTable.ListColumns(eventColumnIndices(eventIdx)).DataBodyRange.Cells(dataRowIndex, 1).value))
            eventDateValue = dataTable.ListColumns(eventDateColumnIndices(eventIdx)).DataBodyRange.Cells(dataRowIndex, 1).value
            eventLabelText = Trim$(CStr(dataTable.ListColumns(eventDescColumnIndices(eventIdx)).DataBodyRange.Cells(dataRowIndex, 1).value))
            If Len(eventName) = 0 Or Not IsDate(eventDateValue) Or Len(eventLabelText) = 0 Then GoTo NextEvCol
            eventDate = CDate(eventDateValue)
            If eventDate < timelineStart Or eventDate > timelineEnd Then GoTo NextEvCol
            evCount = evCount + 1
            If evCount > UBound(evCenters) Then Exit For
            evCenters(evCount) = GetTimelineOffset(timelineData, eventDate + 0.5)
            evLabelTexts(evCount) = eventLabelText
            lblWidths(evCount) = Application.WorksheetFunction.Min(maxLblW, _
                Application.WorksheetFunction.Max(maxLblW * 0.5, MeasureTextWidthWithFont(eventLabelText, fontEventDesc.name, CLng(fontEventDesc.Size), fontEventDesc.bold, fontEventDesc.italic) + 8))
NextEvCol:
        Next eventIdx
        If evCount < 3 Then GoTo NextTaskRow
        ' Сортируем по evCenter
        For k = 1 To evCount - 1
            For t = k + 1 To evCount
                If evCenters(t) < evCenters(k) Then
                    tmpCenter = evCenters(k)
                    tmpWid = lblWidths(k)
                    tmpText = evLabelTexts(k)
                    evCenters(k) = evCenters(t)
                    lblWidths(k) = lblWidths(t)
                    evLabelTexts(k) = evLabelTexts(t)
                    evCenters(t) = tmpCenter
                    lblWidths(t) = tmpWid
                    evLabelTexts(t) = tmpText
                End If
            Next t
        Next k
        For k = 2 To evCount - 1
            overlapLeft = (evCenters(k) - evCenters(k - 1)) < (lblWidths(k) / 2 + lblWidths(k - 1) / 2 + MIN_GAP)
            overlapRight = (evCenters(k + 1) - evCenters(k)) < (lblWidths(k) / 2 + lblWidths(k + 1) / 2 + MIN_GAP)
            If overlapLeft And overlapRight Then
                ' Добавляем только высоту одного яруса подписи (без отступов, они уже учтены в базовом резерве)
                ' Вычисляем высоту подписи по фактическому количеству строк среднего события
                ' Используем фактическую ширину подписи (как в основном цикле) для точного расчета
                Dim secondTierHeight As Double
                Dim actualLines As Long
                Dim fontSize As Long
                Dim availW As Double
                Dim lineCount As Long
                fontSize = CLng(fontEventDesc.Size)
                If fontSize <= 0 Then fontSize = 8
                ' Используем фактическую ширину подписи для точного расчета количества строк
                ' (такая же логика, как в основном цикле вычисления резерва)
                availW = eventLabelWidthPoints
                If availW <= 0 Then availW = 40 * 2.83465
                ' Вычисляем фактическое количество строк для среднего события с учетом реального шрифта
                lineCount = EstimateWrappedLinesForFont(evLabelTexts(k), availW, fontEventDesc)
                If lineCount < 1 Then lineCount = 1
                actualLines = Application.WorksheetFunction.Min(lineCount, eventLabelMaxLines)
                ' Только высота текста без отступов (отступы уже учтены в базовом резерве)
                secondTierHeight = actualLines * EstimateTextHeight(fontSize) + (actualLines - 1) * 1 + 2
                ' Запас: 1 мм от верха самого высокого описания 1-го уровня до низа описания 2-го уровня
                secondTierHeightForTask(i) = secondTierHeight
                ' НЕ увеличиваем eventLabelReserveForTask - резерв первого яруса используется для позиционирования баров
                ' Высота второго яруса будет добавлена напрямую к rowHeights, чтобы не влиять на позиционирование баров
                Exit For
            End If
        Next k
NextTaskRow:
    Next i
End Sub

' Computes horizontal offsets for event labels:
' 1) If two events on same bar are close (distance < "Ширина описания события"), spread descriptions in opposite directions to avoid overlap.
' 2) If one cannot shift (would exit timeline), shift the other toward center until the first fits.
' 3) All descriptions must stay within timeline bounds.
' 4) If three events in a row and middle label still overlaps both neighbours, labelTiers(middle)=1 (draw above).
Private Sub ComputeEventLabelOffsets(ByRef eventItems As Collection, ByVal labelWidthPoints As Double, _
                                     ByVal labelMaxLines As Long, ByRef textOffsets() As Double, _
                                     ByRef labelTiers() As Long, _
                                     ByVal timelineLeft As Double, ByVal timelineWidth As Double, _
                                     ByVal diagramLeft As Double, ByVal diagramWidth As Double, _
                                     ByVal sidebarMode As String, ByVal sidebarEnabled As Boolean, _
                                     ByRef fontEventDesc As fontSettingsType)
    Const TOP_TOL As Double = 2#  ' Увеличен допуск для группировки событий на одной задаче
    Const MIN_GAP As Double = 2
    Dim n As Long, idx As Long, ev As Variant
    Dim evTops() As Double, evCenters() As Double, lblWidths() As Double, hasLabel() As Boolean
    Dim groupStart As Long, groupEnd As Long, i As Long, j As Long, p As Long
    Dim evLeft As Double, evTop As Double, evWid As Double, lblText As String
    Dim maxLblW As Double, availW As Double
    Dim grpCount As Long, grpIdx() As Long, k As Long, tmp As Long
    Dim evCenter As Double, lblW As Double
    Dim leftBound As Double, rightBound As Double
    Dim minOff() As Double, maxOff() As Double
    Dim centerI As Double, centerJ As Double, wi As Double, wj As Double
    Dim required As Double, offI As Double, offJ As Double
    Dim iter As Long, changed As Boolean
    Dim closeDist As Double
    Dim lblLeft As Double, lblRight As Double
    Dim midIdx As Long, leftIdx As Long, rightIdx As Long
    Dim leftL As Double, leftR As Double, midL As Double, midR As Double, rightL As Double, rightR As Double

    If eventItems Is Nothing Or eventItems.count = 0 Then Exit Sub
    n = eventItems.count
    ReDim textOffsets(1 To n)
    ReDim labelTiers(1 To n)
    For i = 1 To n
        labelTiers(i) = 0
    Next i
    ReDim evTops(1 To n)
    ReDim evCenters(1 To n)
    ReDim lblWidths(1 To n)
    ReDim hasLabel(1 To n)

    maxLblW = IIf(labelWidthPoints > 0, labelWidthPoints, 40 * 2.83465)
    closeDist = maxLblW

    leftBound = Application.WorksheetFunction.Max(diagramLeft, timelineLeft)
    rightBound = Application.WorksheetFunction.Min(diagramLeft + diagramWidth, timelineLeft + timelineWidth)

    idx = 0
    For Each ev In eventItems
        idx = idx + 1
        If IsArray(ev) And UBound(ev) >= 8 Then
            evLeft = ev(1)
            evTop = ev(2)
            evWid = ev(3)
            lblText = Trim$(CStr(ev(8)))
            evTops(idx) = evTop
            evCenters(idx) = evLeft + evWid / 2
            hasLabel(idx) = (Len(lblText) > 0)
            If hasLabel(idx) Then
                availW = Application.WorksheetFunction.Max(1, maxLblW)
                lblW = Application.WorksheetFunction.Min(maxLblW, _
                      Application.WorksheetFunction.Max(evWid * 2, MeasureTextWidthWithFont(lblText, fontEventDesc.name, CLng(fontEventDesc.Size), fontEventDesc.bold, fontEventDesc.italic) + 8))
                lblWidths(idx) = lblW
            Else
                lblWidths(idx) = 0
            End If
        Else
            evTops(idx) = 0
            evCenters(idx) = 0
            lblWidths(idx) = 0
            hasLabel(idx) = False
        End If
    Next ev

    ReDim minOff(1 To n)
    ReDim maxOff(1 To n)
    For i = 1 To n
        If hasLabel(i) And lblWidths(i) > 0 Then
            minOff(i) = leftBound - evCenters(i) + lblWidths(i) / 2
            maxOff(i) = rightBound - evCenters(i) - lblWidths(i) / 2
        End If
    Next i

    groupStart = 1
    Do While groupStart <= n
        groupEnd = groupStart
        Do While groupEnd < n
            If Abs(evTops(groupEnd + 1) - evTops(groupStart)) >= TOP_TOL Then Exit Do
            groupEnd = groupEnd + 1
        Loop

        If groupEnd - groupStart >= 1 Then
            grpCount = 0
            For i = groupStart To groupEnd
                If hasLabel(i) Then grpCount = grpCount + 1
            Next i
            If grpCount >= 2 Then
                ReDim grpIdx(1 To grpCount)
                k = 0
                For i = groupStart To groupEnd
                    If hasLabel(i) Then
                        k = k + 1
                        grpIdx(k) = i
                    End If
                Next i
                For i = 1 To grpCount - 1
                    For j = i + 1 To grpCount
                        If evCenters(grpIdx(i)) > evCenters(grpIdx(j)) Then
                            tmp = grpIdx(i)
                            grpIdx(i) = grpIdx(j)
                            grpIdx(j) = tmp
                        End If
                    Next j
                Next i

                For iter = 1 To 20
                    changed = False
                    For p = 1 To grpCount - 1
                        i = grpIdx(p)
                        j = grpIdx(p + 1)
                        centerI = evCenters(i)
                        centerJ = evCenters(j)
                        wi = lblWidths(i)
                        wj = lblWidths(j)
                        If (centerJ - centerI) < closeDist Then
                            required = (wi + wj) / 2 + MIN_GAP - (centerJ - centerI)
                            If required > 0 Then
                                offI = textOffsets(i)
                                offJ = textOffsets(j)
                                If offJ - offI < required Then
                                    offI = -required / 2
                                    offJ = required / 2
                                    offI = Application.WorksheetFunction.Max(minOff(i), Application.WorksheetFunction.Min(maxOff(i), offI))
                                    offJ = Application.WorksheetFunction.Max(minOff(j), Application.WorksheetFunction.Min(maxOff(j), offJ))
                                    If offJ - offI < required Then
                                        offJ = Application.WorksheetFunction.Min(maxOff(j), offI + required)
                                        If offJ - offI < required Then
                                            offI = Application.WorksheetFunction.Max(minOff(i), offJ - required)
                                        End If
                                        offJ = Application.WorksheetFunction.Max(minOff(j), Application.WorksheetFunction.Min(maxOff(j), offI + required))
                                    End If
                                    If textOffsets(i) <> offI Or textOffsets(j) <> offJ Then changed = True
                                    textOffsets(i) = offI
                                    textOffsets(j) = offJ
                                End If
                            End If
                        End If
                    Next p
                    If Not changed Then Exit For
                Next iter
            End If
        End If
        groupStart = groupEnd + 1
    Loop

    ' Среднее из трёх и более подписей: если ДО применения сдвигов перекрывает соседей — второй ярус (выше)
    ' Проверяем ДО применения textOffsets, чтобы определить, нужно ли поднимать среднюю подпись
    groupStart = 1
    Do While groupStart <= n
        groupEnd = groupStart
        Do While groupEnd < n
            If Abs(evTops(groupEnd + 1) - evTops(groupStart)) >= TOP_TOL Then Exit Do
            groupEnd = groupEnd + 1
        Loop
        If groupEnd - groupStart >= 2 Then
            grpCount = 0
            For i = groupStart To groupEnd
                If hasLabel(i) Then grpCount = grpCount + 1
            Next i
            If grpCount >= 3 Then
                ReDim grpIdx(1 To grpCount)
                k = 0
                For i = groupStart To groupEnd
                    If hasLabel(i) Then
                        k = k + 1
                        grpIdx(k) = i
                    End If
                Next i
                For i = 1 To grpCount - 1
                    For j = i + 1 To grpCount
                        If evCenters(grpIdx(i)) > evCenters(grpIdx(j)) Then
                            tmp = grpIdx(i)
                            grpIdx(i) = grpIdx(j)
                            grpIdx(j) = tmp
                        End If
                    Next j
                Next i
                For p = 2 To grpCount - 1
                    midIdx = grpIdx(p)
                    leftIdx = grpIdx(p - 1)
                    rightIdx = grpIdx(p + 1)
                    ' Проверяем перекрытие ДО применения горизонтальных сдвигов (textOffsets ещё = 0)
                    leftL = evCenters(leftIdx) - lblWidths(leftIdx) / 2
                    leftR = evCenters(leftIdx) + lblWidths(leftIdx) / 2
                    midL = evCenters(midIdx) - lblWidths(midIdx) / 2
                    midR = evCenters(midIdx) + lblWidths(midIdx) / 2
                    rightL = evCenters(rightIdx) - lblWidths(rightIdx) / 2
                    rightR = evCenters(rightIdx) + lblWidths(rightIdx) / 2
                    ' Средняя подпись перекрывает обе соседние ДО применения сдвигов
                    ' Перекрытие с левой: правая граница левой подписи > левая граница средней подписи
                    ' Перекрытие с правой: левая граница правой подписи < правая граница средней подписи
                    If (leftR + MIN_GAP > midL) And (rightL - MIN_GAP < midR) Then
                        labelTiers(midIdx) = 1
                        ' Для события на втором ярусе убираем горизонтальное смещение - подпись строго над событием
                        textOffsets(midIdx) = 0
                        Exit For
                    End If
                Next p
            End If
        End If
        groupStart = groupEnd + 1
    Loop

    For i = 1 To n
        If hasLabel(i) And lblWidths(i) > 0 Then
            ' Для событий на втором ярусе не изменяем textOffsets - подпись должна быть строго над событием
            If labelTiers(i) = 1 Then
                textOffsets(i) = 0
            Else
                evCenter = evCenters(i)
                lblW = lblWidths(i)
                lblLeft = evCenter - lblW / 2 + textOffsets(i)
                lblRight = evCenter + lblW / 2 + textOffsets(i)
                If lblLeft < leftBound Then textOffsets(i) = minOff(i)
                lblRight = evCenter + lblW / 2 + textOffsets(i)
                If lblRight > rightBound Then textOffsets(i) = maxOff(i)
            End If
        End If
    Next i
End Sub

'' Возвращает фигуру с листа Справочники по имени (если существует), иначе Nothing.
'' Ищет по точному совпадению и по перебору (на случай Picture 1 и т.п.).
Private Function TryGetCustomEventShape(ByVal shapeName As String) As Shape
    Dim wsRef As Worksheet
    Dim nameTrim As String
    Dim shp As Shape

    On Error Resume Next
    nameTrim = Trim$(shapeName)
    If Len(nameTrim) = 0 Then
        Set TryGetCustomEventShape = Nothing
        Exit Function
    End If
    Set wsRef = ThisWorkbook.Worksheets(SHEET_REF)
    If wsRef Is Nothing Then
        Set TryGetCustomEventShape = Nothing
        Exit Function
    End If
    Set TryGetCustomEventShape = wsRef.Shapes(nameTrim)
    If Err.Number = 0 And Not TryGetCustomEventShape Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear
    For Each shp In wsRef.Shapes
        If StrComp(Trim$(shp.name), nameTrim, vbTextCompare) = 0 Then
            Set TryGetCustomEventShape = shp
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If
    Next shp
    Set TryGetCustomEventShape = Nothing
    Err.Clear
    On Error GoTo 0
End Function

'' Копирует кастомную фигуру на диаграмму по имени (с листа Справочники). Ссылку получаем непосредственно перед Copy, чтобы избежать устаревших ссылок при повторной генерации.
Private Sub DrawCustomEventShape(ByVal ws As Worksheet, ByVal eventShapeName As String, _
                                  ByVal leftPos As Double, ByVal topPos As Double, _
                                  ByVal width As Double, ByVal height As Double, _
                                  ByVal fillTransparency As Double, _
                                  ByRef fontEventDesc As fontSettingsType, _
                                  ByVal labelText As String, _
                                  ByVal labelGapPoints As Double, ByVal labelWidthPoints As Double, _
                                  ByVal labelMaxLines As Long, ByVal textOffsetX As Double, _
                                  Optional ByVal labelTierOffset As Double = 0)
    Dim pastedShp As Shape
    Dim sourceShape As Shape
    Dim srcW As Double
    Dim srcH As Double
    Dim aspect As Double
    Dim newW As Double
    Dim newH As Double
    Dim shpLeft As Double
    Dim shpTop As Double
    Dim shapeCountBefore As Long

    On Error GoTo FallbackAddShape
    Set sourceShape = TryGetCustomEventShape(eventShapeName)
    If sourceShape Is Nothing Then GoTo FallbackAddShape

    srcW = sourceShape.width
    srcH = sourceShape.height
    If srcH <= 0 Then srcH = 1
    aspect = srcW / srcH
    newH = height
    newW = newH * aspect
    shpLeft = leftPos + (width - newW) / 2
    shpTop = topPos

    shapeCountBefore = ws.Shapes.count
    Application.CutCopyMode = False
    ws.Activate
    sourceShape.Copy
    ws.Paste
    Application.CutCopyMode = False
    Set pastedShp = Nothing
    On Error Resume Next
    If ws.Parent.ActiveSheet Is ws Then
        If TypeName(Selection) = "Shape" Then Set pastedShp = Selection
        If pastedShp Is Nothing And TypeName(Selection) = "DrawingObjects" Then
            If Selection.ShapeRange.count >= 1 Then Set pastedShp = Selection.ShapeRange(1)
        End If
    End If
    On Error GoTo 0
    If pastedShp Is Nothing Then
        If ws.Shapes.count > shapeCountBefore Then
            Set pastedShp = ws.Shapes(shapeCountBefore + 1)
        Else
            Set pastedShp = ws.Shapes(ws.Shapes.count)
        End If
    End If
    If pastedShp Is Nothing Then GoTo FallbackAddShape

    pastedShp.Left = shpLeft
    pastedShp.Top = shpTop
    pastedShp.width = newW
    pastedShp.height = newH
    On Error Resume Next
    pastedShp.Fill.transparency = ClampTransparency(fillTransparency) / 100
    ' Копируем настройки прозрачности с исходной фигуры (на Справочниках отображается корректно)
    If (sourceShape.Type = msoPicture Or sourceShape.Type = msoLinkedPicture) And _
       (pastedShp.Type = msoPicture Or pastedShp.Type = msoLinkedPicture) Then
        pastedShp.PictureFormat.TransparentBackground = sourceShape.PictureFormat.TransparentBackground
        pastedShp.PictureFormat.TransparencyColor = sourceShape.PictureFormat.TransparencyColor
    End If
    On Error GoTo 0

    If Len(Trim$(labelText)) > 0 Then
        DrawEventLabelAboveShape ws, pastedShp, labelText, fontEventDesc, labelGapPoints, labelWidthPoints, labelMaxLines, textOffsetX, labelTierOffset
    End If
    pastedShp.ZOrder msoBringToFront
    Exit Sub

FallbackAddShape:
    On Error GoTo 0
    DrawEventShape ws, msoShapeOval, leftPos, topPos, width, height, RGB(91, 155, 213), False, fontEventDesc, fillTransparency, labelText, labelGapPoints, labelWidthPoints, labelMaxLines, textOffsetX, "", IIf(labelTierOffset <> 0, 1, 0), labelTierOffset
End Sub

'' Рисует подпись над фигурой события.
Private Sub DrawEventLabelAboveShape(ByVal ws As Worksheet, ByVal eventShape As Shape, _
                                     ByVal labelText As String, ByRef fontEventDesc As fontSettingsType, _
                                     ByVal labelGapPoints As Double, ByVal labelWidthPoints As Double, _
                                     ByVal labelMaxLines As Long, ByVal textOffsetX As Double, _
                                     Optional ByVal labelTierOffset As Double = 0)
    Dim lbl As Shape
    Dim lblH As Double
    Dim lblW As Double
    Dim leftPos As Double
    Dim topPos As Double
    Dim fsEv As fontSettingsType
    Dim maxLabelWidth As Double
    Dim maxLines As Long
    Dim availW As Double
    Dim lineCount As Long
    Dim actualLines As Long
    Dim effectiveGap As Double

    fsEv = fontEventDesc
    maxLabelWidth = IIf(labelWidthPoints > 0, labelWidthPoints, 40 * 2.83465)
    maxLines = IIf(labelMaxLines >= 1, labelMaxLines, 3)
    lblW = Application.WorksheetFunction.Min(maxLabelWidth, _
          Application.WorksheetFunction.Max(eventShape.width * 2, MeasureTextWidthWithFont(labelText, fsEv.name, CLng(fsEv.Size), fsEv.bold, fsEv.italic) + 8))
    availW = Application.WorksheetFunction.Max(1, lblW)
    lineCount = EstimateWrappedLinesForFont(labelText, availW, fsEv)
    If lineCount < 1 Then lineCount = 1
    actualLines = Application.WorksheetFunction.Min(lineCount, maxLines)
    lblH = actualLines * EstimateTextHeight(CLng(fsEv.Size)) + (actualLines - 1) * 1 + 2
    effectiveGap = labelGapPoints
    leftPos = eventShape.Left + (eventShape.width / 2) - (lblW / 2) + textOffsetX
    ' Описание события всегда над событием
    ' Для первого яруса (labelTierOffset = 0): описание над событием
    ' Для второго яруса (labelTierOffset > 0): описание выше первого яруса на labelTierOffset
    If labelTierOffset > 0 Then
        ' Второй ярус: выше события на labelTierOffset
        topPos = eventShape.Top - lblH - effectiveGap - labelTierOffset
    Else
        ' Первый ярус: над событием
        topPos = eventShape.Top - lblH - effectiveGap
    End If
    Set lbl = ws.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=leftPos, Top:=topPos, width:=lblW, height:=lblH)
    lbl.TextFrame2.TextRange.text = labelText
    ApplyFontToTextRange lbl.TextFrame2.TextRange, fsEv
    With lbl.TextFrame2.TextRange.ParagraphFormat
        .Alignment = msoAlignCenter
        .SpaceBefore = 0
        .SpaceAfter = 0
    End With
    lbl.TextFrame2.VerticalAnchor = msoAnchorTop
    lbl.TextFrame2.WordWrap = msoCTrue
    lbl.TextFrame2.AutoSize = msoAutoSizeNone
    lbl.TextFrame2.MarginTop = 0
    lbl.TextFrame2.MarginBottom = 0
    lbl.TextFrame2.MarginLeft = 0
    lbl.TextFrame2.MarginRight = 0
    lbl.Line.Visible = msoFalse
    lbl.Fill.Visible = msoFalse
    lbl.ZOrder msoBringToFront
End Sub

Private Sub DrawEventShape(ByVal ws As Worksheet, ByVal shapeType As MsoAutoShapeType, _
                           ByVal leftPos As Double, ByVal topPos As Double, _
                           ByVal width As Double, ByVal height As Double, ByVal fillColor As Long, _
                           ByVal showBorder As Boolean, ByRef fontEventDesc As fontSettingsType, _
                           Optional ByVal fillTransparency As Double = 0, Optional ByVal labelText As String = "", _
                           Optional ByVal labelGapPoints As Double = 2.83465, Optional ByVal labelWidthPoints As Double = -1, Optional ByVal labelMaxLines As Long = 3, _
                           Optional ByVal textOffsetX As Double = 0, Optional ByVal eventName As String = "", _
                           Optional ByVal labelTier As Long = 0, Optional ByVal labelTierHeight As Double = 0)
    Dim shp As Shape
    Dim customShp As Shape

    ' Пробуем взять кастомную фигуру с листа Справочники (имя = имя типа события). Передаём имя, чтобы внутри получать свежую ссылку перед Copy.
    If Len(Trim$(eventName)) > 0 Then
        Set customShp = TryGetCustomEventShape(eventName)
        If Not customShp Is Nothing Then
            ' Смещение второго яруса задаётся снаружи (1 мм от верха первого уровня до низа второго)
            Dim customTierOffset As Double
            If labelTier > 0 Then
                customTierOffset = labelTierHeight
            Else
                customTierOffset = 0
            End If
            DrawCustomEventShape ws, CStr(eventName), leftPos, topPos, width, height, fillTransparency, _
                fontEventDesc, labelText, labelGapPoints, labelWidthPoints, labelMaxLines, textOffsetX, _
                customTierOffset
            Exit Sub
        End If
    End If

    Set shp = ws.Shapes.AddShape(shapeType, leftPos, topPos, width, height)
    shp.Fill.ForeColor.RGB = fillColor
    shp.Fill.transparency = ClampTransparency(fillTransparency) / 100
    If showBorder Then
        shp.Line.Visible = msoTrue
        shp.Line.ForeColor.RGB = DarkenColor(fillColor, 0.2)
        shp.Line.Weight = 0.75
    Else
        shp.Line.Visible = msoFalse
    End If

    ' Label above event (optional)
    If Len(Trim$(labelText)) > 0 Then
        Dim lbl As Shape
        Dim lblH As Double, lblW As Double
        Dim maxLabelWidth As Double
        Dim effectiveGap As Double
        Dim maxLines As Long
        Dim fsEv As fontSettingsType
        fsEv = fontEventDesc
        maxLabelWidth = IIf(labelWidthPoints > 0, labelWidthPoints, 40 * 2.83465)
        maxLines = IIf(labelMaxLines >= 1, labelMaxLines, 3)
        lblW = Application.WorksheetFunction.Min(maxLabelWidth, _
              Application.WorksheetFunction.Max(width * 2, MeasureTextWidthWithFont(labelText, fsEv.name, CLng(fsEv.Size), fsEv.bold, fsEv.italic) + 8))
        Dim availW As Double
        Dim lineCount As Long
        Dim actualLines As Long
        availW = Application.WorksheetFunction.Max(1, lblW)
        lineCount = EstimateWrappedLinesForFont(labelText, availW, fsEv)
        If lineCount < 1 Then lineCount = 1
        actualLines = Application.WorksheetFunction.Min(lineCount, maxLines)
        Const LINE_GAP_PT As Double = 1
        lblH = actualLines * EstimateTextHeight(CLng(fsEv.Size)) + (actualLines - 1) * LINE_GAP_PT + 2
        effectiveGap = labelGapPoints
        ' Описание события всегда над событием
        ' Если labelTier = 1, смещение задаётся снаружи (1 мм от верха первого уровня до низа второго)
        Dim tierOffset As Double
        Dim lblTop As Double
        If labelTier > 0 And labelTierHeight > 0 Then
            tierOffset = labelTierHeight
            ' Второй ярус: выше события на tierOffset
            lblTop = topPos - lblH - effectiveGap - tierOffset
        Else
            tierOffset = 0
            ' Первый ярус: над событием
            lblTop = topPos - lblH - effectiveGap
        End If
        Set lbl = ws.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
                                       Left:=leftPos + (width / 2) - (lblW / 2) + textOffsetX, _
                                       Top:=lblTop, _
                                       width:=lblW, height:=lblH)
        lbl.TextFrame2.TextRange.text = labelText
        ApplyFontToTextRange lbl.TextFrame2.TextRange, fsEv
        With lbl.TextFrame2.TextRange.ParagraphFormat
            .Alignment = msoAlignCenter
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
        lbl.TextFrame2.VerticalAnchor = msoAnchorTop
        lbl.TextFrame2.WordWrap = msoCTrue
        lbl.TextFrame2.AutoSize = msoAutoSizeNone
        lbl.TextFrame2.MarginTop = 0
        lbl.TextFrame2.MarginBottom = 0
        lbl.TextFrame2.MarginLeft = 0
        lbl.TextFrame2.MarginRight = 0
        lbl.Line.Visible = msoFalse
        lbl.Fill.Visible = msoFalse
        lbl.ZOrder msoBringToFront
    End If

    shp.ZOrder msoBringToFront
End Sub

Private Sub DrawLegend(ByVal ws As Worksheet, ByVal leftPos As Double, ByVal topPos As Double, _
                       ByVal width As Double, ByVal height As Double, ByVal lineTypesTable As ListObject, _
                       ByVal eventTypesTable As ListObject, ByVal padding As Double, _
                       ByVal barHeightNominal As Double, ByRef usedLineTypes() As String, _
                       ByVal usedLineTypeCount As Long, ByRef usedEventTypes() As String, _
                       ByVal usedEventTypeCount As Long, ByVal showEventBorder As Boolean, _
                       ByRef fontLegend As fontSettingsType)
    Dim symbolSize As Double
    Dim rowHeight As Double
    Dim lineCount As Long
    Dim eventCount As Long
    Dim rowsCount As Long
    Dim rowIndex As Long
    Dim lineRow As ListRow
    Dim eventRow As ListRow
    Dim lineType As String
    Dim lineMeaning As String
    Dim eventName As String
    Dim eventMeaning As String
    Dim lineColor As Long
    Dim lineTransparency As Double
    Dim eventColor As Long
    Dim eventTransparency As Double
    Dim eventShape As MsoAutoShapeType
    Dim eventHeight As Double
    Dim lineRect As Shape
    Dim textBox As Shape
    Dim leftColumnWidth As Double
    Dim rightColumnWidth As Double
    Dim rightColumnLeft As Double
    Dim textLeft As Double
    Dim shapeTop As Double
    Dim rectHeight As Double
    Dim meaningIndex As Long
    Dim eventNameIndex As Long
    Dim eventMeaningIndex As Long
    Dim maxLineTextWidth As Double
    Dim maxEventTextWidth As Double
    Dim maxEventShapeSize As Double
    Dim blocksGap As Double
    Dim rowGapPoints As Double
    Dim rowPitch As Double
    Dim lineMeaningIndex As Long

    lineCount = usedLineTypeCount
    eventCount = usedEventTypeCount
    rowsCount = Application.WorksheetFunction.Max(lineCount, eventCount)
    If rowsCount <= 0 Then
        Exit Sub
    End If

    symbolSize = 5 * 2.83465
    rectHeight = symbolSize * 0.6
    maxEventShapeSize = symbolSize
    maxLineTextWidth = 0
    maxEventTextWidth = 0

    blocksGap = 10 * 2.83465 ' 1 cm

    rowGapPoints = 2 * 2.83465 ' 2 мм межстрочный зазор

    If Not lineTypesTable Is Nothing Then
        lineMeaningIndex = GetTableColumnIndex(lineTypesTable, "Что означает")

        For Each lineRow In lineTypesTable.ListRows
            lineType = Trim$(CStr(lineRow.Range.Cells(1, 1).value))
            If Len(lineType) > 0 And IsKeyUsed(usedLineTypes, usedLineTypeCount, lineType) Then
                If lineMeaningIndex > 0 Then
                    lineMeaning = CStr(lineRow.Range.Cells(1, lineMeaningIndex).value)
                Else
                    lineMeaning = lineType
                End If
                maxLineTextWidth = Application.WorksheetFunction.Max(maxLineTextWidth, EstimateTextWidth(lineMeaning, CLng(fontLegend.Size)))
            End If
        Next lineRow
    End If

    If Not eventTypesTable Is Nothing Then
        eventNameIndex = GetTableColumnIndex(eventTypesTable, "Событие")
        eventMeaningIndex = GetTableColumnIndex(eventTypesTable, "Что означает")
        If eventNameIndex = 0 Then
            Exit Sub
        End If
        For Each eventRow In eventTypesTable.ListRows
            eventName = Trim$(CStr(eventRow.Range.Cells(1, eventNameIndex).value))
            If Len(eventName) > 0 And IsKeyUsed(usedEventTypes, usedEventTypeCount, eventName) Then
                eventHeight = 0
                eventColor = RGB(91, 155, 213)
                eventShape = msoShapeOval
                eventTransparency = GetEventTypeTransparency(eventTypesTable, eventName, 0)
                If GetEventTypeInfo(eventTypesTable, eventName, eventShape, eventColor, eventHeight) = False Then
                    eventHeight = 0
                End If
                If eventHeight <= 0 Then
                    eventHeight = barHeightNominal
                End If
                maxEventShapeSize = Application.WorksheetFunction.Max(maxEventShapeSize, eventHeight)
                If eventMeaningIndex > 0 Then
                    eventMeaning = CStr(eventRow.Range.Cells(1, eventMeaningIndex).value)
                Else
                    eventMeaning = eventName
                End If
                maxEventTextWidth = Application.WorksheetFunction.Max(maxEventTextWidth, EstimateTextWidth(eventMeaning, CLng(fontLegend.Size)))
            End If
        Next eventRow
    End If

    maxLineTextWidth = maxLineTextWidth * 1.2 + 20
    maxEventTextWidth = maxEventTextWidth * 1.2 + 20

    rowHeight = Application.WorksheetFunction.Max(Application.WorksheetFunction.Max(rectHeight, maxEventShapeSize), 12)
    rowPitch = rowHeight + rowGapPoints
    leftColumnWidth = padding + symbolSize + 10 + maxLineTextWidth + padding
    rightColumnWidth = padding + maxEventShapeSize + 10 + maxEventTextWidth + padding
    rightColumnLeft = leftPos + leftColumnWidth + blocksGap

    rowIndex = 0
    If Not lineTypesTable Is Nothing Then
        For Each lineRow In lineTypesTable.ListRows
            lineType = Trim$(CStr(lineRow.Range.Cells(1, 1).value))
            If Len(lineType) > 0 And IsKeyUsed(usedLineTypes, usedLineTypeCount, lineType) Then
                rowIndex = rowIndex + 1
                lineColor = RGB(91, 155, 213)
                lineTransparency = GetLineTypeTransparency(lineTypesTable, lineType, 0)
                If lineRow.Range.Cells(1, 2).Interior.colorIndex <> xlColorIndexNone Then
                    lineColor = lineRow.Range.Cells(1, 2).Interior.Color
                End If

                If lineMeaningIndex > 0 Then
                    lineMeaning = CStr(lineRow.Range.Cells(1, lineMeaningIndex).value)
                Else
                    lineMeaning = lineType
                End If

                shapeTop = topPos + padding + (rowIndex - 1) * rowPitch + (rowHeight - rectHeight) / 2
                Set lineRect = ws.Shapes.AddShape(msoShapeRectangle, leftPos + padding + (symbolSize - symbolSize) / 2, shapeTop, _
                                                  symbolSize, rectHeight)
                lineRect.Fill.ForeColor.RGB = lineColor
                lineRect.Fill.transparency = ClampTransparency(lineTransparency) / 100
                lineRect.Line.ForeColor.RGB = DarkenColor(lineColor, 0.2)
                lineRect.Line.Weight = 0.75

                textLeft = leftPos + padding + symbolSize + 10
        Set textBox = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, textLeft, _
                                           topPos + padding + (rowIndex - 1) * rowPitch, _
                                           leftColumnWidth - (textLeft - leftPos) - padding, rowHeight)
                textBox.TextFrame2.TextRange.text = lineMeaning
        ApplyFontToTextRange textBox.TextFrame2.TextRange, fontLegend
                textBox.TextFrame2.AutoSize = msoAutoSizeNone
        textBox.TextFrame2.WordWrap = msoFalse
        textBox.TextFrame2.TextRange.ParagraphFormat.WordWrap = msoFalse
        textBox.TextFrame2.VerticalAnchor = msoAnchorMiddle
        textBox.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        textBox.TextFrame2.MarginTop = 0
        textBox.TextFrame2.MarginBottom = 0
        textBox.Fill.Visible = msoFalse
        textBox.Line.Visible = msoFalse
            End If
        Next lineRow
    End If

    rowIndex = 0
    If Not eventTypesTable Is Nothing Then
        For Each eventRow In eventTypesTable.ListRows
            eventName = Trim$(CStr(eventRow.Range.Cells(1, eventNameIndex).value))
            If Len(eventName) > 0 And IsKeyUsed(usedEventTypes, usedEventTypeCount, eventName) Then
                rowIndex = rowIndex + 1
                eventShape = msoShapeOval
                eventColor = RGB(91, 155, 213)
                eventHeight = 0
                If GetEventTypeInfo(eventTypesTable, eventName, eventShape, eventColor, eventHeight) = False Then
                    eventShape = msoShapeOval
                    eventColor = RGB(91, 155, 213)
                End If
                If eventHeight <= 0 Then
                    eventHeight = barHeightNominal
                End If

                If eventMeaningIndex > 0 Then
                    eventMeaning = CStr(eventRow.Range.Cells(1, eventMeaningIndex).value)
                Else
                    eventMeaning = eventName
                End If

                shapeTop = topPos + padding + (rowIndex - 1) * rowPitch + (rowHeight - eventHeight) / 2
                DrawEventShape ws, eventShape, rightColumnLeft + padding + (maxEventShapeSize - eventHeight) / 2, shapeTop, eventHeight, _
                    eventHeight, eventColor, showEventBorder, fontLegend, eventTransparency, "", 2.83465, -1, 3, 0, GetEventShapeName(eventTypesTable, eventName)

                textLeft = rightColumnLeft + padding + eventHeight + 10
        Set textBox = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, textLeft, _
                                           topPos + padding + (rowIndex - 1) * rowPitch, _
                                           rightColumnWidth - (textLeft - rightColumnLeft) - padding, rowHeight)
        textBox.TextFrame2.TextRange.text = eventMeaning
        ApplyFontToTextRange textBox.TextFrame2.TextRange, fontLegend
                textBox.TextFrame2.AutoSize = msoAutoSizeNone
        textBox.TextFrame2.WordWrap = msoFalse
        textBox.TextFrame2.TextRange.ParagraphFormat.WordWrap = msoFalse
        textBox.TextFrame2.VerticalAnchor = msoAnchorMiddle
        textBox.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        textBox.TextFrame2.MarginTop = 0
        textBox.TextFrame2.MarginBottom = 0
        textBox.Fill.Visible = msoFalse
        textBox.Line.Visible = msoFalse
            End If
        Next eventRow
    End If
End Sub

Private Sub DrawTimelineDividers(ByVal ws As Worksheet, ByVal diagramLeft As Double, ByVal diagramTop As Double, _
                                 ByVal diagramWidth As Double, ByVal yearHeaderHeight As Double, _
                                 ByVal totalHeaderHeight As Double, ByVal totalHeight As Double, _
                                 ByVal displayRowCount As Long, ByRef rowTops() As Double, _
                                 ByRef rowHeights() As Double, ByVal timelineLeft As Double, _
                                 ByVal timelineWidth As Double, ByRef timelineData As timelineData, _
                                 ByVal sidebarMode As String, ByVal sidebarWidth As Double, _
                                 ByVal sidebarGap As Double, ByVal vertThickness As Double, _
                                 ByVal yearVertThickness As Double, ByVal horizThickness As Double, _
                                 ByVal vertColor As Long, ByVal vertTransparency As Double, _
                                 ByVal yearVertColor As Long, ByVal yearVertTransparency As Double, _
                                 ByVal horizColor As Long, ByVal horizTransparency As Double)
    Dim i As Long
    Dim lineLeft As Double
    Dim lineTop As Double
    Dim lineBottom As Double
    Dim lineRight As Double
    Dim lineShape As Shape
    Dim sidebarLeft As Double
    Dim sidebarRight As Double
    Dim periodTop As Double
    Dim currentYear As Long

    If displayRowCount <= 0 Then
        Exit Sub
    End If

    lineTop = diagramTop
    lineBottom = diagramTop + totalHeaderHeight + totalHeight
    lineRight = diagramLeft + diagramWidth
    periodTop = diagramTop + yearHeaderHeight

    If vertThickness > 0 Then
        Set lineShape = ws.Shapes.AddLine(timelineLeft, lineTop, timelineLeft, lineBottom)
        lineShape.Line.ForeColor.RGB = vertColor
        lineShape.Line.transparency = ClampTransparency(vertTransparency) / 100
        lineShape.Line.Weight = vertThickness

        currentYear = Year(timelineData.PeriodStarts(1))
        For i = 1 To timelineData.count
            lineLeft = timelineLeft + timelineData.CumulativeWidths(i)
            Set lineShape = ws.Shapes.AddLine(lineLeft, periodTop, lineLeft, lineBottom)
            lineShape.Line.ForeColor.RGB = vertColor
            lineShape.Line.transparency = ClampTransparency(vertTransparency) / 100
            lineShape.Line.Weight = vertThickness

            If Year(timelineData.PeriodStarts(i)) <> currentYear Then
                currentYear = Year(timelineData.PeriodStarts(i))
                Set lineShape = ws.Shapes.AddLine(lineLeft, lineTop, lineLeft, lineBottom)
                lineShape.Line.ForeColor.RGB = yearVertColor
                lineShape.Line.transparency = ClampTransparency(yearVertTransparency) / 100
                lineShape.Line.Weight = yearVertThickness
            End If
        Next i

        Set lineShape = ws.Shapes.AddLine(timelineLeft + timelineWidth, lineTop, timelineLeft + timelineWidth, lineBottom)
        lineShape.Line.ForeColor.RGB = vertColor
        lineShape.Line.transparency = ClampTransparency(vertTransparency) / 100
        lineShape.Line.Weight = vertThickness

        If LCase$(sidebarMode) = "слева" Or LCase$(sidebarMode) = "с обеих сторон" Then
            sidebarLeft = diagramLeft
            sidebarRight = diagramLeft + sidebarWidth
            Set lineShape = ws.Shapes.AddLine(sidebarRight, lineTop, sidebarRight, lineBottom)
            lineShape.Line.ForeColor.RGB = vertColor
            lineShape.Line.transparency = ClampTransparency(vertTransparency) / 100
            lineShape.Line.Weight = vertThickness
        End If

        If LCase$(sidebarMode) = "справа" Or LCase$(sidebarMode) = "с обеих сторон" Then
            sidebarLeft = timelineLeft + timelineWidth + sidebarGap
            sidebarRight = sidebarLeft + sidebarWidth
            Set lineShape = ws.Shapes.AddLine(sidebarLeft, lineTop, sidebarLeft, lineBottom)
            lineShape.Line.ForeColor.RGB = vertColor
            lineShape.Line.transparency = ClampTransparency(vertTransparency) / 100
            lineShape.Line.Weight = vertThickness
        End If
    End If

    If horizThickness > 0 Then
        For i = 1 To displayRowCount
            lineLeft = diagramLeft
            lineTop = rowTops(i)
            Set lineShape = ws.Shapes.AddLine(lineLeft, lineTop, lineRight, lineTop)
            lineShape.Line.ForeColor.RGB = horizColor
            lineShape.Line.transparency = ClampTransparency(horizTransparency) / 100
            lineShape.Line.Weight = horizThickness
        Next i
        lineTop = rowTops(displayRowCount) + rowHeights(displayRowCount)
        Set lineShape = ws.Shapes.AddLine(diagramLeft, lineTop, lineRight, lineTop)
        lineShape.Line.ForeColor.RGB = horizColor
        lineShape.Line.transparency = ClampTransparency(horizTransparency) / 100
        lineShape.Line.Weight = horizThickness
    End If
End Sub

' Shades the timeline area to the left of the Today line, excluding section and subsection rows.
Private Sub DrawPastEventsOverlays(ByVal ws As Worksheet, ByVal timelineLeft As Double, _
                                   ByRef timelineData As timelineData, ByVal pastColor As Long, _
                                   ByVal pastTransparency As Double, ByRef outOverlayShapes As Collection, _
                                   ByVal displayRowCount As Long, ByRef rowTops() As Double, _
                                   ByRef rowHeights() As Double, ByRef isSectionRow() As Boolean)
    Dim todayDate As Date
    Dim todayLineX As Double
    Dim overlayWidth As Double
    Dim i As Long
    Dim shp As Shape

    todayDate = Date
    If todayDate < timelineData.PeriodStarts(1) Or todayDate > timelineData.PeriodEnds(timelineData.count) Then
        Exit Sub
    End If
    todayLineX = timelineLeft + GetTimelineOffset(timelineData, todayDate)
    overlayWidth = todayLineX - timelineLeft
    If overlayWidth <= 0 Or displayRowCount < 1 Then Exit Sub

    For i = 1 To displayRowCount
        If Not isSectionRow(i) Then
            Set shp = ws.Shapes.AddShape(msoShapeRectangle, timelineLeft, rowTops(i), overlayWidth, rowHeights(i))
            shp.Fill.ForeColor.RGB = pastColor
            shp.Fill.transparency = ClampTransparency(pastTransparency) / 100
            shp.Line.Visible = msoFalse
            If Not outOverlayShapes Is Nothing Then outOverlayShapes.Add shp
        End If
    Next i
End Sub

Private Sub DrawTodayStripe(ByVal ws As Worksheet, ByVal timelineLeft As Double, ByVal topPos As Double, _
                            ByVal totalHeight As Double, ByVal displayRowCount As Long, _
                            ByRef timelineData As timelineData, ByVal stripeThickness As Double, _
                            ByVal stripeColor As Long, ByVal stripeTransparency As Double, _
                            ByVal stripeMode As String)
    Dim todayDate As Date
    Dim offset As Double
    Dim stripeHeight As Double
    Dim stripeLeft As Double
    Dim stripe As Shape
    Dim dotRadius As Double
    Dim dotLeft As Double
    Dim dotTop As Double
    Dim dotCircle As Shape

    If displayRowCount <= 0 Or stripeThickness <= 0 Then
        Exit Sub
    End If

    todayDate = Date
    If todayDate < timelineData.PeriodStarts(1) Or todayDate > timelineData.PeriodEnds(timelineData.count) Then
        Exit Sub
    End If

    offset = GetTimelineOffset(timelineData, todayDate)
    stripeHeight = totalHeight
    stripeLeft = timelineLeft + offset - stripeThickness / 2

    Set stripe = ws.Shapes.AddShape(msoShapeRectangle, stripeLeft, topPos, stripeThickness, stripeHeight)
    stripe.Fill.ForeColor.RGB = stripeColor
    stripe.Fill.transparency = ClampTransparency(stripeTransparency) / 100
    stripe.Line.Visible = msoFalse
    stripe.ZOrder msoBringToFront

    If StrComp(Trim$(stripeMode), "Полоса с точкой", vbTextCompare) = 0 Then
        dotRadius = Application.WorksheetFunction.Max(stripeThickness, 2.83465)
        dotLeft = timelineLeft + offset - dotRadius
        dotTop = topPos + stripeHeight - dotRadius
        Set dotCircle = ws.Shapes.AddShape(msoShapeOval, dotLeft, dotTop, 2 * dotRadius, 2 * dotRadius)
        dotCircle.Fill.ForeColor.RGB = stripeColor
        dotCircle.Fill.transparency = ClampTransparency(stripeTransparency) / 100
        dotCircle.Line.Visible = msoFalse
        dotCircle.ZOrder msoBringToFront
    End If
End Sub

Private Sub DrawTodayLabel(ByVal ws As Worksheet, ByVal timelineLeft As Double, ByVal topPos As Double, _
                           ByVal totalHeight As Double, ByRef timelineData As timelineData, _
                           ByVal labelMode As String, ByRef fontLabel As fontSettingsType)
    Const LABEL_TEXT As String = "Сегодня"
    Dim todayDate As Date
    Dim offset As Double
    Dim lblW As Double
    Dim lblH As Double
    Dim lblLeft As Double
    Dim lblTop As Double
    Dim lbl As Shape
    Dim gap As Double

    todayDate = Date
    If todayDate < timelineData.PeriodStarts(1) Or todayDate > timelineData.PeriodEnds(timelineData.count) Then
        Exit Sub
    End If
    offset = GetTimelineOffset(timelineData, todayDate)
    gap = 2
    lblW = MeasureTextWidthWithFont(LABEL_TEXT, fontLabel.name, CLng(fontLabel.Size), fontLabel.bold, fontLabel.italic) + 4
    lblH = EstimateTextHeight(CLng(fontLabel.Size)) + 2

    If StrComp(Trim$(labelMode), "Горизонтально", vbTextCompare) = 0 Then
        lblLeft = timelineLeft + offset - lblW / 2
        lblTop = topPos + totalHeight + gap
        Set lbl = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, lblLeft, lblTop, lblW, lblH)
        lbl.TextFrame2.TextRange.text = LABEL_TEXT
        ApplyFontToTextRange lbl.TextFrame2.TextRange, fontLabel
        lbl.Rotation = 0
    ElseIf StrComp(Trim$(labelMode), "Вертикально", vbTextCompare) = 0 Then
        lblLeft = timelineLeft + offset - (lblW + 4) / 2
        lblTop = topPos + totalHeight + gap + (lblW + 4) / 2
        Set lbl = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, lblLeft, lblTop, lblW + 4, lblH + 4)
        lbl.TextFrame2.MarginTop = 0
        lbl.TextFrame2.MarginBottom = 0
        lbl.TextFrame2.MarginLeft = 0
        lbl.TextFrame2.MarginRight = 0
        lbl.TextFrame2.WordWrap = msoFalse
        lbl.TextFrame2.TextRange.text = LABEL_TEXT
        ApplyFontToTextRange lbl.TextFrame2.TextRange, fontLabel
        lbl.Rotation = -90
    Else
        Exit Sub
    End If
    lbl.TextFrame2.MarginTop = 0
    lbl.TextFrame2.MarginBottom = 0
    lbl.TextFrame2.MarginLeft = 0
    lbl.TextFrame2.MarginRight = 0
    lbl.TextFrame2.WordWrap = msoFalse
    lbl.Line.Visible = msoFalse
    lbl.Fill.Visible = msoFalse
    lbl.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    lbl.TextFrame2.VerticalAnchor = msoAnchorMiddle
    lbl.ZOrder msoBringToFront
End Sub

Private Sub ClearGanttShapes(ByVal ws As Worksheet)
    Dim shp As Shape
    Dim toDelete As Collection
    Dim i As Long

    Set toDelete = New Collection

    For Each shp In ws.Shapes
        If shp.name <> BUTTON_NAME Then
            toDelete.Add shp
        End If
    Next shp

    For i = 1 To toDelete.count
        toDelete(i).Delete
    Next i
End Sub

Private Sub EnsureGanttButton(ByVal ws As Worksheet)
    Dim btn As Shape
    Dim cmToPoints As Double
    Dim mmToPoints As Double

    cmToPoints = 28.3465
    mmToPoints = 2.83465

    On Error Resume Next
    Set btn = ws.Shapes(BUTTON_NAME)
    On Error GoTo 0

    If btn Is Nothing Then
        Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, 5 * mmToPoints, 5 * mmToPoints, _
                                     5 * cmToPoints, 3 * cmToPoints)
        btn.name = BUTTON_NAME
        btn.OnAction = "CreateGanttDiagram"
        btn.TextFrame2.TextRange.text = "Создать диаграмму"
        btn.TextFrame2.TextRange.Font.Size = 10
        btn.TextFrame2.VerticalAnchor = msoAnchorMiddle
        btn.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    Else
        btn.Left = 5 * mmToPoints
        btn.Top = 5 * mmToPoints
        btn.width = 5 * cmToPoints
        btn.height = 3 * cmToPoints
        btn.OnAction = "CreateGanttDiagram"
    End If
End Sub

Private Function GetWorksheetByName(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheetByName = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function GetWorksheetByNameOrTable(ByVal sheetName As String, ByVal tableName As String) As Worksheet
    Dim ws As Worksheet
    Dim listObj As ListObject

    Set GetWorksheetByNameOrTable = GetWorksheetByName(sheetName)
    If Not GetWorksheetByNameOrTable Is Nothing Then
        Exit Function
    End If

    For Each ws In ThisWorkbook.Worksheets
        For Each listObj In ws.ListObjects
            If listObj.name = tableName Then
                Set GetWorksheetByNameOrTable = ws
                Exit Function
            End If
        Next listObj
    Next ws
End Function

Private Function GetSettingValue(ByVal settingsTable As ListObject, ByVal parameterName As String, _
                                 ByVal defaultValue As Variant) As Variant
    Dim row As ListRow

    For Each row In settingsTable.ListRows
        If CStr(row.Range.Cells(1, 1).value) = parameterName Then
            GetSettingValue = row.Range.Cells(1, 2).value
            Exit Function
        End If
    Next row

    GetSettingValue = defaultValue
End Function

Private Function GetColorFromTable(ByVal colorsTable As ListObject, ByVal objectName As String, _
                                   ByRef transparency As Double) As Long
    Dim row As ListRow
    Dim colorCell As Range
    Dim transparencyCell As Range

    transparency = -1

    For Each row In colorsTable.ListRows
        If NormalizeKey(CStr(row.Range.Cells(1, 1).value)) = NormalizeKey(objectName) Then
            Set colorCell = row.Range.Cells(1, 2)
            If row.Range.Columns.count >= 3 Then
                Set transparencyCell = row.Range.Cells(1, 3)
                If IsNumeric(transparencyCell.value) Then
                    transparency = CDbl(transparencyCell.value)
                End If
            End If
            If colorCell.Interior.colorIndex <> xlColorIndexNone Then
                GetColorFromTable = colorCell.Interior.Color
            Else
                GetColorFromTable = RGB(255, 255, 255)
            End If
            Exit Function
        End If
    Next row

    GetColorFromTable = -1
End Function

'' Возвращает текст заголовка из таблицы Заголовок на листе Данные.
Private Function GetHeaderTextFromTable(ByVal wsData As Worksheet) As String
    Dim tbl As ListObject
    Dim colIdx As Long
    Dim r As Long
    Dim parts() As String
    Dim i As Long

    GetHeaderTextFromTable = ""
    On Error Resume Next
    Set tbl = wsData.ListObjects(TABLE_HEADER)
    If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then
        On Error GoTo 0
        Exit Function
    End If
    colIdx = GetTableColumnIndex(tbl, "Текст")
    If colIdx <= 0 Then colIdx = GetTableColumnIndex(tbl, "Заголовок")
    If colIdx <= 0 Then colIdx = 1
    For r = 1 To tbl.ListRows.count
        If Len(Trim$(CStr(tbl.ListColumns(colIdx).DataBodyRange.Cells(r, 1).value))) > 0 Then
            If Len(GetHeaderTextFromTable) > 0 Then GetHeaderTextFromTable = GetHeaderTextFromTable & vbLf
            GetHeaderTextFromTable = GetHeaderTextFromTable & Trim$(CStr(tbl.ListColumns(colIdx).DataBodyRange.Cells(r, 1).value))
        End If
    Next r
    On Error GoTo 0
End Function

Private Function NormalizeKey(ByVal value As String) As String
    NormalizeKey = LCase$(Trim$(Replace(Replace(value, vbCr, ""), vbLf, "")))
End Function

Private Function GetFontSettingsFromTable(ByVal fontsTable As ListObject, ByVal objectName As String, _
    ByVal defaultName As String, ByVal defaultSize As Double, ByVal defaultColor As Long, _
    ByVal defaultBold As Boolean, ByVal defaultItalic As Boolean) As fontSettingsType
    Dim row As ListRow
    Dim objCol As Long
    Dim colorCol As Long
    Dim fontCol As Long
    Dim sizeCol As Long
    Dim boldCol As Long
    Dim italicCol As Long
    Dim v As Variant

    GetFontSettingsFromTable.name = defaultName
    GetFontSettingsFromTable.Size = defaultSize
    GetFontSettingsFromTable.Color = defaultColor
    GetFontSettingsFromTable.bold = defaultBold
    GetFontSettingsFromTable.italic = defaultItalic

    If fontsTable Is Nothing Then Exit Function

    objCol = GetTableColumnIndex(fontsTable, "Объект")
    If objCol <= 0 Then objCol = 1
    colorCol = GetTableColumnIndex(fontsTable, "Цвет")
    If colorCol <= 0 Then colorCol = 2
    fontCol = GetTableColumnIndex(fontsTable, "Шрифт")
    If fontCol <= 0 Then fontCol = 3
    sizeCol = GetTableColumnIndex(fontsTable, "Высота шрифта, пт")
    If sizeCol <= 0 Then sizeCol = GetTableColumnIndex(fontsTable, "Высота шрифта пт")
    If sizeCol <= 0 Then sizeCol = 4
    boldCol = GetTableColumnIndex(fontsTable, "Жирный")
    If boldCol <= 0 Then boldCol = 5
    italicCol = GetTableColumnIndex(fontsTable, "Курсив")
    If italicCol <= 0 Then italicCol = 6

    For Each row In fontsTable.ListRows
        If NormalizeKey(CStr(row.Range.Cells(1, objCol).value)) = NormalizeKey(objectName) Then
            If fontCol > 0 And fontCol <= row.Range.Columns.count Then
                v = row.Range.Cells(1, fontCol).value
                If Len(Trim$(CStr(v))) > 0 Then GetFontSettingsFromTable.name = CStr(v)
            End If
            If sizeCol > 0 And sizeCol <= row.Range.Columns.count Then
                v = row.Range.Cells(1, sizeCol).value
                If IsNumeric(v) And CDbl(v) > 0 Then GetFontSettingsFromTable.Size = CDbl(v)
            End If
            If colorCol > 0 And colorCol <= row.Range.Columns.count Then
                If row.Range.Cells(1, colorCol).Interior.colorIndex <> xlColorIndexNone Then
                    GetFontSettingsFromTable.Color = row.Range.Cells(1, colorCol).Interior.Color
                End If
            End If
            If boldCol > 0 And boldCol <= row.Range.Columns.count Then
                v = row.Range.Cells(1, boldCol).value
                GetFontSettingsFromTable.bold = (NormalizeKey(CStr(v)) = "да" Or LCase$(Trim$(CStr(v))) = "yes")
            End If
            If italicCol > 0 And italicCol <= row.Range.Columns.count Then
                v = row.Range.Cells(1, italicCol).value
                GetFontSettingsFromTable.italic = (NormalizeKey(CStr(v)) = "да" Or LCase$(Trim$(CStr(v))) = "yes")
            End If
            Exit Function
        End If
    Next row
End Function

Private Sub ApplyFontToTextRange(ByVal txtRange As Object, ByRef fs As fontSettingsType)
    With txtRange.Font
        If Len(fs.name) > 0 Then .name = fs.name
        If fs.Size > 0 Then .Size = fs.Size
        If fs.Color >= 0 Then .Fill.ForeColor.RGB = fs.Color
        .bold = fs.bold
        .italic = fs.italic
    End With
End Sub

Private Function ClampTransparency(ByVal value As Double) As Double
    If value >= 0 And value <= 1 Then
        value = value * 100
    End If
    ClampTransparency = Application.WorksheetFunction.Min(100, _
        Application.WorksheetFunction.Max(0, value))
End Function

Private Function ResolveTransparency(ByVal primary As Double, ByVal Fallback As Double) As Double
    If primary >= 0 Then
        ResolveTransparency = ClampTransparency(primary)
    Else
        ResolveTransparency = ClampTransparency(Fallback)
    End If
End Function

Private Function DarkenColor(ByVal colorValue As Long, ByVal factor As Double) As Long
    Dim r As Long
    Dim g As Long
    Dim b As Long

    r = colorValue Mod 256
    g = (colorValue \ 256) Mod 256
    b = (colorValue \ 65536) Mod 256

    r = Application.WorksheetFunction.Max(0, r * (1 - factor))
    g = Application.WorksheetFunction.Max(0, g * (1 - factor))
    b = Application.WorksheetFunction.Max(0, b * (1 - factor))

    DarkenColor = RGB(r, g, b)
End Function

Private Sub CollectUsedLegendItems(ByVal dataTable As ListObject, ByVal eventTypesTable As ListObject, _
                                   ByRef taskRowMap() As Long, ByVal rowCount As Long, _
                                   ByVal timelineStart As Date, ByVal timelineEnd As Date, _
                                   ByRef eventColumnIndices() As Long, ByRef eventDateColumnIndices() As Long, _
                                   ByVal barHeightNominal As Double, ByRef usedLineTypes() As String, _
                                   ByRef usedLineTypeCount As Long, ByRef usedEventTypes() As String, _
                                   ByRef usedEventTypeCount As Long, ByRef maxEventShapeSize As Double)
    Dim i As Long
    Dim eventIdx As Long
    Dim dataRowIndex As Long
    Dim startDate As Date
    Dim endDate As Date
    Dim lineType As String
    Dim eventName As String
    Dim eventDateValue As Variant
    Dim eventShape As MsoAutoShapeType
    Dim eventColor As Long
    Dim eventHeight As Double

    usedLineTypeCount = 0
    usedEventTypeCount = 0
    maxEventShapeSize = 5 * 2.83465

    For i = 1 To rowCount
        dataRowIndex = taskRowMap(i)
        startDate = dataTable.ListColumns("Дата начала").DataBodyRange.Cells(dataRowIndex, 1).value
        endDate = dataTable.ListColumns("Дата окончания").DataBodyRange.Cells(dataRowIndex, 1).value
        lineType = Trim$(CStr(dataTable.ListColumns("Тип линии").DataBodyRange.Cells(dataRowIndex, 1).value))

        If endDate >= startDate Then
            If Len(lineType) > 0 Then
                AddUniqueKey usedLineTypes, usedLineTypeCount, lineType
            End If

            For eventIdx = 1 To 10
                If eventColumnIndices(eventIdx) > 0 And eventDateColumnIndices(eventIdx) > 0 Then
                    eventName = Trim$(CStr(dataTable.ListColumns(eventColumnIndices(eventIdx)). _
                        DataBodyRange.Cells(dataRowIndex, 1).value))
                    eventDateValue = dataTable.ListColumns(eventDateColumnIndices(eventIdx)). _
                        DataBodyRange.Cells(dataRowIndex, 1).value
                    If Len(eventName) > 0 And IsDate(eventDateValue) Then
                        If CDate(eventDateValue) >= timelineStart And CDate(eventDateValue) <= timelineEnd Then
                            AddUniqueKey usedEventTypes, usedEventTypeCount, eventName
                            eventHeight = 0
                            eventShape = msoShapeOval
                            eventColor = RGB(91, 155, 213)
                            If GetEventTypeInfo(eventTypesTable, eventName, eventShape, eventColor, eventHeight) = False Then
                                eventHeight = 0
                            End If
                            If eventHeight <= 0 Then
                                eventHeight = barHeightNominal
                            End If
                            If eventHeight > maxEventShapeSize Then
                                maxEventShapeSize = eventHeight
                            End If
                        End If
                    End If
                End If
            Next eventIdx
        End If
    Next i
End Sub

Private Function GetLegendHeightFromCounts(ByVal lineCount As Long, ByVal eventCount As Long, _
                                           ByVal padding As Double, ByVal maxEventShapeSize As Double) As Double
    Dim rowsCount As Long
    Dim rowHeight As Double
    Dim symbolSize As Double
    Dim rowGapPoints As Double
    Dim extraBottomPoints As Double

    rowsCount = Application.WorksheetFunction.Max(lineCount, eventCount)
    If rowsCount <= 0 Then
        GetLegendHeightFromCounts = 0
        Exit Function
    End If

    ' Must match DrawLegend row sizing/gaps so background covers legend fully
    symbolSize = 5 * 2.83465
    rowHeight = Application.WorksheetFunction.Max(Application.WorksheetFunction.Max(symbolSize * 0.6, _
        maxEventShapeSize), 12)

    rowGapPoints = 2 * 2.83465        ' 2 мм межстрочный зазор
    extraBottomPoints = 3 * 2.83465   ' запас снизу от обрезания

    GetLegendHeightFromCounts = padding * 2 + (rowsCount * rowHeight) + ((rowsCount - 1) * rowGapPoints) + extraBottomPoints
End Function

Private Sub AddUniqueKey(ByRef keys() As String, ByRef keyCount As Long, ByVal newKey As String)
    Dim i As Long
    Dim normalizedKey As String

    normalizedKey = NormalizeKey(newKey)
    For i = 1 To keyCount
        If NormalizeKey(keys(i)) = normalizedKey Then
            Exit Sub
        End If
    Next i

    keyCount = keyCount + 1
    ReDim Preserve keys(1 To keyCount)
    keys(keyCount) = newKey
End Sub

Private Function IsKeyUsed(ByRef keys() As String, ByVal keyCount As Long, ByVal targetKey As String) As Boolean
    Dim i As Long
    Dim normalizedTarget As String

    normalizedTarget = NormalizeKey(targetKey)
    For i = 1 To keyCount
        If NormalizeKey(keys(i)) = normalizedTarget Then
            IsKeyUsed = True
            Exit Function
        End If
    Next i
End Function

Private Function GetTableColumnIndex(ByVal tableObj As ListObject, ByVal columnName As String) As Long
    Dim idx As Long

    On Error Resume Next
    idx = tableObj.ListColumns(columnName).Index
    On Error GoTo 0

    GetTableColumnIndex = idx
End Function

Private Function GetLineTypeColor(ByVal lineTypesTable As ListObject, ByVal lineType As String, _
                                  ByVal defaultColor As Long) As Long
    Dim row As ListRow

    For Each row In lineTypesTable.ListRows
        If CStr(row.Range.Cells(1, 1).value) = lineType Then
            If row.Range.Cells(1, 2).Interior.colorIndex <> xlColorIndexNone Then
                GetLineTypeColor = row.Range.Cells(1, 2).Interior.Color
            Else
                GetLineTypeColor = defaultColor
            End If
            Exit Function
        End If
    Next row

    GetLineTypeColor = defaultColor
End Function

Function GetLineTypeTransparency(ByVal lineTypesTable As ListObject, ByVal lineType As String, _
                                 ByVal defaultTransparency As Double) As Double
    Dim row As ListRow
    Dim transpIndex As Long
    Dim v As Variant

    If lineTypesTable Is Nothing Then
        GetLineTypeTransparency = defaultTransparency
        Exit Function
    End If

    transpIndex = GetTableColumnIndex(lineTypesTable, "Прозрачность")
    If transpIndex = 0 Then
        GetLineTypeTransparency = defaultTransparency
        Exit Function
    End If

    For Each row In lineTypesTable.ListRows
        If CStr(row.Range.Cells(1, 1).value) = lineType Then
            v = row.Range.Cells(1, transpIndex).value
            If IsNumeric(v) Then
                GetLineTypeTransparency = ClampTransparency(CDbl(v))
            Else
                GetLineTypeTransparency = defaultTransparency
            End If
            Exit Function
        End If
    Next row

    GetLineTypeTransparency = defaultTransparency
End Function

Function GetEventTypeTransparency(ByVal eventTypesTable As ListObject, ByVal eventName As String, _
                                  ByVal defaultTransparency As Double) As Double
    Dim row As ListRow
    Dim nameIndex As Long
    Dim transpIndex As Long
    Dim v As Variant

    If eventTypesTable Is Nothing Then
        GetEventTypeTransparency = defaultTransparency
        Exit Function
    End If

    nameIndex = GetTableColumnIndex(eventTypesTable, "Событие")
    transpIndex = GetTableColumnIndex(eventTypesTable, "Прозрачность")
    If nameIndex = 0 Or transpIndex = 0 Then
        GetEventTypeTransparency = defaultTransparency
        Exit Function
    End If

    For Each row In eventTypesTable.ListRows
        If NormalizeKey(CStr(row.Range.Cells(1, nameIndex).value)) = NormalizeKey(eventName) Then
            v = row.Range.Cells(1, transpIndex).value
            If IsNumeric(v) Then
                GetEventTypeTransparency = ClampTransparency(CDbl(v))
            Else
                GetEventTypeTransparency = defaultTransparency
            End If
            Exit Function
        End If
    Next row

    GetEventTypeTransparency = defaultTransparency
End Function

'' Возвращает значение колонки "Фигура" для данного типа события (имя фигуры на листе Справочники).
Private Function GetEventShapeName(ByVal eventTypesTable As ListObject, ByVal eventName As String) As String
    Dim row As ListRow
    Dim nameIndex As Long
    Dim shapeIndex As Long

    GetEventShapeName = ""
    If eventTypesTable Is Nothing Then Exit Function
    nameIndex = GetTableColumnIndex(eventTypesTable, "Событие")
    shapeIndex = GetTableColumnIndex(eventTypesTable, "Фигура")
    If nameIndex = 0 Or shapeIndex = 0 Then Exit Function

    For Each row In eventTypesTable.ListRows
        If NormalizeKey(CStr(row.Range.Cells(1, nameIndex).value)) = NormalizeKey(eventName) Then
            GetEventShapeName = Trim$(CStr(row.Range.Cells(1, shapeIndex).value))
            Exit Function
        End If
    Next row
End Function

Private Function GetEventTypeInfo(ByVal eventTypesTable As ListObject, ByVal eventName As String, _
                                  ByRef shapeType As MsoAutoShapeType, ByRef fillColor As Long, _
                                  ByRef heightPoints As Double) As Boolean
    Dim row As ListRow
    Dim nameIndex As Long
    Dim shapeIndex As Long
    Dim colorIndex As Long
    Dim heightIndex As Long
    Dim shapeName As String

    If eventTypesTable Is Nothing Then
        GetEventTypeInfo = False
        Exit Function
    End If

    nameIndex = GetTableColumnIndex(eventTypesTable, "Событие")
    shapeIndex = GetTableColumnIndex(eventTypesTable, "Фигура")
    colorIndex = GetTableColumnIndex(eventTypesTable, "Цвет")
    heightIndex = GetTableColumnIndex(eventTypesTable, "Высота, мм")

    If nameIndex = 0 Then
        GetEventTypeInfo = False
        Exit Function
    End If

    For Each row In eventTypesTable.ListRows
        If NormalizeKey(CStr(row.Range.Cells(1, nameIndex).value)) = NormalizeKey(eventName) Then
            If shapeIndex > 0 Then
                shapeName = CStr(row.Range.Cells(1, shapeIndex).value)
                shapeType = GetEventShapeType(shapeName)
            Else
                shapeType = msoShapeOval
            End If

            If colorIndex > 0 And row.Range.Cells(1, colorIndex).Interior.colorIndex <> xlColorIndexNone Then
                fillColor = row.Range.Cells(1, colorIndex).Interior.Color
            Else
                fillColor = RGB(91, 155, 213)
            End If

            heightPoints = 0
            If heightIndex > 0 And IsNumeric(row.Range.Cells(1, heightIndex).value) Then
                heightPoints = CDbl(row.Range.Cells(1, heightIndex).value) * 2.83465
            End If

            GetEventTypeInfo = True
            Exit Function
        End If
    Next row

    GetEventTypeInfo = False
End Function

Private Function GetEventShapeType(ByVal shapeName As String) As MsoAutoShapeType
    Select Case NormalizeKey(shapeName)
        Case "облако"
            GetEventShapeType = msoShapeCloud
        Case "стрелка вниз"
            GetEventShapeType = msoShapeDownArrow
        Case "стрелка вверх"
            GetEventShapeType = msoShapeUpArrow
        Case "стрелка влево"
            GetEventShapeType = msoShapeLeftArrow
        Case "стрелка вправо"
            GetEventShapeType = msoShapeRightArrow
        Case "звезда 4 луча", "звезда 4 лучей"
            GetEventShapeType = msoShape4pointStar
        Case "звезда 5 луча", "звезда 5 лучей"
            GetEventShapeType = msoShape5pointStar
        Case "ромб"
            GetEventShapeType = msoShapeDiamond
        Case "круг", "овал"
            GetEventShapeType = msoShapeOval
        Case Else
            GetEventShapeType = msoShapeOval
    End Select
End Function

Private Function BuildTimeline(ByVal startDate As Date, ByVal endDate As Date, ByVal periodName As String, _
                               ByVal dayWidth As Double, ByVal periodWidth As Double) As timelineData
    Dim data As timelineData
    Dim currentDate As Date
    Dim periodStart As Date
    Dim periodEnd As Date
    Dim label As String
    Dim count As Long
    Dim width As Double
    Dim quarterIndex As Long

    If endDate < startDate Then
        BuildTimeline = data
        Exit Function
    End If

    currentDate = startDate
    Do While currentDate <= endDate
        count = count + 1
        ReDim Preserve data.PeriodStarts(1 To count)
        ReDim Preserve data.PeriodEnds(1 To count)
        ReDim Preserve data.PeriodLabels(1 To count)
        ReDim Preserve data.PeriodWidths(1 To count)
        ReDim Preserve data.CumulativeWidths(1 To count)

        Select Case LCase$(periodName)
            Case "месяцы"
                periodStart = DateSerial(Year(currentDate), Month(currentDate), 1)
                periodEnd = DateSerial(Year(currentDate), Month(currentDate) + 1, 0)
                label = Format$(periodStart, "mmmm")
                width = periodWidth
                currentDate = DateAdd("m", 1, periodStart)
            Case "кварталы"
                quarterIndex = ((Month(currentDate) - 1) \ 3) + 1
                periodStart = DateSerial(Year(currentDate), (quarterIndex - 1) * 3 + 1, 1)
                periodEnd = DateSerial(Year(currentDate), (quarterIndex - 1) * 3 + 4, 0)
                label = CStr(quarterIndex) & " квартал"
                width = periodWidth
                currentDate = DateAdd("m", 3, periodStart)
            Case Else
                periodStart = currentDate
                periodEnd = currentDate
                label = Format$(currentDate, "dd.mm")
                width = dayWidth
                currentDate = DateAdd("d", 1, currentDate)
        End Select

        If periodEnd > endDate Then
            periodEnd = endDate
        End If

        data.PeriodStarts(count) = periodStart
        data.PeriodEnds(count) = periodEnd
        data.PeriodLabels(count) = label
        data.PeriodWidths(count) = width

        If count = 1 Then
            data.CumulativeWidths(count) = 0
        Else
            data.CumulativeWidths(count) = data.CumulativeWidths(count - 1) + data.PeriodWidths(count - 1)
        End If

        data.TotalWidth = data.TotalWidth + width
    Loop

    data.count = count
    BuildTimeline = data
End Function

Private Sub ScaleTimeline(ByRef timelineData As timelineData, ByVal scaleFactor As Double)
    Dim i As Long

    If scaleFactor <= 0 Then
        Exit Sub
    End If

    For i = 1 To timelineData.count
        timelineData.PeriodWidths(i) = timelineData.PeriodWidths(i) * scaleFactor
    Next i

    timelineData.TotalWidth = 0
    For i = 1 To timelineData.count
        If i = 1 Then
            timelineData.CumulativeWidths(i) = 0
        Else
            timelineData.CumulativeWidths(i) = timelineData.CumulativeWidths(i - 1) + timelineData.PeriodWidths(i - 1)
        End If
        timelineData.TotalWidth = timelineData.TotalWidth + timelineData.PeriodWidths(i)
    Next i
End Sub

Private Sub NormalizeTimelinePeriodWidths(ByRef timelineData As timelineData)
    Dim i As Long
    Dim uniformWidth As Double

    If timelineData.count = 0 Then
        Exit Sub
    End If

    uniformWidth = timelineData.TotalWidth / timelineData.count
    If uniformWidth <= 0 Then
        Exit Sub
    End If

    For i = 1 To timelineData.count
        timelineData.PeriodWidths(i) = uniformWidth
    Next i

    timelineData.TotalWidth = 0
    For i = 1 To timelineData.count
        If i = 1 Then
            timelineData.CumulativeWidths(i) = 0
        Else
            timelineData.CumulativeWidths(i) = timelineData.CumulativeWidths(i - 1) + timelineData.PeriodWidths(i - 1)
        End If
        timelineData.TotalWidth = timelineData.TotalWidth + timelineData.PeriodWidths(i)
    Next i
End Sub

Private Function GetTimelineOffset(ByRef timelineData As timelineData, ByVal targetDate As Date) As Double
    Dim i As Long
    Dim daysInPeriod As Long
    Dim ratio As Double

    If timelineData.count = 0 Then
        GetTimelineOffset = 0
        Exit Function
    End If

    If targetDate <= timelineData.PeriodStarts(1) Then
        GetTimelineOffset = 0
        Exit Function
    End If

    If targetDate >= timelineData.PeriodEnds(timelineData.count) + 1 Then
        GetTimelineOffset = timelineData.TotalWidth
        Exit Function
    End If

    For i = 1 To timelineData.count
        If targetDate >= timelineData.PeriodStarts(i) And targetDate <= timelineData.PeriodEnds(i) + 1 Then
            daysInPeriod = DateDiff("d", timelineData.PeriodStarts(i), timelineData.PeriodEnds(i)) + 1
            If daysInPeriod <= 0 Then
                GetTimelineOffset = timelineData.CumulativeWidths(i)
            Else
                ratio = DateDiff("d", timelineData.PeriodStarts(i), targetDate) / daysInPeriod
                GetTimelineOffset = timelineData.CumulativeWidths(i) + timelineData.PeriodWidths(i) * ratio
            End If
            Exit Function
        End If
    Next i

    GetTimelineOffset = timelineData.TotalWidth
End Function

Private Function GetPeriodHeaderHeight(ByRef timelineData As timelineData, ByVal baseHeight As Double, _
                                       ByVal fontSize As Long, ByVal rotateAllPeriods As Boolean) As Double
    Dim i As Long
    Dim maxTextSize As Double
    Dim currentSize As Double

    maxTextSize = baseHeight
    For i = 1 To timelineData.count
        If rotateAllPeriods Then
            currentSize = EstimateTextWidth(timelineData.PeriodLabels(i), fontSize)
        Else
            currentSize = EstimateTextHeight(fontSize)
        End If
        If currentSize > maxTextSize Then
            maxTextSize = currentSize
        End If
    Next i

    GetPeriodHeaderHeight = Application.WorksheetFunction.Max(baseHeight, maxTextSize + 4)
End Function

Private Function ShouldRotatePeriodLabel(ByVal labelText As String, ByVal blockWidth As Double, _
                                         ByVal fontSize As Long) As Boolean
    Dim textWidth As Double

    textWidth = EstimateTextWidth(labelText, fontSize)
    ShouldRotatePeriodLabel = textWidth > blockWidth - 2
End Function

Private Function ShouldRotateAllPeriodLabels(ByRef timelineData As timelineData, ByVal fontSize As Long) As Boolean
    Dim i As Long

    For i = 1 To timelineData.count
        If ShouldRotatePeriodLabel(timelineData.PeriodLabels(i), timelineData.PeriodWidths(i), fontSize) Then
            ShouldRotateAllPeriodLabels = True
            Exit Function
        End If
    Next i

    ShouldRotateAllPeriodLabels = False
End Function

'' Измеряет ширину текста в пунктах с учётом фактического шрифта (имя, размер, bold, italic).
'' При ошибке API возвращает приблизительную ширину через EstimateTextWidth.
Private Function MeasureTextWidthWithFont(ByVal text As String, ByVal fontName As String, _
                                         ByVal fontSize As Long, ByVal bold As Boolean, ByVal italic As Boolean) As Double
#If VBA7 Then
    Dim hdc As LongPtr
    Dim hFont As LongPtr
    Dim hOld As LongPtr
    Dim lf As LOGFONTW
    Dim sz As API_SIZE
    Dim dpi As Long
    Dim fontPt As Long
    Dim i As Long
    Dim fname As String
#Else
    Dim hdc As Long
    Dim hFont As Long
    Dim hOld As Long
    Dim lf As LOGFONTA
    Dim sz As API_SIZE
    Dim dpi As Long
    Dim fontPt As Long
    Dim i As Long
#End If

    On Error GoTo Fallback

    If Len(text) = 0 Then
        MeasureTextWidthWithFont = 0
        Exit Function
    End If

    If Len(Trim$(fontName)) = 0 Then fontName = "Arial"
    fontPt = CLng(fontSize)
    If fontPt < 1 Then fontPt = 10

#If VBA7 Then
    hdc = GetDC(0)
    If hdc = 0 Then GoTo Fallback
    dpi = GetDeviceCaps(hdc, LOGPIXELSY)
    If dpi < 1 Then dpi = 96
    lf.lfHeight = -MulDiv(fontPt, dpi, 72)
    lf.lfWidth = 0
    lf.lfWeight = IIf(bold, FW_BOLD, FW_NORMAL)
    lf.lfItalic = IIf(italic, 1, 0)
    lf.lfCharSet = 1
    lf.lfFaceName(0) = 0
    fname = Left$(fontName & String$(32, vbNullChar), 32)
    For i = 0 To 31
        lf.lfFaceName(i) = 0
    Next i
    For i = 1 To Len(fname)
        lf.lfFaceName(i - 1) = AscW(Mid$(fname, i, 1))
    Next i
    hFont = CreateFontIndirect(lf)
    If hFont = 0 Then
        ReleaseDC 0, hdc
        GoTo Fallback
    End If
    hOld = SelectObject(hdc, hFont)
    If GetTextExtentPoint32(hdc, StrPtr(text), Len(text), sz) = 0 Then
        SelectObject hdc, hOld
        DeleteObject hFont
        ReleaseDC 0, hdc
        GoTo Fallback
    End If
    SelectObject hdc, hOld
    DeleteObject hFont
    ReleaseDC 0, hdc
    MeasureTextWidthWithFont = sz.cx * 72 / dpi
#Else
    hdc = GetDC(0)
    If hdc = 0 Then GoTo Fallback
    dpi = GetDeviceCaps(hdc, LOGPIXELSY)
    If dpi < 1 Then dpi = 96
    lf.lfHeight = -MulDiv(fontPt, dpi, 72)
    lf.lfWidth = 0
    lf.lfWeight = IIf(bold, FW_BOLD, FW_NORMAL)
    lf.lfItalic = IIf(italic, 1, 0)
    lf.lfCharSet = 1
    For i = 0 To 31
        lf.lfFaceName(i) = 0
    Next i
    For i = 1 To Application.WorksheetFunction.Min(Len(fontName), 31)
        lf.lfFaceName(i - 1) = Asc(Mid$(fontName, i, 1))
    Next i
    hFont = CreateFontIndirect(lf)
    If hFont = 0 Then
        ReleaseDC 0, hdc
        GoTo Fallback
    End If
    hOld = SelectObject(hdc, hFont)
    If GetTextExtentPoint32(hdc, text, Len(text), sz) = 0 Then
        SelectObject hdc, hOld
        DeleteObject hFont
        ReleaseDC 0, hdc
        GoTo Fallback
    End If
    SelectObject hdc, hOld
    DeleteObject hFont
    ReleaseDC 0, hdc
    MeasureTextWidthWithFont = sz.cx * 72 / dpi
#End If
    Exit Function

Fallback:
    MeasureTextWidthWithFont = EstimateTextWidth(text, fontPt)
End Function

Private Function EstimateTextWidth(ByVal labelText As String, ByVal fontSize As Long) As Double
    Dim i As Long
    Dim ch As String
    Dim total As Double
    Dim avgWidth As Double
    Const NARROW As Double = 0.4
    Const NORMAL As Double = 0.55
    Const WIDE As Double = 0.7
    Const SPACE_WIDTH As Double = 0.25

    If Len(labelText) = 0 Then
        EstimateTextWidth = 0
        Exit Function
    End If
    avgWidth = fontSize * NORMAL
    total = 0
    For i = 1 To Len(labelText)
        ch = Mid$(labelText, i, 1)
        Select Case ch
            Case " "
                total = total + fontSize * SPACE_WIDTH
            Case "(", ")", ",", ".", "'", ":", ";", "!", "i", "l", "1", "I", "[", "]", "{", "}", "\", "/", "|", "`", "~", "-", "_", "=", "+"
                total = total + fontSize * NARROW
            Case "M", "W", "m", "w", "@", "0", "O", "Q", "%", "#", "&"
                total = total + fontSize * WIDE
            Case Else
                total = total + avgWidth
        End Select
    Next i
    EstimateTextWidth = total
End Function

Private Function EstimateTextHeight(ByVal fontSize As Long) As Double
    EstimateTextHeight = fontSize * 1.2
End Function

Private Function GetContrastingTextColor(ByVal fillColor As Long) As Long
    Dim r As Long
    Dim g As Long
    Dim b As Long
    Dim luminance As Double

    r = fillColor Mod 256
    g = (fillColor \ 256) Mod 256
    b = (fillColor \ 65536) Mod 256

    luminance = 0.299 * r + 0.587 * g + 0.114 * b
    If luminance < 128 Then
        GetContrastingTextColor = RGB(255, 255, 255)
    Else
        GetContrastingTextColor = RGB(0, 0, 0)
    End If
End Function




