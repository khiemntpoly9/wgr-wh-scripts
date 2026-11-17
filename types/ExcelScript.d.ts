/**
 * Stub type declarations cho ExcelScript namespace.
 * Cung cấp đủ type để IDE không báo lỗi "Cannot find namespace 'ExcelScript'".
 *
 * Lưu ý: Đây là bản khai báo tối giản, chỉ bao gồm các API đang sử dụng.
 * Khi cần thêm API mới, bổ sung khai báo vào đây.
 *
 * Tham khảo đầy đủ:
 * https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript
 */

declare namespace ExcelScript {
  // ─── Core Objects ──────────────────────────────────────────────────

  interface Workbook {
    getActiveWorksheet(): Worksheet;
    getWorksheet(name: string): Worksheet | undefined;
    getWorksheets(): Worksheet[];
    addWorksheet(name?: string): Worksheet;
    /**
     * Gets a workbook-scoped named item. Returns undefined when the name does not exist.
     * https://learn.microsoft.com/javascript/api/office-scripts/excelscript/excelscript.workbook#excelscript-excelscript-workbook-getnameditem-member(1)
     */
    getNamedItem(name: string): NamedItem | undefined;
  }

  interface Worksheet {
    getName(): string;
    setName(name: string): void;
    getRange(address?: string): Range;
    /**
     * Gets a range object containing the rectangular range specified by indexes.
     * @param startRow - Zero-based row index of the top-left cell.
     * @param startColumn - Zero-based column index of the top-left cell.
     * @param rowCount - Number of rows in the range.
     * @param columnCount - Number of columns in the range.
     */
    getRangeByIndexes(
      startRow: number,
      startColumn: number,
      rowCount: number,
      columnCount: number,
    ): Range;
    getUsedRange(valuesOnly?: boolean): Range | undefined;
    getRanges(address: string): RangeAreas;
    getCell(row: number, column: number): Range;
    copy(
      positionType?: WorksheetPositionType,
      relativeTo?: Worksheet,
    ): Worksheet;
    delete(): void;
    getPosition(): number;
    setPosition(position: number): void;
    getVisibility(): SheetVisibility;
    setVisibility(visibility: SheetVisibility): void;
    getAutoFilter(): AutoFilter;
    getFreezePanes(): WorksheetFreezePanes;
    getCharts(): Chart[];
    addChart(
      chartType: ChartType,
      sourceData: Range,
      seriesBy?: ChartSeriesBy,
    ): Chart;
    activate(): void;
    getTables(): Table[];
    addTable(address: Range | string, hasHeaders: boolean): Table;
    /**
     * Gets a worksheet-scoped named item. Returns undefined when the name does not exist.
     * https://learn.microsoft.com/javascript/api/office-scripts/excelscript/excelscript.worksheet#excelscript-excelscript-worksheet-getnameditem-member(1)
     */
    getNamedItem(name: string): NamedItem | undefined;
  }

  /** A defined name (Name Manager). Scope is either the workbook or one worksheet. */
  interface NamedItem {
    getName(): string;
    /** Range the name points at. Throws when the name is not a range. */
    getRange(): Range;
  }

  interface Table {
    delete(): void;
    getName(): string;
    setName(name: string): void;
    setShowFilterButton(showFilterButton: boolean): void;
    getRange(): Range;
    getAutoFilter(): AutoFilter;
  }

  interface Chart {
    delete(): void;
    getTitle(): ChartTitle;
    setPosition(startCell: Range | string, endCell?: Range | string): void;
  }

  interface ChartTitle {
    getText(): string;
    setText(text: string): void;
    setVisible(visible: boolean): void;
  }

  interface WorksheetFreezePanes {
    freezeAt(frozenRange: Range | string): void;
    freezeColumns(count: number): void;
    freezeRows(count: number): void;
    getLocation(): Range;
    unfreeze(): void;
  }

  interface AutoFilter {
    apply(range: Range | string, columnIndex?: number, criteria?: FilterCriteria): void;
    clearCriteria(): void;
    remove(): void;
    reapply(): void;
    getEnabled(): boolean;
    getIsDataFiltered(): boolean;
    getCriteria(): FilterCriteria[];
    getRange(): Range;
  }

  enum ChartType {
    columnClustered = "ColumnClustered",
    columnStacked = "ColumnStacked",
    line = "Line",
    pie = "Pie",
    barClustered = "BarClustered",
  }

  enum ChartSeriesBy {
    auto = "Auto",
    columns = "Columns",
    rows = "Rows",
  }

  interface FilterCriteria {
    color?: string;
    criterion1?: string;
    criterion2?: string;
    dynamicCriteria?: DynamicFilterCriteria;
    filterOn: FilterOn;
    icon?: Icon;
    operator?: FilterOperator;
    subField?: string;
    values?: (string | FilterDatetime)[];
  }

  interface FilterDatetime {
    date: string;
    specificity: FilterDatetimeSpecificity;
  }

  enum FilterOn {
    bottomItems = "BottomItems",
    bottomPercent = "BottomPercent",
    cellColor = "CellColor",
    dynamic = "Dynamic",
    fontColor = "FontColor",
    values = "Values",
    topItems = "TopItems",
    topPercent = "TopPercent",
    icon = "Icon",
    custom = "Custom",
  }

  enum FilterOperator {
    and = "And",
    or = "Or",
  }

  enum DynamicFilterCriteria {
    unknown = "Unknown",
    aboveAverage = "AboveAverage",
    allDatesInPeriodApril = "AllDatesInPeriodApril",
    allDatesInPeriodAugust = "AllDatesInPeriodAugust",
    allDatesInPeriodDecember = "AllDatesInPeriodDecember",
    allDatesInPeriodFebruary = "AllDatesInPeriodFebruary",
    allDatesInPeriodJanuary = "AllDatesInPeriodJanuary",
    allDatesInPeriodJuly = "AllDatesInPeriodJuly",
    allDatesInPeriodJune = "AllDatesInPeriodJune",
    allDatesInPeriodMarch = "AllDatesInPeriodMarch",
    allDatesInPeriodMay = "AllDatesInPeriodMay",
    allDatesInPeriodNovember = "AllDatesInPeriodNovember",
    allDatesInPeriodOctober = "AllDatesInPeriodOctober",
    allDatesInPeriodQuarter1 = "AllDatesInPeriodQuarter1",
    allDatesInPeriodQuarter2 = "AllDatesInPeriodQuarter2",
    allDatesInPeriodQuarter3 = "AllDatesInPeriodQuarter3",
    allDatesInPeriodQuarter4 = "AllDatesInPeriodQuarter4",
    allDatesInPeriodSeptember = "AllDatesInPeriodSeptember",
    belowAverage = "BelowAverage",
    lastMonth = "LastMonth",
    lastQuarter = "LastQuarter",
    lastWeek = "LastWeek",
    lastYear = "LastYear",
    nextMonth = "NextMonth",
    nextQuarter = "NextQuarter",
    nextWeek = "NextWeek",
    nextYear = "NextYear",
    thisMonth = "ThisMonth",
    thisQuarter = "ThisQuarter",
    thisWeek = "ThisWeek",
    thisYear = "ThisYear",
    today = "Today",
    tomorrow = "Tomorrow",
    yearToDate = "YearToDate",
    yesterday = "Yesterday",
  }

  enum FilterDatetimeSpecificity {
    year = "Year",
    month = "Month",
    day = "Day",
    hour = "Hour",
    minute = "Minute",
    second = "Second",
  }

  interface Range {
    getValue(): string | number | boolean;
    getValues(): (string | number | boolean)[][];
    setValue(value: string | number | boolean): void;
    setValues(values: (string | number | boolean)[][]): void;
    getText(): string;
    getTexts(): string[][];
    getAddress(): string;
    /** Worksheet that contains this range. */
    getWorksheet(): Worksheet;
    getRowIndex(): number;
    getColumnIndex(): number;
    getRowCount(): number;
    getColumnCount(): number;
    getFormat(): RangeFormat;
    getNumberFormat(): string;
    setNumberFormat(numberFormat: string): void;
    getNumberFormatLocal(): string;
    setNumberFormatLocal(numberFormatLocal: string): void;
    merge(across?: boolean): void;
    unmerge(): void;
    getExtendedRange(direction: KeyboardDirection, activeCell?: Range): Range;
    getRangeEdge(direction: KeyboardDirection, activeCell?: Range): Range;
    getResizedRange(deltaRows: number, deltaColumns: number): Range;
    getOffsetRange(rowOffset: number, columnOffset: number): Range;
    getEntireColumn(): Range;
    getEntireRow(): Range;
    getColumn(column: number): Range;
    getRow(row: number): Range;
    getLastCell(): Range;
    getLastColumn(): Range;
    getLastRow(): Range;
    clear(applyTo?: ClearApplyTo): void;
    copyFrom(
      sourceRange: Range | string,
      copyType?: RangeCopyType,
      skipBlanks?: boolean,
      transpose?: boolean,
    ): void;
    setColumnHidden(columnHidden: boolean): void;
    getColumnHidden(): boolean;
    setRowHidden(rowHidden: boolean): void;
    getRowHidden(): boolean;
    getIsEntireColumn(): boolean;
    getIsEntireRow(): boolean;
    getFormula(): string;
    setFormula(formula: string): void;
    getFormulas(): string[][];
    setFormulas(formulas: string[][]): void;
    addConditionalFormat(
      type: ConditionalFormatType,
    ): ConditionalFormat;
    getConditionalFormats(): ConditionalFormat[];
    clearAllConditionalFormats(): void;
    calculate(): void;
    delete(shift: DeleteShiftDirection): void;
    insert(shift: InsertShiftDirection): void;
    getHyperlink(): RangeHyperlink;
    setHyperlink(hyperlink: RangeHyperlink): void;
    removeDuplicates(columns: number[], includesHeader: boolean): RemoveDuplicatesResult;
    find(text: string, criteria: SearchCriteria): Range | undefined;
    replaceAll(text: string, replacement: string, criteria: ReplaceCriteria): number;
    getSort(): RangeSort;
    getSpecialCells(
      cellType: SpecialCellType,
      cellValueType?: SpecialCellValueType,
    ): RangeAreas;
    getSurroundingRegion(): Range;
    getAbsoluteResizedRange(numRows: number, numColumns: number): Range;
    getIntersection(anotherRange: Range | string): Range;
  }

  interface RangeAreas {
    getAreas(): Range[];
    getAddress(): string;
  }

  interface RangeFormat {
    getFont(): RangeFont;
    getFill(): RangeFill;
    getBorders(): RangeBorder[];
    getRangeBorder(index: BorderIndex): RangeBorder;
    setHorizontalAlignment(
      horizontalAlignment: HorizontalAlignment,
    ): void;
    getHorizontalAlignment(): HorizontalAlignment;
    setVerticalAlignment(verticalAlignment: VerticalAlignment): void;
    getVerticalAlignment(): VerticalAlignment;
    setWrapText(wrapText: boolean): void;
    getWrapText(): boolean;
    setRowHeight(height: number): void;
    getRowHeight(): number;
    setColumnWidth(width: number): void;
    getColumnWidth(): number;
    autofitColumns(): void;
    autofitRows(): void;
    setIndentLevel(indentLevel: number): void;
    getIndentLevel(): number;
    setShrinkToFit(shrinkToFit: boolean): void;
    getShrinkToFit(): boolean;
    setTextOrientation(textOrientation: number): void;
    getTextOrientation(): number;
    getProtection(): FormatProtection;
    setAutoIndent(autoIndent: boolean): void;
    getAutoIndent(): boolean;
  }

  interface RangeFont {
    setName(name: string): void;
    getName(): string;
    setSize(size: number): void;
    getSize(): number;
    setBold(bold: boolean): void;
    getBold(): boolean;
    setItalic(italic: boolean): void;
    getItalic(): boolean;
    setColor(color: string): void;
    getColor(): string;
    setUnderline(underline: RangeUnderlineStyle): void;
    getUnderline(): RangeUnderlineStyle;
    setStrikethrough(strikethrough: boolean): void;
    getStrikethrough(): boolean;
    setTintAndShade(tintAndShade: number): void;
    getTintAndShade(): number;
    setSuperscript(superscript: boolean): void;
    getSuperscript(): boolean;
    setSubscript(subscript: boolean): void;
    getSubscript(): boolean;
  }

  interface RangeFill {
    setColor(color: string): void;
    getColor(): string;
    clear(): void;
    setPattern(pattern: FillPattern): void;
    getPattern(): FillPattern;
    setPatternColor(patternColor: string): void;
    getPatternColor(): string;
    setPatternTintAndShade(patternTintAndShade: number): void;
    getPatternTintAndShade(): number;
    setTintAndShade(tintAndShade: number): void;
    getTintAndShade(): number;
  }

  interface RangeBorder {
    setStyle(style: BorderLineStyle): void;
    getStyle(): BorderLineStyle;
    setColor(color: string): void;
    getColor(): string;
    setWeight(weight: BorderWeight): void;
    getWeight(): BorderWeight;
    setSideIndex(sideIndex: BorderIndex): void;
    getSideIndex(): BorderIndex;
    setTintAndShade(tintAndShade: number): void;
    getTintAndShade(): number;
  }

  interface FormatProtection {
    setFormulaHidden(formulaHidden: boolean): void;
    getFormulaHidden(): boolean;
    setLocked(locked: boolean): void;
    getLocked(): boolean;
  }

  // ─── Sorting ───────────────────────────────────────────────────────

  interface RangeSort {
    apply(
      fields: SortField[],
      matchCase?: boolean,
      hasHeaders?: boolean,
      orientation?: SortOrientation,
      method?: SortMethod,
    ): void;
  }

  interface SortField {
    ascending?: boolean;
    color?: string;
    dataOption?: SortDataOption;
    icon?: Icon;
    key: number;
    sortOn?: SortOn;
    subField?: string;
  }

  // ─── Conditional Formatting ────────────────────────────────────────

  interface ConditionalFormat {
    getCellValue(): CellValueConditionalFormat;
    getColorScale(): ColorScaleConditionalFormat;
    getCustom(): CustomConditionalFormat;
    getDataBar(): DataBarConditionalFormat;
    getIconSet(): IconSetConditionalFormat;
    getPreset(): PresetCriteriaConditionalFormat;
    getTextComparison(): TextConditionalFormat;
    getTopBottom(): TopBottomConditionalFormat;
    delete(): void;
    getPriority(): number;
    setPriority(priority: number): void;
    getStopIfTrue(): boolean;
    setStopIfTrue(stopIfTrue: boolean): void;
    getType(): ConditionalFormatType;
    getRange(): Range;
    getRanges(): RangeAreas;
  }

  interface CellValueConditionalFormat {
    getFormat(): ConditionalRangeFormat;
    getRule(): ConditionalCellValueRule;
    setRule(rule: ConditionalCellValueRule): void;
  }

  interface CustomConditionalFormat {
    getFormat(): ConditionalRangeFormat;
    getRule(): ConditionalFormatRule;
  }

  interface ConditionalFormatRule {
    getFormula(): string;
    setFormula(formula: string): void;
    getFormulaLocal(): string;
    setFormulaLocal(formulaLocal: string): void;
  }

  interface ConditionalRangeFormat {
    getBorders(): ConditionalRangeBorder[];
    getConditionalRangeBorder(index: ConditionalRangeBorderIndex): ConditionalRangeBorder;
    getFill(): ConditionalRangeFill;
    getFont(): ConditionalRangeFont;
    getNumberFormat(): ConditionalFormatNumberFormat;
  }

  interface ConditionalRangeBorder {
    setColor(color: string): void;
    getColor(): string;
    setStyle(style: ConditionalRangeBorderLineStyle): void;
    getStyle(): ConditionalRangeBorderLineStyle;
    getSideIndex(): ConditionalRangeBorderIndex;
  }

  interface ConditionalRangeFill {
    setColor(color: string): void;
    getColor(): string;
    clear(): void;
  }

  interface ConditionalRangeFont {
    setBold(bold: boolean): void;
    getBold(): boolean;
    setColor(color: string): void;
    getColor(): string;
    setItalic(italic: boolean): void;
    getItalic(): boolean;
    setStrikethrough(strikethrough: boolean): void;
    getStrikethrough(): boolean;
    setUnderline(underline: ConditionalRangeFontUnderlineStyle): void;
    getUnderline(): ConditionalRangeFontUnderlineStyle;
  }

  interface ConditionalFormatNumberFormat {
    setFormat(format: string): void;
    getFormat(): string;
  }

  interface ConditionalCellValueRule {
    formula1: string;
    formula2?: string;
    operator: ConditionalCellValueOperator;
  }

  // Stub interfaces for conditional formatting types not fully detailed
  interface ColorScaleConditionalFormat {}
  interface DataBarConditionalFormat {}
  interface IconSetConditionalFormat {}
  interface PresetCriteriaConditionalFormat {}
  interface TopBottomConditionalFormat {}

  interface TextConditionalFormat {
    getFormat(): ConditionalRangeFormat;
    getRule(): ConditionalTextComparisonRule;
    setRule(rule: ConditionalTextComparisonRule): void;
  }

  interface ConditionalTextComparisonRule {
    operator: ConditionalTextOperator;
    text: string;
  }

  // ─── Search & Replace ──────────────────────────────────────────────

  interface SearchCriteria {
    completeMatch?: boolean;
    matchCase?: boolean;
    searchDirection?: SearchDirection;
  }

  interface ReplaceCriteria {
    completeMatch?: boolean;
    matchCase?: boolean;
  }

  interface RemoveDuplicatesResult {
    getRemoved(): number;
    getUniqueRemaining(): number;
  }

  // ─── Hyperlinks ────────────────────────────────────────────────────

  interface RangeHyperlink {
    address?: string;
    documentReference?: string;
    screenTip?: string;
    textToDisplay?: string;
  }

  // ─── Icon ──────────────────────────────────────────────────────────

  interface Icon {
    index: number;
    set: IconSet;
  }

  // ─── Enums ─────────────────────────────────────────────────────────

  enum HorizontalAlignment {
    general = "General",
    left = "Left",
    center = "Center",
    right = "Right",
    fill = "Fill",
    justify = "Justify",
    centerAcrossSelection = "CenterAcrossSelection",
    distributed = "Distributed",
  }

  enum VerticalAlignment {
    top = "Top",
    center = "Center",
    bottom = "Bottom",
    justify = "Justify",
    distributed = "Distributed",
  }

  enum KeyboardDirection {
    left = "Left",
    right = "Right",
    up = "Up",
    down = "Down",
  }

  enum WorksheetPositionType {
    none = "None",
    before = "Before",
    after = "After",
    beginning = "Beginning",
    end = "End",
  }

  enum SheetVisibility {
    visible = "Visible",
    hidden = "Hidden",
    veryHidden = "VeryHidden",
  }

  enum BorderIndex {
    edgeTop = "EdgeTop",
    edgeBottom = "EdgeBottom",
    edgeLeft = "EdgeLeft",
    edgeRight = "EdgeRight",
    insideVertical = "InsideVertical",
    insideHorizontal = "InsideHorizontal",
    diagonalDown = "DiagonalDown",
    diagonalUp = "DiagonalUp",
  }

  enum BorderLineStyle {
    none = "None",
    continuous = "Continuous",
    dash = "Dash",
    dashDot = "DashDot",
    dashDotDot = "DashDotDot",
    dot = "Dot",
    double = "Double",
    slantDashDot = "SlantDashDot",
  }

  enum BorderWeight {
    hairline = "Hairline",
    thin = "Thin",
    medium = "Medium",
    thick = "Thick",
  }

  enum RangeUnderlineStyle {
    none = "None",
    single = "Single",
    double = "Double",
    singleAccountant = "SingleAccountant",
    doubleAccountant = "DoubleAccountant",
  }

  enum ClearApplyTo {
    all = "All",
    formats = "Formats",
    contents = "Contents",
    hyperlinks = "Hyperlinks",
    removeHyperlinks = "RemoveHyperlinks",
  }

  enum RangeCopyType {
    all = "All",
    formulas = "Formulas",
    values = "Values",
    formats = "Formats",
    link = "Link",
  }

  enum DeleteShiftDirection {
    up = "Up",
    left = "Left",
  }

  enum InsertShiftDirection {
    down = "Down",
    right = "Right",
  }

  enum FillPattern {
    none = "None",
    solid = "Solid",
    gray50 = "Gray50",
    gray75 = "Gray75",
    gray25 = "Gray25",
    horizontal = "Horizontal",
    vertical = "Vertical",
    down = "Down",
    up = "Up",
    checker = "Checker",
    semiGray75 = "SemiGray75",
    lightHorizontal = "LightHorizontal",
    lightVertical = "LightVertical",
    lightDown = "LightDown",
    lightUp = "LightUp",
    grid = "Grid",
    crissCross = "CrissCross",
    gray16 = "Gray16",
    gray8 = "Gray8",
    linearGradient = "LinearGradient",
    rectangularGradient = "RectangularGradient",
  }

  enum ConditionalFormatType {
    custom = "Custom",
    dataBar = "DataBar",
    colorScale = "ColorScale",
    iconSet = "IconSet",
    topBottom = "TopBottom",
    presetCriteria = "PresetCriteria",
    containsText = "ContainsText",
    cellValue = "CellValue",
  }

  enum ConditionalCellValueOperator {
    invalid = "Invalid",
    between = "Between",
    notBetween = "NotBetween",
    equalTo = "EqualTo",
    notEqualTo = "NotEqualTo",
    greaterThan = "GreaterThan",
    lessThan = "LessThan",
    greaterThanOrEqual = "GreaterThanOrEqual",
    lessThanOrEqual = "LessThanOrEqual",
  }

  enum ConditionalTextOperator {
    contains = "Contains",
    notContains = "NotContains",
    beginsWith = "BeginsWith",
    endsWith = "EndsWith",
  }

  enum ConditionalRangeBorderIndex {
    edgeTop = "EdgeTop",
    edgeBottom = "EdgeBottom",
    edgeLeft = "EdgeLeft",
    edgeRight = "EdgeRight",
  }

  enum ConditionalRangeBorderLineStyle {
    none = "None",
    continuous = "Continuous",
    dash = "Dash",
    dashDot = "DashDot",
    dashDotDot = "DashDotDot",
    dot = "Dot",
  }

  enum ConditionalRangeFontUnderlineStyle {
    none = "None",
    single = "Single",
    double = "Double",
  }

  enum SearchDirection {
    forward = "Forward",
    backwards = "Backwards",
  }

  enum SortOrientation {
    rows = "Rows",
    columns = "Columns",
  }

  enum SortMethod {
    pinYin = "PinYin",
    strokeCount = "StrokeCount",
  }

  enum SortDataOption {
    normal = "Normal",
    textAsNumber = "TextAsNumber",
  }

  enum SortOn {
    value = "Value",
    cellColor = "CellColor",
    fontColor = "FontColor",
    icon = "Icon",
  }

  enum IconSet {
    invalid = "Invalid",
    threeArrows = "ThreeArrows",
    threeArrowsGray = "ThreeArrowsGray",
    threeFlags = "ThreeFlags",
    threeTrafficLights1 = "ThreeTrafficLights1",
    threeTrafficLights2 = "ThreeTrafficLights2",
    threeSigns = "ThreeSigns",
    threeSymbols = "ThreeSymbols",
    threeSymbols2 = "ThreeSymbols2",
    fourArrows = "FourArrows",
    fourArrowsGray = "FourArrowsGray",
    fourRedToBlack = "FourRedToBlack",
    fourRating = "FourRating",
    fourTrafficLights = "FourTrafficLights",
    fiveArrows = "FiveArrows",
    fiveArrowsGray = "FiveArrowsGray",
    fiveRating = "FiveRating",
    fiveQuarters = "FiveQuarters",
    threeStars = "ThreeStars",
    threeTriangles = "ThreeTriangles",
    fiveBoxes = "FiveBoxes",
  }

  enum SpecialCellType {
    conditionalFormats = "ConditionalFormats",
    dataValidations = "DataValidations",
    blanks = "Blanks",
    constants = "Constants",
    formulas = "Formulas",
    sameConditionalFormat = "SameConditionalFormat",
    sameDataValidation = "SameDataValidation",
    visible = "Visible",
  }

  enum SpecialCellValueType {
    all = "All",
    errors = "Errors",
    errorsLogical = "ErrorsLogical",
    errorsNumbers = "ErrorsNumbers",
    errorsText = "ErrorsText",
    errorsLogicalNumber = "ErrorsLogicalNumber",
    errorsLogicalText = "ErrorsLogicalText",
    errorsNumberText = "ErrorsNumberText",
    logical = "Logical",
    logicalNumbers = "LogicalNumbers",
    logicalText = "LogicalText",
    logicalNumbersText = "LogicalNumbersText",
    numbers = "Numbers",
    numbersText = "NumbersText",
    text = "Text",
  }
}
