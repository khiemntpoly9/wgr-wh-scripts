export const createMockRange = (
  {
    values = [],
    formulas = null,
    rowIndex = 0,
    rowCount = values.length,
    intersection = null,
  } = {},
) => ({
  getValues() {
    return values;
  },
  getFormulas() {
    if (formulas) return formulas;
    return values.map((row) =>
      row.map((cell) =>
        typeof cell === "string" && cell.startsWith("=") ? cell : "",
      ),
    );
  },
  getRowIndex() {
    return rowIndex;
  },
  getRowCount() {
    return rowCount;
  },
  getIntersection() {
    return intersection;
  },
});

export const createRecordingSheet = () => {
  let calls = [];
  let toColumnName = (columnIndex) => {
    let index = columnIndex + 1;
    let name = "";
    while (index > 0) {
      let mod = (index - 1) % 26;
      name = String.fromCharCode(65 + mod) + name;
      index = Math.floor((index - 1) / 26);
    }
    return name;
  };
  let toA1Address = (rowIndex, columnIndex, rowCount, columnCount) => {
    let startRow = rowIndex + 1;
    let endRow = rowIndex + rowCount;
    let startCol = toColumnName(columnIndex);
    let endCol = toColumnName(columnIndex + columnCount - 1);
    return `${startCol}${startRow}:${endCol}${endRow}`;
  };
  let createRangeRecorder = (address) => {
    let call = {
      address,
      values: null,
      numberFormat: null,
      numberFormatLocal: null,
    };
    calls.push(call);
    return {
      setValues(values) {
        call.values = values;
      },
      setNumberFormat(format) {
        call.numberFormat = format;
      },
      setNumberFormatLocal(format) {
        call.numberFormatLocal = format;
      },
    };
  };

  return {
    calls,
    sheet: {
      getRange(address) {
        return createRangeRecorder(address);
      },
      getRangeByIndexes(rowIndex, columnIndex, rowCount, columnCount) {
        return createRangeRecorder(
          toA1Address(rowIndex, columnIndex, rowCount, columnCount),
        );
      },
    },
  };
};

export const createMappedSheet = ({ usedRange = null, ranges = {} } = {}) => ({
  getUsedRange() {
    return typeof usedRange === "function" ? usedRange() : usedRange;
  },
  getRange(address) {
    let range =
      typeof ranges === "function"
        ? ranges(address)
        : ranges[address];

    if (!range) {
      throw new Error(`Unexpected address: ${address}`);
    }

    return typeof range === "function" ? range(address) : range;
  },
});

export const createWorkbookWithSheets = (sheets = {}) => ({
  getWorksheet(name) {
    return sheets[name];
  },
});

export const createMockExcelScriptEnums = () => ({
  HorizontalAlignment: {
    center: "center",
    left: "left",
    right: "right",
  },
  VerticalAlignment: {
    top: "top",
    center: "center",
    bottom: "bottom",
  },
  SheetVisibility: {
    visible: "visible",
    hidden: "hidden",
    veryHidden: "veryHidden",
  },
  ClearApplyTo: {
    all: "all",
    contents: "contents",
    formats: "formats",
  },
});
