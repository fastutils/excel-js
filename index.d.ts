declare namespace Excel {
    export class ExcelError extends Error {

    }

    export class ExcelValueFormatError extends ExcelError {
        readonly rowIndex: number;
        readonly columnIndex: number;
    }

    export class ExcelMultiValueFormatError extends ExcelError {
        readonly errors: ExcelValueFormatError[];
    }

    function transform<T>(sheet : Sheet, parser : (row: Row) => T) : T[];
    function transform<T>(sheet : Sheet, startIndex : number, parser : (row: Row) => T) : T[];
    function transform<T>(sheet : Sheet, startIndex : number, count : number, parser : (row: Row) => T) : T[];
    function transform<T>(sheetIndex : number, excel : string | ArrayBuffer | Excel, parser : (row: Row) => T) : T[];
    function transform<T>(sheetIndex : number, excel : string | ArrayBuffer | Excel, startIndex : number, parser : (row: Row) => T) : T[];
    function transform<T>(sheetIndex : number, excel : string | ArrayBuffer | Excel, startIndex : number, count : number, parser : (row: Row) => T) : T[];
    function transform<T>(excel : string | ArrayBuffer | Excel, parser : (row: Row) => T) : T[];
    function transform<T>(excel : string | ArrayBuffer | Excel, startIndex : number, parser : (row: Row) => T) : T[];
    function transform<T>(excel : string | ArrayBuffer | Excel, startIndex : number, count : number, parser : (row: Row) => T) : T[];
    function transform<T>(sheet : Sheet, creator : () => T, ...setters : ((T, Cell) => void)[]) : T[];
    function transform<T>(sheet : Sheet, startIndex : number, creator : () => T, ...setters : ((T, Cell) => void)[]) : T[];
    function transform<T>(sheet : Sheet, startIndex : number, count : number, creator : () => T, ...setters : ((T, Cell) => void)[]) : T[];
    function transform<T>(sheetIndex : number, excel : string | ArrayBuffer | Excel, creator : () => T, ...setters : ((T, Cell) => void)[]) : T[];
    function transform<T>(sheetIndex : number, excel : string | ArrayBuffer | Excel, startIndex : number, creator : () => T, ...setters : ((T, Cell) => void)[]) : T[];
    function transform<T>(sheetIndex : number, excel : string | ArrayBuffer | Excel, startIndex : number, count : number, creator : () => T, ...setters : ((T, Cell) => void)[]) : T[];
    function transform<T>(excel : string | ArrayBuffer | Excel, creator : () => T, ...setters : ((T, Cell) => void)[]) : T[];
    function transform<T>(excel : string | ArrayBuffer | Excel, startIndex : number, creator : () => T, ...setters : ((T, Cell) => void)[]) : T[];
    function transform<T>(excel : string | ArrayBuffer | Excel, startIndex : number, count : number, creator : () => T, ...setters : ((T, Cell) => void)[]) : T[];

    function map<T>(sheet : Sheet, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function map<T>(sheet : Sheet, startIndex : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function map<T>(sheet : Sheet, startIndex : number, count : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function map<T>(sheetIndex : number, excel : string | Excel, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function map<T>(sheetIndex : number, excel : string | Excel, startIndex : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function map<T>(sheetIndex : number, excel : string | Excel, startIndex : number, count : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function map<T>(excel : string | ArrayBuffer | Excel, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function map<T>(excel : string | ArrayBuffer | Excel, startIndex : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function map<T>(excel : string | ArrayBuffer | Excel, startIndex : number, count : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];

    function mapCollectErrors<T>(sheet : Sheet, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapCollectErrors<T>(sheet : Sheet, startIndex : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapCollectErrors<T>(sheet : Sheet, startIndex : number, count : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapCollectErrors<T>(sheetIndex : number, excel : string | ArrayBuffer | Excel, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapCollectErrors<T>(sheetIndex : number, excel : string | ArrayBuffer | Excel, startIndex : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapCollectErrors<T>(sheetIndex : number, excel : string | ArrayBuffer | Excel, startIndex : number, count : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapCollectErrors<T>(excel : string | ArrayBuffer | Excel, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapCollectErrors<T>(excel : string | ArrayBuffer | Excel, startIndex : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapCollectErrors<T>(excel : string | ArrayBuffer | Excel, startIndex : number, count : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];

    function mapIgnoreErrors<T>(sheet : Sheet, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapIgnoreErrors<T>(sheet : Sheet, startIndex : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapIgnoreErrors<T>(sheet : Sheet, startIndex : number, count : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapIgnoreErrors<T>(sheetIndex : number, excel : string | ArrayBuffer | Excel, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapIgnoreErrors<T>(sheetIndex : number, excel : string | ArrayBuffer | Excel, startIndex : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapIgnoreErrors<T>(sheetIndex : number, excel : string | ArrayBuffer | Excel, startIndex : number, count : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapIgnoreErrors<T>(excel : string | ArrayBuffer | Excel, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapIgnoreErrors<T>(excel : string | ArrayBuffer | Excel, startIndex : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapIgnoreErrors<T>(excel : string | ArrayBuffer | Excel, startIndex : number, count : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapIgnoreErrors<T>(collectors: ExcelValueFormatError[], sheet : Sheet, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapIgnoreErrors<T>(collectors: ExcelValueFormatError[], sheet : Sheet, startIndex : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapIgnoreErrors<T>(collectors: ExcelValueFormatError[], sheet : Sheet, startIndex : number, count : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapIgnoreErrors<T>(collectors: ExcelValueFormatError[], sheetIndex : number, excel : string | ArrayBuffer | Excel, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapIgnoreErrors<T>(collectors: ExcelValueFormatError[], sheetIndex : number, excel : string | ArrayBuffer | Excel, startIndex : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapIgnoreErrors<T>(collectors: ExcelValueFormatError[], sheetIndex : number, excel : string | ArrayBuffer | Excel, startIndex : number, count : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapIgnoreErrors<T>(collectors: ExcelValueFormatError[], excel : string | ArrayBuffer | Excel, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapIgnoreErrors<T>(collectors: ExcelValueFormatError[], excel : string | ArrayBuffer | Excel, startIndex : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];
    function mapIgnoreErrors<T>(collectors: ExcelValueFormatError[], excel : string | ArrayBuffer | Excel, startIndex : number, count : number, creator : () => T, ...properties : PropertyInfo<T, any>[]) : T[];

    function forDate<T, P>(column : number, setter : (obj : T, value : Date, rowIndex : number, columnIndex : number) => void) : PropertyInfo<T, P>;
    function forDate<T, P>(column : number, setter : (obj : T, value : Date, rowIndex : number, columnIndex : number) => void, required : boolean) : PropertyInfo<T, P>;
    function forDate<T, P>(column : number, setter : (obj : T, value : Date, rowIndex : number, columnIndex : number) => void, validator : (obj : T, value : Date, rowIndex : number, columnIndex : number) => Error) : PropertyInfo<T, P>;
    function forDate<T, P>(minColumn : number, maxColumn : number, setter : (obj : T, value : Date, rowIndex : number, columnIndex : number) => void) : PropertyInfo<T, P>;
    function forDate<T, P>(minColumn : number, maxColumn : number, setter : (obj : T, value : Date, rowIndex : number, columnIndex : number) => void, required : boolean) : PropertyInfo<T, P>;
    function forDate<T, P>(minColumn : number, maxColumn : number, setter : (obj : T, value : Date, rowIndex : number, columnIndex : number) => void, validator : (obj : T, value : Date, rowIndex : number, columnIndex : number) => Error) : PropertyInfo<T, P>;
    function forDate<T, P>(column : number, creator : (obj : T) => P, reader : (obj : T, property : P, value : Date, rowIndex : number, columnIndex : number) => void, setter : (obj : T, property : P) => void) : PropertyInfo<T, P>;
    function forDate<T, P>(column : number, creator : (obj : T) => P, reader : (obj : T, property : P, value : Date, rowIndex : number, columnIndex : number) => void, setter : (obj : T, property : P) => void, required : boolean) : PropertyInfo<T, P>;
    function forDate<T, P>(column : number, creator : (obj : T) => P, reader : (obj : T, property : P, value : Date, rowIndex : number, columnIndex : number) => void, setter : (obj : T, property : P) => void, validator : (obj : T, property : P, value : Date, rowIndex : number, columnIndex : number) => Error) : PropertyInfo<T, P>;
    function forDate<T, P>(minColumn : number, maxColumn : number, creator : (obj : T) => P, reader : (obj : T, property : P, value : Date, rowIndex : number, columnIndex : number) => void, setter : (obj : T, property : P) => void) : PropertyInfo<T, P>;
    function forDate<T, P>(minColumn : number, maxColumn : number, creator : (obj : T) => P, reader : (obj : T, property : P, value : Date, rowIndex : number, columnIndex : number) => void, setter : (obj : T, property : P) => void, required : boolean) : PropertyInfo<T, P>;
    function forDate<T, P>(minColumn : number, maxColumn : number, creator : (obj : T) => P, reader : (obj : T, property : P, value : Date, rowIndex : number, columnIndex : number) => void, setter : (obj : T, property : P) => void, validator : (obj : T, property : P, value : Date, rowIndex : number, columnIndex : number) => Error) : PropertyInfo<T, P>;

    function forString<T, P>(column : number, setter : (obj : T, value : string, rowIndex : number, columnIndex : number) => void) : PropertyInfo<T, P>;
    function forString<T, P>(column : number, setter : (obj : T, value : string, rowIndex : number, columnIndex : number) => void, required : boolean) : PropertyInfo<T, P>;
    function forString<T, P>(column : number, setter : (obj : T, value : string, rowIndex : number, columnIndex : number) => void, validator : (obj : T, value : string, rowIndex : number, columnIndex : number) => Error) : PropertyInfo<T, P>;
    function forString<T, P>(minColumn : number, maxColumn : number, setter : (obj : T, value : string, rowIndex : number, columnIndex : number) => void) : PropertyInfo<T, P>;
    function forString<T, P>(minColumn : number, maxColumn : number, setter : (obj : T, value : string, rowIndex : number, columnIndex : number) => void, required : boolean) : PropertyInfo<T, P>;
    function forString<T, P>(minColumn : number, maxColumn : number, setter : (obj : T, value : string, rowIndex : number, columnIndex : number) => void, validator : (obj : T, value : string, rowIndex : number, columnIndex : number) => Error) : PropertyInfo<T, P>;
    function forString<T, P>(column : number, creator : (obj : T) => P, reader : (obj : T, property : P, value : string, rowIndex : number, columnIndex : number) => void, setter : (obj : T, property : P) => void) : PropertyInfo<T, P>;
    function forString<T, P>(column : number, creator : (obj : T) => P, reader : (obj : T, property : P, value : string, rowIndex : number, columnIndex : number) => void, setter : (obj : T, property : P) => void, required : boolean) : PropertyInfo<T, P>;
    function forString<T, P>(column : number, creator : (obj : T) => P, reader : (obj : T, property : P, value : string, rowIndex : number, columnIndex : number) => void, setter : (obj : T, property : P) => void, validator : (obj : T, property : P, value : string, rowIndex : number, columnIndex : number) => Error) : PropertyInfo<T, P>;
    function forString<T, P>(minColumn : number, maxColumn : number, creator : (obj : T) => P, reader : (obj : T, property : P, value : string, rowIndex : number, columnIndex : number) => void, setter : (obj : T, property : P) => void) : PropertyInfo<T, P>;
    function forString<T, P>(minColumn : number, maxColumn : number, creator : (obj : T) => P, reader : (obj : T, property : P, value : string, rowIndex : number, columnIndex : number) => void, setter : (obj : T, property : P) => void, required : boolean) : PropertyInfo<T, P>;
    function forString<T, P>(minColumn : number, maxColumn : number, creator : (obj : T) => P, reader : (obj : T, property : P, value : string, rowIndex : number, columnIndex : number) => void, setter : (obj : T, property : P) => void, validator : (obj : T, property : P, value : string, rowIndex : number, columnIndex : number) => Error) : PropertyInfo<T, P>;

    function forNumber<T, P>(column : number, setter : (obj : T, value : number, rowIndex : number, columnIndex : number) => void) : PropertyInfo<T, P>;
    function forNumber<T, P>(column : number, setter : (obj : T, value : number, rowIndex : number, columnIndex : number) => void, required : boolean) : PropertyInfo<T, P>;
    function forNumber<T, P>(column : number, setter : (obj : T, value : number, rowIndex : number, columnIndex : number) => void, validator : (obj : T, value : number, rowIndex : number, columnIndex : number) => Error) : PropertyInfo<T, P>;
    function forNumber<T, P>(minColumn : number, maxColumn : number, setter : (obj : T, value : number, rowIndex : number, columnIndex : number) => void) : PropertyInfo<T, P>;
    function forNumber<T, P>(minColumn : number, maxColumn : number, setter : (obj : T, value : number, rowIndex : number, columnIndex : number) => void, required : boolean) : PropertyInfo<T, P>;
    function forNumber<T, P>(minColumn : number, maxColumn : number, setter : (obj : T, value : number, rowIndex : number, columnIndex : number) => void, validator : (obj : T, value : number, rowIndex : number, columnIndex : number) => Error) : PropertyInfo<T, P>;
    function forNumber<T, P>(column : number, creator : (obj : T) => P, reader : (obj : T, property : P, value : number, rowIndex : number, columnIndex : number) => void, setter : (obj : T, property : P) => void) : PropertyInfo<T, P>;
    function forNumber<T, P>(column : number, creator : (obj : T) => P, reader : (obj : T, property : P, value : number, rowIndex : number, columnIndex : number) => void, setter : (obj : T, property : P) => void, required : boolean) : PropertyInfo<T, P>;
    function forNumber<T, P>(column : number, creator : (obj : T) => P, reader : (obj : T, property : P, value : number, rowIndex : number, columnIndex : number) => void, setter : (obj : T, property : P) => void, validator : (obj : T, property : P, value : number, rowIndex : number, columnIndex : number) => Error) : PropertyInfo<T, P>;
    function forNumber<T, P>(minColumn : number, maxColumn : number, creator : (obj : T) => P, reader : (obj : T, property : P, value : number, rowIndex : number, columnIndex : number) => void, setter : (obj : T, property : P) => void) : PropertyInfo<T, P>;
    function forNumber<T, P>(minColumn : number, maxColumn : number, creator : (obj : T) => P, reader : (obj : T, property : P, value : number, rowIndex : number, columnIndex : number) => void, setter : (obj : T, property : P) => void, required : boolean) : PropertyInfo<T, P>;
    function forNumber<T, P>(minColumn : number, maxColumn : number, creator : (obj : T) => P, reader : (obj : T, property : P, value : number, rowIndex : number, columnIndex : number) => void, setter : (obj : T, property : P) => void, validator : (obj : T, property : P, value : number, rowIndex : number, columnIndex : number) => Error) : PropertyInfo<T, P>;

    class PropertyInfo<T, P> {
        set(t : T, row : Row, rowIndex : number) : void
    }

    class Excel {
        getSheet(sheetIndex : number) : Sheet;
    }

    class Sheet {
        readonly rows : number;
        readonly columns : number;

        getRow(rowIndex) : Row;
        getColumn(columnIndex) : Column;
        getCell(rowIndex, columnIndex) : Cell;
    }

    class Cell {
        readonly value : string;
        readonly row : number;
        readonly column : number;
    }

    class Row {
        readonly cells : Cell[];
        readonly index : number;
        readonly sheet : Sheet;
        readonly columns : number;

        getCell(columnIndex : number) : Cell;
    }

    class Column {
        readonly cells : Cell[];
        readonly index : number;
        readonly sheet : Sheet;
        readonly rows : number;

        getCell(rowIndex : number) : Cell;
    }
}

export = Excel;
export as namespace Excel;
