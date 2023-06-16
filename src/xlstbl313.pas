(*
Copyright (C) 2018, Sridharan S

This file is part of xltbl (Table interface for Excel Worksheet).

xlstbl is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

xltbl is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
 MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License version 3
 along with Table interface for Excel Worksheet  If not, see <http://www.gnu.org/licenses/>.
*)
unit xlstbl313;
{$DEFINE SM }
{$IFNDEF SM }
{.$DEFINE LIBXL }
{$DEFINE NativeExcel }
{$ENDIF SM }

interface
uses
  SysUtils, Classes,
  Windows,
{$IFDEF SM }
  XLSFile,
  XLSWorkbook,
{$ENDIF SM }
{$IFDEF LibXL }
  LibXL,
{$ENDIF }
{$IFDEF NativeExcel }
  nExcel,
{$ENDIF NativeExcel }
  bjXml3_1,
  Variants,
  StrUtils,
  VchLib,
  Dialogs;

{$IFDEF LibXL }
type
   TColumn = TFilterColumn;
{$ENDIF }

type
ValidFormat = (vfGeneral, vfInteger, vfDate, vfText);

  TbjXLSTable = class(TInterfacedObject)
  private
    FO_Row: integer;
    FO_Column: integer;
    FxLSFileName: string;
    FSheetName: string;
    FToSaveFile: boolean;
    FOWner: boolean;
    FPageLen: integer;
    FLastRow:  integer;
  protected
    function IsEmpty(aRow: integer = -1): boolean;
    function GetFieldCol(const aName: string): integer;
    function GetLastRow: integer;
  public
  { Public declarations }
    IDate: integer;
    BOF: boolean;
    EOF: boolean;
{$IFDEF SM }
    XL: TXLSFile;
    Workbook: TWorkbook;
    WorkSheet: TSheet;
{$ENDIF SM }
{$IFDEF LIBXL }
    Workbook: TXLbOOK;
    WorkSheet: TXLSheet;
{$ENDIF LIBXL }
{$IFDEF NativeExcel }
    Workbook: IXLSWorkbook;
    WorkSheet: IXLSWorksheet;
{$ENDIF NativeExcel }
    CurrentColumn: integer;
    CurrentRow: integer;
    SkipCount: integer;
    FieldList: TStringList;
    ColumnList: array of integer;
    IDList: TStringList;
    ListColList: TStringList;
    procedure SetXLSFile(const aName: string);
    procedure NewFile(const aName: string);
    procedure SetSheet(const aSheet: string);
    procedure Close;
    procedure SetOrigin(const aRow, aColumn: integer);
    procedure SetOColumn(const aColumn: integer);
    function GetRecVal(const aCol: Integer; aRow: integer = -1): variant;
    function GetRecFloat(const aCol: Integer; aRow: integer = -1): Double;
    function GetRecString(const aCol: Integer; aRow: integer = -1): string;
    function GetFieldVal(const aField: string; aRow: integer = -1): variant;
    function GetFieldCurr(const aField: string): currency;
    function GetFieldFloat(const aField: string; aRow: integer = -1): Double;
    function GetFieldString(const aField: string; aRow: integer = -1): string;
    function GetFieldSDate(const aField: string): string;
    function GetFieldToken(const aField: string): string;
    function GetFieldName(const aName: string): string;
{$IFDEF SM }
    function GetFieldObj(const aField: string): TColumn;
    function GetCellObj(const arow: integer; const aField: string): TCell; overload;
    function GetCellObj(const arow: integer; const aCol: integer): TCell; overload;
    procedure SetFieldFormat(const aField: string; const aFOrmat: Integer);
    procedure SetFormatAt(const aCol: integer; aFOrmat: Integer);
    procedure SetCellFormat(const aField: string; aFOrmat: Integer);
{$ENDIF SM }
{$IFDEF NativeExcel }
    function GetFieldObj(const aField: string): IXLSRange;
    function GetCellObj(const arow: integer; const aField: string): IXLSRange; overload;
    function GetCellObj(const arow: integer; const aCol: integer): IXLSRange; overload;
{$ENDIF NativeExcel }
    procedure SetFieldVal(const aField: string; const aValue: variant);
    procedure SetFieldStr(const aField: string; const aValue: string);
    procedure SetFieldWStr(const aField: string; const aValue: string);
    procedure SetFieldNum(const aField: string; const aValue: double);
    function FindField(const aName: string): pChar;
//    function GetFieldCol(const aName: string): integer;
    procedure SetFields(const aList: TStrings; const ToWrite: boolean);
    function GetFields(const aList: TStrings): Tstrings;
    procedure ParseXml(const aNode: IbjXml; const FldLst: TStringList);
    function IsEmptyField(const aField: string; aRow: integer = -1): boolean;
    procedure Insert;
    procedure Delete;
    procedure ClearRow;
    procedure Next;
    procedure Prior;
    procedure First;
    procedure Last;
    procedure Save;
    procedure SaveAs(const aName: string);
    procedure AtSay(const acol: Integer; const aMsg: Variant);
    constructor Create;
    destructor Destroy; override;

    property O_Row: integer read FO_row;
    property O_Column: integer read FO_Column write SetOColumn;
    property XLSFileName: string write SetXLSFile;
    property ToSaveFile: boolean read FToSaveFile write FToSaveFile;
    property SheetName: string write SetSheet;
    property Owner: boolean read Fowner write Fowner;
    property PageLen: integer read FPageLen write FPagelen;
    property LastRow: integer read GetLastRow;
  end;

implementation


constructor TbjXLSTable.Create;
begin
  inherited;
{
  FO_Row  := 0;
  FO_Column  := 0;
}
{$IFDEF NativeExcel }
  FO_Row  := 1;
  FO_Column  := 1;
{$ENDIF NativeExcel }
  FToSaveFile := False;
  FieldList := THashedStringList.Create;
  IDList := THashedStringList.Create;
  ListColList := TStringList.Create;
  SkipCount := 1;
  FLastRow := -1;
end;

destructor TbjXLSTable.Destroy;
begin
  XLSFileName := '';
  FieldList.Clear;
  FieldList.Free;
  if Assigned(IDList) then
  IDList.Free;
  if Assigned(ListColList) then
  ListColList.Free;
Close;
  Inherited;
end;

procedure TbjXLSTable.Save;
begin
if Length(FXlsFileName) > 0 then
{$IFDEF SM }
    XL.SaveAs(FXlsFileName);
{$ENDIF SM }
{$IFDEF LIBXL }
  Workbook.Save(PChar(FXlsFileName));
{$ENDIF LIBXL}
{$IFDEF NativeExcel }
    Workbook.SaveAs(FXlsFileName);
{$ENDIF NativeExcel }
end;

procedure TbjXLSTable.SaveAs(const aName  : string);
begin
{$IFDEF SM }
  XL.SaveAs(AName);
{$ENDIF SM }
{$IFDEF LIBXL }
  Workbook.Save(PChar(AName));
{$ENDIF LIBXL}
{$IFDEF NativeExcel }
  Workbook.SaveAs(aName);
{$ENDIF NativeExcel }
end;

procedure TbjXLSTable.SetXLSFile(const aName: string);
begin
  Close;
{$IFDEF SM }
  FXLSFileName := aName;
  XL := TXLSFile.Create;
  FOwner := True;

  if FileExists(FXLSFileName) then
    XL.OpenFile(FXLSFileName);
  Workbook := XL.Workbook;
{$ENDIF SM }
{$IFDEF LIBXL }
  FXLSFileName := aName;
  FOwner := True;
  Workbook := TBinBook.Create;
  Workbook.setKey('Name', 'Key');
  if FileExists(aName) then
    Workbook.load(pChar(aName));
{$ENDIF LIBXL}
{$IFDEF NativeExcel }
  FXLSFileName := aName;
  FOwner := True;
  Workbook := TXLSWorkbook.Create;
  if FileExists(aName) then
    Workbook.Open(aName);
{$ENDIF NativeExcel }
end;

procedure TbjXLSTable.NewFile(const aName  : string);
begin
  SetXlsFile('');
  FXLSFileName := aName;
end;

procedure TbjXLSTable.close;
begin
  if Owner then
{$IFDEF SM }
  if Assigned(XL) then
    Xl.Free;
{$ENDIF SM}
{$IFDEF LIBXL }
  if Assigned(Workbook) then
  Workbook.Free;
{$ENDIF LIBXL}
{$IFDEF NativeExcel }
  Workbook := nil;
{$ENDIF NativeExcel }
end;

procedure TbjXLSTable.SetSheet(const aSheet: string);
var
  i: Integer;
begin
  if not Assigned(Workbook) then
    Exit;
  if FSheetName = aSheet then
    Exit;
{$IFDEF SM }
  WorkSheet := Workbook.SheetByName(aSheet);
  if not Assigned(WorkSheet) then
  begin
    Workbook.Sheets.Add(aSheet);
    WorkSheet := Workbook.SheetByName(aSheet);
  end;
{$ENDIF SM}
{$IFDEF LIBXL }
  WorkSheet := Workbook.GetSheetbyName(PChar(aSheet));
  if not Assigned(WorkSheet) then
  begin
    Workbook.addSheet(PChar(aSheet));
    WorkSheet := Workbook.GetSheetByName(PChar(aSheet));
  end;
{$ENDIF LIBXL}
{$IFDEF NativeExcel }
  for i := 1 to Workbook.Sheets.Count do
  begin
    if Workbook.Sheets.Entries[i].Name = aSheet then
    begin
      WorkSheet := Workbook.Sheets.Entries[i];
      Exit;
    end;
  end;
  WorkSheet := nil;
  WorkSheet := Workbook.Sheets.Add;
  Workbook.Sheets[1].Name := aSheet;
{$ENDIF NativeExcel }
end;

procedure TbjXLSTable.SetOrigin(const aRow, aColumn: integer);
begin
  FO_Row := aRow;
  FO_Column :=aColumn;
end;

procedure TbjXLSTable.SetOColumn(const aColumn: integer);
begin
  FO_Column :=aColumn;
end;
function TbjXLSTable.GetRecVal(const aCol: Integer; aRow: integer = -1): variant;
{$IFDEF LIBXL }
var
  wCellType: CellType;
  wDateFormat: TFormat;
{$ENDIF LIBXL}
begin
  if aRow = -1 then
    aRow := Currentrow;
{$IFDEF SM }
  Result := WorkSheet.Cells[FO_row + aRow, FO_Column + aCol].Value;
{$ENDIF SM}
{$IFDEF LIBXL }
  wCellType := WorkSheet.GetCellType(FO_row + aRow, FO_Column + aCol);
  if (wCellType = CELLTYPE_EMPTY) or (wCellType = CELLTYPE_BLANK) then
    Exit;
  if WorkSheet.isDate(FO_row + aRow, FO_Column + aCol) then
  begin
    wDateFormat := Workbook.addFormat();
    wDateFormat.setNumFormat(NUMFORMAT_DATE);
    Result := FormatDateTime('YYYYMMDD', WorkSheet.readNum(FO_row + aRow, FO_Column + aCol, wDateformat));
    Exit;
  end;
  if wCellType = CELLTYPE_STRING then
    Result := WorkSheet.readStr(FO_row + aRow, FO_Column + aCol;
  if wCellType = CELLTYPE_NUMBER then
    Result := WorkSheet.readNum(FO_row + aRow, FO_Column + aCol);
{$ENDIF LIBXL}
{$IFDEF NativeExcel }
  Result := WorkSheet.Cells[FO_row + aRow, FO_Column + aCol].Value;
{$ENDIF NativeExcel }
end;
function TbjXLSTable.GetFieldVal(const aField: string; aRow: integer = -1): variant;
var
  aCol: Integer;
begin
  if FindField(aField) = nil then
  begin
    raise Exception.Create(aField + ' Column not defined');
    Exit;
  end;
  aCol := GetFieldCol(aField);
  Result := GetRecVal(aCol, aRow);
end;

function TbjXLSTable.GetFieldSDate(const aField: string): string;
var
  v_vt: variant;
  VType: Integer;
  mnth: integer;
begin
  v_vt := GetFieldVal(aField);
  if VarIsNull(v_vt) then
    Exit;
  VType := VarType(V_vt) and VarTypeMask;
  case VType of
  varDate:
    Result := FormatDateTime('YYYYMMDD', V_vt);
  varString:
  begin
     Result := V_vt;
    if Result[1] = '''' then
      Result := Copy(Result, 2, Length(Result)-1);
    end;
  end;
end;

function TbjXLSTable.GetFieldCurr(const aField: string): currency;
var
  v_vt: variant;
begin
  v_vt := GetFieldVal(aField);
  if VarIsNull(v_vt) then
    Exit;
  try
    Result := v_vt;
  except
    Result := 0;
  end;
end;

function TbjXLSTable.GetRecFloat(const aCol: Integer; aRow: integer = -1): Double;
var
  v_vt: variant;
begin
  v_vt := GetRecVal(aCol);
  if VarIsNull(v_vt) then
    Exit;
  try
    Result := v_vt;
  except
    Result := 0;
  end;
end;
function TbjXLSTable.GetFieldFloat(const aField: string; aRow: integer = -1): double;
var
  v_vt: variant;
  str: string;
  iValue: double;
  iCode: Integer;
begin
  v_vt := GetFieldVal(aField, aRow);
  if VarIsNull(v_vt) then
    Exit;
  try
    Result := v_vt;
  except
    Result := 0;
  end;
end;

function TbjXLSTable.GetFieldToken(const aField: string): string;
const
  formatChars: array[0..6] of string = ('%', '.00', '.0', '.', ',', '-', '''');
var
  wStr: WideString;
  rCellStr: String;
  ctr: integer;
begin
  Result := '';
{$IFDEF SM }
  wStr := WorkSheet.Cells[FO_Row + CurrentRow,
    FO_Column + GetFieldCol(aField)].ValueAsString;
  rCellStr := wStr;
{$ENDIF SM }
{$IFDEF LIBXL }
  if WorkSheet.GetCellType(FO_row + CurrentRow, FO_Column + GetFieldCol(aField)) = CELLTYPE_NUMBER then
    rCellStr := FloattoStr(WorkSheet.ReadNum(FO_Row + CurrentRow,
    FO_Column + GetFieldCol(aField)));
  if WorkSheet.GetCellType(FO_row + CurrentRow, FO_Column + GetFieldCol(aField)) = CELLTYPE_STRING then
    rCellStr := WorkSheet.readStr(FO_Row + CurrentRow,
    FO_Column + GetFieldCol(aField));
{$ENDIF LIBXL }
{$IFDEF NativeExcel }
  try
  wstr := WorkSheet.Cells[FO_Row + CurrentRow,
    FO_Column + GetFieldCol(aField)].Value;
  except
    wStr := '';
  end;
  rCellStr := wStr;
{$ENDIF NativeExcel }
  rCellStr := Trim(rCellStr);;
  for ctr := low(formatChars) to high(formatChars) do
    if Pos(formatChars[ctr], rCellStr) > 0 then
      rCellStr := stringReplace(rCellStr, formatChars[ctr], '', [rfIgnoreCase, rfReplaceAll]);
  Result := rCellStr;
end;

function TbjXLSTable.GetFieldString(const aField: string; aRow: integer = -1): string;
var
  aCol: Integer;
begin
  if FindField(AField) = nil then
  begin
    raise Exception.Create(aField + ' Column not defined');
    Exit;
  end;
  if IsEmptyField(aField) then
      Exit;
  aCol := GetFieldCol(aField);
  Result := GetRecString(aCol, aRow);
end;
function TbjXLSTable.GetRecString(const aCol: Integer; aRow: integer = -1): string;
var
  wStr: WideString;
begin
  if aRow = -1 then
    aRow := Currentrow;
{$IFDEF SM }
  wStr := UTF8Encode(WorkSheet.Cells[FO_Row + aRow,
    FO_Column + aCol].ValueAsString);
  Result := wStr;
{$ENDIF SM }
{$IFDEF LIBXL }
  if WorkSheet.GetCellType(FO_row + aRow, FO_Column + aCol) = CELLTYPE_NUMBER then
  begin
    Result := FloattoStr(WorkSheet.ReadNum(FO_Row + aRow,
    FO_Column + aCol));
  end;
  if WorkSheet.GetCellType(FO_row + aRow, FO_Column + aCol) = CELLTYPE_STRING then
  begin
    Result := WorkSheet.readStr(FO_Row + aRow,
    FO_Column + aCol);
  end;
{$ENDIF LIBXL }
{$IFDEF NativeExcel }
  try
  wStr := WorkSheet.Cells[FO_Row + aRow,
    FO_Column + aCol].Value;
  except
    wStr := '';
  end;
  Result := UTF8Encode(wStr);
{$ENDIF NativeExcel }
  Result := Trim(Result);
end;

{$IFDEF SM }
function TbjXLSTable.GetFieldObj(const aField: string): TColumn;
{$ENDIF SM }
{$IFDEF NativeExcel }
function TbjXLSTable.GetFieldObj(const aField: string): IXLSRange;
{$ENDIF NativeExcel }
var
  ctr: integer;
begin
  ctr := GetFieldCol(aField);
{$IFDEF SM }
  Result := WorkSheet.Columns[FO_Column + ctr];
{$ENDIF SM }
{.$IFDEF LIBXL }
{.$ENDIF LIBXL }
{$IFDEF NativeExcel }
  Result := WorkSheet.Selection.EntireColumn.Item[FO_Column + ctr]; 
{$ENDIF NativeExcel }
end;

{$IFDEF SM }
procedure TbjXLSTable.SetFieldFormat(const aField: string; const aFOrmat: Integer);
var
  ctr: integer;
  MyColumn: TColumn;
begin
  ctr := GetFieldCol(aField);
  if ctr = -1 then
    Exit;
  MyColumn:= GetFieldObj(AField);
    MyColumn.FormatStringIndex := aFOrmat;
end;

procedure TbjXLSTable.SetFormatAt(const aCol: integer; aFOrmat: Integer);
var
  MyCell: TCell;
begin
  MyCell:= WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + aCol];
  MyCell.FormatStringIndex := aFOrmat;
end;

procedure TbjXLSTable.SetCellFormat(const aField: string; aFOrmat: Integer);
var
  ctr: integer;
{$IFDEF SM }
  MyCell: TCell;
{$ENDIF SM }
{$IFDEF NativeExcel }
  MyCell: TCell;
{$ENDIF SM }
begin
  ctr := GetFieldCol(aField);
  if ctr = -1 then
    Exit;
  MyCell := GetCellObj(FO_Row + CurrentRow, aField);
{$IFDEF SM }
  MyCell.FormatStringIndex := aFOrmat;
{$ENDIF SM }
{$IFDEF NativeExcel }
{$ENDIF SM }
end;
{$ENDIF SM }

{$IFDEF SM }
function TbjXLSTable.GetCellObj(const arow: integer; const aField: string): TCell;
{$ENDIF SM }
{$IFDEF NativeExcel }
function TbjXLSTable.GetCellObj(const arow: integer; const aField: string): IXLSRange;
{$ENDIF NativeExcel }
var
  ctr: integer;
begin
  ctr := GetFieldCol(aField);
  Result := WorkSheet.Cells[FO_Row + aRow, FO_Column + ctr];
end;

{$IFDEF SM }
function TbjXLSTable.GetCellObj(const arow: integer; const aCol: integer): TCell;
{$ENDIF SM }
{$IFDEF NativeExcel }
function TbjXLSTable.GetCellObj(const arow: integer; const aCol: integer): IXLSRange;
{$ENDIF NativeExcel }
begin
  Result := WorkSheet.Cells[FO_Row + aRow, FO_Column + aCol];
end;

procedure TbjXLSTable.SetFieldVal(const aField: string; const aValue: variant);
{$IFDEF LIBXL }
var
  VType  : Integer;
{$ENDIF LIBXL }
begin
  if FindField(AField) = nil then
  begin
    raise Exception.Create(aField + ' Column not defined');
    Exit;
  end;
{$IFDEF SM }
  WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + GetFieldCol(aField)].Value := aValue;
{$ENDIF SM }
{$IFDEF LIBXL }
  VType := VarType(aValue) and VarTypeMask;
  case VType of
  varEmpty:
    Exit;
  varInteger:
    WorkSheet.WriteNum(FO_row + CurrentRow, FO_Column + GetFieldCol(aField), aValue);
  varDate:
    WorkSheet.WriteNum(FO_row + CurrentRow, FO_Column + GetFieldCol(aField), aValue);
  varDouble:
    WorkSheet.WriteNum(FO_row + CurrentRow, FO_Column + GetFieldCol(aField), aValue);
  varCurrency:
    WorkSheet.WriteNum(FO_row + CurrentRow, FO_Column + GetFieldCol(aField), aValue);
  varString:
    WorkSheet.WriteStr(FO_row + CurrentRow, FO_Column + GetFieldCol(aField), pChar(string(aValue)));
  end;
{$ENDIF LIBXL }
{$IFDEF NativeExcel }
  WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + GetFieldCol(aField)].Value := aValue;
{$ENDIF NativeExcel }
end;


procedure TbjXLSTable.SetFieldStr(const aField: string; const aValue: string);
{$IFDEF SM }
var
  ptr: pVarData;
{$ENDIF SM }
begin
  if FindField(AField) = nil then
  begin
    raise Exception.Create(aField + ' Column not defined');
    Exit;
  end;
{$IFDEF SM }
  ptr := FindVarData(WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + GetFieldCol(aField)].Value);
  ptr.VType := varString;
  ptr.VString := pChar(aValue);
{$ENDIF SM }
{$IFDEF LIBXL }
    WorkSheet.WriteStr(FO_row + CurrentRow, FO_Column + GetFieldCol(aField), pChar(string(aValue)));
{$ENDIF LIBXL }
{$IFDEF NativeExcel }
  WorkSheet.Cells[FO_row + CurrentRow, FO_Column + GetFieldCol(aField)].Value := aValue;
{$ENDIF NativeExcel }
end;
procedure TbjXLSTable.SetFieldWStr(const aField: string; const aValue: string);
{$IFDEF SM }
var
  ptr: pVarData;
{$ENDIF SM }
{$IFDEF NativeExcel }
var
  ptr: pVarData;
{$ENDIF NativeExcel }
begin
  if FindField(AField) = nil then
  begin
    raise Exception.Create(aField + ' Column not defined');
    Exit;
  end;
{$IFDEF SM }
  ptr := FindVarData(WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + GetFieldCol(aField)].Value);
  ptr.VType := varOleStr;
  ptr.VOleStr := PWideChar(UTF8Decode(aValue));
{$ENDIF SM }
{$IFDEF NativeExcel }
  ptr := FindVarData(WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + GetFieldCol(aField)].Value);
  ptr.VType := varOleStr;
  ptr.VOleStr := PWideChar(UTF8Decode(aValue));
{$ENDIF NativeExcel }
end;

procedure TbjXLSTable.SetFieldNum(const aField: string; const aValue: double);
{$IFDEF SM }
var
  ptr: pVarData;
{$ENDIF SM }
{$IFDEF NativeExcel }
var
  ptr: pVarData;
{$ENDIF NativeExcel }
begin
  if FindField(AField) = nil then
  begin
    raise Exception.Create(aField + ' Column not defined');
    Exit;
  end;
{$IFDEF SM }
  ptr := FindVarData(WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + GetFieldCol(aField)].Value);
  ptr.VType := VarDouble;
  ptr.VDouble := aValue;
{$ENDIF SM }
{$IFDEF LIBXL }
    WorkSheet.WriteNum(FO_row + CurrentRow, FO_Column + GetFieldCol(aField), aValue);
{$ENDIF LIBXL }
{$IFDEF NativeExcel }
  ptr := FindVarData(WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + GetFieldCol(aField)].Value);
  ptr.VType := VarDouble;
  ptr.VDouble := aValue;
{$ENDIF NativeExcel }
end;

procedure TbjXLSTable.AtSay(const acol: Integer; const aMsg: Variant);
{$IFDEF LIBXL }
var
  VType  : Integer;
{$ENDIF LIBXL }
begin
{$IFDEF SM }
  WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + aCol].Value := aMsg;
{$ENDIF SM }
{$IFDEF LIBXL }
  VType := VarType(VMsg) and VarTypeMask;
  // Set a string to match the type
  case VType of
  varEmpty:
    Exit;
  varInteger:
    WorkSheet.WriteNum(FO_row + CurrentRow, FO_Column + aCol, aMsg);
  varDate:
    WorkSheet.WriteNum(FO_row + CurrentRow, FO_Column + aCol, aMsg);
  varDouble:
    WorkSheet.WriteNum(FO_row + CurrentRow, FO_Column + aCol, aMsg);
  varCurrency:
    WorkSheet.WriteNum(FO_row + CurrentRow, FO_Column + aCol, aMsg);
  varString:
    WorkSheet.WriteStr(FO_row + CurrentRow, FO_Column + aCol, pChar(string(aMsg)));
  end;
{$ENDIF LIBXL }
{$IFDEF NativeExcel }
  WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + aCol].Value := aMsg;
{$ENDIF NativeExcel }
end;

procedure TbjXLSTable.Prior;
var
  wTempRow: integer;
begin
  if CurrentRow = 1 then
  begin
    BOF := True;
    Exit;
  end;
  wTempRow := CurrentRow;
  while wTempRow > 1 do
  begin
    if not IsEmpty(wTempRow-1) then
    begin
      CurrentRow := wTempRow-1;
  	  BOF := False;
      EOF := False;
      Exit;
    end;
//    if SkipCount > 0 then
    if (CurrentRow-wTempRow = SkipCount) then
      break;
    wTempRow := WTempRow - 1;
  end;
  BOF := True;
end;

procedure TbjXLSTable.Next;
var
  wtempRow: integer;
begin
  wTempRow := CurrentRow;
  while True do
  BEGIN
    if not IsEmpty(wTempRow+1) then
    begin
      CurrentRow := wTempRow+1;;
      Eof := False;
      if wTempRow >= LastRow then
        FLastRow := CurrentRow;
      Exit;
    end;
    if wTempRow-CurrentRow = SkipCount then
      break;
    wTempRow := WTempRow + 1;
  END;
  EOF := True;
end;


procedure TbjXLSTable.Insert;
begin
    Last;
  if not IsEmpty then
    CurrentRow := CurrentRow + 1;
end;

procedure TbjXLSTable.Delete;
var
  wTempRow: integer;
begin
{$IFDEF SM }
  WorkSheet.Rows.DeleteRows(FO_row + CurrentRow, FO_row + CurrentRow);
{$ENDIF SM }
{$IFDEF LIBXL }
  WorkSheet.removeRow(FO_row + CurrentRow, FO_row + CurrentRow+1, True);
{$ENDIF LIBXL }
{$IFDEF NativeExcel }
  WorkSheet.RCRange[1, FO_row + CurrentRow, 1, FO_row + CurrentRow].EntireRow.Delete(xlShiftUp);
{$ENDIF NativeExcel }

  if IsEmpty then
  begin
    if CurrentRow = 1 then
      BOF := True;
    wTempRow := CurrentRow;
    while wTempRow < LastRow do
    begin
      if not IsEmpty(wTempRow+1) then
      begin
        CurrentRow := wTempRow+1;
        BOF := False;
  	    EOF := False;
        Exit;
      end;
      wTempRow := WTempRow + 1;
    end;
    EOF := True;
	end;
end;

procedure TbjXLSTable.ClearRow;
begin
{$IFDEF SM }
    WorkSheet.Rows.ClearRows(FO_row + CurrentRow, FO_row + CurrentRow);
{$ENDIF SM }
{$IFDEF NativeExcel }
  WorkSheet.Selection.EntireRow.Item[FO_row + CurrentRow].Clear;
{$ENDIF NativeExcel }
end;

procedure TbjXLSTable.First;
begin
  CurrentRow := 1;
  BOF := False;
  EOF := False;
  if IsEmpty then
  begin
    BOF := True;
    EOF := True;
  end;
end;

procedure TbjXLSTable.Last;
begin
  EOF := False;
  BOF := False;
  CurrentRow := LastRow;
  if IsEmpty then
  begin
    BOF := True;
    EOF := True;
  end;
end;

function TbjXLSTable.GetLastRow: integer;
var
  wTempRow: integer;
begin
  wTempRow := FLastRow;
  if wTempRow = -1 then
  begin
  wTempRow := 1;
  while not IsEmpty(wTempRow) do
      wTempRow := wTempRow+1;;
  end;
    Result := wTempRow;
end;

function TbjXLSTable.GetFieldCol(const aName: string): integer;
var
  idx: integer;
  l_str: string;
begin
  Result := -1;
  l_str := PackStr(aName);
  idx := 0;
  if FieldList.Find(l_str, idx) then
  begin
    Result := ColumnList[idx];
  end;
end;
function TbjXLSTable.GetFieldName(const aName: string): string;
var
  idx: integer;
begin
  Result := '';
  idx := GetFieldCol(aName);
{$IFDEF SM }
  Result := WorkSheet.Cells[FO_Row, FO_Column+Idx].ValueAsString;
{$ENDIF SM }
{$IFDEF NativeExcel }
  Result := WorkSheet.Cells[FO_Row, FO_Column+Idx].Value;
{$ENDIF NativeExcel }
end;

function TbjXLSTable.FindField(const aName: string): pChar;
begin
  Result := nil;
  if GetFieldCol(aName) <> -1 then
    Result := pChar(aName);
end;

procedure TbjXLSTable.SetFields(const aList: TStrings; const ToWrite: boolean);
var
  ctr, j: integer;
  rCellStr: string;
begin
  if not Assigned(aList) then
    Exit;
  FieldList.Clear;
  setLength(ColumnList, aList.Count);
  for ctr := 0 to aList.Count-1 do
  begin
    FieldList.Add(PackStr(aList.Strings[ctr]));
    if ToWrite then
{$IFDEF SM }
      WorkSheet.Cells[FO_Row, FO_Column + ctr].Value := aList.Strings[ctr];
{$ENDIF SM }
{$IFDEF LIBXL }
      WorkSheet.WriteStr(FO_Row, FO_Column + ctr, pChar(aList.Strings[ctr]));
{$ENDIF LIBXL}
{$IFDEF NativeExcel }
      WorkSheet.Cells[FO_Row, FO_Column + ctr].Value := aList.Strings[ctr];
{$ENDIF NativeExcel }
  end;
  FieldList.Sorted := True;
  setLength(ColumnList, FieldList.Count);
  for ctr := 0 to aList.Count + 29 -1 do
  begin
{$IFDEF SM }
    rCellStr := WorkSheet.Cells[FO_Row, FO_Column+ctr].ValueAsString;
{$ENDIF SM }
{$IFDEF LIBXL }
    rCellStr := WorkSheet.readStr(FO_Row, FO_Column+ctr;
{$ENDIF LIBXL}
{$IFDEF NativeExcel }
    rCellStr := WorkSheet.Cells[FO_Row, FO_Column+ctr].Value;
{$ENDIF NativeExcel }
    if Length(rCellStr) = 0 then
      Break;
    for j := 0 to FieldList.Count-1 do
    if PackStr(RCellStr) =
      PackStr(FieldList.Strings[j]) then
    begin
        ColumnList[j] := ctr;
        break;
    end;
  end;
  CurrentRow := 1;
  CurrentColumn := 0;
end;

procedure TbjXLSTable.ParseXml(const aNode: IbjXml; const FldLst: TStringList);
var
  IDAlias, ListAlias: string;
  IDNode, IDAliasNode: IbjXml;
  aliasNode: IbjXml;
  ListNode:  IbjXml;
  k: Integer;
begin
  FldLst.Clear;
  aliasNode := aNode.SearchForTag(nil, 'Alias');
  while Assigned(aliasNode) do
  begin
    FldLst.Add(aliasNode.GetContent);
    aliasNode := aNode.SearchForTag(aliasNode, 'Alias');
  end;
  IDList.Clear;
  IDNode := aNode.SearchForTag(nil, 'KeyCol');
  while Assigned(IDNode) do
  begin
    IDAlias := IDNode.Content;
    IDAliasNode := aNode.SearchForTag(nil, IDAlias);
    if Assigned(IDAliasNode) then
      aliasNode := IDAliasNode.SearchForTag(nil, 'Alias');
    if Assigned(aliasNode) then
      IDList.Add(aliasNode.GetContent);
    IDNode := aNode.SearchForTag(IDnode, 'KeyCol');
  end;
  ListColList.Clear;
  ListNode := aNode.SearchForTag(nil, 'ListCol');
  while Assigned(ListNode) do
  begin
    ListAlias := ListNode.GetContent;
    ListColList.Add(ListAlias);
    for k := 1 to 9 do
    FldLst.Add(PackStr(ListAlias + IntToStr(k)));
    ListNode := aNode.SearchForTag(Listnode, 'ListCol');
  end;
end;

Function TbjXLSTable.GetFields(const aList: TStrings): TStrings;
var
  ctr, j: integer;
{$IFDEF SM }
  rCellStr: String;
{$ENDIF SM }
{$IFDEF NativeExcel }
  rCellStr: WideString;
{$ENDIF NativeExcel }
begin
  FieldList.Clear;
{$IFDEF SM }
  for ctr := 0 to aList.Count + 29 -1 do
  begin
    rCellStr := WorkSheet.Cells[FO_Row, FO_Column+ctr].Value;
    if Length(rCellStr) = 0 then
      Break;
    if Pos('Dr_', rCellStr) > 0 then
    begin
      FieldList.Add(PackStr(rCellStr));
      Continue;
    end;
    if Pos('Cr_', rCellStr) > 0 then
    begin
      FieldList.Add(PackStr(rCellStr));
      Continue;
    end;
    for j := 0 to aList.Count-1 do
    begin
      if PackStr(rCellStr) =
        PackStr(aList.Strings[j]) then
      begin
        FieldList.Add(PackStr(aList.Strings[j]));
        break;
      end;
    end;
  end;
  FieldList.Sorted := True;
  setLength(ColumnList, FieldList.Count);
  for ctr := 0 to aList.Count + 29 -1 do
  begin
    rCellStr := WorkSheet.Cells[FO_Row, FO_Column+ctr].Value;
    if Length(rCellStr) = 0 then
      Break;
    for j := 0 to FieldList.Count-1 do
    begin
      if PackStr(rCellStr) =
        PackStr(FieldList.Strings[j]) then
    begin
        ColumnList[j] := ctr;
        break;
      end;
    end;
  end;
{$ENDIF SM }
{$IFDEF LIBXL }
  for ctr := 0 to aList.Count + 29 -1 do
  begin
    if Length(WorkSheet.readStr(FO_Row, FO_Column+ctr)) = 0 then
      Break;
    if Pos('Dr_', WorkSheet.readStr(FO_Row, FO_Column+ctr)) > 0 then
    begin
      FieldList.Add(WorkSheet.readStr(FO_Row, FO_Column+ctr));
      Continue;
    end;
    if Pos('Cr_', WorkSheet.readStr(FO_Row, FO_Column+ctr)) > 0 then
    begin
      FieldList.Add(WorkSheet.readStr(FO_Row, FO_Column+ctr));
      Continue;
    end;
    for j := 0 to aList.Count-1 do
    begin
      if PackStr(WorkSheet.readStr(FO_Row, FO_Column+ctr)) =
        PackStr(aList.Strings[j]) then
      begin
        FieldList.Add(PackStr(aList.Strings[j]));
        break;
      end;
    end;
  end;
  FieldList.Sorted := True;
  setLength(ColumnList, FieldList.Count);
  for ctr := 0 to aList.Count + 29 -1 do
  begin
    if Length(WorkSheet.readStr(FO_Row, FO_Column+ctr)) = 0 then
      Break;
    for j := 0 to FieldList.Count-1 do
    begin
      if PackStr(WorkSheet.readStr(FO_Row, FO_Column+ctr)) =
        FieldList.Strings[j] then
    begin
        ColumnList[j] := ctr;
        break;
      end;
    end;
  end;
{$ENDIF LIBXL }
{$IFDEF NativeExcel }
  for ctr := 0 to aList.Count + 29-1 do
  begin
  try
    rCellStr := WorkSheet.Cells[FO_Row, FO_Column+ctr].Value;
    except
    rCellStr := '';
    end;
    if Length(rCellStr) = 0 then
      Break;
    if Pos('Dr_', rCellStr) > 0 then
    begin
      FieldList.Add(PackStr(rCellStr));
      Continue;
    end;
    if Pos('Cr_', rCellStr) > 0 then
    begin
      FieldList.Add(PackStr(rCellStr));
      Continue;
    end;
    for j := 0 to aList.Count-1 do
    begin
      if PackStr(rCellStr) =
        PackStr(aList.Strings[j]) then
      begin
        FieldList.Add(PackStr(aList.Strings[j]));
        break;
      end;
    end;
  end;
  FieldList.Sorted := True;
  setLength(ColumnList, FieldList.Count);
  for ctr := 0 to aList.Count + 29-1 do
  begin
    try
    rCellStr := WorkSheet.Cells[FO_Row, FO_Column+ctr].Value;
    except
    rCellStr := '';
    end;
    if Length(rCellStr) = 0 then
      Break;
    for j := 0 to FieldList.Count-1 do
    begin
      if PackStr(rCellStr) =
        PackStr(FieldList.Strings[j]) then
      begin
        ColumnList[j] := ctr;
        break;
      end;
    end;
  end;
{$ENDIF NativeExcel }
  Result := FieldList;
  if IDList.Count > 0 then
  begin
    for  ctr := 0 to FieldList.count-1 do
    begin
      for j := 0 to IDList.Count-1 do
      begin
        if FieldList.Strings[ctr] = PackStr(IDList.Strings[j]) then
        begin
          if ColumnList[ctr] = 0 then
            IDList.Delete(j);
        end;
        Break;
      end;
    end;
  end;
end;

function TbjXLSTable.IsEmptyField(const aField: string; aRow: integer = -1): boolean;
var
  Value: Variant;
begin
  if FindField(aField) = nil then
    raise Exception.Create(aField + ' Column not defined');
  if aRow = -1 then
    aRow := CurrentRow;
{ Comparison with UnAssigned may bot be necessary }
  Value := GetFieldVal(aField, aRow);
  Result := VarIsClear(Value) or VarIsEmpty(Value) or VarIsNull(Value) or (VarCompareValue(Value, Unassigned) = vrEqual);
  if (not Result) and VarIsStr(Value) then
    Result := Value = '';
end;

function TbjXLSTable.IsEmpty(aRow: integer = -1): boolean;
var
  ctr: integer;
  wList: TStringList;
begin
  Result := True;
  if aRow = -1 then
    aRow := CurrentRow;
  wList := FieldList;
  if IDList.Count > 0 then
    wList := IDList;
  for ctr := 0 to wList.Count-1 do
  begin
    if not IsEmptyField(wList[ctr], aRow) then
    begin
      Result := False;
      Exit;
    end;
  end;
end;

end.



