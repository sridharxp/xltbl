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
{.$DEFINE SM }
{.$DEFINE LIBXL }
{.$DEFINE NXL }
{$DEFINE LXW}

interface
uses
  SysUtils, Classes,
  Windows,
{$IFDEF SM }
  XLSFile,
  XLSWorkbook,
{$ENDIF SM }
{$IFDEF LibXL }
  LibXL404,
{$ENDIF }
{$IFDEF NXL }
  nExcel,
{$ENDIF NXL }
{$IFDEF LXW }
  xlsxwriterapi,
{$ENDIF LXW }
  bjXml33,
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
  wDateFormat: TXLFormat;
    wStrFormat: TXLFormat;
{$ENDIF LIBXL }
{$IFDEF NXL }
    Workbook: IXLSWorkbook;
    WorkSheet: IXLSWorksheet;
{$ENDIF NXL }
{$IFDEF LXW }
    Workbook: PLXW_Workbook;
    WorkSheet: PLXW_WorkSheet;
{$ENDIF LXW }
    CurrentColumn: integer;
    CurrentRow: integer;
    SkipCount: integer;
    FieldList: TStringList;
    ColumnList: array of integer;
    IDList: TStringList;
    ListColList: TStringList;
    procedure SetXLSFile(const aName: string);
    function GetFieldCol(const aName: string): integer;
    procedure SetSheet(const aSheet: string);
    procedure Close;
    procedure SetOrigin(const aRow, aColumn: integer);
    procedure SetOColumn(const aColumn: integer);
    function GetRecVal(const aCol: Integer; aRow: integer = -1): variant;
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
{$IFDEF NXL }
    function GetFieldObj(const aField: string): IXLSRange;
    function GetCellObj(const arow: integer; const aField: string): IXLSRange; overload;
    function GetCellObj(const arow: integer; const aCol: integer): IXLSRange; overload;
{$ENDIF NXL }
    procedure SetFieldVal(Const aField: string; const aValue: variant); overload;
    procedure SetFieldVal(const aField: string; const aValue: TVarRec; aFormat: Pointer); overload;
    procedure SetFieldStr(const aField: string; const aValue: string);
    procedure SetFieldWStr(const aField: string; const aValue: string);
    procedure SetFieldNum(const aField: string; const aValue: double);
    procedure SetRecVal(const aCol: Integer; const aValue: TVarRec; aFormat: Pointer); overload;
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
Function TryJulianDateToDateTime(const AValue: Double; out ADateTime: TDateTime): Boolean;

implementation


constructor TbjXLSTable.Create;
begin
  inherited;
{
  FO_Row  := 0;
  FO_Column  := 0;
}
{$IFDEF NXL }
  FO_Row  := 1;
  FO_Column  := 1;
{$ENDIF NXL }
  FToSaveFile := False;
  FieldList := THashedStringList.Create;
  IDList := THashedStringList.Create;
  ListColList := TStringList.Create;
  SkipCount := 1;
  FLastRow := -1;
end;

destructor TbjXLSTable.Destroy;
begin
{  XLSFileName := ''; }
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
{$IFDEF NXL }
    Workbook.SaveAs(FXlsFileName);
{$ENDIF NXL }
end;

procedure TbjXLSTable.SaveAs(const aName  : string);
begin
{$IFDEF SM }
  XL.SaveAs(AName);
{$ENDIF SM }
{$IFDEF LIBXL }
  Workbook.Save(PChar(aName))
{$ENDIF LIBXL}
{$IFDEF NXL }
  Workbook.SaveAs(aName);
{$ENDIF NXL }
end;

procedure TbjXLSTable.SetXLSFile(const aName: string);
begin
  if FXlsFileName = aName then
    Exit;
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
  if (Pos('.xlsx', LowerCase(aName)) = 0) then
  Workbook := TBinBook.Create
  else
  Workbook := TXmlBook.Create;
  Workbook.setLocale('UTF-8');
  Workbook.setKey('Name', 'Key');
  if FileExists(aName) then
    Workbook.load(pChar(aName));
  wDateFormat := Workbook.addFormat();
  wDateFormat.setNumFormat(NUMFORMAT_DATE);
  wStrFormat := Workbook.addFormat();
  wStrFormat.setNumFormat(NUMFORMAT_TEXT);
{$ENDIF LIBXL}
{$IFDEF NXL }
  FXLSFileName := aName;
  FOwner := True;
  Workbook := TXLSWorkbook.Create;
  if FileExists(aName) then
    Workbook.Open(aName);
{$ENDIF NXL }
{$IFDEF LXW }
  FXLSFileName := aName;
  FOwner := True;
  WorkBook := workbook_new(pUTF8Char(aName));
{$ENDIF LXW }
end;

{
procedure TbjXLSTable.NewFile(const aName  : string);
begin
  SetXlsFile('');
  FXLSFileName := aName;
end;
}

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
  if Assigned(wDateFormat) then
    wDateFormat.Free;
  if Assigned(wStrFormat) then
    wStrFormat.Free;
  if Assigned(WorkSheet) then
    WorkSheet.Free;
{$ENDIF LIBXL}
{$IFDEF NXL }
  Workbook := nil;
{$ENDIF NXL }
{$IFDEF LXW }
  workbook_close(Workbook);
{$ENDIF LXW }
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
  if Assigned(WorkSheet) then
    WorkSheet.Free;
  WorkSheet := Workbook.GetSheetbyName(PChar(aSheet));
  if not Assigned(WorkSheet) then
  begin
    WorkSheet := Workbook.addSheet(PChar(aSheet));
  end;
{$ENDIF LIBXL}
{$IFDEF NXL }
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
{$ENDIF NXL }
{$IFDEF LXW }
    Worksheet := workbook_add_worksheet(Workbook, pUTF8Char(aSheet));
{$ENDIF LXW }
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
  rDate: double;
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
    rDate := WorkSheet.readNum(FO_row + aRow, FO_Column + aCol, wDateformat);
    Result := FormatDateTime('YYYYMMDD', rDate);
    Exit;
  end;
  if wCellType = CELLTYPE_STRING then
    Result := WorkSheet.readStr(FO_row + aRow, FO_Column + aCol);
  if wCellType = CELLTYPE_NUMBER then
    Result := WorkSheet.readNum(FO_row + aRow, FO_Column + aCol);
{$ENDIF LIBXL}
{$IFDEF NXL }
  Result := WorkSheet.Cells[FO_row + aRow, FO_Column + aCol].Value;
{$ENDIF NXL }
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
  rDt: TDateTime;
begin
  v_vt := GetFieldVal(aField);
  if VarIsEmpty(v_vt) or VarIsNull(v_vt) then
    Exit;
  VType := VarType(V_vt) and VarTypeMask;
  case VType of
  varDate:
    Result := FormatDateTime('YYYYMMDD', V_vt);
  VarDouble:
    begin
    if TryJulianDateToDateTime(V_Vt, rDt) then
      V_Vt := rDt;
    Result := FormatDateTime('YYYYMMDD', V_Vt);
    end;
  varString:
  begin
     Result := V_vt;
    Result := StringReplace(Result, #10, '', [rfReplaceAll, rfIgnoreCase]);
    if Result[1] = '''' then
      Result := Copy(Result, 2, Length(Result)-1);
    end;
  else
    Result := GetFieldString(aField);
  end;
end;

function TbjXLSTable.GetFieldCurr(const aField: string): currency;
var
  v_vt: variant;
  rStr: string;
  aCol: Integer;
begin
  if FindField(AField) = nil then
  begin
    raise Exception.Create(aField + ' Column not defined');
    Exit;
  end;
  aCol := GetFieldCol(aField);
  v_vt := GetFieldVal(aField);
  if VarIsEmpty(v_vt) or VarIsNull(v_vt) then
    Exit;
  try
    Result := v_vt;
  except
    Result := 0;
  end;
  if Result = 0 then

begin
   rStr := GetFieldString(aField);
    if Length(rStr) = 0 then
      Exit;
    rStr := StringReplace(rStr, ',', '', [rfReplaceAll, rfIgnoreCase]);
    rStr := StringReplace(rStr, #10, '', [rfReplaceAll, rfIgnoreCase]);
    try
    Result := StrtoFloat(rStr);
    except
      Result := 0;
    end;
  end;
end;
function TbjXLSTable.GetFieldFloat(const aField: string; aRow: integer = -1): double;
var
  v_vt: variant;
  str: string;
  iValue: double;
  iCode: Integer;
  aCol: Integer;
begin
  if FindField(AField) = nil then
  begin
    raise Exception.Create(aField + ' Column not defined');
    Exit;
  end;
  aCol := GetFieldCol(aField);
  v_vt := GetRecVal(aCol, aRow);
  if VarIsEmpty(v_vt) or VarIsNull(v_vt) then
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
  rVal: Variant;
  wStr: WideString;
  rCellStr: String;
  ctr: integer;
begin
  Result := '';
{$IFDEF SM }
  WorkSheet.Cells[FO_Row + CurrentRow,
    FO_Column + GetFieldCol(aField)].FormatIndex := 35;
  wStr := WorkSheet.Cells[FO_Row + CurrentRow,
    FO_Column + GetFieldCol(aField)].ValueAsString;
  rCellStr := wStr;
{$ENDIF SM }
{$IFDEF LIBXL }
    rCellStr := WorkSheet.readStr(FO_Row + CurrentRow,
    FO_Column + GetFieldCol(aField), wStrFormat);
{$ENDIF LIBXL }
{$IFDEF NXL }
  WorkSheet.Cells[FO_Row + CurrentRow,
    FO_Column + GetFieldCol(aField)].NumberFormat := '@';
  rVal := WorkSheet.Cells[FO_Row + CurrentRow,
    FO_Column + GetFieldCol(aField)].Value;
  try
  if not (VarIsEmpty(rVal) or VarIsNull(rVal)) then
    wStr := rVal;
  except
    wStr := '';
  end;
  rCellStr := wStr;
{$ENDIF NXL }
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
    raise Exception.Create(aField + ' Column not defined');
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
  WorkSheet.Cells[FO_Row + aRow, FO_Column + aCol].FormatIndex := 35;
  wStr := UTF8Encode(WorkSheet.Cells[FO_Row + aRow,
    FO_Column + aCol].ValueAsString);
  Result := wStr;
{$ENDIF SM }
{$IFDEF LIBXL }
  Result := WorkSheet.readStr(FO_Row + aRow, FO_Column + aCol, wStrFormat);
{$ENDIF LIBXL }
{$IFDEF NXL }
  WorkSheet.Cells[FO_Row + aRow, FO_Column + aCol].NumberFormat := '@';
  try
  wStr := WorkSheet.Cells[FO_Row + aRow,
    FO_Column + aCol].Value;
  except
    wStr := '';
  end;
  Result := UTF8Encode(wStr);
{$ENDIF NXL }
  Result := Trim(Result);
end;


{$IFDEF SM }
function TbjXLSTable.GetFieldObj(const aField: string): TColumn;
var
  ctr: integer;
begin
  ctr := GetFieldCol(aField);
  Result := WorkSheet.Columns[FO_Column + ctr];
 end;
{$ENDIF SM }
{$IFDEF NXL }
function TbjXLSTable.GetFieldObj(const aField: string): IXLSRange;
var
  ctr: integer;
begin
  ctr := GetFieldCol(aField);
{.$IFDEF LIBXL }
{.$ENDIF LIBXL }
  Result := WorkSheet.Selection.EntireColumn.Item[FO_Column + ctr]; 
end;
{$ENDIF NXL }

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
  MyCell: TCell;
begin
  ctr := GetFieldCol(aField);
  if ctr = -1 then
    Exit;
  MyCell := GetCellObj(FO_Row + CurrentRow, aField);
  MyCell.FormatStringIndex := aFOrmat;
end;
{$ENDIF SM }

{$IFDEF SM }
function TbjXLSTable.GetCellObj(const arow: integer; const aField: string): TCell;
var
  ctr: integer;
begin
  ctr := GetFieldCol(aField);
  Result := WorkSheet.Cells[FO_Row + aRow, FO_Column + ctr];
end;
{$ENDIF SM }
{$IFDEF NXL }
function TbjXLSTable.GetCellObj(const arow: integer; const aField: string): IXLSRange;
var
  ctr: integer;
begin
  ctr := GetFieldCol(aField);
  Result := WorkSheet.Cells[FO_Row + aRow, FO_Column + ctr];
end;
{$ENDIF NXL }

{$IFDEF SM }
function TbjXLSTable.GetCellObj(const arow: integer; const aCol: integer): TCell;
begin
  Result := WorkSheet.Cells[FO_Row + aRow, FO_Column + aCol];
end;
{$ENDIF SM }
{$IFDEF NXL }
function TbjXLSTable.GetCellObj(const arow: integer; const aCol: integer): IXLSRange;
begin
  Result := WorkSheet.Cells[FO_Row + aRow, FO_Column + aCol];
end;
{$ENDIF NXL }

procedure TbjXLSTable.SetFieldVal(const aField: string; const aValue: variant);
{$IFDEF LIBXL }
var
  VType  : Integer;
{$ENDIF LIBXL }
{$IFDEF LXW }
var
  VType  : Integer;
{$ENDIF LXW }
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
{$IFDEF NXL }
  WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + GetFieldCol(aField)].Value := aValue;
{$ENDIF NXL }
end;

procedure TbjXLSTable.SetFieldVal(const aField: string; const aValue: TVarRec; aFormat: Pointer);
var
  VType  : Integer;
  rCol: Integer;
begin
  if FindField(AField) = nil then
  begin
    raise Exception.Create(aField + ' Column not defined');
    Exit;
  end;
  rCol := GetFieldCol(aField);
  SetRecVal(rCol, aValue, aFormat);
end;
procedure TbjXLSTable.SetRecVal(const aCol: Integer; const aValue: TVarRec; aFormat: Pointer);
var
  VType  : Integer;
begin
{$IFDEF SM }
    with aValue do
  case aValue.VType of
  vtInteger:
    WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + aCol].Value := vInteger;
  VtExtended:
    WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + aCol].Value := vExtended^;
  vtCurrency:
    WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + aCol].Value := vCurrency^;
  vtString:
    WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + aCol].Value := vString^;
  vtPChar:
    WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + aCol].Value := vPChar^;
  vtPWideChar:
    WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + aCol].Value := vPWideChar^;
  end;
{$ENDIF SM }
{$IFDEF LIBXL }
    with aValue do
  case aValue.VType of
  vtInteger:
    WorkSheet.WriteNum(FO_row + CurrentRow, FO_Column + aCol, vInteger);
  vtExtended:
    WorkSheet.WriteNum(FO_row + CurrentRow, FO_Column + aCol, vExtended^);
  VtCurrency:
    WorkSheet.WriteNum(FO_row + CurrentRow, FO_Column + aCol, vCurrency^);
  vtString:
    WorkSheet.WriteStr(FO_row + CurrentRow, FO_Column + aCol, pChar(vString));
  vtPChar:
    WorkSheet.WriteStr(FO_row + CurrentRow, FO_Column + aCol, vPChar);
  vtPWideChar:
    WorkSheet.WriteStr(FO_row + CurrentRow, FO_Column + aCol, PChar(vPWideChar));
  end;
{$ENDIF LIBXL }
{$IFDEF NXL }
    with aValue do
  case aValue.VType of
  vtInteger:
    WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + aCol].Value := vInteger;
  VtExtended:
    WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + aCol].Value := vExtended^;
  vtCurrency:
    WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + aCol].Value := vCurrency^;
  vtString:
    WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + aCol].Value := vString^;
  vtPChar:
    WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + aCol].Value := vPChar^;
  vtPWideChar:
    WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + aCol].Value := vPWideChar^;
  end;
{$ENDIF NXL }
{$IFDEF LXW }
  with aValue do
  case aValue.VType of
  vtInteger:
    WorkSheet_Write_Number(worksheet, FO_row + CurrentRow, FO_Column + aCol, vInteger, nil);
  VtExtended:
    WorkSheet_Write_Number(worksheet, FO_row + CurrentRow, FO_Column + aCol, vExtended^, nil);
  vtCurrency:
    WorkSheet_Write_Number(worksheet, FO_row + CurrentRow, FO_Column + aCol, vCurrency^, nil);
  vtString:
    WorkSheet_Write_string(worksheet, FO_row + CurrentRow, FO_Column + aCol, pChar(vString), nil);
  vtPChar:
    WorkSheet_Write_string(worksheet, FO_row + CurrentRow, FO_Column + aCol, vPChar, nil);
  vtPWideChar:
    WorkSheet_Write_string(worksheet, FO_row + CurrentRow, FO_Column + aCol, pointer(vPWideChar), nil);
  end;
{$ENDIF LXW }
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
{$IFDEF NXL }
  WorkSheet.Cells[FO_row + CurrentRow, FO_Column + GetFieldCol(aField)].Value := aValue;
{$ENDIF NXL }
end;
procedure TbjXLSTable.SetFieldWStr(const aField: string; const aValue: string);
{$IFDEF SM }
var
  ptr: pVarData;
{$ENDIF SM }
{$IFDEF NXL }
var
  ptr: pVarData;
{$ENDIF NXL }
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
{$IFDEF NXL }
  ptr := FindVarData(WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + GetFieldCol(aField)].Value);
  ptr.VType := varOleStr;
  ptr.VOleStr := PWideChar(UTF8Decode(aValue));
{$ENDIF NXL }
end;

procedure TbjXLSTable.SetFieldNum(const aField: string; const aValue: double);
{$IFDEF SM }
var
  ptr: pVarData;
{$ENDIF SM }
{$IFDEF NXL }
var
  ptr: pVarData;
{$ENDIF NXL }
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
{$IFDEF NXL }
  ptr := FindVarData(WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + GetFieldCol(aField)].Value);
  ptr.VType := VarDouble;
  ptr.VDouble := aValue;
{$ENDIF NXL }
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
  VType := VarType(aMsg) and VarTypeMask;
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
{$IFDEF NXL }
  WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + aCol].Value := aMsg;
{$ENDIF NXL }
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
  WorkSheet.removeRow(FO_row + CurrentRow, FO_row + CurrentRow);
{$ENDIF LIBXL }
{$IFDEF NXL }
  WorkSheet.RCRange[FO_row + CurrentRow, 0, FO_row + CurrentRow, 0].EntireRow.Delete(xlShiftUp);
{$ENDIF NXL }

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
{$IFDEF NXL }
  WorkSheet.Selection.EntireRow.Item[FO_row + CurrentRow].Clear;
{$ENDIF NXL }
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
{$IFDEF NXL }
  Result := WorkSheet.Cells[FO_Row, FO_Column+Idx].Value;
{$ENDIF NXL }
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
  BackupList: TstringList;
begin
  if not Assigned(aList) then
    Exit;
  BackupList := TStringList.Create;
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
{$IFDEF NXL }
      WorkSheet.Cells[FO_Row, FO_Column + ctr].Value := aList.Strings[ctr];
{$ENDIF NXL }
{$IFDEF LXW }
    worksheet_write_string(Worksheet, FO_Row, FO_Column + ctr,
      pUTF8Char(aList.Strings[ctr]), nil);
{$ENDIF LXW }
  end;
  for j := 0 to aList.Count-1 do
    BackupList.Add(aList.Strings[j]);
  FieldList.Sorted := True;
  setLength(ColumnList, FieldList.Count);
  for Ctr := 0 to BackupList.Count-1 do
    for j := 0 to FieldList.Count-1 do
   if PackStr(BackupList.Strings[Ctr]) = FieldList.Strings[j] then
        ColumnList[j] := ctr;
  BackupList.Free;
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
  rCellStr: String;
{.$ENDIF SM }
{.$IFDEF NXL }
//  rCellStr: WideString;
{.$ENDIF NXL }
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
    rCellStr := WorkSheet.readStr(FO_Row, FO_Column+ctr, wStrFormat);
    if Length(rCellStr) = 0 then
      Break;
    if Pos('Dr_', rCellStr) > 0 then
    begin
      FieldList.Add(rCellStr);
      Continue;
    end;
    if Pos('Cr_', rCellStr) > 0 then
    begin
      FieldList.Add(rCellStr);
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
    rCellStr := WorkSheet.readStr(FO_Row, FO_Column+ctr, wStrFormat);
    if Length(rCellStr) = 0 then
      Break;
    for j := 0 to FieldList.Count-1 do
    begin
      if PackStr(rCellStr) =
        FieldList.Strings[j] then
    begin
        ColumnList[j] := ctr;
        break;
      end;
    end;
  end;
{$ENDIF LIBXL }
{$IFDEF NXL }
  for ctr := 0 to aList.Count + 29-1 do
  begin
    rCellStr := VarToWideStr(WorkSheet.Cells[FO_Row, FO_Column+ctr].Value);
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
    rCellStr := VarToWideStr(WorkSheet.Cells[FO_Row, FO_Column+ctr].Value);
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
{$ENDIF NXL }
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
{  Result := VarIsClear(Value) or VarIsEmpty(Value) or VarIsNull(Value) or (VarCompareValue(Value, Unassigned) = vrEqual); }
  Result := VarIsEmpty(Value) or VarIsNull(Value);
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
Function TryJulianDateToDateTime(const AValue: Double; out ADateTime: TDateTime): Boolean;
var
  a,b,c,d,e,m:longint;
  day,month,year: word;
begin
  a := trunc(AValue + 32044.5);
  b := (4*a + 3) div 146097;
  c := a - (146097*b div 4);
  d := (4*c + 3) div 1461;
  e := c - (1461*d div 4);
  m := (5*e+2) div 153;
  day := e - ((153*m + 2) div 5) + 1;
  month := m + 3 - 12 *  ( m div 10 );
  year := (100*b) + d - 4800 + ( m div 10 );
  result := TryEncodeDate ( Year, Month, Day, ADateTime );
  if Result then
    ADateTime:=ADateTime+frac(AValue-0.5);
end;

end.



