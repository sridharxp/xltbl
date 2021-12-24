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
unit xlstbl3;
{$DEFINE SM }
{$IFNDEF SM }
{$DEFINE LIBXL }
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
  bjXml3_1,
  Variants,
  StrUtils,
  Dialogs;

{$IFDEF LibXL }
type
   TColumn = TFilterColumn;
{$ENDIF }

type

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
    function IsEmpty(aRow: integer = 0): boolean;
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
    Workbook: TbOOK;
    WorkSheet: TSheet;
{$ENDIF LIBXL }
    CurrentColumn: integer;
    CurrentRow: integer;
    SkipCount: integer;
    FieldList: TStringList;
    ColumnList: array of integer;
    procedure SetXLSFile(const aName: string);
    procedure NewFile(const aName: string);
    procedure SetSheet(const aSheet: string);
    procedure Close;
    procedure SetOrigin(const aRow, aColumn: integer);
    function GetFieldVal(const aField: string; aRow: integer = 0): variant;
    function GetFieldCurr(const aField: string): currency;
    function GetFieldFloat(const aField: string): double;
    function GetFieldString(const aField: string): string;
    function GetFieldSDate(const aField: string): string;
    function GetFieldToken(const aField: string): string;
{$IFDEF SM }
    function GetFieldObj(const aField: string): TColumn;
    function GetCellObj(const arow: integer; const aField: string): TCell; overload;
    function GetCellObj(const arow: integer; const aCol: integer): TCell; overload;
    procedure SetFieldFormat(const aField: string; const aFOrmat: Integer);
    procedure SetFormatAt(const aCol: integer; aFOrmat: Integer);
    procedure SetCellFormat(const aField: string; aFOrmat: Integer);
{$ENDIF SM }
    procedure SetFieldVal(const aField: string; const aValue: variant);
    procedure SetFieldStr(const aField: string; const aValue: string);
    procedure SetFieldNum(const aField: string; const aValue: double);
    function FindField(const aName: string): pChar;
//    function GetFieldCol(const aName: string): integer;
    procedure SetFields(const aList: TStrings; const ToWrite: boolean);
    function GetFields(const aList: TStrings): Tstrings;
    procedure ParseXml(const aNode: IbjXml; const FldLst: TStringList);
    function IsEmptyField(const aField: string; aRow: integer = 0): boolean;
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
    property O_Column: integer read FO_Column;
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
{
  FO_Row  := 0;
  FO_Column  := 0;
}
  FToSaveFile := False;
  FieldList := TStringList.Create;
  SkipCount := 1;
  FLastRow := -1;
end;

destructor TbjXLSTable.Destroy;
begin
  FieldList.Clear;
  FieldList.Free;
if Owner then
begin
{$IFDEF SM }
    WorkSheet := nil;
    Workbook := nil;
    XL.Clear;
    XL.Free;
{$ENDIF SM }
{$IFDEF LibXL }
  Workbook.Free;
{$ENDIF LIBXL }
end;
  Inherited;
end;

procedure TbjXLSTable.Save;
begin
{$IFDEF SM }
  if Length(FXlsFileName) > 0 then
    XL.SaveAs(FXlsFileName);
{$ENDIF SM }
{$IFDEF LIBXL }
  Workbook.Save(PChar(FXlsFileName));
{$ENDIF LIBXL}
end;

procedure TbjXLSTable.SaveAs(const aName  : string);
begin
{$IFDEF SM }
  XL.SaveAs(AName);
{$ENDIF SM }
{$IFDEF LIBXL }
  Workbook.Save(PChar(AName));
{$ENDIF LIBXL}
end;

procedure TbjXLSTable.SetXLSFile(const aName: string);
begin
{$IFDEF SM }
  if Assigned(XL) then
  if Owner then
  begin
    XL.Clear;
    XL.Free;
  end;
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
  Workbook.setKey('Sri', 'windows-21212c060ecde40e64bf6569abo5p2h2');
  if FileExists(aName) then
    Workbook.load(pChar(aName));
{$ENDIF LIBXL}
end;

procedure TbjXLSTable.NewFile(const aName  : string);
begin
  SetXlsFile('');
  FXLSFileName := aName;
end;

procedure TbjXLSTable.close;
begin
{$IFDEF SM }
  if not Owner then
  begin
    Xl := nil;
    Exit;
  end;
  if Assigned(XL) then
  begin
    Xl.Clear;
    Xl.Free;
  end;
{$ENDIF SM}
{$IFDEF LIBXL }
  if not Owner then
    Workbook := nil;
  Workbook.Free;
{$ENDIF LIBXL}
end;

procedure TbjXLSTable.SetSheet(const aSheet: string);
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
end;

procedure TbjXLSTable.SetOrigin(const aRow, aColumn: integer);
begin
  FO_Row := aRow;
  FO_Column :=aColumn;
end;

function TbjXLSTable.GetFieldVal(const aField: string; aRow: integer = 0): variant;
{$IFDEF LIBXL }
var
  wCellType: CellType;
  wDateFormat: TFormat;
{$ENDIF LIBXL}
begin
  if FindField(aField) = nil then
  begin
    raise Exception.Create(aField + ' Column not defined');
    Exit;
  end;
  if aRow = 0 then
    aRow := Currentrow;
{$IFDEF SM }
  Result := WorkSheet.Cells[FO_row + aRow, FO_Column + GetFieldCol(aField)].Value;
{$ENDIF SM}
{$IFDEF LIBXL }
  wCellType := WorkSheet.GetCellType(FO_row + aRow, FO_Column + GetFieldCol(aField));
  if (wCellType = CELLTYPE_EMPTY) or (wCellType = CELLTYPE_BLANK) then
    Exit;
  if WorkSheet.isDate(FO_row + aRow, FO_Column + GetFieldCol(aField)) then
  begin
    wDateFormat := Workbook.addFormat();
    wDateFormat.setNumFormat(NUMFORMAT_DATE);
    Result := FormatDateTime('YYYYMMDD', WorkSheet.readNum(FO_row + aRow, FO_Column + GetFieldCol(aField), wDateformat));
    Exit;
  end;
  if wCellType = CELLTYPE_STRING then
    Result := WorkSheet.readStr(FO_row + aRow, FO_Column + GetFieldCol(aField));
  if wCellType = CELLTYPE_NUMBER then
    Result := WorkSheet.readNum(FO_row + aRow, FO_Column + GetFieldCol(aField));
{$ENDIF LIBXL}
end;

function TbjXLSTable.GetFieldSDate(const aField: string): string;
var
  v_vt: variant;
  VType: Integer;
  mnth: integer;
  str: string;
begin
  v_vt := GetFieldVal(aField);
  str := V_vt;
  VType := VarType(V_vt) and VarTypeMask;
  case VType of
  varDate:
    Result := FormatDateTime('YYYYMMDD', V_vt);
  varString:
  begin
     Result := V_vt;
    if str[1] = '''' then
      Result := Copy(Str, 2, Length(str)-1);
    end;
  end;
end;

function TbjXLSTable.GetFieldCurr(const aField: string): currency;
var
  v_vt: variant;
begin
  try
    v_vt := GetFieldVal(aField);
    Result := v_vt;
  except
    Result := 0;
  end;
end;

function TbjXLSTable.GetFieldFloat(const aField: string): double;
var
  v_vt: variant;
  str: string;
  iValue: double;
  iCode: Integer;
begin
  try
    v_vt := GetFieldVal(aField);
    Result := v_vt;
  except
    Result := 0;
  end;
end;

function TbjXLSTable.GetFieldToken(const aField: string): string;
const
  formatChars: array[0..6] of string = ('%', '.00', '.0', '.', ',', '-', '''');
var
  str: string;
  ctr: integer;
begin
  Result := '';
{$IFDEF SM }
  str := WorkSheet.Cells[FO_Row + CurrentRow,
    FO_Column + GetFieldCol(aField)].ValueAsString;
{$ENDIF SM }
{$IFDEF LIBXL }
  if WorkSheet.GetCellType(FO_row + CurrentRow, FO_Column + GetFieldCol(aField)) = CELLTYPE_NUMBER then
    str := FloattoStr(WorkSheet.ReadNum(FO_Row + CurrentRow,
    FO_Column + GetFieldCol(aField)));
  if WorkSheet.GetCellType(FO_row + CurrentRow, FO_Column + GetFieldCol(aField)) = CELLTYPE_STRING then
    str := WorkSheet.readStr(FO_Row + CurrentRow,
    FO_Column + GetFieldCol(aField));
{$ENDIF LIBXL }
  str := Trim(str);;
  for ctr := low(formatChars) to high(formatChars) do
    if Pos(formatChars[ctr], str) > 0 then
      str := stringReplace(str, formatChars[ctr], '', [rfIgnoreCase, rfReplaceAll]);
  Result := str;
end;

function TbjXLSTable.GetFieldString(const aField: string): string;
begin
  if FindField(AField) = nil then
  begin
    raise Exception.Create(aField + ' Column not defined');
    Exit;
  end;
{$IFDEF SM }
  Result := WorkSheet.Cells[FO_Row + CurrentRow,
    FO_Column + GetFieldCol(aField)].ValueAsString;
{$ENDIF SM }
{$IFDEF LIBXL }
  if WorkSheet.GetCellType(FO_row + CurrentRow, FO_Column + GetFieldCol(aField)) = CELLTYPE_NUMBER then
  begin
    Result := FloattoStr(WorkSheet.ReadNum(FO_Row + CurrentRow,
    FO_Column + GetFieldCol(aField)));
  end;
  if WorkSheet.GetCellType(FO_row + CurrentRow, FO_Column + GetFieldCol(aField)) = CELLTYPE_STRING then
  begin
    Result := WorkSheet.readStr(FO_Row + CurrentRow,
    FO_Column + GetFieldCol(aField));
  end;
{$ENDIF LIBXL }
  Result := Trim(Result);
end;

{$IFDEF SM }
function TbjXLSTable.GetFieldObj(const aField: string): TColumn;
var
  ctr: integer;
begin
  ctr := GetFieldCol(aField);
{.$IFDEF SM }
  Result := WorkSheet.Columns[FO_Column + ctr];
{.$ENDIF SM }
{.$IFDEF LIBXL }
{.$ENDIF LIBXL }
end;

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

function TbjXLSTable.GetCellObj(const arow: integer; const aField: string): TCell;
var
  ctr: integer;
begin
  ctr := GetFieldCol(aField);
  Result := WorkSheet.Cells[FO_Row + aRow, FO_Column + ctr];
end;

function TbjXLSTable.GetCellObj(const arow: integer; const aCol: integer): TCell;
begin
  Result := WorkSheet.Cells[FO_Row + aRow, FO_Column + aCol];
end;
{$ENDIF SM }

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
end;

procedure TbjXLSTable.SetFieldNum(const aField: string; const aValue: double);
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
  ptr.VType := VarDouble;
  ptr.VDouble := aValue;
{$ENDIF SM }
{$IFDEF LIBXL }
    WorkSheet.WriteNum(FO_row + CurrentRow, FO_Column + GetFieldCol(aField), aValue);
{$ENDIF LIBXL }
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
    if SkipCount > 0 then
    if (CurrentRow-wTempRow >= SkipCount) then
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
    if SkipCount > 0 then
    if wTempRow-CurrentRow >= SkipCount then
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
(*
{$IFDEF SM }
  FLastRow := WorkSheet.Rows.Count;
{$ENDIF SM }
*)
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
  l_str := LowerCase(aName);
  idx := 0;
  if FieldList.Find(l_str, idx) then
  begin
    Result := ColumnList[idx];
  end;
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
begin
  if not Assigned(aList) then
    Exit;
  FieldList.Clear;
  setLength(ColumnList, aList.Count);
  for ctr := 0 to aList.Count-1 do
  begin
    FieldList.Add(LowerCase(aList.Strings[ctr]));
    if ToWrite then
{$IFDEF SM }
      WorkSheet.Cells[FO_Row, FO_Column + ctr].Value := aList.Strings[ctr];
{$ENDIF SM }
{$IFDEF LIBXL }
      WorkSheet.WriteStr(FO_Row, FO_Column + ctr, pChar(aList.Strings[ctr]));
{$ENDIF LIBXL}
  end;
  FieldList.Sorted := True;
  setLength(ColumnList, FieldList.Count);
  for ctr := 0 to aList.Count + 29 -1 do
  begin
{$IFDEF SM }
    if Length(WorkSheet.Cells[FO_Row, FO_Column+ctr].ValueAsString) = 0 then
{$ENDIF SM }
{$IFDEF LIBXL }
    if Length(WorkSheet.readStr(FO_Row, FO_Column+ctr)) = 0 then
{$ENDIF LIBXL}
      Break;
    for j := 0 to FieldList.Count-1 do
{$IFDEF SM }
    if LowerCase(WorkSheet.Cells[FO_Row, FO_Column+ctr].ValueAsString) =
      LowerCase(FieldList.Strings[j]) then
{$ENDIF SM }
{$IFDEF LIBXL }
    if LowerCase(WorkSheet.readStr(FO_Row, FO_Column+ctr)) =
      LowerCase(FieldList.Strings[j]) then
{$ENDIF LIBXL}
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
  aliasNode: IbjXml;
begin
  FldLst.Clear;
  aliasNode := aNode.SearchForTag(nil, 'Alias');
  while Assigned(aliasNode) do
  begin
    FldLst.Add(aliasNode.GetContent);
    aliasNode := aNode.SearchForTag(aliasNode, 'Alias');
  end;
end;

Function TbjXLSTable.GetFields(const aList: TStrings): TStrings;
var
  ctr, j: integer;
begin
  FieldList.Clear;
{$IFDEF SM }
  for ctr := 0 to aList.Count + 29 -1 do
  begin
    if Length(WorkSheet.Cells[FO_Row, FO_Column+ctr].ValueAsString) = 0 then
      Break;
    if Pos('Dr_', WorkSheet.Cells[FO_Row, FO_Column+ctr].ValueAsString) > 0 then
    begin
      FieldList.Add(WorkSheet.Cells[FO_Row, FO_Column+ctr].ValueAsString);
      Continue;
    end;
    if Pos('Cr_', WorkSheet.Cells[FO_Row, FO_Column+ctr].ValueAsString) > 0 then
    begin
      FieldList.Add(WorkSheet.Cells[FO_Row, FO_Column+ctr].ValueAsString);
      Continue;
    end;
    for j := 0 to aList.Count-1 do
    begin
      if LowerCase(WorkSheet.Cells[FO_Row, FO_Column+ctr].ValueAsString) =
        LowerCase(aList.Strings[j]) then
      begin
        FieldList.Add(LowerCase(aList.Strings[j]));
        break;
      end;
    end;
  end;
  FieldList.Sorted := True;
  setLength(ColumnList, FieldList.Count);
  for ctr := 0 to aList.Count + 29 -1 do
  begin
    if Length(WorkSheet.Cells[FO_Row, FO_Column+ctr].ValueAsString) = 0 then
      Break;
    for j := 0 to FieldList.Count-1 do
    if LowerCase(WorkSheet.Cells[FO_Row, FO_Column+ctr].ValueAsString) =
      LowerCase(FieldList.Strings[j]) then
    begin
        ColumnList[j] := ctr;
        break;
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
      if LowerCase(WorkSheet.readStr(FO_Row, FO_Column+ctr)) =
        LowerCase(aList.Strings[j]) then
      begin
        FieldList.Add(LowerCase(aList.Strings[j]));
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
    if LowerCase(WorkSheet.readStr(FO_Row, FO_Column+ctr)) =
      LowerCase(FieldList.Strings[j]) then
    begin
        ColumnList[j] := ctr;
        break;
    end;
  end;
{$ENDIF LIBXL }
  Result := FieldList;
end;

function TbjXLSTable.IsEmptyField(const aField: string; aRow: integer = 0): boolean;
var
  Value: Variant;
begin
  if aRow = 0 then
    aRow := CurrentRow;
{ Comparison with UnAssigned may bot be necessary }
  Value := GetFieldVal(aField, aRow);
  Result := VarIsClear(Value) or VarIsEmpty(Value) or VarIsNull(Value) or (VarCompareValue(Value, Unassigned) = vrEqual);
  if (not Result) and VarIsStr(Value) then
    Result := Value = '';
end;

function TbjXLSTable.IsEmpty(aRow: integer = 0): boolean;
var
  ctr: integer;
begin
  Result := True;
  if aRow = 0 then
    aRow := CurrentRow;
  for ctr := 0 to FieldList.Count-1 do
  begin
    if not IsEmptyField(FieldList[ctr], aRow) then
    begin
      Result := False;
      Exit;
    end;
  end;
end;

end.



