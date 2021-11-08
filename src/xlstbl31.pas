unit xlstbl31;
{
Index (0-based) Format string
0 General
1 0
2 0.00
3 #,##0
4 ($#,##0_);($#,##0)
5 ($#,##0_);[Red]($#,##0)
6 ($#,##0.00_);($#,##0.00)
7 ($#,##0.00_);[Red]($#,##0.00)
8 0%
9 0.00%
10 0.00E+00
11 # ?/?
12 # ??/??
13 m/d/yy
14 d-mmm-yy
15 d-mmm
16 mmm-yy
17 h:mm AM/PM
18 h:mm:ss AM/PM
19 h:mm
20 h:mm:ss
21 m/d/yy h:mm
22 _($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)
23 _(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)
24 _($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)
25 _(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)
26 #,##0.00
27 (#,##0_);(#,##0)
28 (#,##0_);[Red](#,##0)
29 (#,##0.00_);(#,##0.00)
30 (#,##0.00_);[Red](#,##0.00)
31 mm:ss
32 [h]:mm:ss
33 mm:ss.0
34 ##0.0E+0
35 @
}
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
  IniFiles,
 VchLib,
  Dialogs;

{$IFDEF LibXL }
type
   TColumn = TFilterColumn;
{$ENDIF }

type
(*
  IDataset = interface(IInterface)
    ['{FB7722CC-62D4-4400-B93A-AAA4401246BD}']
    function GetFields: IDataFields; stdcall;
    function GetHead: IDatasetHead; stdcall;
    function GetRecordCount: Integer; stdcall;
    function GetValues: IDataValues; stdcall;
    function IsEmpty:Boolean;
    procedure First;
    procedure Last;
    procedure Next;
    procedure Prior;
    procedure Append;
    procedure Post;
    procedure Edit;
    procedure Open;
    procedure Close;
    procedure Clear;
    function Locate(const KeyFields: string; const KeyValues: Variant):Boolean; stdcall;
    function CopyData(const ASource:IDataset): HResult; stdcall;
    function CopyHead(const ASource:IDatasetHead): HResult; stdcall;
    function GetActive: Boolean;
    function GetBof: Boolean;
    function GetEof: Boolean;
    function RecordState: TMemStarUpdateStatus; stdcall;
    procedure ClearState;
    function GetEditState: Boolean; stdcall;
    procedure ReplaceValue(const AFieldName:string; AOldVal, ARepVal:Variant);
        stdcall;
    procedure SetActive(Value: Boolean);
    procedure SetEditState(const Value: Boolean); stdcall;
    property Active: Boolean read GetActive write SetActive;
    property Bof: Boolean read GetBof;
    property EditState: Boolean read GetEditState write SetEditState;
    property Eof: Boolean read GetEof;
    property Fields: IDataFields read GetFields;
    property Head: IDatasetHead read GetHead;
    property RecordCount: Integer read GetRecordCount;
    property Values: IDataValues read GetValues;
  end;
*)

  TbjXLSTable = class(TInterfacedObject)
  private
    FO_Row: integer;
    FO_Column: integer;
{    FMaxRow: integer; }
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
    Workbook: TbOOK;
    WorkSheet: TSheet;
{$ENDIF LIBXL }
    CurrentColumn: integer;
    CurrentRow: integer;
    SkipCount: integer;
    FieldList: TStringList;
    ColumnList: array of integer;
    IDList: TStringList;
    procedure SetXLSFile(const aName: string);
    procedure NewFile(const aName: string);
    procedure SetSheet(const aSheet: string);
    procedure Close;
    procedure SetOrigin(const aRow, aColumn: integer);
    function GetFieldVal(const aField: string; aRow: integer = -1): variant;
//    function GetFieldFmla(const aField: string): string;
//    function GetFieldStr(const aField: string): string;
    function GetFieldCurr(const aField: string): currency;
    function GetFieldFloat(const aField: string): double;
    function GetFieldString(const aField: string): string;
    function GetFieldSDate(const aField: string): string;
    function GetFieldToken(const aField: string): string;
//    procedure SetFieldType(const aField: string; const aType: integer);
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
    property O_Column: integer read FO_Column;
{    property MaxRow: integer read FMaxRow write FMaxRow; }
    property XLSFileName: string write SetXLSFile;
    property ToSaveFile: boolean read FToSaveFile write FToSaveFile;
    property SheetName: string write SetSheet;
    property Owner: boolean read Fowner write Fowner;
//    property FieldVal[aField: string]: variant read GetVal write SetVal;
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
  FieldList := THashedStringList.Create;
  IDList := THashedStringList.Create;
{  FMaxRow := -1; }
  SkipCount := 1;
  FLastRow := -1;
end;

destructor TbjXLSTable.Destroy;
begin
  FieldList.Clear;
  FieldList.Free;
  if Assigned(IDList) then
  IDList.Free;
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
//  WorkSheet := Workbook.Sheets[0];
{$ENDIF SM }
{$IFDEF LIBXL }
  FXLSFileName := aName;
  FOwner := True;
  Workbook := TBinBook.Create;
  Workbook.setKey('Sri', 'windows-21212c060ecde40e64bf6569abo5p2h2');
  if FileExists(aName) then
    Workbook.load(pChar(aName));
//  WorkSheet := Workbook.getSheet(0);
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

function TbjXLSTable.GetFieldVal(const aField: string; aRow: integer = -1): variant;
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
  if aRow = -1 then
    aRow := Currentrow;
{$IFDEF SM }
//  Result := WorkSheet.Cells[FO_row + CurrentRow, FO_Column + GetFieldCol(aField)].Value;
  Result := WorkSheet.Cells[FO_row + aRow, FO_Column + GetFieldCol(aField)].Value;
//  Result := Encode(WorkSheet.Cells[FO_row + CurrentRow, FO_Column + GetFieldCol(aField)].Value);
{$ENDIF SM}
{$IFDEF LIBXL }
  wCellType := WorkSheet.GetCellType(FO_row + aRow, FO_Column + GetFieldCol(aField));
  if (wCellType = CELLTYPE_EMPTY) or (wCellType = CELLTYPE_BLANK) then
//  if (wCellType = CELLTYPE_EMPTY) then
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
//  str: string;
//  iValue: double;
//  iCode: Integer;
begin
  try
    v_vt := GetFieldVal(aField);
    Result := v_vt;
  except
    Result := 0;
  end;
end;

function TbjXLSTable.GetFieldFloat(const aField: string): double;
//const
//  formatChars: array[0..2] of Char = ',%''';
var
  v_vt: variant;
  str: string;
  iValue: double;
  iCode: Integer;
//  VType: Integer;
begin
  try
    v_vt := GetFieldVal(aField);
    Result := v_vt;
  except
{
    VType := VarType(v_vt) and VarTypeMask;
    case VType of
      varString:
      begin
        str := V_vt;
//    for ctr := low(formatChars) to high(formatChars) do
//    if Pos(formatChars[ctr], str) > 0 then
//      str := stringReplace(Str, formatChars[ctr], '', [rfIgnoreCase, rfReplaceAll]);
        Val(Str, iValue, iCode);
        if iCode = 0 then
          Result := iValue;
      end;
    end;
}
    Result := 0;
  end;
end;
{
  VType := VarType(v_vt) and VarTypeMask;
  // Set a string to match the type
  case VType of
  varEmpty:
    begin
    Result := 0;
    Exit;
    end;
  varInteger:
    begin
    Result := V_vt;
    end;
  varDate:
    begin
    Result := V_vt;
    end;
  varDouble:
    begin
    Result := V_vt;
    end;
  varCurrency:
    begin
    Result := V_vt;
    end;
  varString:
    begin
    Val(Str, iValue, iCode);
    if iCode = 0 then
      Result := iValue;
    end;
  else
    Result := 0;
  end;
}

function TbjXLSTable.GetFieldToken(const aField: string): string;
const
  formatChars: array[0..6] of string = ('%', '.00', '.0', '.', ',', '-', '''');
var
  str: string;
  ctr: integer;
begin
  Result := '';
{  v_vt := GetFieldVal(aField);
  VType := VarType(v_vt) and VarTypeMask;
  // Set a string to match the type
  case VType of
  varEmpty:
    Exit;
  end;
  str := V_vt;
  if Pos('%', str) > 0 then
  begin
    str := stringReplace(Str, '%', '', [rfIgnoreCase, rfReplaceAll]);
    Result := str;
    Exit;
  end;
  if RightStr(str, 3) = '.00' then
  begin
    str := stringReplace(Str, '.00', '', [rfIgnoreCase, rfReplaceAll]);
    Result := str;
    Exit;
  end;
  if RightStr(str, 2) = '.0' then
  begin
    str := stringReplace(Str, '.0', '', [rfIgnoreCase, rfReplaceAll]);
    Result := str;
    Exit;
  end;
  if Pos('.', str) > 0 then
  begin
    if STrtoFloat(str) < 1 then
      Result := FormatFloat('##.##', V_vt * 100);
    Exit;
  end;
}
{
  VType := VarType(v_vt) and VarTypeMask;
  // Set a string to match the type
  case VType of
  varEmpty:
    begin
    Result := '';
    Exit;
    end;
  varInteger:
    begin
    Result := InttoStr(V_vt);
    end;
  varDate:
    begin
    Result := InttoStr(V_vt);
    end;
  varDouble:
    begin
      if V_vt <1 then
        Result := FormatFloat('##.##', V_vt * 100);
    Result := FormatFloat('##.##', V_vt);
//    ShowMessage('Double');
    end;
  varCurrency:
    begin
    Result := FormatFloat('##.##', V_vt);
    end;
  varString:
    begin
      Result := V_vt;
//    ShowMessage('Text');
    end;
  else
    Result := '';
  end;
}
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
//var
//  val: variant;
begin
  if FindField(AField) = nil then
  begin
    raise Exception.Create(aField + ' Column not defined');
    Exit;
  end;
{$IFDEF SM }
//  Val := WorkSheet.Cells[FO_Row + CurrentRow,
//    FO_Column + GetFieldCol(aField)].Value;
  Result := UTF8Encode(WorkSheet.Cells[FO_Row + CurrentRow,
    FO_Column + GetFieldCol(aField)].ValueAsString);
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


{
procedure TbjXLSTable.SetFieldType(const aField: string; const aType: integer);
var
  ctr: integer;
  ccolumn: TColumn;
begin
  ctr := GetFieldCol(aField);
  Ccolumn := WorkSheet.Columns[FO_Column + ctr];
  Ccolumn.FormatStringIndex:= aType;
end;
}

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
//  MyColumn:= WorkSheet.Columns[FO_Column + ctr];
  MyColumn:= GetFieldObj(AField);
//  if Assigned(MyColumn)then
//  begin
    MyColumn.FormatStringIndex := aFOrmat;
//  end;
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
//  if WorkSheet.GetCellType(FO_row + CurrentRow, FO_Column + GetFieldCol(aField)) = CELLTYPE_NUMBER then
//    WorkSheet.WriteNum(FO_row + CurrentRow, FO_Column + GetFieldCol(aField), aValue);
//  if WorkSheet.GetCellType(FO_row + CurrentRow, FO_Column + GetFieldCol(aField)) = CELLTYPE_STRING then
//    WorkSheet.WriteStr(FO_row + CurrentRow, FO_Column + GetFieldCol(aField), nil);
  VType := VarType(aValue) and VarTypeMask;
  // Set a string to match the type
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
//  WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + GetFieldCol(aField)].Value := Decode(aValue);
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
//  WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + GetFieldCol(aField)].Value := aValue;
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
//  WorkSheet.Cells[FO_Row + CurrentRow, FO_Column + GetFieldCol(aField)].Value := aValue;
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
//    if SkipCount > 0 then
//    if wTempRow-CurrentRow >= SkipCount then
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
//  wTempRow := 0;
//  if PageLen > 0 then
//    wTempRow := FO_Row + PageLen - 1;
(*
{$IFDEF SM }
  FLastRow := WorkSheet.Rows.Count;
{$ENDIF SM }
*)
  wTempRow := FLastRow;
//  if PageLen > 0 then
//  if wTempRow < FLastRow then
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

function TbjXLSTable.FindField(const aName: string): pChar;
begin
  Result := nil;
  if GetFieldCol(aName) <>  -1 then
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
    FieldList.Add(PackStr(aList.Strings[ctr]));
    if ToWrite then
{$IFDEF SM }
      WorkSheet.Cells[FO_Row, FO_Column + ctr].Value := aList.Strings[ctr];
{$ENDIF SM }
{$IFDEF LIBXL }
      WorkSheet.WriteStr(FO_Row, FO_Column + ctr, pChar(aList.Strings[ctr]));
{$ENDIF LIBXL}
//      WorkSheet.Cells[FO_Row, FO_Column + ctr].Value := Decode(aList.Strings[ctr]);
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
// Check Packstr
{$IFDEF SM }
    if PackStr(WorkSheet.Cells[FO_Row, FO_Column+ctr].ValueAsString) =
      PackStr(FieldList.Strings[j]) then
{$ENDIF SM }
{$IFDEF LIBXL }
    if PackStr(WorkSheet.readStr(FO_Row, FO_Column+ctr)) =
      FieldList.Strings[j] then
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
  IDAlias: string;
  IDNode, IDAliasNode: IbjXml;
  aliasNode: IbjXml;
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
      if PackStr(WorkSheet.Cells[FO_Row, FO_Column+ctr].ValueAsString) =
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
    if Length(WorkSheet.Cells[FO_Row, FO_Column+ctr].ValueAsString) = 0 then
      Break;
    for j := 0 to FieldList.Count-1 do
    begin
// Check Packstr
      if PackStr(WorkSheet.Cells[FO_Row, FO_Column+ctr].ValueAsString) =
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
            IDList.Delete(j-1);
        end;
        Break;
      end;
    end;
  end;
end;

function TbjXLSTable.IsEmptyField(const aField: string; aRow: integer = -1): boolean;
var
//  v_vt: variant;
//  VType  : Integer;
  Value: Variant;
begin
{
  Result := False;
  v_vt := GetFieldVal(aField);
  VType := VarType(V_vt) and VarTypeMask;
  case VType of
  varEmpty:
    Result := True;
  varString:
    if Length(Trim(V_vt)) = 0 then
      Result := True;
  end;
}
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
{
  for ctr := 0 to FieldList.Count-1 do
  begin
    if not IsEmptyField(FieldList[ctr], aRow) then
    begin
      Result := False;
      Exit;
    end;
  end;
}
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


