unit DelphiDevDataSetTools;

interface

uses Classes, DB, IdGlobal;

{procedure CopyDataSet:
 Copies data from dataset ASourceDataSet to dataset ADestDataSet.
 Excludes fields defined in string AExcludeFields. The field names should be
 comma seperated}
procedure CopyDataSet(ASourceDataSet, ADestDataSet: TDataSet;
  AExcludedFields: String = '');

{procedure CopyRecord:
 Copies the current record from dataset ASourceDataSet to dataset ADestDataSet.
 Excludes fields defined in string AExcludeFields. The field names should be
 comma seperated}
procedure CopyRecord(ASourceDataSet, ADestDataSet: TDataSet;
  AExcludedFields: String = '');

{function FieldHasValue
 Returns True if the field of ADataSet specified in AFieldName has the same
 value in all dataset rows}
function FieldHasValue(ADataSet: TDataSet; AFieldName: String;
  AFieldValue: Variant): Boolean;

{procedure CopyValueToDataSet
 Copies AFieldValue to all the rows of ADataSet at field AFieldName}
procedure CopyValueToDataSet(ADataSet: TDataSet; AFieldName: String; AFieldValue: Variant);

function CheckFldIfEmpty(ADataSet: TDataSet; AFieldName: String): Boolean;

implementation

procedure CopyDataSet(ASourceDataSet, ADestDataSet: TDataSet;
  AExcludedFields: String = '');
var
  Fld: TField;
  i: Integer;
  ExcludedFieldsStrList: TStringList;
  bmrk: TBookmark;
begin
  if not Assigned(ASourceDataSet) or not ASourceDataSet.Active then exit;
  if not Assigned(ADestDataSet) or not ADestDataSet.Active then exit;
  ExcludedFieldsStrList := TStringList.Create;
  ExcludedFieldsStrList.CommaText := AExcludedFields;
  try
    with ASourceDataSet do begin
      bmrk := GetBookmark;
      DisableControls;
      try
        First;
        while not eof do begin
          ADestDataSet.Last;
          ADestDataSet.Append;
          for i := 0 to FieldCount - 1 do begin
            if ExcludedFieldsStrList.IndexOf(Fields[i].FieldName) > 0 then begin
              Fld := ADestDataSet.FindField(Fields[i].FieldName);
              if Assigned(Fld) then
                Fld.AsVariant := Fields[i].AsVariant;
            end;
          end;
          Next;
        end;
      finally
        GotoBookmark(bmrk);
        EnableControls;
        FreeAndNil(bmrk);
      end;
    end;
  finally
    ExcludedFieldsStrList.Free;
  end;
end;

procedure CopyRecord(ASourceDataSet, ADestDataSet: TDataSet;
  AExcludedFields: String = '');
var
  Fld: TField;
  i: Integer;
  ExcludedFieldsStrList: TStringList;
begin
  if not Assigned(ASourceDataSet) or not ASourceDataSet.Active then exit;
  if not Assigned(ADestDataSet) or not ADestDataSet.Active then exit;
  ExcludedFieldsStrList := TStringList.Create;
  ExcludedFieldsStrList.CommaText := AExcludedFields;
  try
    for i := 0 to ASourceDataSet.FieldCount - 1 do begin
      if ExcludedFieldsStrList.IndexOf(ASourceDataSet.Fields[i].FieldName) > 0 then begin
        Fld := ADestDataSet.FindField(ASourceDataSet.Fields[i].FieldName);
        if Assigned(Fld) then
          Fld.AsVariant := ASourceDataSet.Fields[i].AsVariant;
      end;
    end;
  finally
    ExcludedFieldsStrList.Free;
  end;
end;

function FieldHasValue(ADataSet: TDataSet; AFieldName: String;
  AFieldValue: Variant): Boolean;
var
  Fld: TField;
  bmrk: TBookmark;
begin
  Result := False;
  if not Assigned(ADataSet) or not ADataSet.Active then Exit;
  Fld := ADataSet.FindField(AFieldName);
  if not Assigned(Fld) then Exit;
  with ADataSet do begin
    bmrk := GetBookmark;
    Result := True;
    try
      DisableControls;
      First;
      while not eof do begin
        if Fld.AsVariant <> AFieldValue then begin
          Result := False;
          Exit;
        end;
        Next;
      end;
    finally
      GotoBookmark(bmrk);
      EnableControls;
      FreeAndNil(bmrk);
    end;
  end;
end;

procedure CopyValueToDataSet(ADataSet: TDataSet; AFieldName: String;
  AFieldValue: Variant);
var
  Fld: TField;
begin
  if not Assigned(ADataSet) or not ADataSet.Active then Exit;
  Fld := ADataSet.FindField(AFieldName);
  if not Assigned(Fld) then Exit;
  with ADataSet do begin
    First;
    while not eof do begin
      Edit;
      Fld.AsVariant := AFieldValue; 
      Next;
    end;
  end;
end;

function CheckFldIfEmpty(ADataSet: TDataSet; AFieldName: String): Boolean;
var
  Fld: TField;
begin
  Result := False;
  if not Assigned(ADataSet) or not ADataSet.Active then Exit;
  Fld := ADataSet.FindField(AFieldName);
  if not Assigned(Fld) then Exit;
  with ADataSet do begin
    try
      Result := True;
      DisableControls;
      First;
      while not eof do begin
        if not Fld.IsNull then begin
          Result := False;
          exit;
        end;
        Next;
      end;
    finally
      EnableControls;
    end;
  end;
end;

end.
