procedure TForm1.Button3Click(Sender: TObject);

var
FileName1   : TDynamicString;
Rpt : TStringList;
i : integer;
Document       : IDocument;
OpenFile      : TOpenDialog;

uses
  StrUtils;
begin
Rpt := TStringList.Create;

   OpenFile := TOpenDialog.Create(Application);

   OpenFile.Title  := 'Import excel file containing classes_and_nets ';
   OpenFile.Filter := 'excel file (*.xls)|*.xls';

   if not OpenFile.Execute then exit;

   FileName1 := OpenFile.FileName;

listbox2.Items.LoadFromFile(FileName1);
end;

procedure TForm1.Button4Click(Sender: TObject);
begin
 listbox2.Clear;
end;

function Get_Iterator_Count(Iterator: IPCB_BoardIterator): Integer;
var
    cnt : Integer;
    cmp : IPCB_Component;
begin
  cnt := 0;

  cmp := Iterator.FirstPCBObject;
  While cmp <> NIL Do
  Begin
       Inc(cnt);
       cmp := Iterator.NextPCBObject;
  End;
  result := cnt;
end;

procedure TForm1.Button1Click(Sender: TObject);
Var
    Net         : IPCB_Net;
    Iterator    : IPCB_BoardIterator;
    i : integer;
    j : integer;
    check : String;
    needle : String;
    NetClass : IPCB_ObjectClass;
    Rpt2 : TStringList;
    FileName2   : TDynamicString;
    PCBObject     : IPCB_Object;
    Board       : IPCB_Board;
    ASetOfObjects : TObjectSet;

Begin
    // Retrieve the current board
    Board := PCBServer.GetCurrentPCBBoard;

    If Board = Nil Then Exit;
    Iterator := Board.BoardIterator_Create;
    Iterator.SetState_FilterAll;
    Iterator.AddFilter_ObjectSet(MkSet(eNetObject));
    Rpt2 := TStringList.Create;
    Rpt2 := listbox2.Items;
    PCBServer.PreProcess;
    ProgressBar1.Position := 0;
    for i := 0 to rpt2.count - 1 Do
       begin
       ProgressBar1.Position := i;
       ProgressBar1.Max := rpt2.count;
       ProgressBar1.Update;
       if AnsiContainsText(Rpt2[i], 'netclass :') then
       begin

       check := Rpt2[i];
       delete(check,1,10);
       Rpt2[i] := check;
       NetClass := PCBServer.PCBClassFactoryByClassMember(eClassMemberKind_Net);
       NetClass.SuperClass := False;
       Board.AddPCBObject(NetClass); // and add it to the PCB

       end;
       NetClass.Name := check;
       Net := Iterator.FirstPCBObject;
       While (Net <> Nil) Do
       begin
       if net.name = Rpt2[i] then
       begin
       NetClass.AddMemberByName(Net.Name);

       end;
       Net := Iterator.NextPCBObject;
       end;
    end;
    PCBServer.PostProcess;
    Board.BoardIterator_Destroy(Iterator);
    showmessage('Carpe diem');
    close;
end;









