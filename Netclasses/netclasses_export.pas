
Var
   Board    : IPCB_Board;
   Iterator : IPCB_BoardIterator;
   NetClass : IPCB_ObjectClass;
   FileName   : String;
   NetInfo    : TStringList;

Procedure GenerateReport(Var ARpt : TStringList; AObjectClass : IPCB_ObjectClass);
Var
    S : TPCBString;
    I : Integer;
    uses FileCtrl;
Begin
        I := 0;
        While AObjectClass.MemberName[I] <> '' Do
        Begin
            ARpt.Add(AObjectClass.MemberName[I]);
            Inc(I);
        End;
    End;

procedure TForm1.Button1Click(Sender: TObject);
var
  I : Integer;
Begin
   Board := PCBServer.GetCurrentPCBBoard;
   if Board = nil then exit;
   Iterator  := Board.BoardIterator_Create;
   Iterator.SetState_FilterAll;
   Iterator.AddFilter_ObjectSet(MkSet(eClassObject));
   NetClass := Iterator.FirstPCBObject;
   i := 1;
   While NetClass <> NIl Do
   Begin
       If ((NetClass.MemberKind = eClassMemberKind_Net) and (NetClass.Name <> 'All Nets')) Then ListBox1.Items.AddObject(NetClass.Name, NetClass);
          NetClass := Iterator.NextPCBObject;
   End;
   Board.BoardIterator_Destroy(Iterator);
end;

procedure TForm1.ListBox1DblClick(Sender: TObject);
Var
    ClassRpt : TStringList;
    I        : Integer;
    FileName : String;
    Document : IServerDocument;
begin
  Begin
    If Form1.ListBox1.Items.Count < 1 Then Exit;
    ClassRpt := TStringList.Create;
    For i := 0 To Form1.ListBox1.Items.Count - 1 Do
        If Form1.ListBox1.Selected[i] Then
            GenerateReport(ClassRpt, Form1.Listbox1.Items.Objects[i]);
    // Display the object class and its members report
    FileName := ChangeFileExt(Board.FileName,'.REP');
    ClassRpt.SaveToFile(Filename);
    ClassRpt.Free;
    Document  := Client.OpenDocument('Text', FileName);
    If Document <> Nil Then
        Client.ShowDocument(Document);
End;
end;



procedure TForm1.Button2Click(Sender: TObject);
uses VCL.FileCtrl;
var
FileName   : TDynamicString;
i        : Integer;
j  : Integer;
Rpt : TStringList;
ObjectC : IPCB_ObjectClass;
chosenDirectory : string;
Document       : IDocument;
SaveFile   : TSaveDialog;

begin
If Form1.ListBox1.Items.Count < 1 Then showmessage('populate list first') ;
Rpt := TStringList.Create;

For i := 0 To Form1.ListBox1.Items.Count - 1 Do
begin
           ObjectC := Form1.Listbox1.Items.Objects[i];
           Rpt.AddObject('netclass :' + ObjectC.Name + '_new', i);
           j := 0;
        While (ObjectC.MemberName[J] <> '')  Do
        Begin
            Rpt.AddObject(ObjectC.MemberName[J], ObjectC);
                        Inc(J);
        End;
end;

if Form1.ListBox1.Items.Count > 0 Then
begin
SaveFile := TSaveDialog.Create(Application);
SaveFile.Title  := 'Export to location';
SaveFile.Filter := 'excel file (*.xls)|*.xls';
SaveFile.DefaultExt := 'xls';

if not SaveFile.Execute then exit;

Filename := SaveFile.FileName;

Rpt.SaveToFile(FileName);

ShowMessage('saved to ' + SaveFile.FileName);
end;
close;
end;
























