unit UfrmGeradorDeRelatorio;

interface

uses
  ShareMem,
  mmSystem,
  MidasLib,
  ShellApi,
  Windows, Messages, SysUtils, Variants, Classes, Graphics,
  Controls, Forms, Dialogs, ComObj, ComCtrls,
  DBClient,
  DateUtils,
  StrUtils, FMTBcd, Provider, DB,
  SqlExpr, StdCtrls, ExtCtrls, Word2000, {DBXpress,} DBXMySQL;

type

   Titag = record
     indiceTag : integer;
     campoCorrespondente : widestring;
   end;

   Ttagcampo = record
     tag : widestring;
     campo : widestring;
   end;

  TfrmGeradorDeRelatorios = class(TForm)
    panelProgressBar: TPanel;
    progressbar: TProgressBar;
    panel1: TPanel;
    labelArqDoc: TLabel;
    editArqDoc: TEdit;
    labelArqDocSaveAs: TLabel;
    editArqDocSaveAs: TEdit;
    panel2: TPanel;
    btnGerarRelatorio: TButton;
    btnCancelar: TButton;
    dtsRelatorio: TSQLDataSet;
    dsRelatorio: TDataSource;
    cdsRelatorio: TClientDataSet;
    dspRelatorio: TDataSetProvider;
    Connection: TSQLConnection;
    lblProgressbar: TLabel;
    query: TSQLQuery;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormCreate(Sender: TObject);

  private
    { Private declarations }

    var WordApp, WordDoc : OleVariant;
        RelatorioEstaPreparado, RelatorioFoiGerado : boolean;
        docModelo, docNovo, ultimoRelatorioGerado : string;

        TituloRelatorio : string;
        ColunasFixas : integer;

        usuariologado, maquinausuariologado : string;

    procedure InicializaProgressBar(n:integer);
    function CriaArquivoBaseRelatorioLocal:boolean;

    procedure AjustaTamanhoFormulario;

    procedure SubstituiTagsParagrafo;
    procedure SubstituiTagsPrivadas;
    procedure SubstituiTagsCabecalho;

    function DataEHoraDoServidor:TDateTime;

  public
    { Public declarations }
    procedure ConexaoBD(drivername,hostname,database,username,password:string);
    procedure setInformacoesDoUsuario(pUsuarioLogado,pMaquinaUsuarioLogado:string);

    procedure SubstituiTodasOcorrencias(tag,txtnovo,tagparagrafo:widestring);
    procedure SubstituiTexto(tag,txtnovo,tagparagrafo:widestring);
    procedure InsereTexto(tag,txtnovo,tagparagrafo:widestring);
    function ExisteTag(tag:widestring):boolean;
    function SelecionaTabela(n:integer):boolean;
    procedure ProximaColuna;
    procedure EscreveTexto(str:widestring);
    procedure EscreveTextoDoCampo(n:integer);
    procedure SalvaRelatorio;
    procedure ProximoRegistro;
    procedure ExibePastaDoRelatorioGerado;
    function PreparaRelatorio(modelo,novodoc,nomerelat,sql:widestring):boolean;
    procedure CancelaRelatorio;

    function Relatorio(modelo,relatorio,nomerelat,sql:widestring):boolean; overload;
    function Relatorio(modelo,relatorio,nomerelat,sql:widestring;indiceTabela:integer):boolean; overload;
    function Relatorio(modelo,relatorio,nomerelat,sql:widestring;indiceTabela:integer;tags:TStringList):boolean; overload;
    function Relatorio(modelo,relatorio,nomerelat,sql:widestring;tags:TStringList):boolean; overload;

    procedure setColunasFixas(n:integer);

  end;


var
  frmGeradorDeRelatorios: TfrmGeradorDeRelatorios;

implementation

{$R *.dfm}

procedure TfrmGeradorDeRelatorios.setInformacoesDoUsuario(pUsuarioLogado,pMaquinaUsuarioLogado:string);
begin
  usuariologado := pUsuarioLogado;
  maquinausuariologado := pMaquinaUsuarioLogado;
end;

function TfrmGeradorDeRelatorios.DataEHoraDoServidor:TDateTime;
begin
  query.Close;
  query.SQL.Clear;
  query.SQL.Add(' select now() ');
  query.Open;
  Result := query.Fields.Fields[0].Value;
end;

procedure TfrmGeradorDeRelatorios.ConexaoBD(drivername,hostname,database,username,password:string);
begin
  Connection.Close;
  Connection.DriverName := drivername;
  Connection.Params.Values['DriverName'] := drivername;
  Connection.Params.Values['HostName'] := hostname;
  Connection.Params.Values['Database'] := database;
  Connection.Params.Values['User_Name'] := username;
  Connection.Params.Values['Password'] := password;
  Connection.Open;
end;


procedure TfrmGeradorDeRelatorios.AjustaTamanhoFormulario;
var
  size1, size2, dif : integer;
begin
  if ( editArqDoc.Width > editArqDocSaveAs.Width ) then dif := frmGeradorDeRelatorios.Width - editArqDoc.Width
  else dif := frmGeradorDeRelatorios.Width - editArqDocSaveAs.Width;

  size1 := Canvas.TextWidth(editArqDoc.Text)+
           editArqDoc.Padding.Left+
           editArqDoc.Padding.Right+
           60;

  size2 := Canvas.TextWidth(editArqDocSaveAs.Text)+
           editArqDocSaveAs.Padding.Left+
           editArqDocSaveAs.Padding.Right+
           60;

  if ( size1 < size2 ) then size1 := size2;
  if ( size1 < labelArqDoc.Width ) then size1 := labelArqDoc.Width;
  if ( size1 < labelArqDocSaveAs.Width ) then size1 := labelArqDocSaveAs.Width;

  frmGeradorDeRelatorios.Width :=
    size1 +
    dif;
end;

procedure TfrmGeradorDeRelatorios.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  docModelo := '';
  docNovo := '';
  TituloRelatorio := '';
  ColunasFixas := 0;

  editArqDoc.Text := '...';
  editArqDocSaveAs.Text := '...';
  RelatorioEstaPreparado := False;
end;

procedure TfrmGeradorDeRelatorios.FormCreate(Sender: TObject);
begin
  usuariologado := 'login';
  maquinausuariologado := 'maquina_cliente';

  docModelo := '';
  docNovo := '';
  TituloRelatorio := '';
  ColunasFixas := 0;

  ultimoRelatorioGerado := '';
  RelatorioEstaPreparado := False;
end;

procedure TfrmGeradorDeRelatorios.FormResize(Sender: TObject);
begin
  btnGerarRelatorio.Margins.Left := Trunc((panel2.Width - btnGerarRelatorio.Width - btnCancelar.Width - btnCancelar.Margins.Left) / 2);
end;

procedure TfrmGeradorDeRelatorios.FormShow(Sender: TObject);
begin
  AjustaTamanhoFormulario;
end;



procedure TfrmGeradorDeRelatorios.InicializaProgressBar(n:integer);
begin
  progressbar.Min := 0;
  progressbar.Max := n;
  progressbar.Step := 1;
  progressbar.Position := 0;
  lblProgressbar.Caption := '0 %';
end;

function TfrmGeradorDeRelatorios.CriaArquivoBaseRelatorioLocal:boolean;
begin
  try

    try
      // Pega uma instância ativa do Word
      WordApp := GetActiveOleObject('Word.Application');
    except
      // Cria objeto principal de controle
      WordApp := CreateOleObject('Word.Application');
    end;

  except
    ShowMessage('MS Word não instalado!'#10#13'Impossível gerar relatório neste modo.');
    Exit;
  end;

  WordApp.DisplayAlerts := 0;
  WordApp.Visible:= False;

  WordDoc := WordApp.Documents.Open(docModelo);

  if ( FileExists( docNovo ) ) then DeleteFile( docNovo );

  WordDoc.SaveAs( docNovo );

  Result := True;
end;

procedure TfrmGeradorDeRelatorios.SubstituiTagsParagrafo;
begin
  SubstituiTodasOcorrencias('<enter()>', '^p', '');
end;

procedure TfrmGeradorDeRelatorios.SubstituiTagsPrivadas;
begin
  WordDoc.Sections.Item(1).Headers.Item(1).Range.Find.Execute('<hoje()>', ReplaceWith := FormatDateTime('dd/mm/yyyy', DataEHoraDoServidor ) );
  WordDoc.Sections.Item(1).Headers.Item(1).Range.Find.Execute('<recordcount()>', ReplaceWith := IntToStr( cdsRelatorio.RecordCount ) );
  WordDoc.Sections.Item(1).Headers.Item(1).Range.Find.Execute('<numeroregistros()>', ReplaceWith := IntToStr( cdsRelatorio.RecordCount ) );

  WordDoc.Sections.Item(1).Headers.Item(1).Range.Find.Execute('<data()>', ReplaceWith := FormatDateTime('dd/mm/yyyy', DataEHoraDoServidor ) );
  WordDoc.Sections.Item(1).Headers.Item(1).Range.Find.Execute('<hora()>', ReplaceWith := FormatDateTime('hh:nn', DataEHoraDoServidor ) );
  WordDoc.Sections.Item(1).Headers.Item(1).Range.Find.Execute('<login()>', ReplaceWith := usuariologado );
  WordDoc.Sections.Item(1).Headers.Item(1).Range.Find.Execute('<clientmachine()>', ReplaceWith := maquinausuariologado );
  WordDoc.Sections.Item(1).Headers.Item(1).Range.Find.Execute('<servidor()>', ReplaceWith := Connection.Params.Values['HostName'] );
  WordDoc.Sections.Item(1).Headers.Item(1).Range.Find.Execute('<titulorelatorio()>', ReplaceWith := TituloRelatorio );
end;

procedure TfrmGeradorDeRelatorios.SubstituiTagsCabecalho;
begin
  //# <data> # <hora> #
  //# <login>@<clientmachine> # - (<servidor>)
  // <titulorelatorio>

  if ( ExisteTag('<data>') ) then
    WordDoc.Sections.Item(1).Headers.Item(1).Shapes.Item(1).TextFrame.TextRange.Find.Execute('<data>', ReplaceWith := FormatDateTime('dd/mm/yyyy', DataEHoraDoServidor ) );
  if ( ExisteTag('<hora>') ) then
    WordDoc.Sections.Item(1).Headers.Item(1).Shapes.Item(1).TextFrame.TextRange.Find.Execute('<hora>', ReplaceWith := FormatDateTime('hh:nn', DataEHoraDoServidor ) );
  if ( ExisteTag('<login>') ) then
    WordDoc.Sections.Item(1).Headers.Item(1).Shapes.Item(2).TextFrame.TextRange.Find.Execute('<login>', ReplaceWith := usuariologado );
  if ( ExisteTag('<clientmachine>') ) then
    WordDoc.Sections.Item(1).Headers.Item(1).Shapes.Item(2).TextFrame.TextRange.Find.Execute('<clientmachine>', ReplaceWith := maquinausuariologado );
  if ( ExisteTag('<servidor>') ) then
    WordDoc.Sections.Item(1).Headers.Item(1).Shapes.Item(2).TextFrame.TextRange.Find.Execute('<servidor>', ReplaceWith := Connection.Params.Values['HostName'] );
  if ( ExisteTag('<titulorelatorio>') ) then
    WordDoc.Sections.Item(1).Headers.Item(1).Shapes.Item(3).TextFrame.TextRange.Find.Execute('<titulorelatorio>', ReplaceWith := TituloRelatorio );
  if ( ExisteTag('<record_count>') ) then
    WordDoc.Sections.Item(1).Headers.Item(1).Shapes.Item(3).TextFrame.TextRange.Find.Execute('<record_count>', ReplaceWith := IntToStr(cdsRelatorio.RecordCount) );
end;

function TfrmGeradorDeRelatorios.PreparaRelatorio(modelo,novodoc,nomerelat,sql:widestring):boolean;
begin
  Result := False;
  RelatorioEstaPreparado := False;
  RelatorioFoiGerado := False;

  docModelo := modelo;
  docNovo := novodoc;
  TituloRelatorio := nomerelat;

  editArqDoc.Enabled := True;
  editArqDocSaveAs.Enabled := True;

  editArqDoc.Text := docModelo;
  editArqDocSaveAs.Text := docNovo;

  AjustaTamanhoFormulario;

  cdsRelatorio.Close;
  dtsRelatorio.Close;
  dtsRelatorio.CommandText := sql;
  cdsRelatorio.Open;

  if ( cdsRelatorio.RecordCount > 0 ) then
  begin

    if ( (docModelo <> '') and (docNovo <> '') and (TituloRelatorio <> '') ) then
    begin

      if ( CriaArquivoBaseRelatorioLocal ) then
      begin
        InicializaProgressBar(cdsRelatorio.RecordCount);

        SubstituiTagsPrivadas;

        SubstituiTagsCabecalho;

        RelatorioEstaPreparado := True;
        Result := True;

        editArqDoc.Enabled := False;
        editArqDocSaveAs.Enabled := False;

      end
      else
      begin
        ShowMessage('Erro na criação do Relatório!'#13#13'Verifique suas permissões de usuário!');
      end;

    end
    else
    begin
      ShowMessage('Ao preparar o Relatório indique :'#13#13+
                  ' - o modelo a ser usado;'#13#13+
                  ' - o nome do relatório a ser gerado;'#13#13+
                  ' - o nome do Relatório; e'#13#13+
                  ' - o comando "sql" a ser consultado');
    end;

  end
  else
  begin
    ShowMessage('Nenhum registro encontrado!');
    cdsRelatorio.Close;
  end;
end;


procedure TfrmGeradorDeRelatorios.CancelaRelatorio;
begin
  if ( RelatorioEstaPreparado ) then
  begin
    cdsRelatorio.Close;
    WordApp.ActiveDocument.Save;
    WordApp.Quit(False);
    RelatorioFoiGerado := False;
    RelatorioEstaPreparado := False;
    if ( FileExists( docNovo ) ) then DeleteFile( docNovo );
    if ( frmGeradorDeRelatorios.Visible ) then frmGeradorDeRelatorios.Close;
  end;
end;


procedure TfrmGeradorDeRelatorios.SubstituiTodasOcorrencias(tag,txtnovo,tagparagrafo:widestring);
begin
  if ( RelatorioEstaPreparado ) then
  begin
    while ( ExisteTag(tag) ) do SubstituiTexto(tag, txtnovo, tagparagrafo);
  end
  else
  begin
    ShowMessage('Relatório não preparado...');
  end;
end;

procedure TfrmGeradorDeRelatorios.SubstituiTexto(tag,txtnovo,tagparagrafo:widestring);
begin
  if ( RelatorioEstaPreparado ) then
  begin
    while ( Length(txtnovo) > 0 ) do
    begin
      if ( Length(txtnovo) > 100 ) then
      begin
        WordDoc.Content.Find.Execute(tag, ReplaceWith := LeftStr(txtnovo,100) + tag ) ;
        txtnovo := RightStr( txtnovo, Length(txtnovo)-100 );
      end
      else
      begin
        WordDoc.Content.Find.Execute(tag, ReplaceWith := txtnovo + tagparagrafo ) ;
        txtnovo := '';
      end;
    end;
  end
  else
  begin
    ShowMessage('Relatório não preparado...');
  end;
end;

procedure TfrmGeradorDeRelatorios.InsereTexto(tag,txtnovo,tagparagrafo:widestring);
begin
  if ( RelatorioEstaPreparado ) then
  begin
    while ( Length(txtnovo) > 0 ) do
    begin
      if ( Length(txtnovo) > 100 ) then
      begin
        WordDoc.Content.Find.Execute(tag, ReplaceWith := LeftStr(txtnovo,100) + tag ) ;
        txtnovo := RightStr( txtnovo, Length(txtnovo)-100 );
      end
      else
      begin
        WordDoc.Content.Find.Execute(tag, ReplaceWith := txtnovo + tagparagrafo + tag ) ;
        txtnovo := '';
      end;
    end;
  end
  else
  begin
    ShowMessage('Relatório não preparado...');
  end;
end;

function TfrmGeradorDeRelatorios.ExisteTag(tag:widestring):boolean;
begin
  Result := False;
  if ( LeftStr(tag,1) <> '<' ) then tag := '<' + tag;
  if ( RightStr(tag,1) <> '>' ) then tag := tag + '>';
  if ( RelatorioEstaPreparado ) then Result := WordDoc.Content.Find.Execute(tag);
end;

function TfrmGeradorDeRelatorios.SelecionaTabela(n:integer):boolean;
begin
  if ( RelatorioEstaPreparado ) then
  begin
    if ( n <= 0 ) then n := 1;
    WordApp.Selection.Tables.Item(n).Select;
    Result := True;
  end
  else
  begin
    ShowMessage('Relatório não preparado...');
  end;
end;

procedure TfrmGeradorDeRelatorios.ProximaColuna;
begin
  if ( RelatorioEstaPreparado ) then
  begin
    WordApp.Selection.MoveRight(12);
  end
  else
  begin
    ShowMessage('Relatório não preparado...');
  end;
end;

procedure TfrmGeradorDeRelatorios.EscreveTexto(str:widestring);
begin
  if ( RelatorioEstaPreparado ) then
  begin
    WordApp.Selection.TypeText(Text := str);
  end
  else
  begin
    ShowMessage('Relatório não preparado...');
  end;
end;

procedure TfrmGeradorDeRelatorios.EscreveTextoDoCampo(n:integer);
begin
  if ( RelatorioEstaPreparado ) then
  begin
    WordApp.Selection.TypeText(Text := cdsRelatorio.Fields.Fields[n].AsString );
  end
  else
  begin
    ShowMessage('Relatório não preparado...');
  end;
end;

procedure TfrmGeradorDeRelatorios.SalvaRelatorio;
begin
  if ( RelatorioEstaPreparado ) then
  begin
    SubstituiTagsParagrafo;

    cdsRelatorio.Close;
    WordApp.ActiveDocument.Save;
    WordApp.Quit(False);

    ultimoRelatorioGerado := docNovo;
    RelatorioFoiGerado := True;
    frmGeradorDeRelatorios.Close;
  end
  else
  begin
    ShowMessage('Relatório não preparado...');
  end;
end;

procedure TfrmGeradorDeRelatorios.ProximoRegistro;
begin
  if ( RelatorioEstaPreparado ) then
  begin
    cdsRelatorio.Next;
    ProgressBar.StepBy(1);
    lblProgressbar.Caption := IntToStr( Trunc( cdsRelatorio.RecNo / cdsRelatorio.RecordCount * 100 ) ) + ' %';
    Application.ProcessMessages;
  end
  else
  begin
    ShowMessage('Relatório não preparado...');
  end;
end;

procedure TfrmGeradorDeRelatorios.ExibePastaDoRelatorioGerado;
begin
  if ( RelatorioFoiGerado ) then
  begin
    if ( ultimoRelatorioGerado <> '' ) then
      ShellExecute(Application.Handle, PChar('open'), PChar('explorer.exe'), PChar('/select, '+ ultimoRelatorioGerado ), nil,SW_NORMAL);
  end
  else
  begin
    ShowMessage('Relatório não foi gerado...');
  end;
end;



function TfrmGeradorDeRelatorios.Relatorio(modelo,relatorio,nomerelat,sql:widestring;indiceTabela:integer):boolean;
var quantCampos, quantColunas : integer;

  procedure GeraConteudoDaTabela;
  var i : integer;
  begin

    while ( not cdsRelatorio.Eof ) do
    begin

      for i := 0 to quantCampos - 1 do
      begin
        ProximaColuna;
        if ( i > (ColunasFixas - 1) ) then EscreveTextoDoCampo( i - ColunasFixas );
      end;
      ProximoRegistro;

    end;

  end;

begin
  Result := False;

  if ( indiceTabela <> 0 ) then //Result := Relatorio(modelo,relatorio,nomerelat,sql)

    if ( PreparaRelatorio(modelo,relatorio,nomerelat,sql) ) then
    begin
      quantCampos := cdsRelatorio.Fields.Count;

      if ( SelecionaTabela(indiceTabela) ) then
      begin

        quantColunas := WordApp.Selection.Tables.Item(indiceTabela).Columns.Count;

        if ( quantCampos = ( quantColunas - ColunasFixas ) ) then
        begin
          frmGeradorDeRelatorios.Show;

          GeraConteudoDaTabela;

          SalvaRelatorio;

          ExibePastaDoRelatorioGerado;

          Result := True;
        end
        else
        begin
          ShowMessage('O modelo indicado possui uma tabela com [' + IntToStr(quantColunas) + '] colunas;'+#10#13+
                      'enquanto a seleção de registros apresenta [' + IntToStr(quantCampos) + '] campos.'+#10#13+#10#13+
                      'Peça ao desenvolvedor do Sistema para verificar'+#10#13+
                      'o "Select" e o modelo de relatório indicado!' );
          Result := False;
          frmGeradorDeRelatorios.Close;;
        end;

      end;

    end
    else Result := False;
end;

function TfrmGeradorDeRelatorios.Relatorio(modelo,relatorio,nomerelat,sql:widestring;indiceTabela:integer;tags:TStringList):boolean;
var i, quantCampos : integer;
    tagcampo : array of Ttagcampo;

  procedure CopiaLinhaTabela;
  begin
    WordApp.Selection.Tables.Item(indiceTabela).Cell(
      WordApp.Selection.Tables.Item(indiceTabela).Rows.Count, 1 ).Select;
    WordApp.Selection.MoveDown( Unit:=wdLine, Count:=1, Extend:=wdExtend );
    WordApp.Selection.MoveLeft( Unit:=wdCharacter, Count:=2, Extend:=wdExtend );
    WordApp.Selection.Copy;
    WordApp.Selection.MoveDown( Unit:=wdLine, Count:=1 );
  end;

  procedure ColaLinhaTabela;
  begin
    WordApp.Selection.Tables.Item(indiceTabela).Cell(
      WordApp.Selection.Tables.Item(indiceTabela).Rows.Count,
      WordApp.Selection.Tables.Item(indiceTabela).Columns.Count).Select;
    ProximaColuna;
    WordApp.Selection.Paste;
    WordApp.Selection.MoveDown( Unit:=wdLine, Count:=1 );
  end;

  procedure GeraConteudoDaTabelaComTags;
  var i : integer;
      tag, conteudoCampo : widestring;
  begin
    SelecionaTabela(indiceTabela);
    CopiaLinhaTabela;

    while ( not cdsRelatorio.Eof ) do
    begin

      for i := 0 to quantCampos - 1 do
      begin
        tag := tagcampo[i].tag;
        conteudoCampo := cdsRelatorio.FieldByName( tagcampo[i].campo ).AsWideString; 
        if ( conteudoCampo = '' ) then conteudoCampo := ' '
        else conteudoCampo := AnsiReplaceStr(conteudoCampo,#10,#13);
        SubstituiTexto( tag, conteudoCampo, '' );
      end;

      if ( cdsRelatorio.RecNo < cdsRelatorio.RecordCount ) then ColaLinhaTabela;

      ProximoRegistro;

    end;

  end;

  procedure MontaListaTagCampo;
  var i, indiceTag : integer;
      correspondem : boolean;

      procedure montaListaTagCampoComTags;
      var i : integer;
      begin
        for i := 0 to quantCampos - 1 do
        begin
          tagcampo[i].tag := tags.ValueFromIndex[i];
          tagcampo[i].campo := cdsRelatorio.Fields.Fields[i].FieldName;
        end;
      end;

  begin
    if ( tags.Count = quantCampos ) then
    begin
      SetLength(tagcampo,tags.Count);

      tags.Sorted := True;
      tags.Sort;

      correspondem := True;
      for i := 0 to quantCampos - 1 do
      begin
        indiceTag := tags.IndexOf( cdsRelatorio.Fields.Fields[i].FieldName );

        correspondem := correspondem and ( indiceTag >= 0 );

        if ( indiceTag >= 0 ) then
        begin
          tagcampo[i].tag := '<'+cdsRelatorio.Fields.Fields[i].FieldName+'>';
          tagcampo[i].campo := cdsRelatorio.Fields.Fields[i].FieldName;
        end;

      end;

      if ( not correspondem ) then montaListaTagCampoComTags;
    end;
  end;

begin
  Result := False;

  for i := 0 to tags.Count - 1 do tags.ValueFromIndex[i] := '<' + tags.ValueFromIndex[i];

  if ( indiceTabela <> 0 ) then
  begin

    if ( PreparaRelatorio(modelo,relatorio,nomerelat,sql) ) then
    begin
      quantCampos := cdsRelatorio.Fields.Count;

      if ( SelecionaTabela(indiceTabela) ) then
      begin
        frmGeradorDeRelatorios.Show;

        MontaListaTagCampo;

        GeraConteudoDaTabelaComTags;

        SalvaRelatorio;

        ExibePastaDoRelatorioGerado;

        Result := True;

        frmGeradorDeRelatorios.Close;;
      end
      else Result := False;

    end
    else Result := False;

  end
  else Result := False;

end;

function TfrmGeradorDeRelatorios.Relatorio(modelo,relatorio,nomerelat,sql:widestring):boolean;
var quantCampos, quantColunas, quantRegistros : integer;
    tag : string;
    conteudoCampo : widestring;

  function CamposCorrespondem : boolean;
  var i : integer;
  begin
    Result := True;
    for i := 0 to cdsRelatorio.Fields.Count - 1 do
      Result := Result and ( ExisteTag(cdsRelatorio.Fields.Fields[i].FieldName) );
  end;


  procedure GeraRelatorioComConteudoDosCampos;
  var  campoi : integer;
  begin
    quantRegistros := cdsRelatorio.RecordCount;

    cdsRelatorio.First;

    while ( not cdsRelatorio.Eof ) do
    begin

      for campoi := 0 to quantCampos - 1 do
      begin
        tag := '<' + cdsRelatorio.Fields.Fields[campoi].FieldName + '>';
        conteudoCampo := cdsRelatorio.Fields.Fields[campoi].AsWideString;
        conteudoCampo := AnsiReplaceStr(conteudoCampo,#10,#13);
        SubstituiTexto( tag, conteudoCampo, '' );
      end;

      if ( cdsRelatorio.RecNo < cdsRelatorio.RecordCount ) then
      begin
        WordApp.Selection.EndKey(wdStory);
        WordApp.Selection.InsertBreak;
        WordApp.Selection.InsertFile(modelo);
      end;

      ProximoRegistro;

    end;

  end;

begin
  Result := False;

  if ( PreparaRelatorio(modelo,relatorio,nomerelat,sql) ) then
  begin
    quantCampos := cdsRelatorio.Fields.Count;

    if ( CamposCorrespondem ) then
    begin
      frmGeradorDeRelatorios.Show;
      GeraRelatorioComConteudoDosCampos;
      SalvaRelatorio;
      ExibePastaDoRelatorioGerado;
      Result := True;
    end
    else
    begin
      ShowMessage('Os campos da pesquisa não correspondem às tag''s encontradas no modelo do documento.');
      Result := False;
      frmGeradorDeRelatorios.Close;;
    end;
  end
  else Result := False;

end;

function TfrmGeradorDeRelatorios.Relatorio(modelo,relatorio,nomerelat,sql:widestring;tags:TStringList):boolean;
var quantCampos, quantColunas, quantRegistros : integer;
    tag : string;
    conteudoCampo : widestring;
    itag : array of Titag;

  function CamposCorrespondem : boolean;
  var i, indiceTag : integer;
  begin
    if ( tags.Count = quantCampos ) then
    begin
      Result := True;

      SetLength(itag,tags.Count);

      for i := 0 to quantCampos - 1 do
      begin
        Result := Result and ( tags.Find( cdsRelatorio.Fields.Fields[i].FieldName, indiceTag ) );
        itag[i].indiceTag := indiceTag;
        itag[i].campoCorrespondente := cdsRelatorio.Fields.Fields[i].FieldName;
      end;
    end
    else Result := False;
  end;


  procedure GeraRelatorioComConteudoDosCampos;
  var  campoi : integer;
  begin
    quantRegistros := cdsRelatorio.RecordCount;

    cdsRelatorio.First;

    while ( not cdsRelatorio.Eof ) do
    begin

      for campoi := 0 to quantCampos - 1 do
      begin
        tag := '<' + itag[campoi].campoCorrespondente + '>';
        conteudoCampo := cdsRelatorio.Fields.FieldByName(itag[campoi].campoCorrespondente).AsWideString;
        conteudoCampo := AnsiReplaceStr(conteudoCampo,#10,#13);
        if ( ExisteTag(tag) ) then
          SubstituiTexto( tag, conteudoCampo, '' );
      end;

      if ( cdsRelatorio.RecNo < cdsRelatorio.RecordCount ) then
      begin
        WordApp.Selection.EndKey(wdStory);
        WordApp.Selection.InsertBreak;
        WordApp.Selection.InsertFile(modelo);
      end;

      ProximoRegistro;

    end;

  end;

begin
  Result := False;

  if ( PreparaRelatorio(modelo,relatorio,nomerelat,sql) ) then
  begin
    quantCampos := cdsRelatorio.Fields.Count;

    if ( CamposCorrespondem ) then
    begin
      frmGeradorDeRelatorios.Show;
      GeraRelatorioComConteudoDosCampos;
      SalvaRelatorio;
      ExibePastaDoRelatorioGerado;
      Result := True;
    end
    else
    begin
      ShowMessage('Os campos da pesquisa não correspondem às tag''s encontradas no modelo do documento.');
      Result := False;
      frmGeradorDeRelatorios.Close;;
    end;

    FreeMemory(itag);
  end
  else Result := False;

end;

procedure TfrmGeradorDeRelatorios.setColunasFixas(n:integer);
begin
  ColunasFixas := n;
end;

end.
