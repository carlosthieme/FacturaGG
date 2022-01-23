unit FacMain;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ComCtrls, Buttons, ExtCtrls, Grids, AdvGrid, vgCtrls, Num2Word,
  Printers, WinSpool, IniFiles, Db, DBTables;

type
  TFacMainForm = class(TForm)
    SelFecha: TDateTimePicker;
    Label1: TLabel;
    Label2: TLabel;
    Edit1: TEdit;
    Label3: TLabel;
    Edit4: TEdit;
    Edit5: TEdit;
    Edit6: TEdit;
    Edit14: TEdit;
    Edit9: TEdit;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Grid: TAdvStringGrid;
    btnCerrar: TSpeedButton;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Impresoras: TComboBox;
    Clientes: TTable;
    DataSource1: TDataSource;
    btnImprimir: TSpeedButton;
    Factura: TTable;
    DataSource2: TDataSource;
    Detalle: TTable;
    DataSource3: TDataSource;
    Edit10: TJustifyEdit;
    Edit11: TJustifyEdit;
    Edit12: TJustifyEdit;
    Edit13: TJustifyEdit;
    Edit7: TJustifyEdit;
    Edit2: TJustifyEdit;
    Edit3: TJustifyEdit;
    Edit8: TJustifyEdit;
    Check: TCheckBox;
    Edit15: TJustifyEdit;
    procedure btnCerrarClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure GridGetAlignment(Sender: TObject; ARow, ACol: Integer; var AAlignment: TAlignment);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure btnImprimirClick(Sender: TObject);
    procedure GridKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Edit3Exit(Sender: TObject);
    procedure GridExit(Sender: TObject);
    procedure CheckClick(Sender: TObject);
    procedure Edit15Exit(Sender: TObject);
  private
    Function FechaLarga(Recibe : String) : String;
    Procedure GuardarDatos;
    Procedure ImprimeFactura;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FacMainForm: TFacMainForm;
  ArchivoCfg : String;
  SetupFiles, DataFiles, BackupFiles : String;

implementation


{$R *.DFM}

Function TFacMainForm.FechaLarga(Recibe : String) : String;
Var
  Mes : Integer;
  DD, MM, YY : String;
Begin
  DD    := FormatDateTime('dd', StrToDateTime(Recibe));
  MM    := FormatDateTime('mm', StrToDateTime(Recibe));
  YY    := FormatDateTime('yyyy', StrToDateTime(Recibe));
  Mes   := StrToInt(MM);
  Case Mes Of
    1 : Begin MM := 'Enero' End;
    2 : Begin MM := 'Febrero' End;
    3 : Begin MM := 'Marzo' End;
    4 : Begin MM := 'Abril' End;
    5 : Begin MM := 'Mayo' End;
    6 : Begin MM := 'Junio' End;
    7 : Begin MM := 'Julio' End;
    8 : Begin MM := 'Agosto' End;
    9 : Begin MM := 'Septiembre' End;
   10 : Begin MM := 'Octubre' End;
   11 : Begin MM := 'Noviembre' End;
   12 : Begin MM := 'Diciembre' End;
  End;
  Result := '          ' + DD + '        ' + MM + '                 ' + YY;
End;

procedure TFacMainForm.FormCreate(Sender: TObject);
Var
  ConfigIni : TIniFile;
begin
  Impresoras.Text := 'Impresora 1=LPT1';
  SelFecha.Date := Now;
  ArchivoCfg := ExtractFileDir(ParamStr(0)) + '\FacturaGG.Ini';
  ConfigIni   := TiniFile.Create(ArchivoCfg);
  SetupFiles  := ConfigIni.ReadString('Archivos', 'Instalacion', 'C:\FacturaGG');
  DataFiles   := ConfigIni.ReadString('Archivos', 'BaseDatos', 'C:\FacturaGG\Data\');
  BackUpFiles := ConfigIni.ReadString('Archivos', 'Respaldo', 'C:\FacturaGG\BackUp\');
  ConfigIni.ReadSectionValues('Printers', Impresoras.Items);
  ConfigIni.Free;

  Clientes.TableName := DataFiles + '\Clientes.Dbf';
  Clientes.IndexName := 'RUT';
  Factura.TableName  := DataFiles + '\Facturas.Dbf';
  Factura.IndexName  := 'FACTURA';
  Detalle.TableName  := DataFiles + '\Detalle.Dbf';
  Detalle.IndexName  := 'Factura';
end;


procedure TFacMainForm.btnCerrarClick(Sender: TObject);
begin
  Close;
end;

procedure TFacMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TFacMainForm.GridGetAlignment(Sender: TObject; ARow, ACol: Integer; var AAlignment: TAlignment);
begin
  If (ACol in [0..5]) And (ARow = 0) Then AAlignment := taCenter;
  If (ACol in [0, 1, 2, 4, 5]) And (ARow > 0) Then AAlignment := taRightJustify;
end;

procedure TFacMainForm.FormKeyPress(Sender: TObject; var Key: Char);
begin
  If Key = #13 Then
    If Not (ActiveControl Is TAdvStringGrid) Then
      Begin
        Key := #0;
        Perform(WM_NEXTDLGCTL, 0, 0);
      End;
end;

procedure TFacMainForm.btnImprimirClick(Sender: TObject);
Begin
  ImprimeFactura;
  GuardarDatos;
End;

Procedure TFacMainForm.Imprimefactura;
Var
  A : TextFile;
  Linea : String[120];
  Largo, I, J, K : Integer;
  Palabra, Palabra2, Texto, Puerta, Fecha : String;
begin
  Fecha := DateToStr(SelFecha.Date);
  Fecha := FechaLarga(Fecha);

  Puerta := Copy(Impresoras.Text, 13, 200);

  // Salida Archivo Datos Cabecera Factura
  AssignFile(A,Puerta);
  ReWrite(A);
  WriteLn(A, #27 + #35 + #48);
  WriteLn(A, #27 + #103);
//  WriteLn(A, '');
//  WriteLn(A, '');
//  WriteLn(A, '');
  WriteLn(A, '');
  WriteLn(A, '');
  WriteLn(A, '');
        //             1         2         3         4         5         6         7         8         9        10        11        12
        //    123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
  WriteLn(A, '                                                                               ' + Edit15.Text);
  WriteLn(A, '');
  WriteLn(A, '');
  // Fecha Emision Linea 9
  WriteLn(A, Fecha);
  WriteLn(A, '');

  // Linea de Nombre y Rut Linea 11
  For I := 1 To 120 Do
    Begin
      Linea[I] := ' ';
    End;
  Texto := Edit1.Text;
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 10 To (10 + Largo) Do
    Begin
      Linea[I] := Texto[J];
      Inc(J);
    End;
  Texto := Edit2.Text + '-' + Edit3.Text;
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 95 To (95 + Largo) Do
    Begin
      Linea[I] := Texto[J];
      Inc(J);
    End;
  For I := 1 To 120 Do
    Write(A, Linea[I]);
  WriteLn(A, '');
  WriteLn(A, '');

  // Linea Direccion y Comuna Linea 13
  For I := 1 To 120 Do
    Begin
      Linea[I] := ' ';
    End;
  Texto := Edit4.Text;
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 11 To (11 + Largo) Do
    Begin
      Linea[I] := Texto[J];
      Inc(J);
    End;
  Texto := Edit5.Text;
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 95 To (95 + Largo) Do
    Begin
      Linea[I] := Texto[J];
      Inc(J);
    End;
  For I := 1 To 120 Do
    Write(A, Linea[I]);
  WriteLn(A, '');
  WriteLn(A, '');

  // Linea Giro, Telefono, Fax, Ciudad Linea 15
  For I := 1 To 120 Do
    Begin
      Linea[I] := ' ';
    End;
  Texto := Edit6.Text;
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 7 To (7 + Largo) Do
    Begin
      Linea[I] := Texto[J];
      Inc(J);
    End;
  Texto := Edit7.Text;
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 60 To (60 + largo) Do
    Begin
      Linea[I] := Texto[J];
      Inc(J);
    End;
  Texto := Edit8.Text;
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 77 To (77 + largo) Do
    Begin
      Linea[I] := Texto[J];
      Inc(J);
    End;
  Texto := Edit9.Text;
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 97 To (97 + largo) Do
    Begin
      Linea[I] := Texto[J];
      Inc(J);
    End;
  For I := 1 To 120 Do
    Write(A, Linea[I]);
  WriteLn(A, '');
  WriteLn(A, '');

  // Linea Guia, Nota Vta., Vendedor, SS Linea 17
  For I := 1 To 120 Do
    Begin
      Linea[I] := ' ';
    End;
  Texto := Edit10.Text;
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 20 To (20 + Largo) Do
    Begin
      Linea[I] := Texto[J];
      Inc(J);
    End;
  Texto := Edit11.Text;
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 57 To (57 + largo) Do
    Begin
      Linea[I] := Texto[J];
      Inc(J);
    End;
  Texto := Edit12.Text;
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 97 To (97 + largo) Do
    Begin
      Linea[I] := Texto[J];
      Inc(J);
    End;
  Texto := Edit13.Text;
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 115 To (115 + largo) Do
    Begin
      Linea[I] := Texto[J];
      Inc(J);
    End;
  For I := 1 To 120 Do
    Write(A, Linea[I]);
  WriteLn(A, '');
  WriteLn(A, '');
  // Linea Condiciones de Venta Linea 19
  WriteLn(A, '                                                                                              ' + Edit14.Text);
  WriteLn(A, '');
  WriteLn(A, '');
  WriteLn(A, '');
  WriteLn(A, '');
  WriteLn(A, '');

  For I := 1 To 26 Do
    Begin
      // Columna Items
      For K := 1 To 120 Do
        Begin
          Linea[K] := ' ';
        End;
      Texto := Grid.Cells[0, I];
      Largo := Length(Texto);
      If Largo < 1 Then Texto := ' ' Else
      If Largo >= 4 Then Begin Largo := 4; Texto := Copy(Grid.Cells[0, I], 1, 4); End;
      J:= Largo;
      For K := 8 DownTo (8 - Largo) Do
        Begin
          Linea[K] := Texto[J];
          Dec(J);
        End;
      // Columna Cantidad
      Texto := Grid.Cells[1, I];
      Largo := Length(Texto);
      If Largo < 1 Then Texto := ' ' Else
      If Largo >= 5 Then Begin Largo := 5; Texto := Copy(Grid.Cells[1, I], 1, 5); End;
      J:= Largo;
      For K := 17 DownTo (17 - Largo) Do
        Begin
          Linea[K] := Texto[J];
          Dec(J);
        End;
      // Columna Codigo
      Texto := Grid.Cells[2, I];
      Largo := Length(Texto);
      If Largo < 1  Then Texto := ' ' Else
      If Largo >= 10 Then Begin Largo := 10; Texto := Copy(Grid.Cells[2, I], 1, 10); End;
      J:= Largo;
      For K := 31 DownTo (31 - Largo) Do
        Begin
          Linea[K] := Texto[J];
          Dec(J);
        End;
      // Columna  Descripcion
      Texto := Grid.Cells[3, I];
      Largo := Length(Texto); J:= 1;
      If Largo < 1  Then Texto := ' ' Else
      If Largo >= 50 Then Begin Largo := 50; Texto := Copy(Grid.Cells[3, I], 1, 50); End;
      For K := 35 To (35 + Largo) Do
        Begin
          Linea[K] := Texto[J];
          Inc(J);
        End;
      // Columna Valor Unitario
      Texto := Grid.Cells[4, I];
      Texto := TrimLeft(Texto); Texto := TrimRight(Texto);
      Largo := Length(Texto);
      If Largo < 1  Then Texto := ' ' Else
      If Largo >= 10 Then
        Begin
          Largo := 10;
          Texto := Copy(Texto, 1, 10);
        End;
      J:= Largo;
      For K := 108 DownTo (108 - Largo) Do
        Begin
          Linea[K] := Texto[J];
          Dec(J);
        End;
      // Columna Valor Total
      Texto := Grid.Cells[5, I];
      Texto := TrimLeft(Texto); Texto := TrimRight(Texto);
      Largo := Length(Texto);
      If Largo < 1  Then Texto := ' ' Else
      If Largo >= 10  Then
        Begin
          Largo := 10;
          Texto := Copy(Texto, 1, 10);
        End;
      J := Largo;
      For K := 120 DownTo (120 - Largo) Do
        Begin
          Linea[K] := Texto[J];
          Dec(J);
        End;
      For K := 1 To 120 Do
        Write(A, Linea[K]);
      WriteLn(A, '');
    End;

  // Impresion de Cantidad en Palabras
  Panel3.Caption := TrimLeft(Panel3.Caption);
  Panel3.Caption := TrimRight(Panel3.Caption);
  Largo := Length(Panel3.Caption);
  Case Largo Of
    1 : Palabra := Unidades(StrToInt(Panel3.Caption));
    2 : Palabra := Decenas(StrToInt(Panel3.Caption));
    3 : Palabra := Centenas(StrToInt(Panel3.Caption));
    4 : Palabra := Millares(StrToInt(Panel3.Caption));
    5 : Palabra := DecenasMil(StrToInt(Panel3.Caption));
    6 : Palabra := CentenasMil(StrToInt(Panel3.Caption));
    7 : Palabra := Millones(StrToInt(Panel3.Caption));
    8 : Palabra := DecenasMillon(StrToInt(Panel3.Caption));
    9 : Palabra := CentenasMillon(StrToInt(Panel3.Caption));
   10 : Palabra := MillaresMillon(StrToInt(Panel3.Caption));
  End;
  If Largo > 100 Then
    Begin
      Palabra2 := Copy(Palabra, 101, Largo);
      Palabra  := Copy(Palabra, 1, 100);
      Palabra  := Palabra + '-';
      Palabra2 := Palabra2 + ' PESOS.';
      WriteLn(A, '   SON ' + Palabra);
      WriteLn(A, '       ' + Palabra2);
//      WriteLn(A, '');
    End
  Else
    Begin
      Palabra := Palabra + ' PESOS.';
      WriteLn(A, '   SON ' + Palabra);
//      WriteLn(A, '');
      WriteLn(A, '');
    End;

  // Linea Valor Neto
  For K := 1 To 120 Do Begin Linea[K] := ' '; End;
  Texto := TrimLeft(Panel1.Caption); Texto := TrimRight(Panel1.Caption);
  Largo := Length(Texto); J:= Largo;
  If Largo < 1 Then Texto := ' ';
  For I := 120 DownTo (120 - Largo) Do
    Begin
      Linea[I] := Texto[J];
      Dec(J);
    End;
  For I := 1 To 120 Do
    Write(A, Linea[I]);
  WriteLn(A, '');
  WriteLn(A, '');
  // Linea Valor IVA
  For K := 1 To 120 Do Begin Linea[K] := ' '; End;
  Texto := TrimLeft(Panel2.Caption); Texto := TrimRight(Panel2.Caption);
  Largo := Length(Texto); J:= Largo;
  If Largo < 1 Then Texto := ' ';
  For I := 120 DownTo (120 - Largo) Do
    Begin
      Linea[I] := Texto[J];
      Dec(J);
    End;
  For I := 1 To 120 Do
    Write(A, Linea[I]);
  WriteLn(A, '');
  WriteLn(A, '');
  // Linea Valor Total
  For K := 1 To 120 Do Begin Linea[K] := ' '; End;
  Texto := TrimLeft(Panel3.Caption); Texto := TrimRight(Panel3.Caption);
  Largo := Length(Texto); J:= Largo;
  If Largo < 1 Then Texto := ' ';
  For I := 120 DownTo (120 - Largo) Do
    Begin
      Linea[I] := Texto[J];
      Dec(J);
    End;
  For I := 1 To 120 Do
    Write(A, Linea[I]);
  WriteLn(A, '');
  WriteLn(A, '');
  CloseFile(A);
end;

procedure TFacMainForm.GridKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var
  Largo : Integer;
  Descripcion : String;
  Cantidad, Unitario, TotalLinea : Real;
begin
  If Key = 13 Then
  Case Grid.Col Of
    4 : Begin
          Largo := Length(Grid.Cells[3, Grid.Row]);
          If Largo >= 60 Then
            Begin
              Grid.Cells[3, Grid.Row] := UpperCase(Grid.Cells[3, Grid.Row]);
              If Grid.Row < 29 Then
                Begin
                  Grid.Cells[3, Grid.Row + 1] := Copy(Grid.Cells[3, Grid.Row], 61, Largo - 60);
                  Grid.Cells[3, Grid.Row] := Copy(Grid.Cells[3, Grid.Row], 1, 60);
                End
              Else
                Begin
                  Grid.Cells[3, Grid.Row] := Copy(Grid.Cells[3, Grid.Row], 1, 60);
                End;
            End
          Else
            Begin
              Grid.Cells[3, Grid.Row] := Copy(Grid.Cells[3, Grid.Row], 1, 60);
              Grid.Cells[3, Grid.Row] := UpperCase(Grid.Cells[3, Grid.Row]);
            End;

          If (Grid.Cells[1, Grid.Row] <> '') And (Grid.Cells[1, Grid.Row] <> '') Then
            Begin
              TotalLinea := 0; Cantidad := 0; Unitario := 0;
              Cantidad   := StrToFloat(Grid.Cells[1, Grid.Row]);
              Unitario   := StrToFloat(Grid.Cells[4, Grid.Row]);
              TotalLinea := (Cantidad * Unitario);
              Grid.Cells[5, Grid.Row] := Format('%10.0f', [TotalLinea]);
            End;
        End;
  End;

end;

procedure TFacMainForm.Edit3Exit(Sender: TObject);
Var
  Rut, Dig : String;
  I, Largo, Acumula, Resto : Integer;
  rRut : Array[1..8] Of Integer;
begin
  Acumula := 0;
  Rut := Edit2.Text;
  Largo := Length(Rut);
  Case Largo Of
    7 : Begin
          rRut[2] := StrToInt(Copy(Rut, 1, 1)); rRut[3] := StrToInt(Copy(Rut, 2, 1)); rRut[4] := StrToInt(Copy(Rut, 3, 1));
          rRut[5] := StrToInt(Copy(Rut, 4, 1)); rRut[6] := StrToInt(Copy(Rut, 5, 1)); rRut[7] := StrToInt(Copy(Rut, 6, 1));
          rRut[8] := StrToInt(Copy(Rut, 7, 1));

          Acumula := Acumula + (rRut[8] * 2);
          Acumula := Acumula + (rRut[7] * 3);
          Acumula := Acumula + (rRut[6] * 4);
          Acumula := Acumula + (rRut[5] * 5);
          Acumula := Acumula + (rRut[4] * 6);
          Acumula := Acumula + (rRut[3] * 7);
          Acumula := Acumula + (rRut[2] * 2);
        End;
    8 : Begin
          rRut[1] := StrToInt(Copy(Rut, 1, 1)); rRut[2] := StrToInt(Copy(Rut, 2, 1)); rRut[3] := StrToInt(Copy(Rut, 3, 1)); rRut[4] := StrToInt(Copy(Rut, 4, 1));
          rRut[5] := StrToInt(Copy(Rut, 5, 1)); rRut[6] := StrToInt(Copy(Rut, 6, 1)); rRut[7] := StrToInt(Copy(Rut, 7, 1)); rRut[8] := StrToInt(Copy(Rut, 8, 1));

          Acumula := Acumula + (rRut[8] * 2);
          Acumula := Acumula + (rRut[7] * 3);
          Acumula := Acumula + (rRut[6] * 4);
          Acumula := Acumula + (rRut[5] * 5);
          Acumula := Acumula + (rRut[4] * 6);
          Acumula := Acumula + (rRut[3] * 7);
          Acumula := Acumula + (rRut[2] * 2);
          Acumula := Acumula + (rRut[1] * 3);
        End;
  End;
  Resto := Acumula Mod 11;
  Dig   := IntToStr(11 - Resto);
  If Dig = '11' Then
    Dig := '0';
  If Dig = '10' Then
    Dig := 'K';

  If Dig <> Edit3.Text Then
    Begin
      Edit2.Text := ''; Edit3.Text := '';
      Edit2.SetFocus;
    End
  Else
    Begin
      Clientes.Open;
      Clientes.SetKey;
      If Not Clientes.FindKey([StrToFloat(Rut)]) Then
        Begin
          Check.Checked := True;
          Check.State   := cbChecked;
          Edit1.Enabled := True; Edit1.Text := '';
          Edit4.Enabled := True; Edit4.Text := '';
          Edit5.Enabled := True; Edit5.Text := '';
          Edit6.Enabled := True; Edit6.Text := '';
          Edit7.Enabled := True; Edit7.Text := '';
          Edit8.Enabled := True; Edit8.Text := '';
          Edit9.Enabled := True; Edit9.Text := '';
          Edit1.SetFocus;
        End
      Else
        Begin
          Edit1.Text  := Clientes.Fields[2].AsString;   // Razon Social;
          Edit4.Text  := Clientes.Fields[3].AsString;   // Direccion
          Edit5.Text  := Clientes.Fields[4].AsString;   // Comuna
          Edit6.Text  := Clientes.Fields[6].AsString;   // Giro
          Edit7.Text  := Clientes.Fields[7].AsString;   // Fono
          Edit8.Text  := Clientes.Fields[8].AsString;   // Fax
          Edit9.Text  := Clientes.Fields[5].AsString;   // Ciudad
          Edit10.SetFocus;
        End;
      Clientes.Close;
    End;
end;

procedure TFacMainForm.GridExit(Sender: TObject);
Var
  I : Integer;
  Acumula : Real;
begin
  Acumula := 0;
  For I := 1 To 29 Do
    Begin
      If (Grid.Cells[5, I] <> '') Then
        Acumula := Acumula + StrToFloat(Grid.Cells[5, I]);
    End;
  Panel1.Caption := Format('%10.0f', [Acumula]) + ' ';
  Panel2.Caption := Format('%10.0f', [Acumula * 0.18]) + ' ';
  Panel3.Caption := Format('%10.0f', [Acumula * 1.18]) + ' ';
  MessageBox(Handle, 'Por Favor, Revise los Antecedentes Antes de Imprimir.', 'Impresión Preparada...', mb_Ok or mb_IconQuestion or mb_DefButton1);
  btnImprimir.Enabled := True;
end;

Procedure TFacMainForm.GuardarDatos;
Var
  I : Integer;
Begin
  If (Check.State = cbChecked) Then
    Begin
    End;
  Factura.Open;
  Factura.SetKey;
  If Not Factura.FindKey([StrToFloat(Edit15.Text)]) Then
    Begin
      Factura.Append;
      Factura.Fields[0].Value  := StrToFloat(Edit2.Text);         // Rut
      Factura.Fields[1].Value  := StrToFloat(Edit15.Text);        // Numero Factura
      Factura.Fields[2].Value  := SelFecha.Date;                  // Fecha Emision Factura
      Factura.Fields[3].Value  := Edit10.Text;                    //
      Factura.Fields[4].Value  := Edit11.Text;
      Factura.Fields[5].Value  := Edit12.Text;
      Factura.Fields[6].Value  := Edit13.Text;
      Factura.Fields[7].Value  := Edit14.Text;
      Factura.Fields[8].Value  := StrToFloat(Panel1.Caption);
      Factura.Fields[9].Value  := StrToFloat(Panel2.Caption);
      Factura.Fields[10].Value := StrToFloat(Panel3.Caption);
      Factura.Post;
      Detalle.Open;
      For I := 1 To 29 Do
        Begin
          Detalle.Append;
          Detalle.Fields[0].Value := StrToFloat(Edit15.Text);
          Detalle.Fields[1].Value := IntToStr(I);
          Detalle.Fields[2].Value := Grid.Cells[0, I];
          Detalle.Fields[3].Value := Grid.Cells[1, I];
          Detalle.Fields[4].Value := Grid.Cells[2, I];
          Detalle.Fields[5].Value := Grid.Cells[3, I];
          Detalle.Fields[6].Value := Grid.Cells[4, I];
          Detalle.Fields[7].Value := Grid.Cells[5, I];
          Detalle.Post
        End;
    End
  Else
    Begin
      MessageBox(Handle, 'FACTURA YA EXISTE. SELECCIONE OTRO NÚMERO', 'ADVERTENCIA...', mb_Ok or mb_IconStop or mb_DefButton1);
      Edit15.SetFocus;
    End;
  Factura.Close;
  Detalle.Close;
End;

procedure TFacMainForm.CheckClick(Sender: TObject);
begin
  If (Check.State = cbChecked) Then
    Begin
//      Check.Checked := False;
      Edit1.Enabled := False;
      Edit4.Enabled := False;
      Edit5.Enabled := False;
      Edit6.Enabled := False;
      Edit7.Enabled := False;
      Edit8.Enabled := False;
      Edit9.Enabled := False;
    End
  Else
    Begin
//      Check.Checked := True;
      Edit1.Enabled := True;
      Edit4.Enabled := True;
      Edit5.Enabled := True;
      Edit6.Enabled := True;
      Edit7.Enabled := True;
      Edit8.Enabled := True;
      Edit9.Enabled := True;
    End;
end;

procedure TFacMainForm.Edit15Exit(Sender: TObject);
begin
  Factura.Open;
  Factura.SetKey;
  If Factura.FindKey([StrToFloat(Edit15.Text)]) Then
    Begin
      MessageBox(Handle, 'FACTURA YA EXISTE. SELECCIONE OTRO NÚMERO', 'ADVERTENCIA...', mb_Ok or mb_IconStop or mb_DefButton1);
      Edit15.Text := '';
      Edit15.SetFocus;
    End;
  Factura.Close;  
end;

end.
