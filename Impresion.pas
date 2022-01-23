unit Impresion;

interface

uses
    Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
    StdCtrls, ComCtrls, Num2Word, Printers;
    
    Function FechaLarga(Recibe : String) : String;
    Function FacturaNacional(Puerta : String) : Boolean;

implementation

Function FechaLarga(Recibe : String) : String;
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
  Result := DD + '       ' + MM + '        ' + YY;
End;


Function FacturaNacional(Puerta : String) : Boolean;
Var
  A, B : TextFile;
  LineaPrint : String[137];
  Texto : String;
  Largo, I, J : Integer;
Begin
  // Lee Cabecera Factura
  AssignFile(A, 'C:\HFactura.Dat');
  Reset(A);
  AssignFile(B, Puerta);
  {$I-}
  ReWrite(B);
  {$I+}
//WriteLn(B, '         1         2         3         4         5         6         7         8         9         0         1          2        3');
//WriteLn(B, '12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567');
  WriteLn(B, #27 + #15);
//  WriteLn(B, '');
  WriteLn(B, '');
  WriteLn(B, '');
  WriteLn(B, '');
  WriteLn(B, '');
  // Numero Factura Linea 6
  ReadLn(A, Texto);
  WriteLn(B, '                                                                                          ' + Texto);
  WriteLn(B, '');
  WriteLn(B, '');
  // Fecha Emision Linea 9
  ReadLn(A, Texto);
  WriteLn(B, '          ' + Texto);
  WriteLn(B, '');

  // Linea de Nombre y Rut Linea 11
  For I := 1 To 137 Do
    Begin
      LineaPrint[I] := ' ';
    End;
  ReadLn(A, Texto);
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 12 To (12 + Largo) Do
    Begin
      LineaPrint[I] := Texto[J];
      Inc(J);
    End;
  ReadLn(A, Texto);
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 110 To (110 + Largo) Do
    Begin
      LineaPrint[I] := Texto[J];
      Inc(J);
    End;
  For I := 1 To 137 Do
    Write(B, LineaPrint[I]);
  WriteLn(B, '');
  WriteLn(B, '');

  // Linea Direccion y Comuna Linea 13
  For I := 1 To 137 Do
    Begin
      LineaPrint[I] := ' ';
    End;
  ReadLn(A, Texto);
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 12 To (12 + Largo) Do
    Begin
      LineaPrint[I] := Texto[J];
      Inc(J);
    End;
  ReadLn(A, Texto);
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 110 To (110 + Largo) Do
    Begin
      LineaPrint[I] := Texto[J];
      Inc(J);
    End;
  For I := 1 To 137 Do
    Write(B, LineaPrint[I]);
  WriteLn(B, '');
  WriteLn(B, '');

  // Linea Giro, Telefono, Fax, Ciudad Linea 15
  For I := 1 To 137 Do
    Begin
      LineaPrint[I] := ' ';
    End;
  ReadLn(A, Texto); // Giro
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 12 To (12 + Largo) Do
    Begin
      LineaPrint[I] := Texto[J];
      Inc(J);
    End;
  ReadLn(A, Texto); // Telefono
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 67 To (67 + largo) Do
    Begin
      LineaPrint[I] := Texto[J];
      Inc(J);
    End;
  ReadLn(A, Texto); // Fax
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 88 To (88 + largo) Do
    Begin
      LineaPrint[I] := Texto[J];
      Inc(J);
    End;
  ReadLn(A, Texto); // Ciudad
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 112 To (112 + largo) Do
    Begin
      LineaPrint[I] := Texto[J];
      Inc(J);
    End;
  For I := 1 To 137 Do
    Write(B, LineaPrint[I]);
  WriteLn(B, '');
  WriteLn(B, '');

  // Linea Guia, Nota Vta., Vendedor, SS Linea 17
  For I := 1 To 137 Do
    Begin
      LineaPrint[I] := ' ';
    End;
  ReadLn(A, Texto); // Guia
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 24 To (24 + Largo) Do
    Begin
      LineaPrint[I] := Texto[J];
      Inc(J);
    End;
  ReadLn(A, Texto); // Nota Vta.
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 66 To (66 + largo) Do
    Begin
      LineaPrint[I] := Texto[J];
      Inc(J);
    End;
  ReadLn(A, Texto); // Vendedor
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 112 To (112 + largo) Do
    Begin
      LineaPrint[I] := Texto[J];
      Inc(J);
    End;
  ReadLn(A, Texto); // SS
  Largo := Length(Texto); J:= 1;
  If Largo < 1 Then Texto := ' ';
  For I := 130 To (130 + largo) Do
    Begin
      LineaPrint[I] := Texto[J];
      Inc(J);
    End;
  For I := 1 To 137 Do
    Write(B, LineaPrint[I]);
  WriteLn(B, '');
  WriteLn(B, '');
  // Linea Condiciones de Venta Linea 19
  ReadLn(A, Texto);
  WriteLn(B, '                                                                                                             ' + Texto);
  WriteLn(B, '');
  WriteLn(B, '');
  WriteLn(B, '');
  WriteLn(B, '');
//  WriteLn(B, '');
  CloseFile(A);
  CloseFile(B);
  Result := True;
End;



{

  WriteLn(B, '    Solicitada Por : ' + NombreRet);
  WriteLn(B, '    R.U.T.         : ' + RutRet);
  WriteLn(B, '    Firma          :_____________________');
  WriteLn(B, '');

  Largo := Length(Total);
  Case Largo Of
    1 : Palabra := Unidades(StrToInt(Total));
    2 : Palabra := Decenas(StrToInt(Total));
    3 : Palabra := Centenas(StrToInt(Total));
    4 : Palabra := Millares(StrToInt(Total));
    5 : Palabra := DecenasMil(StrToInt(Total));
    6 : Palabra := CentenasMil(StrToInt(Total));
    7 : Palabra := Millones(StrToInt(Total));
    8 : Palabra := DecenasMillon(StrToInt(Total));
    9 : Palabra := CentenasMillon(StrToInt(Total));
   10 : Palabra := MillaresMillon(StrToInt(Total));
  End;

  WriteLn(B, '   ' + Palabra + ' PESOS.-');
  WriteLn(B, '');
  WriteLn(B, '                   EXENTO');
  WriteLn(B, '');
  For I := 1 To 80 Do
    Begin
      LineaPrint[I] := ' ';
    End;
  Largo := Length(Total); J := Largo;
  For I := 80 DownTo (80 - Largo) Do
    Begin
      LineaPrint[I] := Total[J];
      Dec(J);
    End;

  Largo := Length(Iva); J := Largo;
  LineaPrint := '';
  For I := 67 DownTo (67 - Largo) Do
    Begin
      LineaPrint[I] := Iva[J];
      Dec(J);
    End;

  Largo := Length(Neto); J := Largo;
  LineaPrint := '';
  For I := 40 DownTo (40 - Largo) Do
    Begin
      LineaPrint[I] := Neto[J];
      Dec(J);
    End;

  Largo := Length(Exento); J := Largo;
  LineaPrint := '';
  For I := 27 DownTo (27 - Largo) Do
    Begin
      LineaPrint[I] := Exento[J];
      Dec(J);
    End;

  LineaPrint[55] := '1';
  LineaPrint[56] := '8';
  LineaPrint[57] := '%';

  For I := 1 To 80 Do
    Write(B, LineaPrint[I]);

  WriteLn(B, '');
  WriteLn(B, '');
  WriteLn(B, '');
  WriteLn(B, '');
  WriteLn(B, '');
  WriteLn(B, '');
  WriteLn(B, '');
  WriteLn(B, '');
  WriteLn(B, '');

  {Impresion codigos de reposicionamiento  de la hoja}
//WriteLn(B, #12); {Alimentar Papel}

//  CloseFile(A);
//  CloseFile(B);
//  Result := True;
//End;





end.

