unit Num2Word;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Dialogs,
  StdCtrls;

  Function Unidades(Recibe : LongInt) : String;
  Function Decenas(Recibe  : LongInt) : String;
  Function Centenas(Recibe : longInt) : String;
  Function Millares(Recibe : LongInt) : String;
  Function DecenasMil(Recibe : LongInt) : String;
  Function CentenasMil(Recibe : LongInt) : String;
  Function Millones(Recibe : LongInt) : String;
  Function DecenasMillon(Recibe : LongInt) : String;
  Function CentenasMillon(Recibe : LongInt) : String;
  Function MillaresMillon(Recibe : LongInt) : String;

Const
 Unidad  : Array[1..9] Of String = ('UN', 'DOS', 'TRES', 'CUATRO', 'CINCO', 'SEIS', 'SIETE', 'OCHO', 'NUEVE');
 Decena  : Array[1..9] Of String = ('ONCE', 'DOCE', 'TRECE', 'CATORCE', 'QUINCE', 'DIECISEIS', 'DIECISIETE', 'DIECIOCHO', 'DIECINUEVE');
 Decena1 : Array[1..9] Of String = ('VEINTIUNO', 'VEINTIDOS', 'VEINTITRES', 'VINTICUATRO', 'VEINTICINCO', 'VEINTISEIS', 'VEINTISIETE', 'VEINTIOCHO',
                                    'VEINTINUEVE');
 Decena2 : Array[1..9] Of String = ('DIEZ', 'VEINTE', 'TREINTA', 'CUARENTA', 'CINCUENTA', 'SESENTA', 'SETENTA', 'OCHENTA', 'NOVENTA');
 Centena : Array[1..9] Of string = ('CIENTO', 'DOSCIENTOS', 'TRESCIENTOS', 'CUATROCIENTOS', 'QUINIENTOS', 'SEISCIENTOS',
                                    'SETECIENTOS', 'OCHOCIENTOS', 'NOVECIENTOS');

implementation

Function Unidades(Recibe : LongInt) : String;
Begin
  Case Recibe Of
    1..9 : Result := Unidad[Recibe];
  End;

End;

Function Decenas(Recibe : LongInt) : String;
Var
  Entero, Resto : Integer;
Begin
  Case Recibe Of
    10..99 : Begin
               Entero := Recibe Div 10;
               Resto := Recibe Mod 10;
               If (Resto > 0) Then
                 Begin
                   Case Entero Of
                     1 : Result := Decena[Resto];
                     2 : Result := Decena1[Resto]
                     Else
                       Result := Decena2[Entero] + ' Y ' + Unidades(Resto);
                   End;
                 End
               Else
                 Begin
                   Result := Decena2[Entero];
                 End;
             End;
  End;
End;

Function Centenas(Recibe : longInt) : String;
Var
  Entero, Resto : Integer;
  Palabra : String;
Begin
  Case Recibe Of
        100 : Result := 'CIEN';
   101..999 : Begin
                Entero  := Recibe Div 100;
                Resto   := Recibe Mod 100;
                Palabra := Centena[Entero];
                If (Resto = 0) Then
                  Result := Palabra
                Else
                  Begin
                    Entero := Recibe - (Entero * 100);
                    Case Entero Of
                       1..9 : Result := Palabra + ' ' + Unidades(Entero);
                     10..99 : Result := Palabra + ' ' + Decenas(Entero);
                    End;
                  End;
              End;
  End;
End;

Function Millares(Recibe : LongInt) : String;
Var
  Entero, Resto : Integer;
  Palabra :  String;
Begin
  Case Recibe Of
  1000..9999 : Begin
                 Entero  := Recibe Div 1000;
                 Resto   := Recibe Mod 1000;
                 Palabra := Unidades(Entero) + ' MIL ';
                 If (Resto = 0) Then
                   Result := Palabra
                 Else
                   Begin
                     Entero := Recibe - (Entero * 1000);
                     Case Entero Of
                             0 : Result := Palabra;
                          1..9 : Result := Palabra + Unidades(Entero);
                        10..99 : Result := Palabra + Decenas(Entero);
                      100..999 : Result := Palabra + Centenas(Entero);
                     End;
                   End;
               End;
  End;
End;

Function DecenasMil(Recibe : LongInt) : String;
Var
  Entero, Resto : Integer;
  Palabra :  String;
Begin
  Case Recibe Of
  10000..99999 : Begin
                   Entero  := Recibe Div 1000;
                   Resto   := Recibe Mod 1000;
                   Palabra := Decenas(Entero) + ' MIL ';
                   If (Resto = 0) Then
                     Result := Palabra
                   Else
                     Begin
                       Entero := Recibe - (Entero * 1000);
                       Case Entero Of
                           1..9 : Result := Palabra + Unidades(Entero);
                         10..99 : Result := Palabra + Decenas(Entero);
                       100..999 : Result := Palabra + Centenas(Entero);
                       End;
                     End;
                 End;
  End;
End;

Function CentenasMil(Recibe : LongInt) : String;
Var
  Entero, Resto : Integer;
  Palabra :  String;
Begin
  Case Recibe Of
   100000..999999 :
     Begin
       Entero  := Recibe Div 1000;
       Resto   := Recibe Mod 1000;
       Palabra := Centenas(Entero) + ' MIL ';
       If (Resto = 0) Then
         Result := Palabra
       Else
         Begin
           Entero := Recibe - (Entero * 1000);
           Case Entero Of
               1..9 : Result := Palabra + Unidades(Entero);
             10..99 : Result := Palabra + Decenas(Entero);
           100..999 : Result := Palabra + Centenas(Entero);
           End;
         End;
     End;
  End;
End;

Function Millones(Recibe : LongInt) : String;
Var
  Entero  : LongInt;
  Palabra : String;
Begin
  Case Recibe Of
    1000000..9999999 : Begin
                         Entero  := Recibe Div 1000000;
                         If (Entero = 1) Then
                           Palabra := Unidades(Entero) + ' MILLON '
                         Else
                           Palabra := Unidades(Entero) + ' MILLONES ';
                         Entero  := Recibe - (Entero * 1000000);
                         Case Entero Of
                           100000..999999 : Result := Palabra + CentenasMil(Entero);
                             10000..99999 : Result := Palabra + DecenasMil(Entero);
                               1000..9999 : Result := Palabra + Millares(Entero);
                                 100..999 : Result := Palabra + Centenas(Entero);
                                   10..99 : Result := Palabra + Decenas(Entero);
                                     1..9 : Result := Palabra + Unidades(Entero);
                                        0 : Result := Palabra;
                         End;
                       End;
  End;
End;

Function DecenasMillon(Recibe : LongInt) : String;
Var
  Entero  : LongInt;
  Palabra : String;
Begin
  Case Recibe Of
    10000000..99999999 : Begin
                           Entero  := Recibe Div 1000000;
                           Palabra := Decenas(Entero) + ' MILLONES ';
                           Entero  := Recibe - (Entero * 1000000);
                           Case Entero Of
                            100000..999999 : Result := Palabra + CentenasMil(Entero);
                              10000..99999 : Result := Palabra + DecenasMil(Entero);
                                1000..9999 : Result := Palabra + Millares(Entero);
                                  100..999 : Result := Palabra + Centenas(Entero);
                                    10..99 : Result := Palabra + Decenas(Entero);
                                      1..9 : Result := Palabra + unidades(Entero);
                                         0 : Result := Palabra;
                           End;
                         End;
  End;
End;

Function CentenasMillon(Recibe : LongInt) : String;
Var
  Entero  : LongInt;
  Palabra : String;
Begin
  Case Recibe Of
    100000000..999999999 : Begin
                             Entero  := Recibe Div 1000000;
                             Palabra := Centenas(Entero) + ' MILLONES ';
                             Entero  := Recibe - (Entero * 1000000);
                             Case Entero Of
                              100000..999999 : Result := Palabra + CentenasMil(Entero);
                                10000..99999 : Result := Palabra + DecenasMil(Entero);
                                  1000..9999 : Result := Palabra + Millares(Entero);
                                    100..999 : Result := Palabra + Centenas(Entero);
                                      10..99 : Result := Palabra + Decenas(Entero);
                                        1..9 : Result := Palabra + unidades(Entero);
                                           0 : Result := Palabra;
                             End;
                           End;
  End;
End;

Function MillaresMillon(Recibe : LongInt) : String;
Var
  Entero  : LongInt;
  Palabra : String;
Begin
 Case Recibe Of
   1000000000..1999999999 : Begin
                              Entero  := Recibe Div 1000000;
                              Palabra :=  Millares(Entero) + ' MILLONES ';
                              Entero  := Recibe - (Entero * 1000000);
                              Case Entero Of
                               100000..999999 : Result := Palabra + CentenasMil(Entero);
                                 10000..99999 : Result := Palabra + DecenasMil(Entero);
                                   1000..9999 : Result := Palabra + Millares(Entero);
                                     100..999 : Result := Palabra + Centenas(Entero);
                                       10..99 : Result := Palabra + Decenas(Entero);
                                         1..9 : Result := Palabra + unidades(Entero);
                                            0 : Result := Palabra;
                              End;
                            End;
 End;
End;

end.
