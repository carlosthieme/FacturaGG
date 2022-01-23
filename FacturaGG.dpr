program FacturaGG;

uses
  Forms,
  FacMain in 'FacMain.pas' {FacMainForm};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TFacMainForm, FacMainForm);
  Application.CreateForm(TFacMainForm, FacMainForm);
  Application.Run;
end.
