// *******************************************************
// ** Delphi object for dual SpreadSheet managing using **
// ** Excel or OpenOffice in a transparent way.         **
// ** By: Sergio Hernandez (oficina(at)hcsoft.net)      **
// ** Version 1.02 07-02-2012 (DDMMYYYY)                **
// ** Use it freely, change it, etc. at will.           **
// *******************************************************

//Latest version, questions, modifications:
// http://user.services.openoffice.org/en/forum/viewtopic.php?f=21&t=47644&p=219641
// https://github.com/sergio-hcsoft/Delphi-SpreadSheets

{EXAMPLE OF USE
  //Create object: We have two flavours:
  //(A) from an existing file...
  HCalc:= THojaCalc.create(OpenDialog.FileName, false);
  //(B) from a blank document...
  HCalc:= THojaCalc.create(thcOpenOffice, true); //OpenOffice doc if possible, please
  HCalc.FileName:= 'C:\MyNewDoc'; //Needs a file name before you SaveDoc!
  //--end of creation.
  HCalc.ActivateSheetByIndex(2); //Activate second sheet
  if HCalc.IsActiveSheetProtected then
    ShowMessage('2nd sheet of name "'+HCalc.ActiveSheetName+'" IS protected');
  //Change a cell value.
  if HCalc.CellText[i,2] = '' then HCalc.CellText[i,2] := 'Hello world!';
  HCalc.AddNewSheet('New Sheet');
  HCalc.PrintDoc;
  HCalc.SaveDoc;
  HCalc.Free;
}

{TODO LIST:
  -Test on LibreOffice: Some nasty screens opens asking things that should show!
  -PrintActiveSheet is not working for OpenOffice (is it even possible?)
  -Listener for OpenOffice so I can be notified if user visually close the doc.
}

{CHANGE LOG:
 V1.02: Creating from a exiting file didn't set the AmericanFormat (thanxs Malte).
 V1.01:
   ***********************
   ** By Malte Tüllmann **
   ***********************
   -Excel2000/2003 save .xls files in a different way than 2007.
 V1.00:
   -Saving in Excel2007 will use Excel97 .xls file format instead of .xlsx
 V0.99:
   -Added a funtion by Alex Smith to set a cell text into italic.
 V0.98:
   -Added two procedures to easily send a number or a date to a cell position:
   SendDate(Row, Col, Date) and SendNumber(Row, Col, Float), if you look at
   the code you will notice that this is not so trivial as one could spect.
   -I have added (as comments) some useful code found on forums (copy-paste rows)
 V0.97:
   -Added CellFormula(col, row), similar to CellText, but allows to set a cell
   to a number wihout the efect of being considered by excel like a "text that
   looks like a number" (doesn't affect OpenOffice). Use it like this:
   CellFormula(1,1):= '=A2*23211.66';
   Note1: Excel will always spect numbers in this shape: no thousand separator
          and dot as decimal separator, regardless of your local configuration.
   Note2: Date is also bad interpreted in Excel, in this case you can use
          CellText but the date must be in american format: MM/DD/YYYY, if you
          use other format, it will try to interpret as an american date and
          only if it fails will use your local date format to "decode" it.
 V0.96:
   -Added PrintSheetsUntil(LastSheetName: string) -only works on excel- to print
   out all tabs from 1 until -excluded- the one with the given name in such a
   way that only one print job is created instead of one per tab (only way to do
   this in previous versions, so converting part of a excel to a single PDF
   using a printer like PDFCreator was not posible).
 V0.95:
   -ActivateSheetByIndex detect imposible index and allows to insert sheet 100 (it will create all necesary sheets)
   -SaveDocAs added a second optional parameter for OOo to use Excel97 format (rescued from V0.93 by Rômulo)
   -A little stronger ValidateSheetName() (filter away \ and " too).
 V0.94:
   -OpenOffice V2 compatible (small changes)
   -A lot of "try except" to avoid silly errors.
   -SaveDocAs(Name: string): boolean; (Added by Massimiliano Gozzi)
   -New function FileName2URL(Name) to convert from FileName to URL (OOo SaveDosAs)
   -New function ooCreateValue to hide all internals of OOo params creation
 V0.93:
   ***************************
   ** By Rômulo Silva Ramos **
   ***************************
   -FontSize(Row, Col, Size): change font size in that cell.
   -BackgroundColor(row, col: integer; color:TColor);
   -Add ValidateSheetName to validate sheet names when adding or renaming a sheet
   REVERTED FUNCTIONS (not neccesary in newer version V0.95 anymore)
   -Change AddNewSheet to add a new sheet in end at sheet list
   *REVERTED IN V0.95*
       It creates sheet following the active one, so to add at the end:
       ActivateSheetByIndex(CountSheets);
       AddNewSheet('Sheet '+IntToStr(CountSheets+1));
   -Change in SaveDoc to use SaveAs/StoreAsUrl
   *REVERTED V0.95*
       Use SaveDocAs(Name, true) for StoreAsUrl in Excel97 format.
 V0.92:
   -SetActiveSheetName didn't change the name to the right sheet on OpenOffice.
   -PrintPreview: New procedure to show up the print preview window.
   -Bold(Row, Col): Make bold the text in that cell.
   -ColumnWidth(col, width): To change a column width.
 V0.91:
   -NewDoc: New procedure for creating a blank doc (used in create)
   -Create from empty doc adds a blank document and take visibility as parameter.
   -New functions ooCreateValue and ooDispatch to clean up the code.
   -ActiveSheetName: Now is a read-write property, not a read-only function.
   -Visible: Now is a read-write property instead of a create param only.
 V0.9:
  -Create from empty doc now tries both programs (if OO fails try to use Excel).
  -CellTextByName: Didn't work on Excel docs.
}

{  PIECES OF CODE FOUND ON FORUMS WORTH COPYING HERE FOR FUTURE USE

  -Interesting "copy-paste one row to another" delphi code from PauLita posted
  on the OO forum (www.oooforum.org/forum/viewtopic.phtml?t=8878):

  OpenOffice version:
         Programa     := CreateOleObject( 'com.sun.star.ServiceManager' );
         ooParams     := VarArrayCreate([0,0],varVariant);
         ooParams[0]  := Programa.Bridge_GetStruct('com.sun.star.beans.PropertyValue');
         ooView       := Document.getCurrentController;
         ooFrame      := ooView.getFrame;
         ooDispatcher := Programa.createInstance('com.sun.star.frame.DispatchHelper');
         // copy to clipboard
         oRange := Sheet.GetRows.GetByIndex(rl-1);
         ooView.Select( oRange );
         ooDispatcher.executeDispatch( ooFrame, '.uno:Copy',  '', 0, ooParams );
         // add one row to the table
         Sheet.GetRows.InsertByIndex(rl,1);
         // paste from clipboard
         oRange := Sheet.GetRows.GetByIndex(rl);
         ooView.Select( oRange );
         ooDispatcher.executeDispatch( ooFrame, '.uno:Paste',  '', 0, ooParams );
  Excel version:
         Sheet.Rows[r].Copy;
         Sheet.Rows[r+1].Insert(xlDown);
}

unit UHojaCalc;

interface

uses Variants, SysUtils, ComObj, Classes, Graphics, Controls;

//thcError: Tried to open but both failes
//thcNone:  Haven't tried still to open any
type TTipoHojaCalc = (thcError, thcNone, thcExcel, thcOpenOffice);

type THojaCalc = class(TObject)
private
  fVisible:  boolean;
  fKeepAlive: boolean;
  AmericanFormat: TFormatSettings;
  //Program loaded stuff...
  procedure  LoadProg;
  procedure  CloseProg;
  function   GetProgLoaded: boolean;
  procedure  NewDoc;
  procedure  LoadDoc;
  procedure  CloseDoc;
  function   GetDocLoaded: boolean;
  function   GetIsExcel: boolean;
  function   GetIsOpenOffice: boolean;
  procedure  SetVisible(v: boolean);
  //Sheets stuff..
  function   GetCountSheets: integer;
  function   GetActiveSheetName: string;
  procedure  SetActiveSheetName(NewName: string);
  //Cells stuff...
  //General input/output of cell content:
  function   GetCellText(row,col: integer): string;
  procedure  SetCellText(row,col: integer; Txt: string);
  function   GetCellFormula(row,col: integer): string;
  procedure  SetCellFormula(row,col: integer; Txt: string);
  function   GetCellTextByName(Range: string): string;
  procedure  SetCellTextByName(Range: string; Txt: string);
  //OpenOffice only stuff...
  function   FileName2URL(FileName: string): string;
  procedure  ooDispatch(ooCommand: string; ooParams: variant);
  function   ooCreateValue(ooName: string; ooData: variant): variant;
  //Aux functions
  function   ValidateSheetName(Name:string): string;
public
  Tipo: TTipoHojaCalc;    //Witch program was used to manage the doc?
  FileName:    string;    //In windows FileName format C:\MyDoc.XXX
  Programa:    variant;   //Excel or OpenOfice instance created.
  DeskTop:     variant;   //OpenOffice desktop reference (not used now).
  Document:    variant;   //Document opened.
  ActiveSheet: variant;   //Active sheet.
  //Object internals...
  constructor  Create(Name: string; MakeVisible: boolean); overload;
  constructor  Create(MyTipo: TTipoHojaCalc; MakeVisible: boolean); overload;
  destructor   Destroy; override;
  //Program loaded stuff...
  function     SaveDoc: boolean;
  function     SaveDocAs(Name: string; AsExcel97: boolean = false): boolean;
  function     PrintDoc: boolean;
  procedure    ShowPrintPreview;
  property     ProgLoaded: boolean     read GetProgLoaded;
  property     DocLoaded:  boolean     read GetDocLoaded;
  property     IsExcel: boolean        read GetIsExcel;
  property     IsOpenOffice: boolean   read GetIsOpenOffice;
  property     Visible: boolean        read fVisible           write SetVisible;
  property     KeepAlive: boolean      read fKeepAlive         write fKeepAlive;
  //Sheets stuff...
  function     ActivateSheetByIndex(nIndex: integer): boolean;
  function     ActivateSheetByName(SheetName: string; CaseSensitive: boolean): boolean;
  function     IsActiveSheetProtected: boolean;
  function     PrintActiveSheet: boolean;
  function     PrintSheetsUntil(LastSheetName: string): boolean;
  procedure    AddNewSheet(NewName: string);
  property     CountSheets:  integer   read GetCountSheets;
  property     ActiveSheetName: string read GetActiveSheetName write SetActiveSheetName;
  //Cells stuff...
  //Sending numbers and date to a cell, the easy way:
  procedure    SendNumber(row, col: integer; v: double);
  procedure    SendDate(row, col: integer; v: TDate);
  //Properties:
  procedure    Bold(row,col: integer);
  procedure    Italic(row,col: integer);
  procedure    BackgroundColor(row, col: integer; color: TColor);
  procedure    FontSize(row, col, size: integer);
  procedure    ColumnWidth(col, width: integer); //Width in 1/100 of mm.
  //Accesing to the cell content:
  property     CellText[f,c: integer]: string read GetCellText write SetCellText;
  property     CellFormula[f,c: integer]: string read GetCellFormula write SetCellFormula;
  property     CellTextByName[Range: string]: string read GetCellTextByName write SetCellTextByName;
end;

implementation

// ************************
// ** Create and destroy **
// ************************

//Create with an empty doc of requested type (use thcExcel or thcOpenOffice)
//Remember to define FileName before calling to SaveDoc
constructor THojaCalc.Create(MyTipo: TTipoHojaCalc; MakeVisible: boolean);
var
  i: integer;
  IsFirstTry: boolean;
begin
  KeepAlive:= false;
  //Close all opened things first...
  CloseDoc;
  CloseProg;
  //I will try to open twice, so if Excel fails, OpenOffice is used instead
  IsFirstTry:= true;
  for i:= 1 to 2 do begin
    //Try to open Excel...
    if (MyTipo = thcExcel) or (MyTipo = thcNone) then begin
      try
        Programa:= CreateOleObject('Excel.Application');
      except
      end;
      if ProgLoaded then begin
        Tipo:= thcExcel;
        break;
      end else begin
        if IsFirstTry then begin
          //Try OpenOffice as my second choice
          MyTipo:= thcOpenOffice;
          IsFirstTry:= false;
        end else begin
          //Both failed!
          break;
        end;
      end;
    end;
    //Try to open OpenOffice...
    if (MyTipo = thcOpenOffice) or (MyTipo = thcNone)then begin
      try
        Programa:= CreateOleObject('com.sun.star.ServiceManager');
      except
      end;
      if ProgLoaded then begin
        Tipo:= thcOpenOffice;
        break;
      end else begin
        if IsFirstTry then begin
          //Try Excel as my second choice
          MyTipo:= thcExcel;
          IsFirstTry:= false;
        end else begin
          //Both failed!
          break;
        end;
      end;
    end;
  end;
  //Was it able to open any of them?
  if Tipo = thcNone then begin
    Tipo:= thcError;
    raise Exception.Create('THojaCalc.create failed, may be no Office is installed?');
  end;
  //Add a blank document...
  fVisible:= MakeVisible;
  NewDoc;
  //Create an American format to use when sending numbers or dates to excel
  GetLocaleFormatSettings( 0, AmericanFormat);
  AmericanFormat.ThousandSeparator:= ',';
  AmericanFormat.DecimalSeparator:=  '.';
  AmericanFormat.ShortDateFormat:=   'mm/dd/yyyy';
end;

constructor THojaCalc.Create(Name: string; MakeVisible: boolean);
begin
  Tipo:= thcNone;
  //Store values...
  FileName:= Name;
  fVisible:=  MakeVisible;
  //Create an American format to use when sending numbers or dates to excel
  GetLocaleFormatSettings( 0, AmericanFormat);
  AmericanFormat.ThousandSeparator:= ',';
  AmericanFormat.DecimalSeparator:=  '.';
  AmericanFormat.ShortDateFormat:=   'mm/dd/yyyy';
  //Open program and document...
  LoadProg;
  LoadDoc;
end;

destructor THojaCalc.Destroy;
begin
  if not KeepAlive then begin
    CloseDoc;
    CloseProg;
  end;  
  inherited;
end;

// *************************
// ** Loading the program **
// ** Excel or OpenOffice **
// *************************

procedure THojaCalc.LoadProg;
begin
  if ProgLoaded then CloseProg;
  if (UpperCase(ExtractFileExt(FileName))='.XLS') then begin
    //Excel is the primary choice...
    try
      Programa:= CreateOleObject('Excel.Application');
    except end;
    if ProgLoaded then Tipo:= thcExcel;
  end;
  //Not lucky with Excel? Another filetype? Let's go with OpenOffice...
  if Tipo = thcNone then begin
    //Try with OpenOffice...
    try
      Programa:= CreateOleObject('com.sun.star.ServiceManager');
    except end;
    if ProgLoaded then Tipo:= thcOpenOffice;
  end;
  //Still no program loaded?
  if not ProgLoaded then begin
    Tipo:= thcError;
    raise Exception.Create('THojaCalc.create failed, may be no Office is installed?');
  end;
end;

procedure THojaCalc.CloseProg;
begin
  if not Visible then CloseDoc;
  if ProgLoaded then begin
    try
      if IsExcel then      Programa.Quit;
      //Next line made OO V2 not to work anymore as the next call to
      //CreateOleObject('com.sun.star.ServiceManager') failed.
      //if IsOpenOffice then Programa.Dispose;
      Programa:= Unassigned;
    finally end;
  end;
  Tipo:= thcNone;
end;

//Is there any prog loaded? Witch one?
function THojaCalc.GetProgLoaded: boolean;
begin
  result:= not (VarIsEmpty(Programa) or VarIsNull(Programa));
end;
function  THojaCalc.GetIsExcel: boolean;
begin
  result:= (Tipo=thcExcel);
end;
function  THojaCalc.GetIsOpenOffice: boolean;
begin
  result:= (Tipo=thcOpenOffice);
end;

// ************************
// ** Loading a document **
// ************************

procedure THojaCalc.NewDoc;
var ooParams: variant;
begin
  //Is the program running? (Excel or OpenOffice)
  if not ProgLoaded then raise Exception.Create('No program loaded for the new document.');
  //Is there a doc already loaded?
  CloseDoc;
  DeskTop:= Unassigned;
  //OK, now try to create the doc...
  if IsExcel then begin
    Programa.WorkBooks.Add;
    Programa.Visible:= Visible;
    Document:= Programa.ActiveWorkBook;
    ActiveSheet:= Document.ActiveSheet;
  end;
  if IsOpenOffice then begin
    Desktop:=  Programa.CreateInstance('com.sun.star.frame.Desktop');
    //Optional parameters (visible)...
    ooParams:=    VarArrayCreate([0, 0], varVariant);
    ooParams[0]:= ooCreateValue('Hidden', not Visible);
    //Create the document...
    Document:= Desktop.LoadComponentFromURL('private:factory/scalc', '_blank', 0, ooParams);
    ActivateSheetByIndex(1);
  end;
end;

procedure THojaCalc.LoadDoc;
var ooParams: variant;
begin
  if FileName='' then exit;
  //Is the program running? (Excel or OpenOffice)
  if not ProgLoaded then LoadProg;
  //Is there a doc already loaded?
  CloseDoc;
  DeskTop:= Unassigned;
  //OK, now try to open the doc...
  if IsExcel then begin
    Programa.WorkBooks.Open(FileName,3);
    Programa.Visible:= Visible;
    Document:= Programa.ActiveWorkBook;
    ActiveSheet:= Document.ActiveSheet;
  end;
  if IsOpenOffice then begin
    Desktop:=  Programa.CreateInstance('com.sun.star.frame.Desktop');
    //Optional parameters (visible)...
    ooParams:=    VarArrayCreate([0, 0], varVariant);
    //Next line stop working OK on OOo V2: Created blind, always blind!
    //so now it is create as visible, then set to non visible if requested
    //ooParams[0]:= ooCreateValue('Hidden', not Visible);
    ooParams[0]:= ooCreateValue('Hidden', false); //Create as visible, then make it not visible if necesary
    //Open the document...
    Document:= Desktop.LoadComponentFromURL(FileName2URL(FileName), '_blank', 0, ooParams);
    ActiveSheet:= ActivateSheetByIndex(1);
    //If has to be non visible, set it now...
    if not visible then
      Document.getCurrentController.getFrame.getContainerWindow.setVisible(false);
  end;
  if Tipo=thcNone then
    raise Exception.Create('No puedo leer el fichero "'+FileName+'" al no estar presente el programa necesario.');
end;

function THojaCalc.SaveDoc: boolean;
begin
  result:= false;
  if DocLoaded then begin
    if IsExcel then begin
      Document.Save;
      result:= true;
    end;
    if IsOpenOffice then begin
      Document.Store;
      result:= true;
    end;
  end;
end;

//Function added by Massimiliano Gozzi on V0.92
//AsEXcel97 taken form V0.93 by Rômulo Silva Ramos
//Saving as .xls on Excel 2000/2003 trick by Malte Tüllmann on V1.01
function THojaCalc.SaveDocAs(Name: string; AsExcel97: boolean = false): boolean;
var ooParams: variant;
begin
  result:= false;
  if DocLoaded then begin
    if IsExcel then begin
      if (StrToFloat(Programa.Application.Version, AmericanFormat) < 12) then
        //Before Excel 2007 this was the method to force SaveAs Excel97 .xls
        //by Malte Tüllmann on V1.01
        Document.Saveas(Name, -4143, EmptyParam, EmptyParam, EmptyParam, EmptyParam)
      else
        // From Excel 2003 this is the way to force .xls file format (excel8)
        // for back compatibility with older excel version and OO.
        //
        // 51 = xlOpenXMLWorkbook (without macro's in 2007-2010, xlsx)
        // 52 = xlOpenXMLWorkbookMacroEnabled (with or without macro's in 2007-2010, xlsm)
        // 50 = xlExcel12 (Excel Binary Workbook in 2007-2010 with or without macro's, xlsb)
        // 56 = xlExcel8 (97-2003 format in Excel 2007-2010, xls)
        // More on this here: http://www.rondebruin.nl/saveas.htm
        Document.Saveas(Name, 56);
      FileName:= Name;
      result:= true;
    end;
    if IsOpenOffice then begin
      //I may need 1 or 2 params...
      if AsExcel97 then
        ooParams:= VarArrayCreate([0, 1], varVariant)
      else
        ooParams:= VarArrayCreate([0, 0], varVariant);
      //First one for prompting on overwrite (good idea!)
      ooParams[0]:= ooCreateValue('Overwrite', false);
      //Optionally tell OpenOffie to use Excel97 .xls format
      if AsExcel97 then
        ooParams[1]:= ooCreateValue('FilterName', 'MS Excel 97');
      //Do the save!
      Document.StoreAsUrl(FileName2URL(Name), ooParams);
      FileName:= Name;
      result:= true;
    end;
  end;
end;

//Print the Doc...
function THojaCalc.PrintDoc: boolean;
var ooParams: variant;
begin
  result:= false;
  if DocLoaded then begin
    if IsExcel then begin
      Document.PrintOut;
      result:= true;
    end;
    if IsOpenOffice then begin
      //NOTE: OpenOffice will print all sheets with Printable areas, but if no
      //printable areas are defined in the doc, it will print all entire sheets.
      //Optional parameters (wait until fully sent to printer)...
      ooParams:=   VarArrayCreate([0, 0], varVariant);
      ooParams[0]:= ooCreateValue('Wait', true);
      Document.Print(ooParams);
      result:= true;
    end;
  end;
end;

procedure THojaCalc.ShowPrintPreview;
begin
  if DocLoaded then begin
    //Force visibility of the doc...
    Visible:= true;
    if IsExcel then
      Document.PrintOut(,,,true);
    if IsOpenOffice then
      ooDispatch('.uno:PrintPreview', Unassigned);
  end;
end;

procedure THojaCalc.SetVisible(v: boolean);
begin
  if DocLoaded and (v<>fVisible) then begin
    if IsExcel then
      Programa.Visible:= v;
    if IsOpenOffice then
      Document.getCurrentController.getFrame.getContainerWindow.setVisible(v);
    fVisible:= v;
  end;
end;

procedure THojaCalc.CloseDoc;
begin
  if DocLoaded then begin
    //Close it...
    try
      if IsOpenOffice then Document.Dispose;
      if IsExcel      then Document.close;
    finally end;
    //Clean up both "pointer"...
    Document:= Null;
    ActiveSheet:= Null;
  end;
end;

function THojaCalc.GetDocLoaded: boolean;
begin
  result:= not (VarIsEmpty(Document) or VarIsNull(Document));
end;

// *********************
// ** Managing sheets **
// *********************

function THojaCalc.GetCountSheets: integer;
begin
  result:= 0;
  if DocLoaded then begin
    if IsExcel      then result:= Document.Sheets.count;
    if IsOpenOffice then result:= Document.getSheets.GetCount;
  end;
end;

//Index is 1 based in Excel, but OpenOffice uses it 0-based
//Here we asume 1-based so OO needs to activate (nIndex-1)
function THojaCalc.ActivateSheetByIndex(nIndex: integer): boolean;
begin
  result:= false;
  if DocLoaded then begin
    //Exists this sheet number?
    if (nIndex<1) then
      raise Exception.Create('Can not activate sheet #'+IntToStr(nIndex));
    while (nIndex>CountSheets) do begin
      ActivateSheetByIndex(CountSheets);
      AddNewSheet('Nueva hoja '+IntToStr(CountSheets+1));
      sleep(100); //Needs time to do it!
    end;
    //Activate it now...
    if IsExcel then begin
      Document.Sheets[nIndex].activate;
      ActiveSheet:= Document.ActiveSheet;
      result:= true;
    end;
    if IsOpenOffice then begin
      ActiveSheet:= Document.getSheets.getByIndex(nIndex-1);
      result:= true;
    end;
    sleep(100); //Asyncronus, so better give it time to make the change
  end;
end;

//Find a sheet by its name...
function THojaCalc.ActivateSheetByName(SheetName: string; CaseSensitive: boolean): boolean;
var
  OldActiveSheet: variant;
  i: integer;
begin
  result:= false;
  if DocLoaded then begin
    if CaseSensitive then begin
      //Find the EXACT name...
      if IsExcel then begin
        Document.Sheets[SheetName].Select;
        ActiveSheet:= Document.ActiveSheet;
        result:= true;
      end;
      if IsOpenOffice then begin
        ActiveSheet:= Document.getSheets.getByName(SheetName);
        result:= true;
      end;
    end else begin
      //Find the Sheet regardless of the case...
      OldActiveSheet:= ActiveSheet;
      for i:= 1 to GetCountSheets do begin
        ActivateSheetByIndex(i);
        if UpperCase(ActiveSheetName)=UpperCase(SheetName) then begin
          result:= true;
          Exit;
        end;
      end;
      //If not found, let the old active sheet active...
      ActiveSheet:= OldActiveSheet;
    end;
  end;
end;

//Name of the active sheet?
function THojaCalc.GetActiveSheetName: string;
begin
  if DocLoaded then begin
    if IsExcel then
      result:= ActiveSheet.Name;
    if IsOpenOffice then
      result:= ActiveSheet.GetName;
  end;
end;
procedure THojaCalc.SetActiveSheetName(NewName: string);
begin
  if DocLoaded then begin
    //Clean name first...
    NewName := ValidateSheetName(NewName);
    if IsExcel then
      Programa.ActiveSheet.Name:= NewName;
    if IsOpenOffice then begin
      ActiveSheet.setName(NewName);
      //This code always changes the name of "visible" sheet, not active one!
      //ooParams:= VarArrayCreate([0, 0], varVariant);
      //ooParams[0]:= ooCreateValue('Name', NewName);
      //ooDispatch('.uno:RenameTable', ooParams);
    end;
  end;
end;

//Check for sheet protection (password)...
function THojaCalc.IsActiveSheetProtected: boolean;
begin
  result:= false;
  if DocLoaded then begin
    if IsExcel then
      result:= ActiveSheet.ProtectContents;
    if IsOpenOffice then
      result:= ActiveSheet.IsProtected;
  end;
end;

//WARNING: This function is NOT dual, only works for Excel docs!
//Send active sheet to default printer (as seen in preview window)...
function THojaCalc.PrintActiveSheet: boolean;
begin
  result:= false;
  if DocLoaded then begin
    if IsExcel then begin
      ActiveSheet.PrintOut;
      result:= true;
    end;
    if IsOpenOffice then begin
      raise Exception.Create('Function "PrintActiveSheet" still not working in OpenOffice!');
      //ActiveSheet.Print;
      result:= false;
    end;
  end;
end;

//WARNING: This function is NOT dual, only works for Excel docs!
//Select and print sheets from 1 upto -excluded- the one with that name.
//It is interesting for understanding how to pass an array of objects to excel.
function THojaCalc.PrintSheetsUntil(LastSheetName: string): boolean;
var
  i, last: integer;
  Hojas: variant;
begin
  result:= false;
  if DocLoaded then begin
    if IsExcel then begin
      //Macro from Excel:
      //  Sheets(Array("Hoja1", "Hoja2")).Select
      //  ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
      //
      //Witch sheet number correspond to the one previous to "LastSheetName"?
      Last:= 0;
      for i:= 2 to CountSheets do begin
        ActivateSheetByIndex(i);
        if UpperCase(ActiveSheetName)=UpperCase(LastSheetName) then begin
          Last:= i-1;
          break;
        end;
      end;
      //Not found?
      if Last=0 then exit;
      //Create an array of variants -windows standard type- this big...
      Hojas:= VarArrayCreate([1, Last], varVariant);
      //Fill it with the Sheet names...
      for i:= 1 to Last do begin
        ActivateSheetByIndex(i);
        Hojas[i]:= ActiveSheetName;
      end;
      //Print all this array of sheets...
      Programa.Sheets[Hojas].Select;
      Programa.ActiveWindow.SelectedSheets.PrintOut;
      //Done!
      result:= true;
    end;
    if IsOpenOffice then begin
      raise Exception.Create('Function "PrintSheetsUntil" not working in OpenOffice!');
      result:= false;
    end;
  end;
end;

//Add a new sheet, name it, and make it the active sheet...
procedure THojaCalc.AddNewSheet(NewName: string);
var
  ooSheets: variant;
begin
  NewName := ValidateSheetName(NewName);
  if DocLoaded then begin
    if IsExcel then begin
      Document.WorkSheets.Add(null, Document.ActiveSheet, 1);
      Document.ActiveSheet.Name:= NewName;
      //Active sheet has move to this new one, so I need to update the var
      ActiveSheet:= Document.ActiveSheet;
    end;
    if IsOpenOffice then begin
      ooSheets:= Document.getSheets;
      ooSheets.insertNewByName(NewName, 1);
      //Redefine active sheet to this new one
      ActiveSheet:= ooSheets.getByName(NewName);
    end;
  end;
end;

//Clean a Sheet name so it will not cause problems
function THojaCalc.ValidateSheetName(Name: string): string;
begin
  result := StringReplace(name,   ':', '_',[rfReplaceAll]);
  result := StringReplace(result, '/', '_',[rfReplaceAll]);
  result := StringReplace(result, '\', '_',[rfReplaceAll]);
  result := StringReplace(result, '?', '_',[rfReplaceAll]);
  result := StringReplace(result, '*', '_',[rfReplaceAll]);
  result := StringReplace(result, '[', '_',[rfReplaceAll]);
  result := StringReplace(result, ']', '_',[rfReplaceAll]);
  result := StringReplace(result, '"', '_',[rfReplaceAll]);
  if (Trim(result)='') then
    result:= 'Plan' + IntToStr(CountSheets);
  result:= Copy(result, 1, 31);
end;

// ************************
// ** Manage  the  cells **
// ** in the ActiveSheet **
// ************************

//Read/Write cell text (formula en Excel) by index
//OpenOffice start at cell (0,0) while Excel at (1,1)
//Also, Excel uses (row, col) and OpenOffice uses (col, row)
function THojaCalc.GetCellText(row, col: integer): string;
begin
  if DocLoaded then begin
    if IsExcel then      result:= ActiveSheet.Cells[row, col].Text;
    if IsOpenOffice then result:= ActiveSheet.getCellByPosition(col-1, row-1).getFormula;
  end;
end;
procedure  THojaCalc.SetCellText(row, col: integer; Txt: string);
begin
  if DocLoaded then begin
    if IsExcel then begin
      ActiveSheet.Cells[row, col].Select;
      Programa.ActiveCell.Value:= Txt;
    end;
    if IsOpenOffice then
      ActiveSheet.getCellByPosition(col-1, row-1).setFormula(Txt);
  end;
end;
//Same, but assuming Txt contains a formula, like '=12.00'
function THojaCalc.GetCellFormula(row, col: integer): string;
begin
  if DocLoaded then begin
    if IsExcel then      result:= ActiveSheet.Cells[row, col].Formula;
    if IsOpenOffice then result:= ActiveSheet.getCellByPosition(col-1, row-1).getFormula;
  end;
end;
procedure  THojaCalc.SetCellFormula(row, col: integer; Txt: string);
var Rango: string;
begin
  if DocLoaded then begin
    if IsExcel then begin
      Rango:= Programa.Range[ActiveSheet.Cells[row, col], ActiveSheet.Cells[row, col]].Address;
      Programa.Range[Rango].Formula:= Txt;
    end;
    if IsOpenOffice then
      ActiveSheet.getCellByPosition(col-1, row-1).setFormula(Txt);
  end;
end;

//Sending numbers and date to a cell, the easy way:
procedure THojaCalc.SendNumber(row, col: integer; v: double);
begin
  //The only way to efectively send a number to excel is by setting the formula
  //to something like '=12323.50', where the number must be formated like this
  //example. If you try to send it via CellText, even using the correct format,
  //excel will consider it a text (not always) instead of a number, so you can't
  //operate on those "texts".
  //It is weird, excel sais "this number is stored as a text"... WTF! If you
  //click on the formula edit and just press enter, it is converted to number,
  //or if you copy a number 1, then select the cells and do a special paste
  //multiplying... can't understand the point for this.
  //
  //OpenOffice is happy just recieving a text with the format '12323.50', but
  //notice that the SetCellText procedure send the text as a formula instead of
  //a text in OpenOffice, as it always work OK, with or without the '=' char.
  //
  CellFormula[row, col]:= '='+Format('%f', [v], AmericanFormat);
  //Note: format string '%f' doesn't show any ThousanSeparator, don't use '%n'!
end;
procedure THojaCalc.SendDate(row, col: integer; v: TDate);
begin
  //Excel and OO need to recieve dates as CellText in American format MM/DD/YYYY
  CellText[row, col]:= FormatDateTime('ddddd', v, AmericanFormat);
  //OpenOffice need to be set to Date Format or it will show an integer
  if IsOpenOffice then begin
    //37 = Short date format, usually DD/MM/YY
    //36 = Long date format, usually DD/MM/YYYY
    ActiveSheet.getCellByPosition(col-1, row-1).NumberFormat:= 36;
    //
    //Note: This code format all selected cells and is NOT tested
    //
    //var ooParams: variant;
    //
    //ooParams:= VarArrayCreate([0, 0], varVariant);
    //ooParams[0]:= ooCreateValue('NumberFormatValue', 37);
    //ooDispatch('NumberFormatValue', ooParams);
    //
  end;
end;

//Read/Write cell text (formula in OO) by name instead of position
//For instance, you can set the value for cell 'NewSheet!A12' or similar
//NOTE: If range contains several cells, first one will be used.
function THojaCalc.GetCellTextByName(Range: string): string;
var OldActiveSheet: variant;
begin
  if DocLoaded then begin
    if IsExcel then begin
      result:=  Programa.Range[Range].Text; //Set 'Formula' but Get 'Text';
    end;
    if IsOpenOffice then begin
      OldActiveSheet:= ActiveSheet;
      //If range is in the form 'NewSheet!A1' then first change sheet to 'NewSheet'
      if pos('!', Range) > 0 then begin
        //Activate the proper sheet...
        if not ActivateSheetByName(Copy(Range, 1, pos('!', Range)-1), false) then
          raise Exception.Create('Sheet "'+Copy(Range, 1, pos('!', Range)-1)+'" not present in the document.');
        Range:= Copy(Range, pos('!', Range)+1, 999);
      end;
      result:= ActiveSheet.getCellRangeByName(Range).getCellByPosition(0,0).getFormula;
      ActiveSheet:= OldActiveSheet;
    end;
  end;
end;
procedure  THojaCalc.SetCellTextByName(Range: string; Txt: string);
var OldActiveSheet: variant;
begin
  if DocLoaded then begin
    if IsExcel then begin
      Programa.Range[Range].formula:= Txt;
    end;
    if IsOpenOffice then begin
      OldActiveSheet:= ActiveSheet;
      //If range is in the form 'NewSheet!A1' then first change sheet to 'NewSheet'
      if pos('!', Range) > 0 then begin
        //Activate the proper sheet...
        if not ActivateSheetByName(Copy(Range, 1, pos('!', Range)-1), false) then
          raise Exception.Create('Sheet "'+Copy(Range, 1, pos('!', Range)-1)+'" not present in the document.');
        Range:= Copy(Range, pos('!', Range)+1, 999);
      end;
      ActiveSheet.getCellRangeByName(Range).getCellByPosition(0,0).SetFormula(Txt);
      ActiveSheet:= OldActiveSheet;
    end;
  end;
end;

procedure THojaCalc.BackgroundColor(row,col: integer; color: TColor);
begin
  if DocLoaded then begin
    if IsExcel then begin
      Programa.ActiveSheet.Cells[row,col].Interior.Color := color;
    end;
    if IsOpenOffice then begin
      ActiveSheet.getCellByPosition(col-1, row-1).CellBackColor := color;
    end;
  end;
end;

procedure THojaCalc.FontSize(row, col, size: integer);
begin
  if DocLoaded then begin
    if IsExcel then begin
      Programa.ActiveSheet.Cells[row,col].Font.Size := size;
    end;
    if IsOpenOffice then begin
      ActiveSheet.getCellByPosition(col-1, row-1).getText.createTextCursor.CharHeight:= size;
    end;
  end;
end;

procedure THojaCalc.Bold(row,col: integer);
const ooBold: integer = 150; //150 = com.sun.star.awt.FontWeight.BOLD
begin
  if DocLoaded then begin
    if IsExcel then begin
      Programa.ActiveSheet.Cells[row,col].Font.Bold:= true;
    end;
    if IsOpenOffice then begin
      ActiveSheet.getCellByPosition(col-1, row-1).getText.createTextCursor.CharWeight:= ooBold;
    end;
  end;
end;

procedure THojaCalc.Italic(row,col: integer);
const ooItalic = 2; //2 = com.sun.star.awt.FontSlant.ITALIC
begin
  if DocLoaded then begin
    if IsExcel then begin
      Programa.ActiveSheet.Cells[row,col].Font.Italic:= true;
    end;
    if IsOpenOffice then begin
      ActiveSheet.getCellByPosition(col-1, row-1).getText.createTextCursor.CharPosture:= ooItalic;
    end;
  end;
end;

procedure THojaCalc.ColumnWidth(col, width: integer); //Width in 1/100 of mm.
const
  MaxWidthExcel: integer = 255;
begin
  if DocLoaded then begin
    if IsExcel then begin
      //Excel use the width of '0' as the unit, we do an aproximation: Width '0' = 3 mm.
      Width:= round(Width/100/3);
      if Width > MaxWidthExcel then
        Width:= MaxWidthExcel;
      Programa.ActiveSheet.Cells[1, col].ColumnWidth:= Width;
    end;
    if IsOpenOffice then begin
      ActiveSheet.getCellByPosition(col-1, 0).getColumns.getByIndex(0).Width:= width;
    end;
  end;
end;

// ***************************
// ** OpenOffice only stuff **
// ***************************

//Change 'C:\File.txt' into 'file:///c:/File.txt' (for OpenOffice OpenURL)
function THojaCalc.FileName2URL(FileName: string): string;
begin
  result:= '';
  if LowerCase(copy(FileName,1,8))<>'file:///' then
    result:= 'file:///';
  result:= result + StringReplace(FileName, '\', '/', [rfReplaceAll, rfIgnoreCase]);
end;

function THojaCalc.ooCreateValue(ooName: string; ooData: variant): variant;
var
  ooReflection: variant;
begin
  if IsOpenOffice then begin
    ooReflection:= Programa.createInstance('com.sun.star.reflection.CoreReflection');
    ooReflection.forName('com.sun.star.beans.PropertyValue').createObject(result);
    result.Name := ooName;
    result.Value:= ooData;
  end else begin
    raise Exception.Create('ooValue imposible to create, load OpenOffice first!');
  end;
end;

procedure THojaCalc.ooDispatch(ooCommand: string; ooParams: variant);
var
  ooDispatcher, ooFrame: variant;
begin
  if DocLoaded and IsOpenOffice then begin
    if (VarIsEmpty(ooParams) or VarIsNull(ooParams)) then
      ooParams:= VarArrayCreate([0, -1], varVariant);
    ooFrame:= Document.getCurrentController.getFrame;
    ooDispatcher:= Programa.createInstance('com.sun.star.frame.DispatchHelper');
    ooDispatcher.executeDispatch(ooFrame, ooCommand, '', 0, ooParams);
  end else begin
    raise Exception.Create('Dispatch imposible, load a OpenOffice doc first!');
  end;
end;

end.
