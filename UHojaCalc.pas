{[w=1-125,e=3,k+] for formatter} //
{[f-] stop formatting}
// *******************************************************
// ** Delphi object for dual SpreadSheet managing using **
// ** Excel/OpenOffice/LibreOffice in a transparent way **
// ** By: Sergio Hernandez (oficina(at)hcsoft.net)      **
// ** Version 1.08 18-10-2013 (DDMMYYYY)                **
// ** Use it freely, change it, etc. at will.           **
// *******************************************************

//Latest version, questions, modifications:
//
// https://github.com/sergio-hcsoft/Delphi-SpreadSheets
// http://user.services.openoffice.org/en/forum/viewtopic.php?f=21&t=47644&p=219641
// http://forum.openoffice.org/en/forum/viewtopic.php?f=21&t=47644&p=288656#p219641

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
  IF HCalc.CellText[i,2] = '' THEN HCalc.CellText[i,2] := 'Hello world!';
  HCalc.AddNewSheet('New Sheet');
  HCalc.PrintDoc;
  HCalc.SaveDoc;
  HCalc.Free;
}

{TODO LIST:
  -PrintActiveSheet is not working for OpenOffice/LibreOffice (even possible?)
}

{CHANGE LOG:
 V1.08: (18-10-2013 DD/MM/YYY)
   ***************************
   ** By user MARCELVK from **
   ** forum.openoffice.org  **
   ***************************
   -SetTextCell in OpenOffice/LibreOfice case use .string not setFormula().
   -Added properties LastCol and LastRow to get the bounds of used cells.
 V1.07: (15-05-2013 DD/MM/YYYY)
   -From V1.03, trying to open Excel without Excel installed doesn't try to open
   OO instead, just raise an error. Fixed in create().
 V1.06: (08-04-2013 DD/MM/YYYY)
   *******************
   ** Joseph Gordon **
   *******************
   -New function Orientation(row, Col, Angle) to rotate the text in a cell.
   -Auto adjust a column's width using AutoFit(col)
 V1.05: (22-02-2013 DDMMYYYY)
   -Restored "$INCLUDE Compilers.inc" from V1.03 so code is suitable for other
   versions of delphi (Philipe did this works, I just deleted this line ;-).
   -Restored 3 commented lines with params. for a code formatter Philipe use. It
   has no use for others don't using formatters, but it won't harm us!
 V1.04:
   -New function StillConnectedToApp() to check if user closed app. manually.
   Note: Useful for previewing doc. in OO using code like this:
   //
   HCalc.Visible:= true;
   if HCalc.IsExcel then begin
     //Preview of all sheets, one after the user closes the other...
     for i:= 1 to HCalc.Document.Sheets.count do
       HCalc.Document.Sheets[i].PrintOut(,,,true);
   end else begin
     //ooDispatcher is just a variant
     ooDispatcher:= HCalc.Programa.createInstance('com.sun.star.frame.DispatchHelper');
     ooDispatcher.executeDispatch(HCalc.Document.getCurrentController.getFrame, '.uno:PrintPreview', '', 0, VarArrayCreate([0, -1], varVariant));
     //OO returns control just after showing, while excel waits for user to close it.
     //If you don't wait for user to close preview, you will see just a flash:
     while HCalc.StillConnectedToApp() do
       sleep(1000);
     //User has manually closed the preview window at this point.
   end;
   //
 V1.03:
   *****************************
   ** Van Coppenolle Philippe **
   *****************************
   -Reformated code and renamed vars. with more TRY-EXCEPT zones, more robust.
   -New function to encapsulate creation of OLE object ConnectToApp()
   -Create admits new flag to reuse the last created instance of excel.
   -NewDoc admits new flag to add or not a sheet to the new doc (related to the previous one).
   -RemoveSheet added in 3 flavours: by index, byname, and all except one name.
   -Added list of OLE constant for both Excel and OO so you can use them by name in your code.
   -Added FontColor, Underline and HorizontalAlignment properties to cells.
   Note: I added some properties so old vars names are still usable:
         Document, ActiveSheet and Programa.
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
   CellFormula(1,1) := '=A2*23211.66';
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
         Programa     := CreateOleObject('com.sun.star.ServiceManager');
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
{[f?] restore formatting}

UNIT UHojaCalc;

//Find this file searching on google, or just try here:
// https://code.google.com/p/virtual-treeview/source/browse/trunk/Common/Compilers.inc?r=235
{$INCLUDE Compilers.inc}

INTERFACE

USES
  Windows,
  {$IFDEF COMPILER_6_UP}
  Variants,
  {$ELSE}
  {$ENDIF}
  SysUtils, ComObj, Classes, Graphics, Controls;

CONST
  {--------------------------------------------------------------------------------}
  { Excel constants ( OLE support )                                                }
  { http://www.koders.com/noncode/fid793D4B61A1DF009ACD6544001B50528A598EB275.aspx }
  {--------------------------------------------------------------------------------}
  { XlApplicationInternational }
   xlCountryCode = 1;
   xlCountrySetting = 2;
   xlDecimalSeparator = 3;
   xlThousandsSeparator = 4;
   xlListSeparator = 5;
   xlUpperCaseRowLetter = 6;
   xlUpperCaseColumnLetter = 7;
   xlLowerCaseRowLetter = 8;
   xlLowerCaseColumnLetter = 9;
   xlLeftBracket = 10;
   xlRightBracket = 11;
   xlLeftBrace = 12;
   xlRightBrace = 13;
   xlColumnSeparator = 14;
   xlRowSeparator = 15;
   xlAlternateArraySeparator = 16; //(&H10)
   xlDateSeparator = 17; //(&H11)
   xlTimeSeparator = 18; //(&H12)
   xlYearCode = 19; //(&H13)
   xlMonthCode = 20; //(&H14)
   xlDayCode = 21; //(&H15)
   xlHourCode = 22; //(&H16)
   xlMinuteCode = 23; //(&H17)
   xlSecondCode = 24; //(&H18)
   xlCurrencyCode = 25; //(&H19)
   xlGeneralFormatName = 26; //(&H1A)
   xlCurrencyDigits = 27; //(&H1B)
   xlCurrencyNegative = 28; //(&H1C)
   xlNocurrencyDigits = 29; //(&H1D)
   xlMonthNameChars = 30; //(&H1E)
   xlWeekdayNameChars = 31; //(&H1F)
   xlDateOrder = 32; //(&H20)
   xl24HourClock = 33; //(&H21)
   xlNonEnglishFunctions = 34; //(&H22)
   xlMetric = 35; //(&H23)
   xlCurrencySpaceBefore = 36; //(&H24)
   xlCurrencyBefore = 37; //(&H25)
   xlCurrencyMinusSign = 38; //(&H26)
   xlCurrencyTrailingZeros = 39; //(&H27)
   xlCurrencyLeadingZeros = 40; //(&H28)
   xlMonthLeadingZero = 41; //(&H29)
   xlDayLeadingZero = 42; //(&H2A)
   xl4DigitYears = 43; //(&H2B)
   xlMDY = 44; //(&H2C)
   xlTimeLeadingZero = 45; //(&H2D)
   xlCellTypeLastCell = 11;
   
  { typedef enum XlVAlign }
  xlVAlignBottom = -4107;
  xlVAlignCenter = -4108;
  xlVAlignDistributed = -4117;
  xlVAlignJustify = -4130;
  xlVAlignTop = -4160;

  { typedef enum XlHAlign }
  xlHAlignCenter = - 4108;
  xlHAlignCenterAcrossSelection = 7;
  xlHAlignDistributed = -4117;
  xlHAlignFill = 5;
  xlHAlignGeneral = 1;
  xlHAlignJustify = -4130;
  xlHAlignLeft = -4131;
  xlHAlignRight = - 4152;

  { XlSheetType }
  xlChart = - 4109;
  xlDialogSheet = - 4116;
  xlExcel4IntlMacroSheet = 4;
  xlExcel4MacroSheet = 3;
  xlWorksheet = - 4167;

  { XlWBATemplate }
  xlWBATChart = - 4109;
  xlWBATExcel4IntlMacroSheet = 4;
  xlWBATExcel4MacroSheet = 3;
  xlWBATWorksheet = - 4167;

  { XlUnderlineStyle }
  xlUnderlineStyleNone = - 4142; // (&HFFFFEFD2)
  xlUnderlineStyleSingle = 2;
  xlUnderlineStyleDouble = - 4119; // (&HFFFFEFE9)
  xlUnderlineStyleSingleAccounting = 4;
  xlUnderlineStyleDoubleAccounting = 5;

  {------------------------------------------------------------------------------}
  { OpenOffice constants ( OLE support )                                         }
  {------------------------------------------------------------------------------}
  { NumberFormat, see                                                                 }
  { http://www.openoffice.org/api/docs/common/ref/com/sun/star/util/NumberFormat.html }
  nfALL = 0;  // Description selects all number formats. 
  nfDEFINED = 1;  // Description selects only user-defined number formats. 
  nfDATE = 2;  // Description selects date formats. 
  nfTIME = 4;  // Description selects time formats.
  nfCURRENCY = 8;  // Description selects currency formats. 
  nfNUMBER = 16;  // Description selects decimal number formats.
  nfSCIENTIFIC = 32;  // Description selects scientific number formats. 
  nfFRACTION = 64;  // Description selects number formats for fractions.
  nfPERCENT = 128;  // Description selects percentage number formats. 
  nfTEXT = 256;  // Description selects text number formats. 
  nfDATETIME = 6;  // Description selects number formats which contain date and time. 
  nfLOGICAL = 1024;  // Description selects boolean number formats.
  nfUNDEFINED = 2048;  // Description is used as a return value if no format exists.

TYPE
  { thcError: Tried to open but both fails }
  { thcNone:  Haven't tried still to open any }
  TTipoHojaCalc =
                    (thcError, thcNone, thcExcel, thcOpenOffice);
  TOoUnderlineStyle =
                       (ulNone,           // =  0
                        ulSingle,         // =  1
                        ulDouble,         // =  2
                        ulDotted,         // =  3
                        ulUndef4,         // =  4
                        ulDash,           // =  5
                        ulLongDash,       // =  6
                        ulDashDot,        // =  7
                        ulDashDotDot,     // =  8
                        ulWave,           // =  9
                        ulSmallWave,      // = 10
                        ulDoubleWave,     // = 11
                        ulBold,           // = 12
                        ulBoldDotted,     // = 13
                        ulBoldDash,       // = 14
                        ulBoldLongDash,   // = 15
                        ulBoldDashDot,    // = 16
                        ulBoldDashDotDot, // = 17
                        ulBold_wave       // = 18
                       );
  THojaCalc = CLASS (TObject)
              PRIVATE
                m_bVisible: boolean;
                m_bKeepAlive: boolean;
                m_eTipo: TTipoHojaCalc; //Which program was used to manage the doc?
                m_strFileName: string; //In windows FileName format C:\MyDoc.XXX
                m_bReUseExisting: boolean; // re-use an existing instance of the program
                m_bFirstAddedSheet: boolean; // becomes false after adding a new sheet
                m_vActiveSheet: variant; //Active sheet.
                m_vPrograma: variant; //Excel or OpenOfice instance created.
                m_vDocument: variant; //Document opened.
              {$IFDEF COMPILER_7_UP}
                m_AmericanFormat: TFormatSettings;
              {$ENDIF COMPILER_7_UP} //
                //Object internals...
                FUNCTION ConnectToApp (eMyTipo, eReqTipo: TTipoHojaCalc; bReUseExisting: boolean): TTipoHojaCalc;
                //Program loaded stuff...
                PROCEDURE LoadProg;
                PROCEDURE CloseProg;
                FUNCTION GetProgLoaded: boolean;
                PROCEDURE NewDoc (bAddNewSheet: boolean);
                PROCEDURE LoadDoc;
                PROCEDURE CloseDoc;
                FUNCTION GetDocLoaded: boolean;
                FUNCTION GetIsExcel: boolean;
                FUNCTION GetIsOpenOffice: boolean;
                PROCEDURE SetVisible (v: boolean); //
                //Sheets stuff..
                FUNCTION GetCountSheets: integer;
                FUNCTION GetActiveSheetName: string;
                PROCEDURE SetActiveSheetName (strNewName: string); //
                function  GetLastRow: integer;
                function  GetLastCol: integer;
                //Cells stuff...
                //General input/output of cell content:
                FUNCTION GetCellText (row, col: integer): string;
                PROCEDURE SetCellText (row, col: integer; strTxt: string);
                FUNCTION GetCellFormula (row, col: integer): string;
                PROCEDURE SetCellFormula (row, col: integer; strTxt: string);
                FUNCTION GetCellTextByName (strRange: string): string;
                PROCEDURE SetCellTextByName (strRange: string; strTxt: string); //
                //OpenOffice only stuff...
                FUNCTION FileName2URL (strFileName: string): string;
                PROCEDURE ooDispatch (strOoCommand: string; vOoParams: variant);
                FUNCTION ooCreateValue (strOoName: string; vOoData: variant): variant; //
                //Aux functions
                FUNCTION ValidateSheetName (strName: string): string;
              PUBLIC
                m_vDeskTop: variant; //OpenOffice desktop reference.
                //Object internals...
                CONSTRUCTOR Create (strName: string; bMakeVisible: boolean; bReUseExisting: boolean = false); overload;
                CONSTRUCTOR Create (eMyTipo: TTipoHojaCalc; bMakeVisible: boolean; bReUseExisting: boolean = false); overload;
                DESTRUCTOR Destroy; override; //
                FUNCTION StillConnectedToApp: boolean;
                //Program loaded stuff...
                FUNCTION SaveDoc: boolean;
                FUNCTION SaveDocAs (strName: string; bAsExcel97: boolean = false): boolean;
                FUNCTION PrintDoc: boolean;
                PROCEDURE ShowPrintPreview;
                PROPERTY Programa: variant read m_vPrograma write m_vPrograma;
                PROPERTY ProgLoaded: boolean read GetProgLoaded;
                PROPERTY Document: variant read m_vDocument write m_vDocument;
                PROPERTY DocLoaded: boolean read GetDocLoaded;
                PROPERTY IsExcel: boolean read GetIsExcel;
                PROPERTY IsOpenOffice: boolean read GetIsOpenOffice;
                PROPERTY Visible: boolean read m_bVisible write SetVisible;
                PROPERTY KeepAlive: boolean read m_bKeepAlive write m_bKeepAlive;
                //Sheets stuff...
                FUNCTION ActivateSheetByIndex (nIndex: integer): boolean;
                FUNCTION ActivateSheetByName (strSheetName: string; bCaseSensitive: boolean): boolean;
                FUNCTION IsActiveSheetProtected: boolean;
                FUNCTION PrintActiveSheet: boolean;
                FUNCTION PrintSheetsUntil (strLastSheetName: string): boolean;
                PROCEDURE AddNewSheet (strNewName: string; bRemoveDummySheets: boolean = false);
                PROCEDURE RemoveSheetByName (strOldName: string);
                PROCEDURE RemoveSheetByIndex (nIndex: integer);
                PROCEDURE RemoveAllSheetsExcept (strOldName: string; bCaseSensitive: boolean);
                PROPERTY eTipo: TTipoHojaCalc read m_eTipo {write m_eTipo};
                PROPERTY FileName: string read m_strFileName write m_strFileName;
                PROPERTY CountSheets: integer read GetCountSheets;
                PROPERTY ActiveSheet: variant read m_vActiveSheet write m_vActiveSheet;
                PROPERTY ActiveSheetName: string read GetActiveSheetName write SetActiveSheetName;
                PROPERTY FirstAddedSheet: boolean read m_bFirstAddedSheet; //
                property LastRow: integer read GetLastRow;
                property LastCol: integer read GetLastCol;
                //Cells stuff...
                //Sending numbers and date to a cell, the easy way:
                PROCEDURE SendNumber (row, col: integer; v: double);
                PROCEDURE SendDate (row, col: integer; v: TDate); //Properties:
                PROCEDURE Bold (row, col: integer);
                PROCEDURE Italic (row, col: integer);
                PROCEDURE Underline (row, col: integer; eOoUnderlineStyle: TOoUnderlineStyle);
                PROCEDURE FontColor (row, col: integer; color: TColor);
                PROCEDURE BackgroundColor (row, col: integer; color: TColor);
                PROCEDURE FontSize (row, col, size: integer);
                PROCEDURE HorizontalAlignment (row, col: integer; ha: TAlignment);
                PROCEDURE ColumnWidth (col, width: integer); //Width in 1/100 of mm.
                PROCEDURE NumberFormat (col, width: integer; strNumberFormat: string);
                PROCEDURE Orientation(row,Col: integer; Angle: integer);
                PROCEDURE AutoFit(col: integer); //AutoFix/OptimumWidth
                //Accesing to the cell content:
                PROPERTY CellText[f, c: integer]: string read GetCellText write SetCellText;
                PROPERTY CellFormula[f, c: integer]: string read GetCellFormula write SetCellFormula;
                PROPERTY CellTextByName[Range: string]: string read GetCellTextByName write SetCellTextByName;
                //Aux functions
                FUNCTION SwapColor (nColor: TColor): TColor;
              END {THojaCalc};


IMPLEMENTATION //

CONST
  strOleExcel = 'Excel.Application';
  strOleOpenOffice = 'com.sun.star.ServiceManager';
  strOleOoDesktop = 'com.sun.star.frame.Desktop';

{ ************************ }
{ ** Create and destroy ** }
{ ************************ }


CONSTRUCTOR THojaCalc.Create (eMyTipo: TTipoHojaCalc; bMakeVisible: boolean; bReUseExisting: boolean = false);

{ Create with an empty doc of requested type (use thcExcel or thcOpenOffice) }
{ Remember to define FileName before calling to SaveDoc }

  VAR
    nTryal: integer;


  BEGIN
    m_bKeepAlive := false; //
  //Close all opened things first...
    CloseDoc;
    CloseProg; //
  //I will try to open twice, so if Excel fails, OpenOffice is used instead
    m_eTipo := thcNone;
    m_bReUseExisting := bReUseExisting;
    FOR nTryal := 1 TO 2 DO
      BEGIN //
      //Try to open Excel...
        m_eTipo := ConnectToApp(eMyTipo, thcExcel, bReUseExisting);
        IF m_eTipo = thcExcel THEN
          break;
      //Try to open OpenOffice...
        m_eTipo := ConnectToApp(eMyTipo, thcOpenOffice, bReUseExisting);
        IF m_eTipo = thcOpenOffice THEN
          break;
      //Unlucky? Then let it use whatever it finds on the second try:
        eMyTipo:= thcNone;
      END {FOR}; //

  { Was it able to open any of them? }
    IF eTipo = thcNone THEN
      BEGIN
        m_eTipo := thcError;
        RAISE Exception.Create('THojaCalc.create failed, may be no Office is installed?');
      END {IF}; //

  { Add a blank document... }
    m_bVisible := bMakeVisible;
    NewDoc(NOT m_bReUseExisting); // Do NOT add a new sheet an existing WB must be re-used
  //Create an American format to use when sending numbers or dates to excel
  {$IFDEF COMPILER_12_UP}
    m_AmericanFormat := TFormatSettings.Create(Windows.LOCALE_NEUTRAL);

  {$ELSE}
  {$IFDEF COMPILER_8_UP}
    GetLocaleFormatSettings(Windows.LOCALE_NEUTRAL, m_AmericanFormat);

  {$ELSE}

  {$ENDIF}
  {$ENDIF}
  {$IFDEF COMPILER_7_UP}
    GetLocaleFormatSettings( 0, m_AmericanFormat);
    m_AmericanFormat.ThousandSeparator := ',';
    m_AmericanFormat.DecimalSeparator := '.';
    m_AmericanFormat.ShortDateFormat := 'mm/dd/yyyy';

  {$ELSE}
  //Will be updated where needed, as they must be saved before and restored afterwards
  //SysUtils.ThousandSeparator := ',';
  //SysUtils.DecimalSeparator := '.';
  //SysUtils.ShortDateFormat := 'mm/dd/yyyy';

  {$ENDIF}
  END {THojaCalc.Create};


CONSTRUCTOR THojaCalc.Create (strName: string; bMakeVisible: boolean; bReUseExisting: boolean = false);


  BEGIN
    m_eTipo := thcNone; //
  //Store values...
    m_strFileName := strName;
    m_bVisible := bMakeVisible; //
  //Create an American format to use when sending numbers or dates to excel
  {$IFDEF COMPILER_12_UP}
    m_AmericanFormat := TFormatSettings.Create(Windows.LOCALE_NEUTRAL);

  {$ELSE}
  {$IFDEF COMPILER_8_UP}
    GetLocaleFormatSettings(Windows.LOCALE_NEUTRAL, m_AmericanFormat);

  {$ELSE}

  {$ENDIF}
  {$ENDIF}
  {$IFDEF COMPILER_7_UP}
  //Will be updated where needed, as they must be saved before and restored afterwards
  //SysUtils.ThousandSeparator := ',';
  //SysUtils.DecimalSeparator := '.';
  //SysUtils.ShortDateFormat := 'mm/dd/yyyy';

  {$ELSE}

  {$ENDIF} //
  //Open program and document...
    LoadProg;
    LoadDoc;
  END {THojaCalc.Create};


FUNCTION THojaCalc.ConnectToApp (eMyTipo, eReqTipo: TTipoHojaCalc; bReUseExisting: boolean): TTipoHojaCalc;

  VAR
    strOleName: string;


  BEGIN
    result := thcNone;
    CASE eReqTipo OF
      thcExcel:
        strOleName := strOleExcel;
      thcOpenOffice:
        strOleName := strOleOpenOffice;
      ELSE
        RAISE Exception.Create('THojaCalc.ConnectToApp failed, invalid requested type');
    END {CASE};
    IF (eMyTipo = thcNone) OR (eMyTipo = eReqTipo) THEN
      BEGIN
        IF bReUseExisting THEN
          TRY
            TRY
              m_vPrograma := GetActiveOleObject(strOleName);
            EXCEPT
            END {TRY};
          FINALLY
          END {TRY};
        IF NOT ProgLoaded THEN
          TRY
            TRY
              m_vPrograma := CreateOleObject(strOleName);
            EXCEPT
            END {TRY};
          FINALLY
          END {TRY};
        IF ProgLoaded THEN
          result := eReqTipo;
      END {IF};
  END {THojaCalc.ConnectToApp};

//After you call a preview, for instance, you can check if the user closed the doc.
FUNCTION THojaCalc.StillConnectedToApp: boolean;

  //VAR
  //  strOleName: string;
  //  tmp_Programa: variant;

  BEGIN
    result := false;
    CASE m_eTipo OF
      thcExcel:
        TRY
          GetActiveOleObject(strOleExcel);
          result:= true;
        EXCEPT
        END {TRY};
      thcOpenOffice:
        TRY
          m_vDocument.getCurrentController.getFrame.getContainerWindow;
          result:= true;
        EXCEPT
        END {TRY};
      ELSE
        exit;
    END; {CASE};

    TRY

    FINALLY
    END {TRY};

    //result:= NOT (VarIsEmpty(tmp_Programa) OR VarIsNull(tmp_Programa));
  END {THojaCalc.StillConnectedToApp};


DESTRUCTOR THojaCalc.Destroy;


  BEGIN
    IF NOT m_bKeepAlive THEN
      TRY
        TRY
          CloseDoc;
        EXCEPT
        END;
      FINALLY
        CloseProg;
      END {IF};
    INHERITED;
  END {THojaCalc.Destroy};

{ ************************* }
{ ** Loading the program ** }
{ ** Excel or OpenOffice ** }
{ ************************* }


PROCEDURE THojaCalc.LoadProg;


  BEGIN
    IF ProgLoaded THEN
      CloseProg;
    m_eTipo := thcNone;
    IF (UpperCase(ExtractFileExt(m_strFileName)) = '.XLS') THEN
      BEGIN //Excel is the primary choice...
        m_eTipo := ConnectToApp(thcNone, thcExcel, m_bReUseExisting);
      END {IF}; //
  //Not lucky with Excel? Another filetype? Let's go with OpenOffice...
    IF eTipo = thcNone THEN
      BEGIN //Try with OpenOffice...
        m_eTipo := ConnectToApp(thcNone, thcOpenOffice, m_bReUseExisting);
      END {IF}; //
  //Still no program loaded?
    IF NOT ProgLoaded THEN
      BEGIN
        m_eTipo := thcError;
        RAISE Exception.Create('THojaCalc.create failed, may be no Office is installed?');
      END {IF};
  END {THojaCalc.LoadProg};


PROCEDURE THojaCalc.CloseProg;


  BEGIN
    IF NOT Visible THEN
      CloseDoc;
    IF ProgLoaded THEN
      BEGIN
        TRY
          IF IsExcel THEN
            m_vPrograma.Quit; //
        //Next line made OO V2 not to work anymore as the next call to
        //CreateOleObject(strOleOpenOffice) failed.
        //IF IsOpenOffice THEN Programa.Dispose;
          m_vPrograma := Unassigned;
        FINALLY
        END {TRY};
      END {IF};
    m_eTipo := thcNone;
  END {THojaCalc.CloseProg};


FUNCTION THojaCalc.GetProgLoaded: boolean;

{ Is there any prog loaded? Which one? }


  BEGIN
    result := NOT (VarIsEmpty(m_vPrograma) OR VarIsNull(m_vPrograma));
  END {THojaCalc.GetProgLoaded};


FUNCTION THojaCalc.GetIsExcel: boolean;


  BEGIN
    result := (m_eTipo = thcExcel);
  END {THojaCalc.GetIsExcel};


FUNCTION THojaCalc.GetIsOpenOffice: boolean;


  BEGIN
    result := (m_eTipo = thcOpenOffice);
  END {THojaCalc.GetIsOpenOffice};

{ ************************ }
{ ** Loading a document ** }
{ ************************ }


PROCEDURE THojaCalc.NewDoc (bAddNewSheet: boolean);

  VAR
    vOoParams: variant;
    AttrPtr: pointer;


  BEGIN //
  //Is the program running? (Excel or OpenOffice)
    IF NOT ProgLoaded THEN
      RAISE Exception.Create('No program loaded for the new document.'); //
  //Is there a doc already loaded?
    CloseDoc;
    m_vDeskTop := Unassigned; //
  //OK, now try to create the doc...
    IF IsExcel THEN
      BEGIN
        m_vDocument := Unassigned;
        IF m_bReUseExisting  THEN
          BEGIN
            m_vDocument := m_vPrograma.ActiveWorkBook;
            AttrPtr := TVarData(m_vDocument).VDispatch;
            IF NOT assigned(AttrPtr) THEN
              m_vDocument := Unassigned;
          END {IF};
        IF (VarIsEmpty(m_vDocument) OR VarIsNull(m_vDocument))
        THEN // no WorkBook present --> create a new one, including sheets
          BEGIN
            m_vDocument := m_vPrograma.WorkBooks.Add;
            m_bFirstAddedSheet := true;
          END {IF}
        ELSE // Workbook present --> create only a new sheet
          BEGIN
            IF bAddNewSheet THEN
              m_vActiveSheet := m_vDocument.Sheets.Add;
            m_bFirstAddedSheet := false;
          END {ELSE};
        m_vPrograma.Visible := Visible;
        m_vDocument := m_vPrograma.ActiveWorkBook;
        m_vActiveSheet := m_vDocument.ActiveSheet;
      END {IF};
    IF IsOpenOffice THEN
      BEGIN
        m_vDesktop := m_vPrograma.CreateInstance(strOleOoDesktop); //
      //Optional parameters (visible)...
        vOoParams := VarArrayCreate([0, 0], varVariant);
        vOoParams[0] := ooCreateValue('Hidden', NOT Visible); //
      //Create the document...
        m_vDocument := m_vDesktop.LoadComponentFromURL('private:factory/scalc', '_blank', 0, vOoParams);
        ActivateSheetByIndex(1);
        m_bFirstAddedSheet := true;
      END {IF};

  //{ Keep only 1 sheet in the workbook }
  //  WHILE CountSheets > 1 DO
  //    RemoveSheetByIndex(2);
  END {THojaCalc.NewDoc};


PROCEDURE THojaCalc.LoadDoc;

  VAR
    vOoParams: variant;


  BEGIN
    IF m_strFileName = '' THEN
      exit; //
  //Is the program running? (Excel or OpenOffice)
    IF NOT ProgLoaded THEN
      LoadProg; //
  //Is there a doc already loaded?
    CloseDoc;
    m_vDeskTop := Unassigned; //
  //OK, now try to open the doc...
    IF IsExcel THEN
      BEGIN
        m_vPrograma.WorkBooks.Open(m_strFileName, 3);
        m_vPrograma.Visible := Visible;
        m_vDocument := m_vPrograma.ActiveWorkBook;
        m_vActiveSheet := m_vDocument.ActiveSheet;
      END {IF};
    IF IsOpenOffice THEN
      BEGIN
        m_vDesktop := m_vPrograma.CreateInstance(strOleOoDesktop); //
      //Optional parameters (visible)...
        vOoParams := VarArrayCreate([0, 0], varVariant); //
      //Next line stop working OK on OOo V2: Created blind, always blind!
      //so now it is create as visible, then set to non visible if requested
      //vOoParams[0] := ooCreateValue('Hidden', not Visible);
        vOoParams[0] := ooCreateValue('Hidden', false); //Create as visible, then make it not visible if necesary
      //Open the document...
        m_vDocument := m_vDesktop.LoadComponentFromURL(FileName2URL(m_strFileName), '_blank', 0, vOoParams);
        m_vActiveSheet := ActivateSheetByIndex(1); //
      //If has to be non visible, set it now...
        IF NOT visible THEN
          m_vDocument.getCurrentController.getFrame.getContainerWindow.setVisible(false);
      END {IF};
    IF m_eTipo = thcNone THEN
      RAISE Exception.Create('Cannot read file "' + m_strFileName + '" because the needed program is not available.');
  END {THojaCalc.LoadDoc};


FUNCTION THojaCalc.SaveDoc: boolean;


  BEGIN
    result := false;
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          BEGIN
            m_vDocument.Save;
            result := true;
          END {IF};
        IF IsOpenOffice THEN
          BEGIN
            m_vDocument.Store;
            result := true;
          END {IF};
      END {IF};
  END {THojaCalc.SaveDoc};


FUNCTION THojaCalc.SaveDocAs (strName: string; bAsExcel97: boolean = false): boolean;

{ Function added by Massimiliano Gozzi on V0.92 }
{ AsEXcel97 taken form V0.93 by Rômulo Silva Ramos }
{ Saving as .xls on Excel 2000/2003 trick by Malte Tüllmann on V1.01 }

  VAR
    vOoParams: variant;
    exVersion: Extended;
    saveThousandSeparator, saveDecimalSeparator: char;


  BEGIN
    result := false;
    IF DocLoaded
    THEN
      BEGIN
        IF IsExcel THEN
          BEGIN
          {$IFDEF COMPILER_7_UP}
            exVersion := StrToFloat(m_vPrograma.Application.Version, m_AmericanFormat);

          {$ELSE COMPILER_7_UP}
            saveThousandSeparator := SysUtils.ThousandSeparator;
            saveDecimalSeparator := SysUtils.DecimalSeparator;
            SysUtils.ThousandSeparator := ',';
            SysUtils.DecimalSeparator := '.';
            exVersion := StrToFloat(m_vPrograma.Application.Version);
            SysUtils.ThousandSeparator := saveThousandSeparator;
            SysUtils.DecimalSeparator := saveDecimalSeparator;

          {$ENDIF COMPILER_7_UP}
            IF (exVersion < 12)
            THEN //
            //Before Excel 2007 this was the method to force SaveAs Excel97 .xls
            //by Malte Tüllmann on V1.01
              m_vDocument.Saveas(strName, - 4143, EmptyParam, EmptyParam, EmptyParam, EmptyParam)
            ELSE //
            // From Excel 2003 this is the way to force .xls file format (excel8)
            // for back compatibility with older excel version and OO.
            //
            // 51 = xlOpenXMLWorkbook (without macro's in 2007-2010, xlsx)
            // 52 = xlOpenXMLWorkbookMacroEnabled (with or without macro's in 2007-2010, xlsm)
            // 50 = xlExcel12 (Excel Binary Workbook in 2007-2010 with or without macro's, xlsb)
            // 56 = xlExcel8 (97-2003 format in Excel 2007-2010, xls)
            // More on this here: http://www.rondebruin.nl/saveas.htm
              m_vDocument.Saveas(strName, 56);
            m_strFileName := strName;
            result := true;
          END {IF};
        IF IsOpenOffice THEN
          BEGIN //
          //I may need 1 or 2 params...
            IF bAsExcel97
            THEN
              vOoParams := VarArrayCreate([0, 1], varVariant)
            ELSE
              vOoParams := VarArrayCreate([0, 0], varVariant); //
          //First one for prompting on overwrite (good idea!)
            vOoParams[0] := ooCreateValue('Overwrite', false); //
          //Optionally tell OpenOffie to use Excel97 .xls format
            IF bAsExcel97 THEN
              vOoParams[1] := ooCreateValue('FilterName', 'MS Excel 97'); //
          //Do the save!
            m_vDocument.StoreAsUrl(FileName2URL(strName), vOoParams);
            m_strFileName := strName;
            result := true;
          END {IF};
      END {IF};
  END {THojaCalc.SaveDocAs};


FUNCTION THojaCalc.PrintDoc: boolean;

{ Print the Doc... }

  VAR
    vOoParams: variant;


  BEGIN
    result := false;
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          BEGIN
            m_vDocument.PrintOut;
            result := true;
          END {IF};
        IF IsOpenOffice THEN
          BEGIN //
          //NOTE: OpenOffice will print all sheets with Printable areas, but if no
          //printable areas are defined in the doc, it will print all entire sheets.
          //Optional parameters (wait until fully sent to printer)...
            vOoParams := VarArrayCreate([0, 0], varVariant);
            vOoParams[0] := ooCreateValue('Wait', true);
            m_vDocument.Print(vOoParams);
            result := true;
          END {IF};
      END {IF};
  END {THojaCalc.PrintDoc};


PROCEDURE THojaCalc.ShowPrintPreview;


  BEGIN
    IF DocLoaded THEN
      BEGIN //
      //Force visibility of the doc...
        Visible := true;
        IF IsExcel THEN
          m_vDocument.PrintOut(,,, true);
        IF IsOpenOffice THEN
          ooDispatch('.uno:PrintPreview', Unassigned);
      END {IF};
  END {THojaCalc.ShowPrintPreview};


PROCEDURE THojaCalc.SetVisible (v: boolean);


  BEGIN
    IF DocLoaded AND (v <> m_bVisible) THEN
      BEGIN
        IF IsExcel THEN
          m_vPrograma.Visible := v;
        IF IsOpenOffice THEN
          m_vDocument.getCurrentController.getFrame.getContainerWindow.setVisible(v);
        m_bVisible := v;
      END {IF};
  END {THojaCalc.SetVisible};


PROCEDURE THojaCalc.CloseDoc;


  BEGIN
    IF DocLoaded THEN
      BEGIN //
      //Close it...
        TRY
          IF IsOpenOffice THEN
            TRY
              m_vDocument.Dispose;
            EXCEPT
            END;
          IF IsExcel THEN
            TRY
              m_vDocument.close;
            EXCEPT
            END;
        FINALLY
        END {TRY}; //
      //Clean up both "pointer"...
        m_vDocument := Null;
        m_vActiveSheet := Null;
      END {IF};
  END {THojaCalc.CloseDoc};


FUNCTION THojaCalc.GetDocLoaded: boolean;

  VAR
    AttrPtr: Pointer;



  BEGIN
    IF IsExcel AND m_bVisible THEN
      BEGIN
        m_vPrograma := GetActiveOleObject(strOleExcel);
        IF VarIsEmpty(m_vPrograma) OR VarIsNull(m_vPrograma)
        THEN
          m_vDocument := Unassigned
        ELSE
          BEGIN
            m_vDocument := m_vPrograma.ActiveWorkBook;
            AttrPtr := TVarData(m_vDocument).VDispatch;
            IF NOT assigned(AttrPtr) THEN
              m_vDocument := Unassigned
          END {ELSE};
      END {IF};
    IF IsOpenOffice THEN
      BEGIN
      END {IF};
    result := NOT (VarIsEmpty(m_vDocument) OR VarIsNull(m_vDocument));
  END {THojaCalc.GetDocLoaded};

{ ********************* }
{ ** Managing sheets ** }
{ ********************* }


FUNCTION THojaCalc.GetCountSheets: integer;

  VAR
    vActiveSheet, vOoSheets: variant;
    AttrPtr: Pointer;


  BEGIN
    result := 0;
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          TRY
            TRY
              IF VarIsEmpty(m_vDocument) OR VarIsNull(m_vDocument)
              THEN
                vActiveSheet := Unassigned
              ELSE
                vActiveSheet := m_vDocument.ActiveSheet;
              IF VarIsEmpty(vActiveSheet) OR VarIsNull(vActiveSheet)
              THEN
                result := 0
              ELSE
                result := m_vDocument.Sheets.count;
            EXCEPT
              result := 0;
            END {TRY};
          FINALLY
          END {TRY};
        IF IsOpenOffice THEN
          TRY
            TRY
              vOoSheets := m_vDocument.getSheets;
              AttrPtr := TVarData(vOoSheets).VDispatch;
              IF NOT assigned(AttrPtr) THEN
                vOoSheets := Unassigned;
              IF VarIsEmpty(vOoSheets) OR VarIsNull(vOoSheets)
              THEN
                result := 0
              ELSE
                result := vOoSheets.GetCount;
            EXCEPT
              result := 0;
            END {TRY};
          FINALLY
          END {TRY};
      END {IF};
  END {THojaCalc.GetCountSheets};


FUNCTION THojaCalc.ActivateSheetByIndex (nIndex: integer): boolean;

{ Index is 1 based in Excel, but OpenOffice uses it 0-based }
{ Here we asume 1-based so OO needs to activate (nIndex-1) }


  BEGIN
    result := false;
    IF DocLoaded THEN
      BEGIN //
      //Exists this sheet number?
        IF (nIndex < 1) THEN
          RAISE Exception.Create('Can not activate sheet #' + IntToStr(nIndex));
        WHILE (nIndex > CountSheets) DO
          BEGIN
            ActivateSheetByIndex(CountSheets);
            AddNewSheet('New sheet ' + IntToStr(CountSheets + 1));
            sleep(100); //Needs time to do it!
          END {WHILE}; //
      //Activate it now...
        IF IsExcel THEN
          BEGIN
            m_vDocument.Sheets[nIndex].activate;
            m_vActiveSheet := m_vDocument.ActiveSheet;
            result := true;
          END {IF};
        IF IsOpenOffice THEN
          BEGIN
            m_vActiveSheet := m_vDocument.getSheets.getByIndex(nIndex - 1);
            IF m_bVisible THEN
              m_vDocument.getCurrentController.setactivesheet(m_vActiveSheet);
            result := true;
          END {IF};
        sleep(100); //Asyncronus, so better give it time to make the change
      END {IF};
  END {THojaCalc.ActivateSheetByIndex};


FUNCTION THojaCalc.ActivateSheetByName (strSheetName: string; bCaseSensitive: boolean): boolean;

{ Find a sheet by its name... }

  VAR
    vOldActiveSheet: variant;
    i: integer;


  BEGIN
    result := false;
    IF DocLoaded THEN
      BEGIN
        IF bCaseSensitive
        THEN
          BEGIN //
          //Find the EXACT name...
            IF IsExcel THEN
              BEGIN
                m_vDocument.Sheets[strSheetName].Select;
                m_vActiveSheet := m_vDocument.ActiveSheet;
                result := true;
              END {IF};
            IF IsOpenOffice THEN
              BEGIN
                m_vActiveSheet := m_vDocument.getSheets.getByName(strSheetName);
                IF m_bVisible THEN
                  m_vDocument.getCurrentController.setactivesheet(m_vActiveSheet);
                result := true;
              END {IF};
          END {IF}
        ELSE
          BEGIN //
          //Find the Sheet regardless of the case...
            vOldActiveSheet := m_vActiveSheet;
            FOR i := 1 TO GetCountSheets DO
              BEGIN
                ActivateSheetByIndex(i);
                IF UpperCase(ActiveSheetName) = UpperCase(strSheetName) THEN
                  BEGIN
                    result := true;
                    Exit;
                  END {IF};
              END {FOR}; //
          //IF NOT found, let the old active sheet active...
            m_vActiveSheet := vOldActiveSheet;
          END {ELSE};
      END {IF};
  END {THojaCalc.ActivateSheetByName};


FUNCTION THojaCalc.GetActiveSheetName: string;

{ Name of the active sheet? }


  BEGIN
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          result := m_vActiveSheet.Name;
        IF IsOpenOffice THEN
          result := m_vActiveSheet.GetName;
      END {IF};
  END {THojaCalc.GetActiveSheetName};


PROCEDURE THojaCalc.SetActiveSheetName (strNewName: string);


  BEGIN
    IF DocLoaded THEN
      BEGIN //
      //Clean name first...
        strNewName := ValidateSheetName(strNewName);
        IF IsExcel THEN
          m_vPrograma.ActiveSheet.Name := strNewName;
        IF IsOpenOffice THEN
          BEGIN
            m_vActiveSheet.setName(strNewName); //
          //This code always changes the name of "visible" sheet, not active one!
          //ooParams := VarArrayCreate([0, 0], varVariant);
          //ooParams[0] := ooCreateValue('Name', strNewName);
          //ooDispatch('.uno:RenameTable', ooParams);
          END {IF};
      END {IF};
  END {THojaCalc.SetActiveSheetName};


FUNCTION THojaCalc.IsActiveSheetProtected: boolean;

{ Check for sheet protection (password)... }


  BEGIN
    result := false;
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          result := m_vActiveSheet.ProtectContents;
        IF IsOpenOffice THEN
          result := m_vActiveSheet.IsProtected;
      END {IF};
  END {THojaCalc.IsActiveSheetProtected};


FUNCTION THojaCalc.PrintActiveSheet: boolean;

{ WARNING: This function is NOT dual, only works for Excel docs!      }
{ Send active sheet to default printer (as seen in preview window)... }


  BEGIN
    result := false;
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          BEGIN
            m_vActiveSheet.PrintOut;
            result := true;
          END {IF};
        IF IsOpenOffice THEN
          BEGIN
            RAISE Exception.Create('Function "PrintActiveSheet" still not working in OpenOffice!');//
          //ActiveSheet.Print;
            result := false;
          END {IF};
      END {IF};
  END {THojaCalc.PrintActiveSheet};


FUNCTION THojaCalc.PrintSheetsUntil (strLastSheetName: string): boolean;

{ WARNING: This function is NOT dual, only works for Excel docs! }
{ Select and print sheets from 1 upto -excluded- the one with that name. }
{ It is interesting for understanding how to pass an array of objects to excel. }

  VAR
    i, last: integer;
    vHojas: variant;


  BEGIN
    result := false;
    IF DocLoaded
    THEN
      BEGIN
        IF IsExcel THEN
          BEGIN //
          //Macro from Excel:
          //  Sheets(Array("Hoja1", "Hoja2")).Select
          //  ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
          //
          //Which sheet number correspond to the one previous to "LastSheetName"?
            Last := 0;
            FOR i := 2 TO CountSheets DO
              BEGIN
                ActivateSheetByIndex(i);
                IF UpperCase(ActiveSheetName) = UpperCase(strLastSheetName) THEN
                  BEGIN
                    Last := i - 1;
                    break;
                  END {IF};
              END {FOR}; //
          //Not found?
            IF Last = 0 THEN
              exit; //
          //Create an array of variants -windows standard type- this big...
            vHojas := VarArrayCreate([1, Last], varVariant); //
          //Fill it with the Sheet names...
            FOR i := 1 TO Last DO
              BEGIN
                ActivateSheetByIndex(i);
                vHojas[i] := ActiveSheetName;
              END {FOR}; //
          //Print all this array of sheets...
            m_vPrograma.Sheets[vHojas].Select;
            m_vPrograma.ActiveWindow.SelectedSheets.PrintOut; //
          //Done!
            result := true;
          END {IF};
        IF IsOpenOffice THEN
          BEGIN
            RAISE Exception.Create('Function "PrintSheetsUntil" not working in OpenOffice!');
            result := false;
          END {IF};
      END {IF};
  END {THojaCalc.PrintSheetsUntil};


PROCEDURE THojaCalc.AddNewSheet (strNewName: string; bRemoveDummySheets: boolean = false);

{ Add a new sheet, name it, and make it the active sheet... }

  VAR
    vOoSheets: variant;


  BEGIN
    strNewName := ValidateSheetName(strNewName);
    IF NOT DocLoaded THEN
      NewDoc(true);
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          BEGIN
            m_vDocument.WorkSheets.Add(null, m_vDocument.ActiveSheet, 1);
            m_vDocument.ActiveSheet.Name := strNewName;
            IF bRemoveDummySheets AND m_bFirstAddedSheet THEN
              RemoveAllSheetsExcept(strNewName, true); //
          //Active sheet has move to this new one, so I need to update the VAR
            m_vActiveSheet := m_vDocument.ActiveSheet;
            m_bFirstAddedSheet := false;
          END {IF};
        IF IsOpenOffice THEN
          BEGIN
            vOoSheets := m_vDocument.getSheets;
            vOoSheets.insertNewByName(strNewName, 1);
            IF bRemoveDummySheets AND m_bFirstAddedSheet THEN
              RemoveAllSheetsExcept(strNewName, true); //
          //Redefine active sheet to this new one
            m_vActiveSheet := vOoSheets.getByName(strNewName);
            IF m_bVisible THEN
              m_vDocument.getCurrentController.setactivesheet(m_vActiveSheet);
            m_bFirstAddedSheet := false;
          END {IF};
      END {IF};
  END {THojaCalc.AddNewSheet};


PROCEDURE THojaCalc.RemoveSheetByName (strOldName: string);

{ Remove an existing sheet by its name }

  VAR
    vOoSheets: variant;


  BEGIN
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          BEGIN
            m_vDocument.WorkSheets[strOldName].Delete;
          //Active sheet might have moved, so I need to update the VAR
            m_vActiveSheet := m_vDocument.ActiveSheet;
          END {IF};
        IF IsOpenOffice THEN
          BEGIN
            vOoSheets := m_vDocument.getSheets;
            vOoSheets.removeByName(strOldName); //
          //Redefine active sheet to the current one
            m_vActiveSheet := m_vDocument.getCurrentController.getActiveSheet;
          END {IF};
      END {IF};
  END {THojaCalc.RemoveSheetByName};


PROCEDURE THojaCalc.RemoveSheetByIndex (nIndex: integer);

{ Remove an existing sheet by its index }

  VAR
    vOoSheets, vOoSheet: variant;
    strOldName: string;


  BEGIN
    IF DocLoaded THEN
      BEGIN
        IF (nIndex < 1) THEN
          RAISE Exception.Create('Can not remove sheet #' + IntToStr(nIndex));
        IF IsExcel THEN
          BEGIN
            m_vDocument.Sheets[nIndex].Delete;
          //Active sheet might have moved, so I need to update the VAR
            m_vActiveSheet := m_vDocument.ActiveSheet;
          END {IF};
        IF IsOpenOffice THEN
          BEGIN
            vOoSheets := m_vDocument.getSheets;
            vOoSheet := vOoSheets.getByIndex(nIndex - 1);
            strOldName := vOoSheet.GetName;
            vOoSheets.removeByName(strOldName); //
          //Redefine active sheet to the current one
            m_vActiveSheet := m_vDocument.getCurrentController.getActiveSheet;
          END {IF};
      END {IF};
  END {THojaCalc.RemoveSheetByIndex};


PROCEDURE THojaCalc.RemoveAllSheetsExcept (strOldName: string; bCaseSensitive: boolean);

{ Remove all sheets from the workbook, except the specified one }

  VAR
    nIndex: integer;
    vSheet, vOoSheets, vOoSheet: variant;
    strName: string;


  BEGIN
    IF DocLoaded THEN
      BEGIN
        nIndex := 1;
        IF NOT bCaseSensitive THEN
          strOldName := UpperCase(strOldName);
        IF IsExcel THEN
          BEGIN
            WHILE nIndex <= CountSheets DO
              BEGIN
                vSheet := m_vDocument.Sheets[nIndex];
                strName := vSheet.Name;
                IF (bCaseSensitive AND (strName = strOldName)) //
                    OR (NOT bCaseSensitive AND (UpperCase(strName) = strOldName))
                THEN
                  inc(nIndex)
                ELSE
                  m_vDocument.Sheets[nIndex].Delete;
              END {WHILE}; //
          //Activate remaining sheet and update the VAR
            m_vActiveSheet := m_vDocument.ActiveSheet;
          END {IF};
        IF IsOpenOffice THEN
          BEGIN
            WHILE nIndex <= CountSheets DO
              BEGIN
                vOoSheets := m_vDocument.getSheets;
                vOoSheet := vOoSheets.getByIndex(nIndex - 1);
                strName := vOoSheet.GetName;
                IF (bCaseSensitive AND (strName = strOldName)) //
                    OR (NOT bCaseSensitive AND (UpperCase(strName) = strOldName))
                THEN
                  inc(nIndex)
                ELSE
                  BEGIN
                    vOoSheets.removeByName(strName); //
                  END {ELSE};
              END {WHILE}; //
          //Activate remaining sheet
            m_vActiveSheet := m_vDocument.getCurrentController.getActiveSheet;
          END {IF};
      END {IF};
  END {THojaCalc.RemoveAllSheetsExcept};


function  THojaCalc.GetLastRow: integer;
var
  oCursor: Variant;
begin
  result := 0;
  if DocLoaded then begin
    if IsExcel then begin
      result := Programa.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Row;
    end;
    if IsOpenOffice then begin
      oCursor := ActiveSheet.createCursor;
      oCursor.gotoEndOfUsedArea(False);
      result := oCursor.RangeAddress.EndRow;
    end;
  end;
end;

function  THojaCalc.GetLastCol: integer;
var
  oCursor: Variant;
begin
  result := 0;
  if DocLoaded then begin
    if IsExcel then begin
      result := Programa.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Column;
    end;
    if IsOpenOffice then begin
      oCursor := ActiveSheet.createCursor;
      oCursor.gotoEndOfUsedArea(False);
      result := oCursor.RangeAddress.EndColumn;
    end;
  end;
end;


FUNCTION THojaCalc.ValidateSheetName (strName: string): string;

{ Clean a Sheet name so it will not cause problems }


  BEGIN
    result := strName;
    result := StringReplace(result, ':', '_', [rfReplaceAll]);
    result := StringReplace(result, '/', '_', [rfReplaceAll]);
    result := StringReplace(result, '\', '_', [rfReplaceAll]);
    result := StringReplace(result, '?', '_', [rfReplaceAll]);
    result := StringReplace(result, '*', '_', [rfReplaceAll]);
    result := StringReplace(result, '[', '_', [rfReplaceAll]);
    result := StringReplace(result, ']', '_', [rfReplaceAll]);
    result := StringReplace(result, '"', '_', [rfReplaceAll]);
    IF (Trim(result) = '') THEN
      result := 'Plan' + IntToStr(CountSheets);
    result := Copy(result, 1, 31);
  END {THojaCalc.ValidateSheetName};

{ ************************ }
{ ** Manage  the  cells ** }
{ ** in the ActiveSheet ** }
{ ************************ }


FUNCTION THojaCalc.GetCellText (row, col: integer): string;

{ Read/Write cell text (formula en Excel) by index           }
{ OpenOffice start at cell (0,0) while Excel at (1,1)        }
{ Also, Excel uses (row, col) and OpenOffice uses (col, row) }


  BEGIN
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          result := m_vActiveSheet.Cells[row, col].Text;
        IF IsOpenOffice THEN
          result := m_vActiveSheet.getCellByPosition(col - 1, row - 1).getFormula;
      END {IF};
  END {THojaCalc.GetCellText};


PROCEDURE THojaCalc.SetCellText (row, col: integer; strTxt: string);


  BEGIN
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          BEGIN
            m_vActiveSheet.Cells[row, col].Select;
            m_vPrograma.ActiveCell.Value := strTxt;
          END {IF};
        IF IsOpenOffice THEN
          m_vActiveSheet.getCellByPosition(col - 1, row - 1).string := strTxt;
          //m_vActiveSheet.getCellByPosition(col - 1, row - 1).setFormula(strTxt);
      END {IF};
  END {THojaCalc.SetCellText};


FUNCTION THojaCalc.GetCellFormula (row, col: integer): string;

{ Same, but assuming Txt contains a formula, like '=12.00' }


  BEGIN
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          result := m_vActiveSheet.Cells[row, col].Formula;
        IF IsOpenOffice THEN
          result := m_vActiveSheet.getCellByPosition(col - 1, row - 1).getFormula;
      END {IF};
  END {THojaCalc.GetCellFormula};


PROCEDURE THojaCalc.SetCellFormula (row, col: integer; strTxt: string);

  VAR
    strRange: string;


  BEGIN
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          BEGIN
            strRange := m_vPrograma.Range[m_vActiveSheet.Cells[row, col], m_vActiveSheet.Cells[row, col]].Address;
            m_vPrograma.Range[strRange].Formula := strTxt;
          END {IF};
        IF IsOpenOffice THEN
          m_vActiveSheet.getCellByPosition(col - 1, row - 1).setFormula(strTxt);
      END {IF};
  END {THojaCalc.SetCellFormula};


PROCEDURE THojaCalc.SendNumber (row, col: integer; v: double);

{ Sending numbers and date to a cell, the easy way: }
{$IFNDEF COMPILER_7_UP}
  VAR
    saveThousandSeparator, saveDecimalSeparator: char;
{$ENDIF COMPILER_7_UP}


  BEGIN //
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
  {$IFDEF COMPILER_7_UP}
    CellFormula[row, col] := '=' + Format('%f', [v], m_AmericanFormat); 

  {$ELSE COMPILER_7_UP}
    saveThousandSeparator := SysUtils.ThousandSeparator;
    saveDecimalSeparator := SysUtils.DecimalSeparator;
    SysUtils.ThousandSeparator := ',';
    SysUtils.DecimalSeparator := '.';
    CellFormula[row, col] := '=' + Format('%f', [v]);
    SysUtils.ThousandSeparator := saveThousandSeparator;
    SysUtils.DecimalSeparator := saveDecimalSeparator;

  {$ENDIF COMPILER_7_UP} //
  //Note: format string '%f' doesn't show any ThousanSeparator, don't use '%n'!
  END {THojaCalc.SendNumber};


PROCEDURE THojaCalc.SendDate (row, col: integer; v: TDate);

{$IFNDEF COMPILER_7_UP}
  VAR
    saveShortDateFormat: string;
{$ENDIF COMPILER_7_UP}


  BEGIN //
  //Excel and OO need to recieve dates as CellText in American format MM/DD/YYYY
  {$IFDEF COMPILER_7_UP}
    CellText[row, col] := FormatDateTime('ddddd', v, m_AmericanFormat);

  {$ELSE COMPILER_7_UP}
    saveShortDateFormat := SysUtils.ShortDateFormat;
    SysUtils.ShortDateFormat := 'mm/dd/yyyy';
    CellText[row, col] := FormatDateTime('ddddd', v);
    SysUtils.ShortDateFormat := saveShortDateFormat;

  {$ENDIF COMPILER_7_UP} //
  //OpenOffice need to be set to Date Format or it will show an integer
    IF IsOpenOffice THEN
      BEGIN //
      // http://www.oooforum.org/forum/viewtopic.phtml?t=59095
      // 30   10/30/06
      // 31   Mon 30/Oct 06
      // 32   10/06
      // 33   Oct 30
      // 34   October
      // 35   4th quarter 06
      // 36   10/30/2006 
      // 37   10/30/06 
      // 38   Monday, October 30, 2006 
      // 39   Oct 30, 06 
      // 40   05:45 
      // 41   05:45:36 
      // 42   05:45 AM 
      // 43   05:45:36 AM 
      // 44   936485:45:36 
      // 45   45:36.00 
      // 46   936485:45:36.00      //37 = Short date format, usually DD/MM/YY
      //36 = Long date format, usually DD/MM/YYYY
        m_vActiveSheet.getCellByPosition(col - 1, row - 1).NumberFormat := 36; //
      //
      //Note: This code format all selected cells and is NOT tested
      //
      //VAR vOoParams: variant;
      //
      //ooParams := VarArrayCreate([0, 0], varVariant);
      //ooParams[0] := ooCreateValue('NumberFormatValue', 37);
      //ooDispatch('NumberFormatValue', ooParams);
      //
      END {IF};
  END {THojaCalc.SendDate};


FUNCTION THojaCalc.GetCellTextByName (strRange: string): string;

{ Read/Write cell text (formula in OO) by name instead of position }
{ For instance, you can set the value for cell 'NewSheet!A12' or similar }
{ NOTE: If range contains several cells, first one will be used. }

  VAR
    vOldActiveSheet: variant;
    nPosExcl: integer;


  BEGIN
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          BEGIN
            result := m_vPrograma.Range[strRange].Text; //Set 'Formula' but Get 'Text';
          END {IF};
        IF IsOpenOffice THEN
          BEGIN
            vOldActiveSheet := m_vActiveSheet; //
          //If range is in the form 'NewSheet!A1' then first change sheet to 'NewSheet'
            nPosExcl := pos('!', strRange);
            IF nPosExcl > 0 THEN
              BEGIN //
              //Activate the proper sheet...
                IF NOT ActivateSheetByName(Copy(strRange, 1, nPosExcl - 1), false) THEN
                  RAISE Exception.Create('Sheet "' + Copy(strRange, 1, nPosExcl - 1) + '" not present in the document.');
                strRange := Copy(strRange, nPosExcl + 1, 999);
              END {IF};
            result := m_vActiveSheet.getCellRangeByName(strRange).getCellByPosition(0, 0).getFormula;
            m_vActiveSheet := vOldActiveSheet;
          END {IF};
      END {IF};
  END {THojaCalc.GetCellTextByName};


PROCEDURE THojaCalc.SetCellTextByName (strRange: string; strTxt: string);

  VAR
    vOldActiveSheet: variant;
    nPosExcl: integer;


  BEGIN
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          BEGIN
            m_vPrograma.Range[strRange].formula := strTxt;
          END {IF};
        IF IsOpenOffice THEN
          BEGIN
            vOldActiveSheet := m_vActiveSheet; //
          //If range is in the form 'NewSheet!A1' then first change sheet to 'NewSheet'
            nPosExcl := pos('!', strRange);
            IF nPosExcl > 0 THEN
              BEGIN //
              //Activate the proper sheet...
                IF NOT ActivateSheetByName(Copy(strRange, 1, nPosExcl - 1), false) THEN
                  RAISE Exception.Create('Sheet "' + Copy(strRange, 1, nPosExcl - 1) + '" not present in the document.');
                strRange := Copy(strRange, nPosExcl + 1, 999);
              END {IF};
            m_vActiveSheet.getCellRangeByName(strRange).getCellByPosition(0, 0).SetFormula(strTxt);
            m_vActiveSheet := vOldActiveSheet;
          END {IF};
      END {IF};
  END {THojaCalc.SetCellTextByName};


PROCEDURE THojaCalc.FontColor (row, col: integer; color: TColor);


  BEGIN
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          BEGIN
            m_vPrograma.ActiveSheet.Cells[row, col].Font.Color := color;
          END {IF};
        IF IsOpenOffice THEN
          BEGIN {swap bytes of color}
            m_vActiveSheet.getCellByPosition(col - 1, row - 1).CharColor := SwapColor(color);
          END {IF};
      END {IF};
  END {THojaCalc.FontColor};


PROCEDURE THojaCalc.BackgroundColor (row, col: integer; color: TColor);


  BEGIN
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          BEGIN
            m_vPrograma.ActiveSheet.Cells[row, col].Interior.Color := color;
          END {IF};
        IF IsOpenOffice THEN
          BEGIN {swap bytes of color}
            m_vActiveSheet.getCellByPosition(col - 1, row - 1).CellBackColor := SwapColor(color);
          END {IF};
      END {IF};
  END {THojaCalc.BackgroundColor};


PROCEDURE THojaCalc.FontSize (row, col, size: integer);


  BEGIN
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          BEGIN
            m_vPrograma.ActiveSheet.Cells[row, col].Font.Size := size;
          END {IF};
        IF IsOpenOffice THEN
          BEGIN
            m_vActiveSheet.getCellByPosition(col - 1, row - 1).getText.createTextCursor.CharHeight := size;
          END {IF};
      END {IF};
  END {THojaCalc.FontSize};


PROCEDURE THojaCalc.HorizontalAlignment (row, col: integer; ha: TAlignment);

  CONST
    ooHAlignStd = 0; //    com.sun.star.table.CellHoriJustify.STANDARD
    ooHAlignLeft = 1; //   com.sun.star.table.CellHoriJustify.LEFT
    ooHAlignCenter = 2; // com.sun.star.table.CellHoriJustify.CENTER
    ooHAlignRight = 3; //  com.sun.star.table.CellHoriJustify.RIGHT



  BEGIN
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          BEGIN
            CASE ha OF
              taLeftJustify:
                m_vPrograma.ActiveSheet.Cells[row, col].HorizontalAlignment := xlHAlignLeft;
              taRightJustify:
                m_vPrograma.ActiveSheet.Cells[row, col].HorizontalAlignment := xlHAlignRight;
              taCenter:
                m_vPrograma.ActiveSheet.Cells[row, col].HorizontalAlignment := xlHAlignCenter;
              ELSE
            END {CASE};
          END {IF};
        IF IsOpenOffice THEN
          BEGIN
            CASE ha OF
              taLeftJustify:
                m_vActiveSheet.getCellByPosition(col - 1, row - 1).HoriJustify := ooHAlignLeft;
              taRightJustify:
                m_vActiveSheet.getCellByPosition(col - 1, row - 1).HoriJustify := ooHAlignRight;
              taCenter:
                m_vActiveSheet.getCellByPosition(col - 1, row - 1).HoriJustify := ooHAlignCenter;
            END {CASE};
          END {IF};
      END {IF};
  END {THojaCalc.HorizontalAlignment};


PROCEDURE THojaCalc.Bold (row, col: integer);

  CONST
    ooBold: integer
          = 150; //150 = com.sun.star.awt.FontWeight.BOLD


  BEGIN
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          BEGIN
            m_vPrograma.ActiveSheet.Cells[row, col].Font.Bold := true;
          END {IF};
        IF IsOpenOffice THEN
          BEGIN
            m_vActiveSheet.getCellByPosition(col - 1, row - 1).getText.createTextCursor.CharWeight := ooBold;
          END {IF};
      END {IF};
  END {THojaCalc.Bold};


PROCEDURE THojaCalc.Italic (row, col: integer);

  CONST
    ooItalic = 2; //2 = com.sun.star.awt.FontSlant.ITALIC


  BEGIN
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          BEGIN
            m_vPrograma.ActiveSheet.Cells[row, col].Font.Italic := true;
          END {IF};
        IF IsOpenOffice THEN
          BEGIN
            m_vActiveSheet.getCellByPosition(col - 1, row - 1).getText.createTextCursor.CharPosture := ooItalic;
          END {IF};
      END {IF};
  END {THojaCalc.Italic};


PROCEDURE THojaCalc.Underline (row, col: integer; eOoUnderlineStyle: TOoUnderlineStyle);


  BEGIN
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          BEGIN
            CASE eOoUnderlineStyle OF
              ulNone:
                m_vPrograma.ActiveSheet.Cells[row, col].Font.Underline := xlUnderlineStyleNone;
              ulSingle:
                m_vPrograma.ActiveSheet.Cells[row, col].Font.Underline := xlUnderlineStyleSingle;
              ulDouble:
                m_vPrograma.ActiveSheet.Cells[row, col].Font.Underline := xlUnderlineStyleDouble;
              ELSE {map all other values to sinlge}
                m_vPrograma.ActiveSheet.Cells[row, col].Font.Underline := xlUnderlineStyleSingle;
            END {CASE};
          END {IF};
        IF IsOpenOffice THEN
          BEGIN
            m_vActiveSheet.getCellByPosition(col - 1, row - 1).getText.createTextCursor.CharUnderline := ord(eOoUnderlineStyle);
          END {IF};
      END {IF};
  END {THojaCalc.Italic};


PROCEDURE THojaCalc.ColumnWidth (col, width: integer);

{ Width in 1/100 of mm. }

  CONST
    MaxWidthExcel = 255;


  BEGIN
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          BEGIN //
          //Excel uses the width of '0' as the unit, we do an aproximation: Width '0' = 3 mm.
            Width := round(Width / 100 / 3);
            IF Width > MaxWidthExcel THEN
              Width := MaxWidthExcel;
            m_vPrograma.ActiveSheet.Cells[1, col].ColumnWidth := Width;
          END {IF};
        IF IsOpenOffice THEN
          BEGIN
            m_vActiveSheet.getCellByPosition(col - 1, 0).getColumns.getByIndex(0).Width := width;
          END {IF};
      END {IF};
  END {THojaCalc.ColumnWidth};


PROCEDURE THojaCalc.NumberFormat (col, width: integer; strNumberFormat: string);


  BEGIN
    IF DocLoaded THEN
      BEGIN
        IF IsExcel THEN
          BEGIN //
            m_vPrograma.ActiveSheet.Cells[1, col].NumberFormat := strNumberFormat;
          END {IF};
        IF IsOpenOffice THEN
          BEGIN
          END {IF};
      END {IF};
  END {THojaCalc.NumberFormat};

{  *************************** }
{  ** OpenOffice only stuff ** }
{  *************************** }


FUNCTION THojaCalc.FileName2URL (strFileName: string): string;

{ Change 'C:\File.txt' into 'file:///c:/File.txt' (for OpenOffice OpenURL) }


  BEGIN
    result := '';
    IF LowerCase(copy(strFileName, 1, 8)) <> 'file:///' THEN
      result := 'file:///';
    result := result + StringReplace(strFileName, '\', '/', [rfReplaceAll, rfIgnoreCase]);
  END {THojaCalc.FileName2URL};


FUNCTION THojaCalc.ooCreateValue (strOoName: string; vOoData: variant): variant;

  VAR
    vOoReflection: variant;


  BEGIN
    IF IsOpenOffice
    THEN
      BEGIN
        vOoReflection := m_vPrograma.createInstance('com.sun.star.reflection.CoreReflection');
        vOoReflection.forName('com.sun.star.beans.PropertyValue').createObject(result);
        result.Name := strOoName;
        result.Value := vOoData;
      END {IF}
    ELSE
      BEGIN
        RAISE Exception.Create('ooValue imposible to create, load OpenOffice first!');
      END {ELSE};
  END {THojaCalc.ooCreateValue};


PROCEDURE THojaCalc.ooDispatch (strOoCommand: string; vOoParams: variant);

  VAR
    vOoDispatcher, vOoFrame: variant;


  BEGIN
    IF DocLoaded AND IsOpenOffice
    THEN
      BEGIN
        IF (VarIsEmpty(vOoParams) OR VarIsNull(vOoParams)) THEN
          vOoParams := VarArrayCreate([0, - 1], varVariant);
        vOoFrame := m_vDocument.getCurrentController.getFrame;
        vOoDispatcher := m_vPrograma.createInstance('com.sun.star.frame.DispatchHelper');
        vOoDispatcher.executeDispatch(vOoFrame, strOoCommand, '', 0, vOoParams);
      END {IF}
    ELSE
      BEGIN
        RAISE Exception.Create('Dispatch imposible, load a OpenOffice doc first!');
      END {ELSE};
  END {THojaCalc.ooDispatch};


FUNCTION THojaCalc.SwapColor (nColor: TColor): TColor;

  VAR
    c1, c2, c3: byte;


  BEGIN
    c1 := (nColor AND $000000FF); //        $..0000FF
    c2 := (nColor AND $0000FF00) SHR 8; //  $..00FF00
    c3 := (nColor AND $00FF0000) SHR 16; // $..FF0000
    result := (c1 SHL 16) + (c2 SHL 8) + c3;
  END {THojaCalc.SwapColor};


procedure THojaCalc.Orientation(row,Col: integer; Angle: integer);
begin
  if DocLoaded then begin
    if IsExcel then begin
      Programa.ActiveSheet.Cells[row,Col].Orientation:= Angle;
    end;
    if IsOpenOffice then begin
      ActiveSheet.getCellByPosition(col-1, row-1).RotateAngle:= Angle*100;
    end;
  end;
end;


procedure THojaCalc.AutoFit(col: integer); //AutoFix/OptimumWidth
begin
  if DocLoaded then begin
    if IsExcel then begin
      Programa.ActiveSheet.Columns[col].AutoFit;
    end;
    if IsOpenOffice then begin
      ActiveSheet.getColumns.getByIndex(col-1).OptimalWidth:=true;
    end;
  end;
end;

END {UHojaCalc}.
