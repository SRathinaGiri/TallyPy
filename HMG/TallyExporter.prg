#include "hmg.ch"

#define APP_TITLE "Tally HMG Exporter"
#define PROJECT_URL "https://github.com/SRathinaGiri/TallyPy/tree/main/HMG"
STATIC s_hBundle := {=>}
STATIC s_lBusy := .F.
STATIC s_lCancel := .F.

FUNCTION Main()
   SET MULTIPLE OFF
   SET DATE BRITISH
   SET CENTURY ON

   SET WINDOW MAIN OFF
   ShowStartupSplash()
   SET WINDOW MAIN ON

   DEFINE WINDOW MainForm ;
      AT 20, 20 ;
      WIDTH 1260 ;
      HEIGHT 820 ;
      TITLE APP_TITLE ;
      ICON "MAINICON" ;
      MAIN ;
      ON INIT AppInit()

      @ 12, 12 FRAME FrameConnection CAPTION "Tally connection and period" WIDTH 612 HEIGHT 170
      @ 40, 25 LABEL LblHost VALUE "Host" WIDTH 80 HEIGHT 24
      @ 38, 105 TEXTBOX TxtHost VALUE "localhost" WIDTH 180 HEIGHT 24
      @ 40, 310 LABEL LblPort VALUE "Port" WIDTH 55 HEIGHT 24
      @ 38, 365 TEXTBOX TxtPort VALUE "9000" WIDTH 90 HEIGHT 24 NUMERIC MAXLENGTH 5
      @ 74, 25 LABEL LblCompany VALUE "Company" WIDTH 80 HEIGHT 24
      @ 72, 105 TEXTBOX TxtCompany VALUE "" WIDTH 350 HEIGHT 24
      @ 108, 25 LABEL LblFrom VALUE "From date" WIDTH 80 HEIGHT 24
      @ 106, 105 TEXTBOX TxtFrom VALUE "" WIDTH 140 HEIGHT 24 TOOLTIP "YYYYMMDD, YYYY-MM-DD or DD-MM-YYYY"
      @ 108, 270 LABEL LblTo VALUE "To date" WIDTH 70 HEIGHT 24
      @ 106, 335 TEXTBOX TxtTo VALUE "" WIDTH 120 HEIGHT 24 TOOLTIP "YYYYMMDD, YYYY-MM-DD or DD-MM-YYYY"
      @ 140, 25 BUTTON BtnConnect CAPTION "Connect" WIDTH 100 HEIGHT 28 ACTION ConnectTally()
      @ 140, 135 BUTTON BtnLoad CAPTION "Load tables" WIDTH 110 HEIGHT 28 ACTION LoadTables()
      @ 140, 255 BUTTON BtnCancel CAPTION "Cancel" WIDTH 90 HEIGHT 28 ACTION RequestCancel()
      @ 140, 355 BUTTON BtnAbout CAPTION "About" WIDTH 100 HEIGHT 28 ACTION ShowAbout()

      @ 12, 640 FRAME FrameDetected CAPTION "Detected / export" WIDTH 592 HEIGHT 170
      @ 40, 653 LABEL LblDetected VALUE "Detected" WIDTH 75 HEIGHT 24
      @ 38, 730 TEXTBOX TxtDetected VALUE "" WIDTH 480 HEIGHT 24 READONLY
      @ 74, 653 LABEL LblFolder VALUE "Folder" WIDTH 75 HEIGHT 24
      @ 72, 730 TEXTBOX TxtFolder VALUE "" WIDTH 380 HEIGHT 24
      @ 70, 1118 BUTTON BtnBrowse CAPTION "Browse" WIDTH 92 HEIGHT 28 ACTION ChooseFolder()
      @ 106, 653 CHECKBOX ChkOverwrite CAPTION "Overwrite existing CSV files" VALUE .F. WIDTH 240 HEIGHT 24
      @ 138, 653 BUTTON BtnExport CAPTION "Export all CSVs" WIDTH 130 HEIGHT 28 ACTION ExportAll()
      @ 138, 793 BUTTON BtnClear CAPTION "Clear results" WIDTH 110 HEIGHT 28 ACTION ClearResults()

      DEFINE TAB TabData ;
         AT 194, 12 ;
         WIDTH 1220 ;
         HEIGHT 430 ;
         ON CHANGE RefreshActiveVirtualGrid()

         PAGE "Vouchers"
            @ 35, 10 GRID GridVouchers WIDTH 1190 HEIGHT 340 ;
               HEADERS { "Date", "Type", "Base type", "Number", "Ledger", "MasterID", "Amount", "Dr/Cr", "Debit", "Credit", "Parent", "Primary group", "Nature", "Group nature", "PAN", "Party", "Party GSTIN", "Ledger GSTIN", "Narration", "Optional", "Company", "From", "To" } ;
               WIDTHS { 85, 110, 110, 90, 170, 75, 95, 55, 95, 95, 130, 130, 80, 95, 90, 150, 110, 110, 220, 65, 180, 85, 85 } ;
               VIRTUAL ;
               ITEMCOUNT 0 ;
               ON QUERYDATA VirtualGridData( "vouchers" )
         END PAGE

         PAGE "All vouchers"
            @ 35, 10 GRID GridAllVouchers WIDTH 1190 HEIGHT 340 ;
               HEADERS { "Date", "Type", "Base type", "Number", "Ledger", "MasterID", "Amount", "Dr/Cr", "Debit", "Credit", "Parent", "Primary group", "Nature", "Group nature", "PAN", "Party", "Party GSTIN", "Ledger GSTIN", "Narration", "Optional", "Company", "From", "To", "Category" } ;
               WIDTHS { 85, 110, 110, 90, 170, 75, 95, 55, 95, 95, 130, 130, 80, 95, 90, 150, 110, 110, 220, 65, 180, 85, 85, 90 } ;
               VIRTUAL ;
               ITEMCOUNT 0 ;
               ON QUERYDATA VirtualGridData( "allvouchers" )
         END PAGE

         PAGE "Ledgers"
            @ 35, 10 GRID GridLedgers WIDTH 1190 HEIGHT 340 ;
               HEADERS { "MasterID", "Name", "Primary group", "Nature", "Group nature", "PAN", "Starting", "Currency", "State", "Parent", "GSTIN", "Opening", "Closing", "Company", "From", "To" } ;
               WIDTHS { 75, 190, 140, 85, 100, 90, 85, 90, 100, 150, 120, 100, 100, 180, 85, 85 } ;
               VIRTUAL ;
               ITEMCOUNT 0 ;
               ON QUERYDATA VirtualGridData( "ledgers" )
         END PAGE

         PAGE "Stock items"
            @ 35, 10 GRID GridStock WIDTH 1190 HEIGHT 340 ;
               HEADERS { "Name", "Parent", "Category", "Ledger", "Opening qty", "Opening value", "Basic value", "Basic qty", "Opening rate", "Closing qty", "Closing value", "Closing rate", "Company", "From", "To" } ;
               WIDTHS { 190, 140, 120, 160, 100, 100, 100, 100, 100, 100, 100, 100, 180, 85, 85 } ;
               VIRTUAL ;
               ITEMCOUNT 0 ;
               ON QUERYDATA VirtualGridData( "stock" )
         END PAGE

         PAGE "Stock vouchers"
            @ 35, 10 GRID GridInventory WIDTH 1190 HEIGHT 340 ;
               HEADERS { "Date", "Type", "Number", "Stock item", "Quantity", "Rate", "Amount", "Godown", "Batch", "Narration", "Company", "From", "To" } ;
               WIDTHS { 85, 120, 90, 190, 100, 95, 100, 130, 110, 220, 180, 85, 85 } ;
               VIRTUAL ;
               ITEMCOUNT 0 ;
               ON QUERYDATA VirtualGridData( "inventory" )
         END PAGE
      END TAB

      @ 636, 12 LABEL LblSummary VALUE "No data loaded" WIDTH 1220 HEIGHT 24 BOLD
      @ 664, 12 EDITBOX EdtLog VALUE "" WIDTH 1220 HEIGHT 100 READONLY NOHSCROLL

      DEFINE STATUSBAR
         STATUSITEM "Ready" WIDTH 1040
         STATUSITEM "HMG 64-bit native build" WIDTH 190
      END STATUSBAR
   END WINDOW

   MainForm.BtnCancel.Enabled := .F.
   CENTER WINDOW MainForm
   ACTIVATE WINDOW MainForm
RETURN NIL

/*
 * Startup artwork is compiled into TallyExporter.exe as the SPLASH PNG
 * resource. Clicking the banner, pressing Escape, or waiting three seconds
 * continues to MainForm.
 */
STATIC PROCEDURE ShowStartupSplash()
   DEFINE WINDOW SplashForm ;
      AT 0, 0 ;
      WIDTH 900 ;
      HEIGHT 300 ;
      NOCAPTION ;
      NOSYSMENU ;
      NOSIZE ;
      NOMINIMIZE ;
      NOMAXIMIZE ;
      TOPMOST ;
      CHILD

      @ 0, 0 IMAGE ImgSplash PICTURE "SPLASH" WIDTH 900 HEIGHT 300 STRETCH ;
         ACTION SplashForm.Release

      DEFINE TIMER SplashTimer INTERVAL 3000 ACTION {|| SplashForm.SplashTimer.Release, SplashForm.Release }
      ON KEY ESCAPE ACTION SplashForm.Release
   END WINDOW

   CENTER WINDOW SplashForm
   ACTIVATE WINDOW SplashForm
RETURN

STATIC PROCEDURE AppInit()
   LOCAL cDocs := GetEnv( "USERPROFILE" ) + "\Documents\Tally Exports"
   IF Empty( GetEnv( "USERPROFILE" ) )
      cDocs := GetCurrentFolder()
   ENDIF
   MainForm.TxtFolder.Value := cDocs
   LogMessage( "Ready." )
RETURN

STATIC PROCEDURE SetBusy( lBusy, cStatus )
   s_lBusy := lBusy
   MainForm.TxtHost.Enabled := ! lBusy
   MainForm.TxtPort.Enabled := ! lBusy
   MainForm.TxtCompany.Enabled := ! lBusy
   MainForm.TxtFrom.Enabled := ! lBusy
   MainForm.TxtTo.Enabled := ! lBusy
   MainForm.BtnConnect.Enabled := ! lBusy
   MainForm.BtnLoad.Enabled := ! lBusy
   MainForm.BtnBrowse.Enabled := ! lBusy
   MainForm.BtnExport.Enabled := ! lBusy
   MainForm.BtnClear.Enabled := ! lBusy
   MainForm.BtnCancel.Enabled := lBusy
   MainForm.TabData.Enabled := ! lBusy
   MainForm.StatusBar.Item( 1 ) := cStatus
   DO EVENTS
RETURN

STATIC PROCEDURE ConnectTally()
   LOCAL hResult
   IF s_lBusy
      RETURN
   ENDIF
   s_lCancel := .F.
   SetBusy( .T., "Connecting to Tally..." )
   LogMessage( "Connecting to " + AllTrim( MainForm.TxtHost.Value ) + ":" + LTrim( Str( MainForm.TxtPort.Value ) ) )
   hResult := TallyCompanyInfo( AllTrim( MainForm.TxtHost.Value ), LTrim( Str( MainForm.TxtPort.Value ) ) )
   IF hResult[ "ok" ]
      MainForm.TxtDetected.Value := hResult[ "name" ] + "  (" + FormatTallyDate( hResult[ "from" ] ) + " to " + FormatTallyDate( hResult[ "to" ] ) + ")"
      IF Empty( AllTrim( MainForm.TxtCompany.Value ) )
         MainForm.TxtCompany.Value := hResult[ "name" ]
      ENDIF
      IF Empty( AllTrim( MainForm.TxtFrom.Value ) )
         MainForm.TxtFrom.Value := hResult[ "from" ]
      ENDIF
      IF Empty( AllTrim( MainForm.TxtTo.Value ) )
         MainForm.TxtTo.Value := hResult[ "to" ]
      ENDIF
      LogMessage( "Connected. Active company: " + hResult[ "name" ] )
      SetBusy( .F., "Connected to " + hResult[ "name" ] )
   ELSE
      LogMessage( "Connection failed: " + hResult[ "error" ] )
      SetBusy( .F., "Connection failed" )
      MsgStop( hResult[ "error" ], APP_TITLE )
   ENDIF
RETURN

STATIC PROCEDURE LoadTables()
   LOCAL hResult, cPort
   IF s_lBusy
      RETURN
   ENDIF
   s_lCancel := .F.
   cPort := LTrim( Str( MainForm.TxtPort.Value ) )
   SetBusy( .T., "Loading Tally tables..." )
   LogMessage( "Load started. TDL collections will be injected into Tally." )
   hResult := TallyLoadAll( AllTrim( MainForm.TxtHost.Value ), cPort, ;
      AllTrim( MainForm.TxtCompany.Value ), AllTrim( MainForm.TxtFrom.Value ), ;
      AllTrim( MainForm.TxtTo.Value ), { |c| LogMessage( c ) }, { || s_lCancel } )
   IF hResult[ "ok" ]
      s_hBundle := hResult
      MainForm.TxtDetected.Value := hResult[ "company" ] + "  (" + FormatTallyDate( hResult[ "from" ] ) + " to " + FormatTallyDate( hResult[ "to" ] ) + ")"
      RefreshActiveVirtualGrid()
      MainForm.LblSummary.Value := "Rows: vouchers " + NToC( Len( hResult[ "vouchers" ] ) ) + ;
         " | all vouchers " + NToC( Len( hResult[ "allvouchers" ] ) ) + ;
         " | ledgers " + NToC( Len( hResult[ "ledgers" ] ) ) + ;
         " | stock items " + NToC( Len( hResult[ "stock" ] ) ) + ;
         " | stock vouchers " + NToC( Len( hResult[ "inventory" ] ) )
      SetBusy( .F., "Loaded data for " + hResult[ "company" ] )
      LogMessage( "Load completed successfully." )
   ELSE
      SetBusy( .F., iif( s_lCancel, "Cancelled", "Load failed" ) )
      LogMessage( hResult[ "error" ] )
      IF ! s_lCancel
         MsgStop( hResult[ "error" ], APP_TITLE )
      ENDIF
   ENDIF
RETURN

/*
 * HMG5 owner-data grids inside a TAB must receive ItemCount only while their
 * page is visible. Setting it on a hidden page makes the later ShowWindow()
 * transition request data recursively and can freeze the UI.
 */
STATIC PROCEDURE RefreshActiveVirtualGrid()
   LOCAL nPage, cGrid, cTable, nRows := 0
   LOCAL aGrids := { "GridVouchers", "GridAllVouchers", "GridLedgers", "GridStock", "GridInventory" }
   LOCAL aTables := { "vouchers", "allvouchers", "ledgers", "stock", "inventory" }

   IF ! IsControlDefined( GridInventory, MainForm )
      RETURN
   ENDIF
   nPage := MainForm.TabData.Value
   IF nPage < 1 .OR. nPage > Len( aGrids )
      RETURN
   ENDIF
   cGrid := aGrids[ nPage ]
   cTable := aTables[ nPage ]
   IF Len( s_hBundle ) > 0 .AND. hb_HHasKey( s_hBundle, cTable )
      nRows := Len( s_hBundle[ cTable ] )
   ENDIF
   SetProperty( "MainForm", cGrid, "ItemCount", nRows )
RETURN

STATIC PROCEDURE VirtualGridData( cTable )
   LOCAL nRow := This.QueryRowIndex
   LOCAL nCol := This.QueryColIndex
   LOCAL aRows

   This.QueryData := ""
   IF Len( s_hBundle ) == 0 .OR. ! hb_HHasKey( s_hBundle, cTable )
      RETURN
   ENDIF
   aRows := s_hBundle[ cTable ]
   IF nRow >= 1 .AND. nRow <= Len( aRows ) .AND. ;
      nCol >= 1 .AND. nCol <= Len( aRows[ nRow ] )
      This.QueryData := hb_CStr( aRows[ nRow ][ nCol ] )
   ENDIF
RETURN

STATIC PROCEDURE RequestCancel()
   IF s_lBusy
      s_lCancel := .T.
      MainForm.StatusBar.Item( 1 ) := "Cancellation requested..."
      LogMessage( "Cancellation requested; the current WinHTTP call must finish first." )
   ENDIF
RETURN

STATIC PROCEDURE ChooseFolder()
   LOCAL cFolder := GetFolder( "Select export folder", MainForm.TxtFolder.Value )
   IF ! Empty( cFolder )
      MainForm.TxtFolder.Value := cFolder
   ENDIF
RETURN

STATIC PROCEDURE ExportAll()
   LOCAL hResult
   IF Len( s_hBundle ) == 0 .OR. ! hb_HHasKey( s_hBundle, "vouchers" )
      MsgExclamation( "Load data before exporting.", APP_TITLE )
      RETURN
   ENDIF
   hResult := TallyExportCsv( s_hBundle, AllTrim( MainForm.TxtFolder.Value ), MainForm.ChkOverwrite.Value )
   IF hResult[ "ok" ]
      LogMessage( "Exported five CSV files to " + hResult[ "folder" ] )
      MainForm.StatusBar.Item( 1 ) := "Export completed"
      MsgInfo( "All CSV files were exported to:" + Chr(13) + Chr(10) + hResult[ "folder" ], APP_TITLE )
   ELSE
      LogMessage( "Export failed: " + hResult[ "error" ] )
      MsgStop( hResult[ "error" ], APP_TITLE )
   ENDIF
RETURN

STATIC PROCEDURE ClearResults()
   s_hBundle := {=>}
   RefreshActiveVirtualGrid()
   MainForm.LblSummary.Value := "No data loaded"
   MainForm.TxtDetected.Value := ""
   LogMessage( "Results cleared." )
RETURN

STATIC PROCEDURE ShowAbout()
   DEFINE WINDOW AboutForm ;
      AT 0, 0 ;
      WIDTH 916 ;
      HEIGHT 455 ;
      TITLE "About " + APP_TITLE ;
      MODAL ;
      NOSIZE

      @ 0, 0 IMAGE ImgAboutSplash PICTURE "SPLASH" WIDTH 900 HEIGHT 300 STRETCH
      @ 312, 20 LABEL LblAboutTitle VALUE APP_TITLE + " 1.0.0" WIDTH 860 HEIGHT 24 CENTERALIGN BOLD
      @ 338, 20 LABEL LblAboutDetails ;
         VALUE "Native HMG/Harbour exporter using WinHTTP, MSXML and TDL-injected exports." ;
         WIDTH 860 HEIGHT 24 CENTERALIGN
      @ 364, 20 LABEL LblAboutSource ;
         VALUE PROJECT_URL ;
         WIDTH 860 HEIGHT 24 CENTERALIGN ;
         FONT "Segoe UI" SIZE 9 UNDERLINE ;
         FONTCOLOR { 0, 102, 204 } ;
         TOOLTIP "Open the HMG source folder on GitHub" ;
         ACTION ShellExecute( 0, "open", PROJECT_URL, NIL, NIL, 1 )
      @ 395, 400 BUTTON BtnAboutClose CAPTION "Close" WIDTH 100 HEIGHT 28 ACTION AboutForm.Release
   END WINDOW

   CENTER WINDOW AboutForm
   ACTIVATE WINDOW AboutForm
RETURN

STATIC PROCEDURE LogMessage( cMessage )
   LOCAL cLine := "[" + Time() + "] " + cMessage
   IF IsWindowDefined( MainForm )
      MainForm.EdtLog.Value := MainForm.EdtLog.Value + iif( Empty( MainForm.EdtLog.Value ), "", Chr(13) + Chr(10) ) + cLine
      MainForm.EdtLog.CaretPos := Len( MainForm.EdtLog.Value )
      DO EVENTS
   ENDIF
RETURN

STATIC FUNCTION NToC( nValue )
RETURN LTrim( Str( nValue ) )

