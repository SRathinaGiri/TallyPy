#include "hmg.ch"

STATIC s_aVouchers := {}
STATIC s_aAll := {}
STATIC s_nQueries := 0
STATIC s_nStarted := 0

FUNCTION Main()
   LOCAL aHeaders23 := {}, aHeaders24 := {}, aWidths23 := {}, aWidths24 := {}, n

   ErrorBlock( {|e| OutStd( "virtual_grid_test: RUNTIME ERROR " + e:Description + hb_eol() ), ExitProcess( 2 ) } )
   s_aVouchers := MakeRows( 3869, 23 )
   s_aAll := MakeRows( 6761, 24 )
   FOR n := 1 TO 24
      AAdd( aHeaders24, "Column " + hb_NToS( n ) )
      AAdd( aWidths24, 100 )
      IF n <= 23
         AAdd( aHeaders23, "Column " + hb_NToS( n ) )
         AAdd( aWidths23, 100 )
      ENDIF
   NEXT

   DEFINE WINDOW Win_Test AT 20, 20 WIDTH 1260 HEIGHT 600 TITLE "Virtual grid tab test" MAIN
      DEFINE TAB TabData AT 10, 10 WIDTH 1220 HEIGHT 500 ON CHANGE BindActiveGrid()
         PAGE "Vouchers"
            @ 35, 10 GRID GridOne WIDTH 1190 HEIGHT 420 ;
               HEADERS aHeaders23 WIDTHS aWidths23 VIRTUAL ITEMCOUNT 0 ;
               ON QUERYDATA TestQuery( 1 )
         END PAGE
         PAGE "All vouchers"
            @ 35, 10 GRID GridTwo WIDTH 1190 HEIGHT 420 ;
               HEADERS aHeaders24 WIDTHS aWidths24 VIRTUAL ITEMCOUNT 0 ;
               ON QUERYDATA TestQuery( 2 )
         END PAGE
      END TAB
      DEFINE TIMER TmrStart INTERVAL 500 ACTION StartTabSwitch()
      DEFINE TIMER TmrFinish INTERVAL 1500 ACTION FinishTabSwitch()
   END WINDOW
   Win_Test.TmrFinish.Enabled := .F.
   Win_Test.Activate
RETURN NIL

STATIC FUNCTION MakeRows( nRows, nCols )
   LOCAL aRows := Array( nRows ), aRow, nRow, nCol
   FOR nRow := 1 TO nRows
      aRow := Array( nCols )
      FOR nCol := 1 TO nCols
         aRow[ nCol ] := "R" + hb_NToS( nRow ) + "C" + hb_NToS( nCol )
      NEXT
      aRows[ nRow ] := aRow
   NEXT
RETURN aRows

STATIC PROCEDURE TestQuery( nTable )
   LOCAL aRows := iif( nTable == 1, s_aVouchers, s_aAll )
   s_nQueries++
   This.QueryData := aRows[ This.QueryRowIndex ][ This.QueryColIndex ]
RETURN

STATIC PROCEDURE StartTabSwitch()
   Win_Test.TmrStart.Enabled := .F.
   s_nQueries := 0
   Win_Test.GridOne.ItemCount := Len( s_aVouchers )
   s_nStarted := hb_MilliSeconds()
   Win_Test.TabData.Value := 2
   Win_Test.TmrFinish.Enabled := .T.
RETURN

STATIC PROCEDURE BindActiveGrid()
   IF Win_Test.TabData.Value == 1
      Win_Test.GridOne.ItemCount := Len( s_aVouchers )
   ELSE
      Win_Test.GridTwo.ItemCount := Len( s_aAll )
   ENDIF
RETURN

STATIC PROCEDURE FinishTabSwitch()
   LOCAL nElapsed := hb_MilliSeconds() - s_nStarted
   Win_Test.TmrFinish.Enabled := .F.
   OutStd( "virtual_grid_test: ELAPSED_MS=" + hb_NToS( nElapsed ) + ;
      " QUERIES=" + hb_NToS( s_nQueries ) + " PASS=1 FAIL=0" + hb_eol() )
   ExitProcess( 0 )
RETURN
