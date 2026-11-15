#include "hmg.ch"

STATIC s_cLog := ""

FUNCTION Main()
   LOCAL hResult := TallyLoadAll( "localhost", "9000", "Venkateshwara Traders", ;
      "20200401", "20210331", {|cText| TestLog( cText ) }, {|| .F. } )
   LOCAL cOutput := "OK=" + iif( hResult[ "ok" ], "TRUE", "FALSE" ) + Chr( 13 ) + Chr( 10 ) + ;
      "ERROR=" + hResult[ "error" ] + Chr( 13 ) + Chr( 10 ) + s_cLog

   IF hResult[ "ok" ]
      cOutput += "VOUCHERS=" + LTrim( Str( Len( hResult[ "vouchers" ] ) ) ) + Chr( 13 ) + Chr( 10 ) + ;
         "LEDGERS=" + LTrim( Str( Len( hResult[ "ledgers" ] ) ) ) + Chr( 13 ) + Chr( 10 )
   ENDIF
   hb_MemoWrit( hb_DirBase() + "load_test_result.txt", cOutput )
RETURN NIL

STATIC PROCEDURE TestLog( cText )
   s_cLog += cText + Chr( 13 ) + Chr( 10 )
RETURN
