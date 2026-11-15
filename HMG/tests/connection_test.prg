#include "hmg.ch"

FUNCTION Main()
   LOCAL hResult := TallyCompanyInfo( "localhost", "9000" )
   LOCAL lValidName := ! Empty( hResult[ "name" ] ) .AND. Upper( AllTrim( hResult[ "name" ] ) ) != "NIL"
   LOCAL cOutput := "OK=" + iif( hResult[ "ok" ], "TRUE", "FALSE" ) + Chr( 13 ) + Chr( 10 ) + ;
      "NAME_VALID=" + iif( lValidName, "TRUE", "FALSE" ) + Chr( 13 ) + Chr( 10 ) + ;
      "NAME_LENGTH=" + LTrim( Str( Len( hResult[ "name" ] ) ) ) + Chr( 13 ) + Chr( 10 ) + ;
      "FROM_PRESENT=" + iif( ! Empty( hResult[ "from" ] ), "TRUE", "FALSE" ) + Chr( 13 ) + Chr( 10 ) + ;
      "TO_PRESENT=" + iif( ! Empty( hResult[ "to" ] ), "TRUE", "FALSE" ) + Chr( 13 ) + Chr( 10 ) + ;
      "ERROR=" + hResult[ "error" ] + Chr( 13 ) + Chr( 10 )
   hb_MemoWrit( hb_DirBase() + "connection_test_result.txt", cOutput )
RETURN NIL

