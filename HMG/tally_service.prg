/*
 * Native Tally XML/TDL service for HMG/Harbour.
 * Uses only Windows WinHTTP and MSXML COM components that ship with Windows.
 */

#include "hmg.ch"

#define CRLF ( Chr( 13 ) + Chr( 10 ) )

FUNCTION TallyCompanyInfo( cHost, cPort )
   LOCAL hPost, hResult := {=>}, oDoc, oNode, oNodes, n
   LOCAL cRequest := "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>" + ;
      "<TYPE>COLLECTION</TYPE><ID>MyCompanyInfo</ID></HEADER><BODY><DESC>" + ;
      "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>" + ;
      '<TDL><TDLMESSAGE><COLLECTION NAME="MyCompanyInfo"><TYPE>Company</TYPE>' + ;
      "<FETCH>Name, StartingFrom, EndingAt, Guid</FETCH><FILTER>IsActiveCompany</FILTER></COLLECTION>" + ;
      '<SYSTEM TYPE="Formulae" NAME="IsActiveCompany">$Name = ##SVCURRENTCOMPANY</SYSTEM>' + ;
      "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"

   hResult[ "ok" ] := .F.
   hResult[ "name" ] := ""
   hResult[ "from" ] := ""
   hResult[ "to" ] := ""
   hResult[ "error" ] := ""
   hPost := PostToTally( cHost, cPort, cRequest, 15000 )
   IF ! hPost[ "ok" ]
      hResult[ "error" ] := hPost[ "error" ]
      RETURN hResult
   ENDIF
   oDoc := ParseTallyXml( hPost[ "text" ], @hResult )
   IF oDoc == NIL
      RETURN hResult
   ENDIF
   /*
    * The active-company collection must be posted as actual UTF-8 bytes;
    * sending a COM BSTR while declaring
    * UTF-8 can corrupt this request inside TallyPrime.
    */
   oNodes := oDoc:selectNodes( "//COMPANY" )
   FOR n := 0 TO oNodes:length - 1
      oNode := oNodes:item( n )
      hResult[ "name" ] := XmlAttribute( oNode, "NAME" )
      IF Empty( hResult[ "name" ] )
         hResult[ "name" ] := FirstValue( oNode, { "NAME", "COMPANYNAME" } )
      ENDIF
      IF Empty( hResult[ "name" ] )
         hResult[ "name" ] := XmlDescendantText( oNode, "NAME" )
      ENDIF
      IF ! Empty( hResult[ "name" ] )
         hResult[ "from" ] := FirstValue( oNode, { "STARTINGFROM" } )
         hResult[ "to" ] := FirstValue( oNode, { "ENDINGAT" } )
         EXIT
      ENDIF
   NEXT
   IF Empty( hResult[ "name" ] )
      hResult[ "name" ] := XmlDocText( oDoc, "//SVCURRENTCOMPANY" )
   ENDIF
   IF Empty( hResult[ "name" ] )
      hResult[ "name" ] := XmlDocText( oDoc, "//COMPANYNAME" )
   ENDIF
   IF Empty( hResult[ "name" ] )
      hResult[ "name" ] := XmlDocText( oDoc, "//SERVERCOMPANYNAME" )
   ENDIF
   IF Empty( hResult[ "from" ] )
      hResult[ "from" ] := XmlDocText( oDoc, "//STARTINGFROM" )
   ENDIF
   IF Empty( hResult[ "to" ] )
      hResult[ "to" ] := XmlDocText( oDoc, "//ENDINGAT" )
   ENDIF
   IF Empty( hResult[ "name" ] )
      hResult[ "error" ] := "Tally answered, but no active company was found. Open a company in Tally and try again."
   ELSE
      hResult[ "ok" ] := .T.
   ENDIF
RETURN hResult

FUNCTION TallyLoadAll( cHost, cPort, cCompany, cFrom, cTo, bLog, bCancel )
   LOCAL hResult := {=>}, hInfo, hMeta, hPost, oDoc
   LOCAL aLedgers, aStock, aAll := {}, aVouchers := {}, aInventory := {}
   LOCAL aChunks, aRows, aInvRows, aRow, aChunk, nMerge

   hResult[ "ok" ] := .F.
   hResult[ "error" ] := ""
   EvalLog( bLog, "Target: http://" + cHost + ":" + cPort )

   IF Empty( cCompany ) .OR. Empty( cFrom ) .OR. Empty( cTo )
      hInfo := TallyCompanyInfo( cHost, cPort )
      IF ! hInfo[ "ok" ]
         hResult[ "error" ] := hInfo[ "error" ]
         RETURN hResult
      ENDIF
      IF Empty( cCompany ); cCompany := hInfo[ "name" ]; ENDIF
      IF Empty( cFrom ); cFrom := hInfo[ "from" ]; ENDIF
      IF Empty( cTo ); cTo := hInfo[ "to" ]; ENDIF
   ENDIF
   cFrom := TallyRequestDate( cFrom )
   cTo := TallyRequestDate( cTo )
   IF Empty( cCompany )
      hResult[ "error" ] := "Company is empty. Connect to Tally or enter the exact company name."
      RETURN hResult
   ENDIF
   IF Empty( cFrom ) .OR. Empty( cTo )
      hResult[ "error" ] := "Enter a valid From and To date."
      RETURN hResult
   ENDIF
   IF Cancelled( bCancel ); RETURN CancelResult(); ENDIF

   EvalLog( bLog, "Using company='" + cCompany + "', period=" + FormatTallyDate( cFrom ) + " to " + FormatTallyDate( cTo ) + "." )
   EvalLog( bLog, "Fetching voucher-type and group metadata." )
   hMeta := FetchMetadata( cHost, cPort, cCompany )
   IF ! hMeta[ "ok" ]
      hResult[ "error" ] := hMeta[ "error" ]
      RETURN hResult
   ENDIF
   IF Cancelled( bCancel ); RETURN CancelResult(); ENDIF

   EvalLog( bLog, "Fetching ledgers." )
   hPost := PostToTally( cHost, cPort, BuildLedgerRequest( cCompany ), 120000 )
   IF ! hPost[ "ok" ]; hResult[ "error" ] := hPost[ "error" ]; RETURN hResult; ENDIF
   oDoc := ParseTallyXml( hPost[ "text" ], @hResult )
   IF oDoc == NIL; RETURN hResult; ENDIF
   aLedgers := ParseLedgers( oDoc, hMeta[ "groups" ], cCompany, cFrom, cTo )
   EvalLog( bLog, "Ledgers parsed: " + LTrim( Str( Len( aLedgers ) ) ) + "." )

   aChunks := PeriodChunks( cFrom, cTo )
   EvalLog( bLog, "Fetching voucher rows in " + LTrim( Str( Len( aChunks ) ) ) + " monthly chunk(s)." )
   FOR EACH aChunk IN aChunks
      IF Cancelled( bCancel ); RETURN CancelResult(); ENDIF
      EvalLog( bLog, "Vouchers: " + FormatTallyDate( aChunk[ 1 ] ) + " to " + FormatTallyDate( aChunk[ 2 ] ) + "." )
      hPost := PostToTally( cHost, cPort, BuildVoucherRequest( cCompany, aChunk[ 1 ], aChunk[ 2 ] ), 180000 )
      IF ! hPost[ "ok" ]; hResult[ "error" ] := hPost[ "error" ]; RETURN hResult; ENDIF
      oDoc := ParseTallyXml( hPost[ "text" ], @hResult )
      IF oDoc == NIL; RETURN hResult; ENDIF
      aRows := ParseVoucherRows( oDoc, hMeta[ "vtypes" ], cCompany, cFrom, cTo )
      nMerge := 0
      FOR EACH aRow IN aRows
         nMerge++
         IF nMerge % 200 == 0
            DO EVENTS
            IF Cancelled( bCancel ); RETURN CancelResult(); ENDIF
         ENDIF
         AAdd( aAll, aRow )
         IF aRow[ 24 ] == "Accounting"
            AAdd( aVouchers, ASizeCopy( aRow, 23 ) )
         ENDIF
      NEXT
   NEXT
   EvalLog( bLog, "All voucher rows parsed: " + LTrim( Str( Len( aAll ) ) ) + "." )

   IF Cancelled( bCancel ); RETURN CancelResult(); ENDIF
   EvalLog( bLog, "Fetching stock items." )
   hPost := PostToTally( cHost, cPort, BuildStockRequest( cCompany ), 120000 )
   IF ! hPost[ "ok" ]; hResult[ "error" ] := hPost[ "error" ]; RETURN hResult; ENDIF
   oDoc := ParseTallyXml( hPost[ "text" ], @hResult )
   IF oDoc == NIL; RETURN hResult; ENDIF
   aStock := ParseStockItems( oDoc, cCompany, cFrom, cTo )

   EvalLog( bLog, "Fetching stock-voucher rows." )
   FOR EACH aChunk IN aChunks
      IF Cancelled( bCancel ); RETURN CancelResult(); ENDIF
      EvalLog( bLog, "Stock vouchers: " + FormatTallyDate( aChunk[ 1 ] ) + " to " + FormatTallyDate( aChunk[ 2 ] ) + "." )
      hPost := PostToTally( cHost, cPort, BuildInventoryRequest( cCompany, aChunk[ 1 ], aChunk[ 2 ] ), 180000 )
      IF ! hPost[ "ok" ]; hResult[ "error" ] := hPost[ "error" ]; RETURN hResult; ENDIF
      oDoc := ParseTallyXml( hPost[ "text" ], @hResult )
      IF oDoc == NIL; RETURN hResult; ENDIF
      aInvRows := ParseInventoryRows( oDoc, cCompany, cFrom, cTo )
      nMerge := 0
      FOR EACH aRow IN aInvRows
         nMerge++
         IF nMerge % 200 == 0
            DO EVENTS
            IF Cancelled( bCancel ); RETURN CancelResult(); ENDIF
         ENDIF
         AAdd( aInventory, aRow )
      NEXT
   NEXT

   hResult[ "ok" ] := .T.
   hResult[ "company" ] := cCompany
   hResult[ "from" ] := cFrom
   hResult[ "to" ] := cTo
   hResult[ "vouchers" ] := aVouchers
   hResult[ "allvouchers" ] := aAll
   hResult[ "ledgers" ] := aLedgers
   hResult[ "stock" ] := aStock
   hResult[ "inventory" ] := aInventory
   EvalLog( bLog, "Stock items: " + LTrim( Str( Len( aStock ) ) ) + "; stock vouchers: " + LTrim( Str( Len( aInventory ) ) ) + "." )
RETURN hResult

FUNCTION TallyExportCsv( hBundle, cFolder, lOverwrite )
   LOCAL hResult := {=>}, aFiles, aFile, cPath
   hResult[ "ok" ] := .F.
   hResult[ "error" ] := ""
   hResult[ "folder" ] := cFolder
   IF Empty( cFolder )
      hResult[ "error" ] := "Select an export folder."
      RETURN hResult
   ENDIF
   IF ! hb_DirExists( cFolder )
      IF ! hb_DirBuild( cFolder )
         hResult[ "error" ] := "Unable to create export folder: " + cFolder
         RETURN hResult
      ENDIF
   ENDIF
   aFiles := { ;
      { "vouchers.csv", VoucherColumns(), hBundle[ "vouchers" ] }, ;
      { "allvouchers.csv", AllVoucherColumns(), hBundle[ "allvouchers" ] }, ;
      { "ledgers.csv", LedgerColumns(), hBundle[ "ledgers" ] }, ;
      { "stock_items.csv", StockColumns(), hBundle[ "stock" ] }, ;
      { "stock_vouchers.csv", InventoryColumns(), hBundle[ "inventory" ] } }
   FOR EACH aFile IN aFiles
      cPath := hb_DirSepAdd( cFolder ) + aFile[ 1 ]
      IF File( cPath ) .AND. ! lOverwrite
         hResult[ "error" ] := "File already exists and overwrite is disabled: " + cPath
         RETURN hResult
      ENDIF
      IF ! WriteCsv( cPath, aFile[ 2 ], aFile[ 3 ] )
         hResult[ "error" ] := "Unable to write: " + cPath
         RETURN hResult
      ENDIF
   NEXT
   hResult[ "ok" ] := .T.
RETURN hResult

STATIC FUNCTION PostToTally( cHost, cPort, cXml, nTimeout )
   LOCAL hResult := {=>}, aHttp, oError
   hResult[ "ok" ] := .F.
   hResult[ "text" ] := ""
   hResult[ "error" ] := ""
   DO EVENTS
   BEGIN SEQUENCE WITH {|oErr| Break( oErr ) }
      aHttp := TallyHttpPostUtf8( cHost, Val( cPort ), cXml, nTimeout )
      DO EVENTS
      IF aHttp[ 1 ]
         hResult[ "text" ] := aHttp[ 3 ]
         hResult[ "ok" ] := .T.
      ELSEIF aHttp[ 2 ] > 0
         hResult[ "error" ] := "Tally HTTP error " + LTrim( Str( aHttp[ 2 ] ) )
      ELSE
         hResult[ "error" ] := "Tally request failed. Windows error " + LTrim( Str( aHttp[ 4 ] ) ) + ;
            ". Check that Tally is running and its HTTP server is enabled on port " + cPort + "."
      ENDIF
   RECOVER USING oError
      hResult[ "error" ] := "Tally request failed: " + ErrorText( oError ) + ". Check that Tally is running and its HTTP server is enabled on port " + cPort + "."
   END SEQUENCE
   DO EVENTS
RETURN hResult

STATIC FUNCTION ParseTallyXml( cXml, hResult )
   LOCAL oDoc := NIL, oError, cReason, cStatus
   DO EVENTS
   cXml := SanitizeTallyXml( cXml )
   BEGIN SEQUENCE WITH {|oErr| Break( oErr ) }
      oDoc := CreateObject( "Msxml2.DOMDocument.6.0" )
      oDoc:async := .F.
      oDoc:validateOnParse := .F.
      oDoc:resolveExternals := .F.
      IF ! oDoc:loadXML( cXml )
         cReason := AllTrim( oDoc:parseError:reason )
         hResult[ "error" ] := "Tally returned invalid XML: " + cReason
         oDoc := NIL
      ELSE
         cReason := XmlDocText( oDoc, "//LINEERROR" )
         cStatus := XmlDocText( oDoc, "//STATUS" )
         IF ! Empty( cReason ) .OR. cStatus == "0"
            hResult[ "error" ] := "Tally rejected the request: " + iif( Empty( cReason ), "STATUS=0", cReason )
            oDoc := NIL
         ENDIF
      ENDIF
   RECOVER USING oError
      hResult[ "error" ] := "Unable to parse Tally XML: " + ErrorText( oError )
      oDoc := NIL
   END SEQUENCE
   DO EVENTS
RETURN oDoc

/*
 * Tally uses XML numeric references for internal display markers such as
 * "&#4; Primary".  XML 1.0 forbids those control characters, and MSXML 6
 * rejects the complete response.  They carry no exported business data, so
 * remove only the XML 1.0-forbidden C0 controls before parsing.
 */
STATIC FUNCTION SanitizeTallyXml( cXml )
   LOCAL nCode, cDecimal, cHex
   FOR nCode := 0 TO 31
      IF nCode == 9 .OR. nCode == 10 .OR. nCode == 13
         LOOP
      ENDIF
      cDecimal := LTrim( Str( nCode ) )
      cHex := Lower( hb_NumToHex( nCode ) )
      cXml := StrTran( cXml, "&#" + cDecimal + ";", "" )
      cXml := StrTran( cXml, "&#x" + cHex + ";", "" )
      cXml := StrTran( cXml, "&#x" + Upper( cHex ) + ";", "" )
      cXml := StrTran( cXml, Chr( nCode ), "" )
   NEXT
RETURN cXml

STATIC FUNCTION FetchMetadata( cHost, cPort, cCompany )
   LOCAL hResult := {=>}, hPost, oDoc, oNodes, oNode, n
   LOCAL hTypes := {=>}, hGroups := {=>}, hGroup, cName, cStatic := StaticVars( cCompany, "", "" )
   LOCAL cVType := Envelope( "AllVoucherTypes", cStatic, ;
      "<COLLECTION NAME='AllVoucherTypes'><TYPE>VoucherType</TYPE><FETCH>Name, Parent</FETCH></COLLECTION>" )
   LOCAL cGroups := Envelope( "AllGroups", cStatic, ;
      "<COLLECTION NAME='AllGroups'><TYPE>Group</TYPE><FETCH>Name, Parent, Nature, _PrimaryGroup</FETCH></COLLECTION>" )

   hResult[ "ok" ] := .F.
   hResult[ "error" ] := ""
   hPost := PostToTally( cHost, cPort, cVType, 120000 )
   IF ! hPost[ "ok" ]; hResult[ "error" ] := hPost[ "error" ]; RETURN hResult; ENDIF
   oDoc := ParseTallyXml( hPost[ "text" ], @hResult )
   IF oDoc == NIL; RETURN hResult; ENDIF
   oNodes := oDoc:selectNodes( "//VOUCHERTYPE" )
   FOR n := 0 TO oNodes:length - 1
      oNode := oNodes:item( n )
      cName := FirstValue( oNode, { "NAME" } )
      IF Empty( cName ); cName := XmlAttribute( oNode, "NAME" ); ENDIF
      IF ! Empty( cName ); hTypes[ Upper( cName ) ] := FirstValue( oNode, { "PARENT" } ); ENDIF
   NEXT

   hPost := PostToTally( cHost, cPort, cGroups, 120000 )
   IF ! hPost[ "ok" ]; hResult[ "error" ] := hPost[ "error" ]; RETURN hResult; ENDIF
   oDoc := ParseTallyXml( hPost[ "text" ], @hResult )
   IF oDoc == NIL; RETURN hResult; ENDIF
   oNodes := oDoc:selectNodes( "//GROUP" )
   FOR n := 0 TO oNodes:length - 1
      oNode := oNodes:item( n )
      cName := FirstValue( oNode, { "NAME" } )
      IF Empty( cName ); cName := XmlAttribute( oNode, "NAME" ); ENDIF
      IF ! Empty( cName )
         hGroup := {=>}
         hGroup[ "parent" ] := FirstValue( oNode, { "PARENT" } )
         hGroup[ "nature" ] := FirstValue( oNode, { "NATURE" } )
         hGroup[ "primary" ] := FirstValue( oNode, { "_PRIMARYGROUP", "PRIMARYGROUP" } )
         hGroups[ Upper( cName ) ] := hGroup
      ENDIF
   NEXT
   hResult[ "ok" ] := .T.
   hResult[ "vtypes" ] := hTypes
   hResult[ "groups" ] := hGroups
RETURN hResult

STATIC FUNCTION ParseLedgers( oDoc, hGroups, cCompany, cFrom, cTo )
   LOCAL aRows := {}, oNodes := oDoc:selectNodes( "//LEDGER" ), oNode, n
   LOCAL cName, cParent, cPrimary, cNature, aNature
   FOR n := 0 TO oNodes:length - 1
      IF n % 200 == 0; DO EVENTS; ENDIF
      oNode := oNodes:item( n )
      cName := FirstValue( oNode, { "NAME" } )
      IF Empty( cName ); cName := XmlAttribute( oNode, "NAME" ); ENDIF
      IF Empty( cName ); LOOP; ENDIF
      cParent := FirstValue( oNode, { "PARENT" } )
      cPrimary := FirstValue( oNode, { "PRIMARYGROUP", "_PRIMARYGROUP" } )
      IF Empty( cPrimary ); cPrimary := ResolvePrimaryGroup( cParent, hGroups ); ENDIF
      cNature := GroupNature( cParent, hGroups )
      aNature := NatureFromPrimary( cPrimary, cNature )
      AAdd( aRows, { ;
         FirstValue( oNode, { "MASTERID" } ), cName, cPrimary, aNature[ 1 ], aNature[ 2 ], ;
         FirstValue( oNode, { "INCOMETAXNUMBER", "PAN" } ), FormatTallyDate( FirstValue( oNode, { "STARTINGFROM" } ) ), ;
         FirstValue( oNode, { "CURRENCYSYMBOL", "CURRENCYORIGINALSYMBOL", "CURRENCYNAME" } ), ;
         FirstValue( oNode, { "STATENAME" } ), cParent, FirstValue( oNode, { "PARTYGSTIN" } ), ;
         NumberText( FirstValue( oNode, { "OPENINGBALANCE" } ) ), NumberText( FirstValue( oNode, { "CLOSINGBALANCE" } ) ), ;
         cCompany, FormatTallyDate( cFrom ), FormatTallyDate( cTo ) } )
   NEXT
RETURN aRows

STATIC FUNCTION ParseVoucherRows( oDoc, hVTypes, cCompany, cFrom, cTo )
   LOCAL aRows := {}, oNodes := oDoc:selectNodes( "//*[LEDGERNAME and AMOUNT]" ), oNode, n
   LOCAL cLedger, cType, cBase, cCategory, cRaw, nRaw, nAmount, lDebit
   LOCAL cOptional, cParty, cPrimary, aNature
   FOR n := 0 TO oNodes:length - 1
      IF n % 200 == 0; DO EVENTS; ENDIF
      oNode := oNodes:item( n )
      cLedger := FirstValue( oNode, { "LEDGERNAME" } )
      cRaw := FirstValue( oNode, { "TXMLSIGNEDAMOUNT", "SIGNEDAMOUNT", "AMOUNT" } )
      nRaw := AmountValue( FirstValue( oNode, { "AMOUNT" } ) )
      nAmount := AmountValue( cRaw )
      IF Empty( cLedger ) .OR. ( Abs( nRaw ) < 0.000001 .AND. Abs( nAmount ) < 0.000001 ); LOOP; ENDIF
      lDebit := Upper( FirstValue( oNode, { "ISDEEMEDPOSITIVE" } ) ) == "YES" .OR. nRaw < 0
      nAmount := Abs( iif( Abs( nRaw ) > 0.000001, nRaw, nAmount ) ) * iif( lDebit, -1, 1 )
      cType := CanonicalVoucherType( FirstValue( oNode, { "TXMLVOUCHERTYPENAME", "VOUCHERTYPENAME" } ) )
      cBase := CanonicalVoucherType( hb_HGetDef( hVTypes, Upper( cType ), cType ) )
      IF Empty( cBase ); cBase := cType; ENDIF
      cCategory := VoucherCategory( cBase )
      cOptional := FirstValue( oNode, { "TXMLSTATUSOPTIONAL", "STATUSOPTIONAL", "ISOPTIONAL" } )
      cOptional := iif( Upper( cOptional ) == "YES", "Yes", "No" )
      cParty := FirstValue( oNode, { "TXMLPARTYLEDGERNAME", "PARTYLEDGERNAME" } )
      IF Empty( cParty ); cParty := "N/A"; ENDIF
      cPrimary := FirstValue( oNode, { "TXMLENTRYPRIMARYGROUP", "ENTRYPRIMARYGROUP", "PRIMARYGROUP" } )
      aNature := NatureFromPrimary( cPrimary, "" )
      AAdd( aRows, { ;
         FormatTallyDate( FirstValue( oNode, { "TXMLDATE", "DATE" } ) ), cType, cBase, ;
         FirstValue( oNode, { "TXMLVOUCHERNUMBER", "VOUCHERNUMBER" } ), cLedger, ;
         FirstValue( oNode, { "TXMLENTRYLEDGERMASTERID", "ENTRYLEDGERMASTERID", "LEDMASTERID" } ), ;
         NumToText( nAmount ), iif( lDebit, "Dr", "Cr" ), ;
         NumToText( iif( lDebit, Abs( nAmount ), 0 ) ), NumToText( iif( lDebit, 0, Abs( nAmount ) ) ), ;
         FirstValue( oNode, { "TXMLENTRYPARENTLEDGER", "ENTRYPARENTLEDGER", "PARENTLEDGER" } ), ;
         cPrimary, aNature[ 1 ], aNature[ 2 ], "", cParty, ;
         FirstValue( oNode, { "TXMLPARTYGSTIN", "PARTYGSTIN" } ), ;
         FirstValue( oNode, { "TXMLENTRYLEDGERGSTIN", "ENTRYLEDGERGSTIN", "LEDGERGSTIN" } ), ;
         FirstValue( oNode, { "TXMLVOUCHERNARRATION", "NARRATION", "VOUCHERNARRATION" } ), cOptional, ;
         cCompany, FormatTallyDate( cFrom ), FormatTallyDate( cTo ), cCategory } )
   NEXT
RETURN aRows

STATIC FUNCTION ParseStockItems( oDoc, cCompany, cFrom, cTo )
   LOCAL aRows := {}, oNodes := oDoc:selectNodes( "//STOCKITEM" ), oNode, n, cName
   FOR n := 0 TO oNodes:length - 1
      IF n % 200 == 0; DO EVENTS; ENDIF
      oNode := oNodes:item( n )
      cName := FirstValue( oNode, { "NAME" } )
      IF Empty( cName ); cName := XmlAttribute( oNode, "NAME" ); ENDIF
      IF Empty( cName ); LOOP; ENDIF
      AAdd( aRows, { cName, FirstValue( oNode, { "PARENT" } ), FirstValue( oNode, { "CATEGORY" } ), ;
         FirstValue( oNode, { "LEDGERNAME" } ), NumberText( FirstValue( oNode, { "OPENINGBALANCE" } ) ), ;
         NumberText( FirstValue( oNode, { "OPENINGVALUE" } ) ), NumberText( FirstValue( oNode, { "BASICVALUE" } ) ), ;
         NumberText( FirstValue( oNode, { "BASICQTY" } ) ), NumberText( FirstValue( oNode, { "OPENINGRATE" } ) ), ;
         NumberText( FirstValue( oNode, { "CLOSINGBALANCE" } ) ), NumberText( FirstValue( oNode, { "CLOSINGVALUE" } ) ), ;
         NumberText( FirstValue( oNode, { "CLOSINGRATE" } ) ), cCompany, FormatTallyDate( cFrom ), FormatTallyDate( cTo ) } )
   NEXT
RETURN aRows

STATIC FUNCTION ParseInventoryRows( oDoc, cCompany, cFrom, cTo )
   LOCAL aRows := {}, oNodes := oDoc:selectNodes( "//*[STOCKITEMNAME]" ), oNode, n, cItem, cType
   FOR n := 0 TO oNodes:length - 1
      IF n % 200 == 0; DO EVENTS; ENDIF
      oNode := oNodes:item( n )
      cItem := FirstValue( oNode, { "STOCKITEMNAME" } )
      cType := FirstValue( oNode, { "TXMLVOUCHERTYPENAME", "VOUCHERTYPENAME" } )
      IF Empty( cItem ) .OR. "ORDER" $ Upper( cType ); LOOP; ENDIF
      AAdd( aRows, { FormatTallyDate( FirstValue( oNode, { "TXMLDATE", "DATE" } ) ), cType, ;
         FirstValue( oNode, { "TXMLVOUCHERNUMBER", "VOUCHERNUMBER" } ), cItem, ;
         NumberText( FirstValue( oNode, { "TXMLSIGNEDQTY", "BILLEDQTY" } ) ), ;
         NumberText( FirstValue( oNode, { "RATE" } ) ), NumberText( FirstValue( oNode, { "TXMLSIGNEDAMOUNT", "AMOUNT" } ) ), ;
         FirstValue( oNode, { "TXMLGODOWNNAME", "GODOWNNAME" } ), FirstValue( oNode, { "TXMLBATCHNAME", "BATCHNAME" } ), ;
         FirstValue( oNode, { "TXMLVOUCHERNARRATION", "NARRATION", "VOUCHERNARRATION" } ), ;
         cCompany, FormatTallyDate( cFrom ), FormatTallyDate( cTo ) } )
   NEXT
RETURN aRows

STATIC FUNCTION BuildLedgerRequest( cCompany )
RETURN Envelope( "MyLedgers", StaticVars( cCompany, "", "" ), ;
   "<COLLECTION NAME='MyLedgers'><TYPE>Ledger</TYPE>" + ;
   "<FETCH>Name, Parent, PartyGSTIN, MasterID, StartingFrom, CurrencyName, StateName, OpeningBalance, ClosingBalance, IncomeTaxNumber</FETCH>" + ;
   "<COMPUTE>PrimaryGroup:$_PrimaryGroup</COMPUTE><COMPUTE>CurrencySymbol:$UnicodeSymbol:Currency:$CurrencyName</COMPUTE>" + ;
   "<COMPUTE>CurrencyOriginalSymbol:$OriginalSymbol:Currency:$CurrencyName</COMPUTE></COLLECTION>" )

STATIC FUNCTION BuildVoucherRequest( cCompany, cFrom, cTo )
RETURN Envelope( "TXMLFlatVoucherRows", StaticVars( cCompany, cFrom, cTo ), ;
   "<COLLECTION NAME='TXMLBaseVouchers'><TYPE>Voucher</TYPE>" + ;
   "<FETCH>Date, VoucherTypeName, VoucherNumber, Narration, PartyLedgerName, PartyGSTIN, IsOptional</FETCH></COLLECTION>" + ;
   "<COLLECTION NAME='TXMLFlatVoucherRows'><SOURCECOLLECTION>TXMLBaseVouchers</SOURCECOLLECTION><WALK>All Ledger Entries</WALK>" + ;
   "<FETCH>LedgerName, Amount, IsDeemedPositive, TXMLDate, TXMLVoucherTypeName, TXMLVoucherNumber, TXMLPartyLedgerName, TXMLPartyGSTIN, TXMLVoucherNarration, TXMLSignedAmount, TXMLDebitAmount, TXMLCreditAmount, TXMLDrCr, TXMLEntryLedgerMasterID, TXMLEntryParentLedger, TXMLEntryPrimaryGroup, TXMLEntryLedgerGSTIN, TXMLStatusOptional, TXMLCompanyName</FETCH>" + ;
   "<COMPUTE>TXMLDate:$..Date</COMPUTE><COMPUTE>TXMLVoucherTypeName:$..VoucherTypeName</COMPUTE>" + ;
   "<COMPUTE>TXMLVoucherNumber:$..VoucherNumber</COMPUTE><COMPUTE>TXMLPartyLedgerName:If NOT $$IsEmpty:$..PartyLedgerName Then $..PartyLedgerName Else &quot;N/A&quot;</COMPUTE>" + ;
   "<COMPUTE>TXMLPartyGSTIN:$..PartyGSTIN</COMPUTE><COMPUTE>TXMLVoucherNarration:$..Narration</COMPUTE>" + ;
   "<COMPUTE>TXMLSignedAmount:If ($IsDeemedPositive OR $Amount &lt; 0) Then $$Abs:$Amount * -1 Else $$Abs:$Amount</COMPUTE>" + ;
   "<COMPUTE>TXMLDebitAmount:If ($IsDeemedPositive OR $Amount &lt; 0) Then $$Abs:$Amount Else 0</COMPUTE>" + ;
   "<COMPUTE>TXMLCreditAmount:If ($IsDeemedPositive OR $Amount &lt; 0) Then 0 Else $$Abs:$Amount</COMPUTE>" + ;
   "<COMPUTE>TXMLDrCr:If ($IsDeemedPositive OR $Amount &lt; 0) Then &quot;Dr&quot; Else &quot;Cr&quot;</COMPUTE>" + ;
   "<COMPUTE>TXMLEntryLedgerMasterID:$MasterID:Ledger:$LedgerName</COMPUTE><COMPUTE>TXMLEntryParentLedger:$Parent:Ledger:$LedgerName</COMPUTE>" + ;
   "<COMPUTE>TXMLEntryPrimaryGroup:$_PrimaryGroup:Ledger:$LedgerName</COMPUTE><COMPUTE>TXMLEntryLedgerGSTIN:$PartyGSTIN:Ledger:$LedgerName</COMPUTE>" + ;
   "<COMPUTE>TXMLStatusOptional:If $..IsOptional Then &quot;Yes&quot; Else &quot;No&quot;</COMPUTE><COMPUTE>TXMLCompanyName:##SVCurrentCompany</COMPUTE>" + ;
   "</COLLECTION>" )

STATIC FUNCTION BuildStockRequest( cCompany )
RETURN Envelope( "MyStockItems", StaticVars( cCompany, "", "" ), ;
   "<COLLECTION NAME='MyStockItems'><TYPE>StockItem</TYPE>" + ;
   "<FETCH>Name, Parent, Category, LedgerName, OpeningBalance, OpeningValue, BasicValue, BasicQty, OpeningRate</FETCH>" + ;
   "<COMPUTE>ClosingBalance:$_ClosingBalance</COMPUTE><COMPUTE>ClosingValue:$_ClosingValue</COMPUTE>" + ;
   "<COMPUTE>ClosingRate:$_ClosingRate</COMPUTE></COLLECTION>" )

STATIC FUNCTION BuildInventoryRequest( cCompany, cFrom, cTo )
RETURN Envelope( "TXMLFlatInventoryRows", StaticVars( cCompany, cFrom, cTo ), ;
   "<COLLECTION NAME='TXMLBaseInventoryVouchers'><TYPE>Voucher</TYPE><FETCH>Date, VoucherTypeName, VoucherNumber, Narration</FETCH></COLLECTION>" + ;
   "<COLLECTION NAME='TXMLFlatInventoryRows'><SOURCECOLLECTION>TXMLBaseInventoryVouchers</SOURCECOLLECTION><WALK>All Inventory Entries</WALK>" + ;
   "<FETCH>StockItemName, BilledQty, Rate, Amount, TXMLDate, TXMLVoucherTypeName, TXMLVoucherNumber, TXMLVoucherNarration, TXMLCompanyName, TXMLSignedQty, TXMLSignedAmount, TXMLGodownName, TXMLBatchName</FETCH>" + ;
   "<COMPUTE>TXMLDate:$..Date</COMPUTE><COMPUTE>TXMLVoucherTypeName:$..VoucherTypeName</COMPUTE><COMPUTE>TXMLVoucherNumber:$..VoucherNumber</COMPUTE>" + ;
   "<COMPUTE>TXMLVoucherNarration:$..Narration</COMPUTE><COMPUTE>TXMLCompanyName:##SVCurrentCompany</COMPUTE>" + ;
   "<COMPUTE>TXMLSignedQty:If $IsDeemedPositive Then $$Abs:$BilledQty Else $$Abs:$BilledQty * -1</COMPUTE>" + ;
   "<COMPUTE>TXMLSignedAmount:If $IsDeemedPositive Then $$Abs:$Amount Else $$Abs:$Amount * -1</COMPUTE>" + ;
   "<COMPUTE>TXMLGodownName:$GodownName</COMPUTE><COMPUTE>TXMLBatchName:$BatchName</COMPUTE></COLLECTION>" )

STATIC FUNCTION Envelope( cId, cStatic, cCollections )
RETURN "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>" + cId + ;
   "</ID></HEADER><BODY><DESC><STATICVARIABLES>" + cStatic + "</STATICVARIABLES><TDL><TDLMESSAGE>" + ;
   cCollections + "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"

STATIC FUNCTION StaticVars( cCompany, cFrom, cTo )
   LOCAL cXml := "<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
   IF ! Empty( cCompany ); cXml += "<SVCURRENTCOMPANY>" + XmlEscape( cCompany ) + "</SVCURRENTCOMPANY>"; ENDIF
   IF ! Empty( cFrom ); cXml += "<SVFROMDATE TYPE='Date'>" + TallyRequestDate( cFrom ) + "</SVFROMDATE>"; ENDIF
   IF ! Empty( cTo ); cXml += "<SVTODATE TYPE='Date'>" + TallyRequestDate( cTo ) + "</SVTODATE>"; ENDIF
RETURN cXml

FUNCTION FormatTallyDate( cValue )
   cValue := AllTrim( hb_CStr( cValue ) )
   IF Len( cValue ) == 8 .AND. Val( cValue ) > 0
      RETURN Left( cValue, 4 ) + "-" + SubStr( cValue, 5, 2 ) + "-" + Right( cValue, 2 )
   ENDIF
RETURN cValue

STATIC FUNCTION TallyRequestDate( cValue )
   LOCAL c := AllTrim( hb_CStr( cValue ) ), aParts
   c := StrTran( c, "/", "-" )
   IF Len( c ) == 8 .AND. Val( c ) > 0; RETURN c; ENDIF
   aParts := hb_ATokens( c, "-" )
   IF Len( aParts ) == 3
      IF Len( aParts[ 1 ] ) == 4
         RETURN StrZero( Val( aParts[ 1 ] ), 4 ) + StrZero( Val( aParts[ 2 ] ), 2 ) + StrZero( Val( aParts[ 3 ] ), 2 )
      ENDIF
      RETURN StrZero( Val( aParts[ 3 ] ), 4 ) + StrZero( Val( aParts[ 2 ] ), 2 ) + StrZero( Val( aParts[ 1 ] ), 2 )
   ENDIF
RETURN c

STATIC FUNCTION PeriodChunks( cFrom, cTo )
   LOCAL aChunks := {}, dStart := hb_SToD( TallyRequestDate( cFrom ) ), dEnd := hb_SToD( TallyRequestDate( cTo ) )
   LOCAL dCurrent, dNext, dChunkEnd, nYear, nMonth
   IF Empty( dStart ) .OR. Empty( dEnd ) .OR. dStart > dEnd
      RETURN { { cFrom, cTo } }
   ENDIF
   dCurrent := dStart
   DO WHILE dCurrent <= dEnd
      nYear := Year( dCurrent ); nMonth := Month( dCurrent ) + 1
      IF nMonth == 13; nMonth := 1; nYear++; ENDIF
      dNext := hb_SToD( StrZero( nYear, 4 ) + StrZero( nMonth, 2 ) + "01" )
      dChunkEnd := Min( dEnd, dNext - 1 )
      AAdd( aChunks, { DToS( dCurrent ), DToS( dChunkEnd ) } )
      dCurrent := dNext
   ENDDO
RETURN aChunks

STATIC FUNCTION XmlFirst( oNode, cXPath )
   LOCAL oResult := NIL
   BEGIN SEQUENCE WITH {|oErr| Break( oErr ) }
      oResult := oNode:selectSingleNode( cXPath )
   RECOVER
      oResult := NIL
   END SEQUENCE
RETURN oResult

STATIC FUNCTION XmlDocText( oDoc, cXPath )
   LOCAL oNode := XmlFirst( oDoc, cXPath )
RETURN iif( oNode == NIL, "", CleanText( oNode:text ) )

STATIC FUNCTION XmlNodeText( oNode, cName )
   LOCAL oChild := XmlFirst( oNode, "./" + cName )
RETURN iif( oChild == NIL, "", CleanText( oChild:text ) )

STATIC FUNCTION XmlDescendantText( oNode, cName )
   LOCAL oChild := XmlFirst( oNode, ".//" + cName )
RETURN iif( oChild == NIL, "", CleanText( oChild:text ) )

STATIC FUNCTION FirstValue( oNode, aNames )
   LOCAL cName, cValue
   FOR EACH cName IN aNames
      cValue := XmlNodeText( oNode, cName )
      IF ! Empty( cValue ); RETURN cValue; ENDIF
   NEXT
RETURN ""

STATIC FUNCTION XmlAttribute( oNode, cName )
   LOCAL cValue := ""
   BEGIN SEQUENCE WITH {|oErr| Break( oErr ) }
      cValue := CleanText( oNode:getAttribute( cName ) )
   RECOVER
      cValue := ""
   END SEQUENCE
RETURN cValue

STATIC FUNCTION CleanText( xValue )
   IF HB_ISNIL( xValue )
      RETURN ""
   ENDIF
RETURN AllTrim( hb_CStr( xValue ) )

STATIC FUNCTION XmlEscape( cValue )
   cValue := StrTran( hb_CStr( cValue ), "&", "&amp;" )
   cValue := StrTran( cValue, "<", "&lt;" )
   cValue := StrTran( cValue, ">", "&gt;" )
   cValue := StrTran( cValue, '"', "&quot;" )
   cValue := StrTran( cValue, "'", "&apos;" )
RETURN cValue

STATIC FUNCTION ErrorText( oError )
   IF HB_ISOBJECT( oError )
      RETURN AllTrim( oError:Description )
   ENDIF
RETURN hb_CStr( oError )

STATIC FUNCTION ResolvePrimaryGroup( cGroup, hGroups )
   LOCAL cCurrent := Upper( AllTrim( cGroup ) ), hGroup, nGuard := 0, cPrimary
   DO WHILE ! Empty( cCurrent ) .AND. nGuard < 100
      nGuard++
      IF ! hb_HHasKey( hGroups, cCurrent ); EXIT; ENDIF
      hGroup := hGroups[ cCurrent ]
      cPrimary := hb_HGetDef( hGroup, "primary", "" )
      IF ! Empty( cPrimary ); RETURN cPrimary; ENDIF
      IF IsPrimaryGroup( cCurrent ); RETURN ProperGroupName( cCurrent ); ENDIF
      cCurrent := Upper( hb_HGetDef( hGroup, "parent", "" ) )
   ENDDO
RETURN ProperGroupName( Upper( cGroup ) )

STATIC FUNCTION GroupNature( cGroup, hGroups )
   LOCAL hGroup
   IF hb_HHasKey( hGroups, Upper( cGroup ) )
      hGroup := hGroups[ Upper( cGroup ) ]
      RETURN hb_HGetDef( hGroup, "nature", "" )
   ENDIF
RETURN ""

STATIC FUNCTION NatureFromPrimary( cPrimary, cNature )
   LOCAL c := Upper( AllTrim( cPrimary ) )
   IF c $ "CAPITAL ACCOUNT|RESERVES & SURPLUS|LOANS (LIABILITY)|BANK OD A/C|SECURED LOANS|UNSECURED LOANS|CURRENT LIABILITIES|DUTIES & TAXES|PROVISIONS|SUNDRY CREDITORS"
      RETURN { "BS", "Liabilities" }
   ENDIF
   IF c $ "FIXED ASSETS|INVESTMENTS|CURRENT ASSETS|STOCK-IN-HAND|DEPOSITS (ASSET)|LOANS & ADVANCES (ASSET)|BANK ACCOUNTS|CASH-IN-HAND|SUNDRY DEBTORS|MISC. EXPENSES (ASSET)"
      RETURN { "BS", "Assets" }
   ENDIF
   IF c $ "SALES ACCOUNTS|DIRECT INCOMES|INDIRECT INCOMES"
      RETURN { "PL", "Income" }
   ENDIF
   IF c $ "PURCHASE ACCOUNTS|DIRECT EXPENSES|INDIRECT EXPENSES"
      RETURN { "PL", "Expenses" }
   ENDIF
   IF ! Empty( cNature ); RETURN { cNature, cNature }; ENDIF
RETURN { "Unknown", "Unknown" }

STATIC FUNCTION IsPrimaryGroup( cUpper )
RETURN cUpper $ "CAPITAL ACCOUNT|RESERVES & SURPLUS|LOANS (LIABILITY)|BANK OD A/C|SECURED LOANS|UNSECURED LOANS|CURRENT LIABILITIES|DUTIES & TAXES|PROVISIONS|SUNDRY CREDITORS|FIXED ASSETS|INVESTMENTS|CURRENT ASSETS|STOCK-IN-HAND|DEPOSITS (ASSET)|LOANS & ADVANCES (ASSET)|BANK ACCOUNTS|CASH-IN-HAND|SUNDRY DEBTORS|MISC. EXPENSES (ASSET)|SUSPENSE ACCOUNT|BRANCH / DIVISIONS|SALES ACCOUNTS|PURCHASE ACCOUNTS|DIRECT INCOMES|INDIRECT INCOMES|DIRECT EXPENSES|INDIRECT EXPENSES"

STATIC FUNCTION ProperGroupName( cUpper )
   LOCAL aGroups := { "Capital Account", "Reserves & Surplus", "Loans (Liability)", "Bank OD A/c", "Secured Loans", "Unsecured Loans", "Current Liabilities", "Duties & Taxes", "Provisions", "Sundry Creditors", "Fixed Assets", "Investments", "Current Assets", "Stock-in-hand", "Deposits (Asset)", "Loans & Advances (Asset)", "Bank Accounts", "Cash-in-hand", "Sundry Debtors", "Misc. Expenses (ASSET)", "Suspense Account", "Branch / Divisions", "Sales Accounts", "Purchase Accounts", "Direct Incomes", "Indirect Incomes", "Direct Expenses", "Indirect Expenses" }, c
   FOR EACH c IN aGroups
      IF Upper( c ) == cUpper; RETURN c; ENDIF
   NEXT
RETURN ""

STATIC FUNCTION CanonicalVoucherType( cType )
   IF Upper( AllTrim( cType ) ) $ "REJECTION IN|REJECTIONS IN"; RETURN "Rejections In"; ENDIF
   IF Upper( AllTrim( cType ) ) $ "REJECTION OUT|REJECTIONS OUT"; RETURN "Rejections Out"; ENDIF
RETURN AllTrim( cType )

STATIC FUNCTION VoucherCategory( cType )
   LOCAL c := Upper( AllTrim( cType ) )
   IF c $ "SALES|PURCHASE|JOURNAL|RECEIPT|PAYMENT|DEBIT NOTE|CREDIT NOTE|CONTRA|MEMORANDUM|REVERSING JOURNAL"; RETURN "Accounting"; ENDIF
   IF c $ "DELIVERY NOTE|RECEIPT NOTE|REJECTIONS IN|REJECTIONS OUT|STOCK JOURNAL|PHYSICAL STOCK|MATERIAL IN|MATERIAL OUT"; RETURN "Inventory"; ENDIF
   IF c $ "SALES ORDER|PURCHASE ORDER|JOB WORK IN ORDER|JOB WORK OUT ORDER"; RETURN "Orders"; ENDIF
   IF c $ "PAYROLL|ATTENDANCE"; RETURN "Payroll"; ENDIF
RETURN "Unknown"

STATIC FUNCTION AmountValue( cValue )
   LOCAL c := StrTran( AllTrim( hb_CStr( cValue ) ), ",", "" ), aTokens, cToken
   c := StrTran( c, "Dr", "" ); c := StrTran( c, "Cr", "" )
   aTokens := hb_ATokens( c, " " )
   IF Len( aTokens ) > 0
      cToken := aTokens[ Len( aTokens ) ]
   ELSE
      cToken := c
   ENDIF
RETURN Val( cToken )

STATIC FUNCTION NumToText( nValue )
   IF Abs( nValue ) < 0.0000001; nValue := 0; ENDIF
RETURN LTrim( Str( nValue, 20, 2 ) )

STATIC FUNCTION NumberText( cValue )
RETURN NumToText( AmountValue( cValue ) )

STATIC FUNCTION ASizeCopy( aSource, nSize )
   LOCAL aResult := {}, n
   FOR n := 1 TO Min( Len( aSource ), nSize ); AAdd( aResult, aSource[ n ] ); NEXT
RETURN aResult

STATIC FUNCTION Cancelled( bCancel )
RETURN HB_ISBLOCK( bCancel ) .AND. Eval( bCancel )

STATIC FUNCTION CancelResult()
   LOCAL h := {=>}
   h[ "ok" ] := .F.; h[ "error" ] := "Operation cancelled by user."
RETURN h

STATIC PROCEDURE EvalLog( bLog, cText )
   IF HB_ISBLOCK( bLog ); Eval( bLog, cText ); ENDIF
RETURN

STATIC FUNCTION WriteCsv( cPath, aHeaders, aRows )
   LOCAL nHandle, aRow, lOk := .T., nRow := 0
   nHandle := FCreate( cPath )
   IF nHandle < 0
      RETURN .F.
   ENDIF
   IF FWrite( nHandle, Chr( 239 ) + Chr( 187 ) + Chr( 191 ) + CsvLine( aHeaders ) + CRLF ) <= 0
      lOk := .F.
   ENDIF
   FOR EACH aRow IN aRows
      nRow++
      IF nRow % 200 == 0; DO EVENTS; ENDIF
      IF FWrite( nHandle, CsvLine( aRow ) + CRLF ) <= 0
         lOk := .F.
         EXIT
      ENDIF
   NEXT
   FClose( nHandle )
RETURN lOk

STATIC FUNCTION CsvLine( aValues )
   LOCAL cLine := "", xValue
   FOR EACH xValue IN aValues
      IF ! Empty( cLine ); cLine += ","; ENDIF
      cLine += CsvEscape( hb_CStr( xValue ) )
   NEXT
RETURN cLine

STATIC FUNCTION CsvEscape( cValue )
   cValue := StrTran( cValue, '"', '""' )
   IF "," $ cValue .OR. '"' $ cValue .OR. Chr( 10 ) $ cValue .OR. Chr( 13 ) $ cValue
      RETURN '"' + cValue + '"'
   ENDIF
RETURN cValue

STATIC FUNCTION VoucherColumns()
RETURN { "Date", "VoucherTypeName", "BaseVoucherType", "VoucherNumber", "LedgerName", "MasterID", "Amount", "DrCr", "DebitAmount", "CreditAmount", "ParentLedger", "PrimaryGroup", "Nature", "NatureOfGroup", "PAN", "PartyLedgerName", "PartyGSTIN", "LedgerGSTIN", "VoucherNarration", "IsOptional", "CompanyName", "FromDate", "ToDate" }

STATIC FUNCTION AllVoucherColumns()
   LOCAL a := VoucherColumns(); AAdd( a, "VoucherCategory" )
RETURN a

STATIC FUNCTION LedgerColumns()
RETURN { "MasterID", "Name", "PrimaryGroup", "Nature", "NatureOfGroup", "PAN", "StartingFrom", "CurrencyName", "StateName", "Parent", "PartyGSTIN", "OpeningBalance", "ClosingBalance", "CompanyName", "FromDate", "ToDate" }

STATIC FUNCTION StockColumns()
RETURN { "Name", "Parent", "Category", "LedgerName", "OpeningBalance", "OpeningValue", "BasicValue", "BasicQty", "OpeningRate", "ClosingBalance", "ClosingValue", "ClosingRate", "CompanyName", "FromDate", "ToDate" }

STATIC FUNCTION InventoryColumns()
RETURN { "Date", "VoucherTypeName", "VoucherNumber", "StockItemName", "BilledQty", "Rate", "Amount", "GodownName", "BatchName", "VoucherNarration", "CompanyName", "FromDate", "ToDate" }

