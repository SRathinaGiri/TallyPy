/* Byte-exact UTF-8 HTTP transport for Tally.
 *
 * The OLE WinHttpRequest Send() method accepts a VARIANT.  Passing a Harbour
 * string makes it a COM BSTR, leaving the transmitted encoding ambiguous even
 * when the Content-Type claims UTF-8.  TallyHttpPostUtf8 sends the original
 * Harbour UTF-8 bytes directly through the native WinHTTP API, matching the
 * QByteArray/requests behavior of the working Qt and Python exporters.
 */

#include "hbapi.h"
#include "hbapiitm.h"
#include <windows.h>
#include <winhttp.h>
#include <string.h>

static WCHAR * utf8_to_wide( const char * text )
{
   int count;
   WCHAR * wide;

   if( text == NULL )
      return NULL;

   count = MultiByteToWideChar( CP_UTF8, 0, text, -1, NULL, 0 );
   if( count <= 0 )
      return NULL;

   wide = ( WCHAR * ) hb_xgrab( ( HB_SIZE ) count * sizeof( WCHAR ) );
   if( MultiByteToWideChar( CP_UTF8, 0, text, -1, wide, count ) <= 0 )
   {
      hb_xfree( wide );
      return NULL;
   }
   return wide;
}

static char * response_as_utf8( const char * data, HB_SIZE size, HB_SIZE * outSize )
{
   HB_SIZE offset = 0;
   int wideCount;
   int utf8Count;
   char * output;

   *outSize = size;
   if( size >= 2 && ( unsigned char ) data[ 0 ] == 0xFF &&
                    ( unsigned char ) data[ 1 ] == 0xFE )
      offset = 2;
   else if( size >= 4 && data[ 1 ] == 0 && data[ 3 ] == 0 )
      offset = 0;
   else
   {
      output = ( char * ) hb_xgrab( size + 1 );
      if( size > 0 )
         memcpy( output, data, size );
      output[ size ] = '\0';
      return output;
   }

   wideCount = ( int ) ( ( size - offset ) / sizeof( WCHAR ) );
   utf8Count = WideCharToMultiByte( CP_UTF8, 0,
      ( const WCHAR * ) ( data + offset ), wideCount,
      NULL, 0, NULL, NULL );
   if( utf8Count <= 0 )
   {
      output = ( char * ) hb_xgrab( size + 1 );
      memcpy( output, data, size );
      output[ size ] = '\0';
      return output;
   }

   output = ( char * ) hb_xgrab( ( HB_SIZE ) utf8Count + 1 );
   WideCharToMultiByte( CP_UTF8, 0,
      ( const WCHAR * ) ( data + offset ), wideCount,
      output, utf8Count, NULL, NULL );
   output[ utf8Count ] = '\0';
   *outSize = ( HB_SIZE ) utf8Count;
   return output;
}

HB_FUNC( TALLYHTTPPOSTUTF8 )
{
   const char * host = hb_parc( 1 );
   int port = hb_parni( 2 );
   const char * body = hb_parc( 3 );
   HB_SIZE bodySize = hb_parclen( 3 );
   int timeout = hb_parni( 4 );
   WCHAR * wideHost = NULL;
   HINTERNET session = NULL;
   HINTERNET connection = NULL;
   HINTERNET request = NULL;
   LPCWSTR acceptTypes[] = { L"text/xml", NULL };
   const WCHAR * headers = L"Content-Type: text/xml; charset=utf-8\r\nAccept: text/xml\r\n";
   DWORD status = 0;
   DWORD statusSize = sizeof( status );
   DWORD available = 0;
   DWORD received = 0;
   DWORD errorCode = 0;
   HB_SIZE dataSize = 0;
   HB_SIZE utf8Size = 0;
   char * data = NULL;
   char * utf8 = NULL;
   BOOL sent = FALSE;
   BOOL ok = FALSE;
   PHB_ITEM result = hb_itemArrayNew( 4 );

   if( timeout <= 0 )
      timeout = 120000;
   if( port <= 0 || port > 65535 )
      port = 9000;

   wideHost = utf8_to_wide( host );
   if( wideHost == NULL || body == NULL )
   {
      errorCode = ERROR_INVALID_PARAMETER;
      goto cleanup;
   }

   session = WinHttpOpen( L"Tally HMG Exporter/1.0",
      WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
      WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0 );
   if( session == NULL )
      goto winerror;

   WinHttpSetTimeouts( session, 10000, 10000, timeout, timeout );
   connection = WinHttpConnect( session, wideHost, ( INTERNET_PORT ) port, 0 );
   if( connection == NULL )
      goto winerror;

   request = WinHttpOpenRequest( connection, L"POST", L"/", NULL,
      WINHTTP_NO_REFERER, acceptTypes, 0 );
   if( request == NULL )
      goto winerror;

   sent = WinHttpSendRequest( request, headers, ( DWORD ) -1L,
      ( LPVOID ) body, ( DWORD ) bodySize, ( DWORD ) bodySize, 0 );
   if( ! sent )
      goto winerror;
   if( ! WinHttpReceiveResponse( request, NULL ) )
      goto winerror;

   WinHttpQueryHeaders( request,
      WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
      WINHTTP_HEADER_NAME_BY_INDEX, &status, &statusSize,
      WINHTTP_NO_HEADER_INDEX );

   data = ( char * ) hb_xgrab( 1 );
   data[ 0 ] = '\0';
   do
   {
      available = 0;
      if( ! WinHttpQueryDataAvailable( request, &available ) )
         goto winerror;
      if( available > 0 )
      {
         data = ( char * ) hb_xrealloc( data, dataSize + available + 1 );
         received = 0;
         if( ! WinHttpReadData( request, data + dataSize, available, &received ) )
            goto winerror;
         dataSize += received;
         data[ dataSize ] = '\0';
      }
   }
   while( available > 0 );

   utf8 = response_as_utf8( data, dataSize, &utf8Size );
   ok = status >= 200 && status < 300;
   goto cleanup;

winerror:
   errorCode = GetLastError();

cleanup:
   hb_arraySetL( result, 1, ok ? HB_TRUE : HB_FALSE );
   hb_arraySetNInt( result, 2, status );
   if( utf8 != NULL )
      hb_arraySetCL( result, 3, utf8, utf8Size );
   else if( data != NULL )
      hb_arraySetCL( result, 3, data, dataSize );
   else
      hb_arraySetC( result, 3, "" );
   hb_arraySetNInt( result, 4, errorCode );

   if( request != NULL )
      WinHttpCloseHandle( request );
   if( connection != NULL )
      WinHttpCloseHandle( connection );
   if( session != NULL )
      WinHttpCloseHandle( session );
   if( wideHost != NULL )
      hb_xfree( wideHost );
   if( data != NULL )
      hb_xfree( data );
   if( utf8 != NULL )
      hb_xfree( utf8 );

   hb_itemReturnRelease( result );
}

