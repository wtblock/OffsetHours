/////////////////////////////////////////////////////////////////////////////
// Copyright © by W. T. Block, all rights reserved
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "OffsetHours.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

/////////////////////////////////////////////////////////////////////////////
// The one and only application object
CWinApp theApp;

/////////////////////////////////////////////////////////////////////////////
// given an image pointer and an ASCII property ID, return the propery value
CString GetStringProperty( Gdiplus::Image* pImage, PROPID id )
{
	CString value;

	// get the size of the date property
	const UINT uiSize = pImage->GetPropertyItemSize( id );

	// if the property exists, it will have a non-zero size 
	if ( uiSize > 0 )
	{
		// using a smart pointer which will release itself
		// when it goes out of context
		unique_ptr<Gdiplus::PropertyItem> pItem =
			unique_ptr<Gdiplus::PropertyItem>
			(
			(PropertyItem*)malloc( uiSize )
				);

		// Get the property item.
		pImage->GetPropertyItem( id, uiSize, pItem.get() );

		// the property should be ASCII
		if ( pItem->type == PropertyTagTypeASCII )
		{
			value = (LPCSTR)pItem->value;
		}
	}

	return value;
} // GetStringProperty

/////////////////////////////////////////////////////////////////////////////
// given a vector of tokens from a string date, look for a possible string
// month token and return its index in the vector or -1 if not found
int FindTextMonthIndex( vector<CString>& tokens )
{
	int value = -1;
	int nIndex = 0;

	// loop through the tokens for a text month (non-numeric)
	for ( CString token : tokens )
	{
		// a non-numeric value will return zero
		const int nMonth = _tstol( token );
		if ( nMonth == 0 )
		{
			value = nIndex;
			break;
		}

		nIndex++;
	}

	return value;
} // FindTextMonthIndex

/////////////////////////////////////////////////////////////////////////////
// given a date string that does not parse into a valid date, try to create
// a valid date from its tokens
bool CreateValidDate( CString csDate )
{
	bool value = false;

	// current date and time
	COleDateTime oNow = COleDateTime::GetCurrentTime();
	const int nCurrentYear = oNow.GetYear();
	
	// a "pm" string indicates the hour needs have 11 added to it 
	bool bPM = false;

	// an "am" string indicates the hour needs have 1 subtracted from it 
	bool bAM = false;

	// original index of AM or PM token
	int nAmOrPm = -1;

	// the month index
	int nMonthIndex = -1;

	// the value of the month
	int nMonth = 0;

	// the year index
	int nYearIndex = -1;

	// the value of the year
	int nYear = -1;

	// if the date did not parse, break the tokens up into a vector
	// using colons, dashs, spaces, commas, periods, and underscores.
	const CString csDelim( _T( ":- ,._" ) );
	int nStart = 0;
	int nToken = 0;
	vector<CString> tokens;
	do
	{
		CString csToken = 
				csDate.Tokenize( csDelim, nStart ).MakeLower();
		if ( csToken.IsEmpty() )
		{
			break;
		}

		// check for "am" or "pm"
		if ( csToken == _T( "pm" ))
		{
			bPM = true;
			continue;
		}

		// check for "am"
		if ( csToken == _T( "pm" ))
		{
			bAM = true;
			continue;
		}

		const bool bNumeric = GetNumeric( csToken );
		if ( bNumeric )
		{
			const int nLen = csToken.GetLength();

			// a four digit number has to be the year
			if ( nLen == 4 )
			{
				nYearIndex = nToken;
				nYear = _tstol( csToken );
			}
		} else // non-numeric
		{
			// non-numeric has to be the month
			nMonth = m_Date.GetMonthOfTheYear( csToken );
			if ( nMonth != 0 )
			{
				nMonthIndex = nToken;
				csToken.Format( _T( "%02s" ), nMonth );
			}
		}

		tokens.push_back( csToken );

		nToken++;

	} while ( true );

	// there should be six tokens
	const size_t tTokens = tokens.size();
	if ( tTokens != 6 )
	{
		return value;
	}

	// these formats should parse okay
	// case A: January 25, 1996 8:30:00
	// case B: 25 January 1996 8:30:00
	// case C: 8:30:00 Jan. 25, 1996
	// case D: 8:30:00 25 Jan. 1996
	// case E: 8:30:00 AM Jan. 25, 1996
	// case F: 8:30:00 PM 25 Jan. 1996
	// case G: 1/25/1996 23:30:00
	// case H: 1/25/1996 11:30:00 PM
	// case I: 1996/1/25 23:30:00
	// case L: 1/25/96 23:30:00
	// case k: 1/25/96 11:30:00 PM

	// Years can be two digit numbers but four digit
	// numbers can only be the year

	// I have observed this not parsing
	// "1:25:1996 23:30:00"

	// reassemble the tokens in a format that we know
	// will work
	int nHour = _tstol( tokens[ 3 ] );
	int nMinute = _tstol( tokens[ 4 ] );
	int nSecond = _tstol( tokens[ 5 ] );
	int nDay = 0;

	// The only non-numberic tokens are months which
	// (AM and PM have been removed) are in possible 
	// month indices of 0, 1, 3 or 4
	switch ( nMonthIndex )
	{
		case 0: // January 25, 1996 8:30:00
		{
			nDay = _tstol( tokens[ 1 ] );
			nYear = _tstol( tokens[ 2 ] );
			break;
		}
		case 1: // 25 January 1996 8:30:00
		{
			nDay = _tstol( tokens[ 0 ] );
			nYear = _tstol( tokens[ 2 ] );
			break;
		}
		case 3: // 8:30:00 Jan. 25, 1996
		{
			nHour = _tstol( tokens[ 0 ] );
			nMinute = _tstol( tokens[ 1 ] );
			nSecond = _tstol( tokens[ 2 ] );
			nDay = _tstol( tokens[ 4 ] );
			nYear = _tstol( tokens[ 5 ] );

			break;
		}
		case 4: // 8:30:00 25 Jan. 1996
		{
			nHour = _tstol( tokens[ 0 ] );
			nMinute = _tstol( tokens[ 1 ] );
			nSecond = _tstol( tokens[ 2 ] );
			nDay = _tstol( tokens[ 3 ] );
			nYear = _tstol( tokens[ 5 ] );
			break;
		}
		default :
		{
			// 
			
			switch ( nYearIndex )
			{
				case 0: // 1996/1/25 23:30:00
				{
					nMonth = _tstol( tokens[ 1 ] );
					nDay = _tstol( tokens[ 2 ] );
					break;
				}
				case 2: // 11/25/1996 23:30:00
				{
					nMonth = _tstol( tokens[ 0 ] );
					nDay = _tstol( tokens[ 1 ] );
					break;
				}
				case 3: // 23:30:00 1996/1/25 
				{
					nHour = _tstol( tokens[ 0 ] );
					nMinute = _tstol( tokens[ 1 ] );
					nSecond = _tstol( tokens[ 2 ] );
					nMonth = _tstol( tokens[ 1 ] );
					nDay = _tstol( tokens[ 2 ] );
					break;
				}
				case 5: // 23:30:00 1996/1/25 
				{
					nHour = _tstol( tokens[ 0 ] );
					nMinute = _tstol( tokens[ 1 ] );
					nSecond = _tstol( tokens[ 2 ] );
					nMonth = _tstol( tokens[ 0 ] );
					nDay = _tstol( tokens[ 1 ] );
					break;
				}
				default :
				{
					// 11/25/96 23:30:00 
					nMonth = _tstol( tokens[ 0 ] );
					nDay = _tstol( tokens[ 1 ] );
					nYear = _tstol( tokens[ 2 ] );
					if 
					( 
						nMonth < 1 || nMonth > 12 ||
						nDay < 1 || nDay > 31
					)
					{
						// 96/1/25 23:30:00
						nYear = _tstol( tokens[ 0 ] );
						nMonth = _tstol( tokens[ 1 ] );
						nDay = _tstol( tokens[ 2 ] );
					} 
				}
			}
		}
	}

	if ( bAM )
	{
		nHour--;

	} else if ( bPM )
	{
		nHour += 11;
	}

	if ( nYear < 100 )
	{
		if ( nYear > nCurrentYear )
		{
			nYear += 1900;

		} else
		{
			nYear += 2000;
		}
	}

	m_Date.Year = nYear;
	m_Date.Month = nMonth;
	m_Date.Day = nDay;
	m_Date.Hour = nHour;
	m_Date.Minute = nMinute;
	m_Date.Second = nSecond;

	value = m_Date.Okay;

	return value;
} // CreateValidDate

/////////////////////////////////////////////////////////////////////////////
// sets the time based on the given string which may not abide by the locale
bool SetTime( CString csDate )
{
	bool value = false;

	COleDateTime oDT;
	COleDateTime::DateTimeStatus eStatus = COleDateTime::invalid;
	if ( !csDate.IsEmpty() )
	{
		// this will fail if not formatted in the correct locale
		oDT.ParseDateTime( csDate );
		eStatus = oDT.GetStatus();

		if ( eStatus != COleDateTime::valid )
		{
			// test to see if the string ends in AM or PM
			const bool bAmPm = csDate.Right( 1 ).MakeLower() == _T( "m" );
			CString csTime;
			if ( bAmPm )
			{
				// something like 03:15:30 PM
				csTime = csDate.Right( 11 );

			} else // something like 03:15:30
			{
				csTime = csDate.Right( 8 );
			}

			// time alone should parse correctly and since time
			// is all we are interested in at this point, that 
			// is okay
			oDT.ParseDateTime( csTime );
			eStatus = oDT.GetStatus();
		}

		// if we have a valid status, update the time
		if ( eStatus == COleDateTime::valid )
		{
			value = true;
			m_Date.Hour = oDT.GetHour();
			m_Date.Minute = oDT.GetMinute();
			m_Date.Second = oDT.GetSecond();
		}
	}

	return value;
} // SetTime

/////////////////////////////////////////////////////////////////////////////
// get the current date taken, if any, from the given filename
CString GetCurrentDateTaken( LPCTSTR lpszPathName )
{
	USES_CONVERSION;

	CString value;

	// smart pointer to the image representing this file
	unique_ptr<Gdiplus::Image> pImage =
		unique_ptr<Gdiplus::Image>
		(
			Gdiplus::Image::FromFile( T2CW( lpszPathName ) )
		);

	// test the date properties stored in the given image
	CString csOriginal =
		GetStringProperty( pImage.get(), PropertyTagExifDTOrig );
	CString csDigitized =
		GetStringProperty( pImage.get(), PropertyTagExifDTDigitized );

	if ( SetTime( csOriginal ) )
	{
		value = csOriginal;

	} else if ( SetTime( csDigitized ) )
	{
		value = csDigitized;
	}

	return value;
} // GetCurrentDateTaken

/////////////////////////////////////////////////////////////////////////////
// Save the data inside pImage to the given lpszPathName
bool Save( LPCTSTR lpszPathName, Gdiplus::Image* pImage )
{
	USES_CONVERSION;

	CString csExt = GetExtension( lpszPathName );

	// save and overwrite the selected image file with current page
	int iValue =
		EncoderValue::EncoderValueVersionGif89 |
		EncoderValue::EncoderValueCompressionLZW |
		EncoderValue::EncoderValueFlush;

	EncoderParameters param;
	param.Count = 1;
	param.Parameter[ 0 ].Guid = EncoderSaveFlag;
	param.Parameter[ 0 ].Value = &iValue;
	param.Parameter[ 0 ].Type = EncoderParameterValueTypeLong;
	param.Parameter[ 0 ].NumberOfValues = 1;

	// writing to the same file will fail, so save to a corrected folder
	// below the image being corrected
	const CString csCorrected = GetCorrectedFolder();
	const CString csFolder = GetFolder( lpszPathName ) + csCorrected;
	if ( !::PathFileExists( csFolder ) )
	{
		if ( !CreatePath( csFolder ) )
		{
			return false;
		}
	}

	// filename plus extension
	const CString csData = GetDataName( lpszPathName );
	const CString csPath = csFolder + _T( "\\" ) + csData;

	CLSID clsid = m_Extension.ClassID;
	Status status = pImage->Save( T2CW( csPath ), &clsid, &param );
	return status == Ok;
} // Save

/////////////////////////////////////////////////////////////////////////////
// crawl through the directory tree
void RecursePath( LPCTSTR path )
{
	USES_CONVERSION;

	// valid file extensions
	const CString csValidExt = _T( ".jpg;.jpeg;.png;.gif;.bmp;.tif;.tiff" );

	// the new folder under the image folder to contain the corrected images
	const CString csCorrected = GetCorrectedFolder();
	const int nCorrected = GetCorrectedFolderLength();

	CString csPathname( path );
	csPathname.TrimRight( _T( "\\" ) );

	// build a string with wildcards
	CString strWildcard;
	strWildcard.Format( _T( "%s\\*.*" ), path );

	// start trolling for files we are interested in
	CFileFind finder;
	BOOL bWorking = finder.FindFile( strWildcard );
	while ( bWorking )
	{
		bWorking = finder.FindNextFile();

		// skip "." and ".." folder names
		if ( finder.IsDots() )
		{
			continue;
		}

		// if it's a directory, recursively search it
		if ( finder.IsDirectory() )
		{
			const CString str = finder.GetFilePath();

			// if the user did not specify recursing into sub-folders
			// then we can ignore any directories found
			if ( m_bRecurse == false )
			{
				continue;
			}

			// do not recurse into the corrected folder
			if ( str.Right( nCorrected ) == csCorrected )
			{
				continue;
			}

			// recurse into the new directory
			RecursePath( str );

		} else // write the properties if it is a valid extension
		{
			const CString csPath = finder.GetFilePath();
			const CString csExt = GetExtension( csPath ).MakeLower();
			const CString csFile = GetFileName( csPath );

			if ( -1 != csValidExt.Find( csExt ) )
			{
				m_Extension.FileExtension = csExt;

				CStdioFile fout( stdout );
				fout.WriteString( csPath + _T( "\n" ) );

				// read the current date and time from the metadata
				const CString csDateTaken = GetCurrentDateTaken( csPath );

				// if the date taken is empty, there is nothing for us
				// to do
				CString csOutput;
				if ( csDateTaken.IsEmpty() )
				{
					fout.WriteString( _T( ".\n") );
					fout.WriteString( _T( "Date Taken is missing.\n" ));
					fout.WriteString( _T( ".\n" ) );
					continue;
				}

				bool bValid = CreateValidDate( csDateTaken );
				if ( !bValid )
				{
					csOutput.Format
					( 
						_T( "Date taken is invalid: %s.\n" ), csDateTaken 
					);
					fout.WriteString( _T( ".\n" ) );
					fout.WriteString( csOutput );
					fout.WriteString( _T( ".\n" ) );
					continue;

				} else
				{
					csOutput.Format
					(
						_T( "Date taken is: %s.\n" ), csDateTaken
					);
					fout.WriteString( _T( ".\n" ) );
					fout.WriteString( csOutput );
					fout.WriteString( _T( ".\n" ) );
				}

				COleDateTime oDT = m_Date.GetDateAndTime();

				// convert the offset in hours to days
				const double dOffset = double( m_nHourOffset ) / 24;

				// modify the date and time with the offset
				oDT.m_dt += dOffset; 

				// change the date
				m_Date.DateAndTime = oDT;

				CString csDate = m_Date.Date;
				bValid = m_Date.Okay;
				if ( !bValid )
				{
					csOutput.Format
					( 
						_T( "New Date taken is invalid: %s.\n" ), csDate
					);
					fout.WriteString( _T( ".\n" ) );
					fout.WriteString( csOutput );
					continue;
					fout.WriteString( _T( ".\n" ) );
					continue;
				}

				csOutput.Format( _T( "New Date: %s\n" ), csDate );
				fout.WriteString( csOutput );

				// smart pointer to the image representing this element
				unique_ptr<Gdiplus::Image> pImage =
					unique_ptr<Gdiplus::Image>
					(
						Gdiplus::Image::FromFile( T2CW( csPath ) )
					);

				// smart pointer to the original date property item
				unique_ptr<Gdiplus::PropertyItem> pOriginalDateItem =
					unique_ptr<Gdiplus::PropertyItem>( new Gdiplus::PropertyItem );
				pOriginalDateItem->id = PropertyTagExifDTOrig;
				pOriginalDateItem->type = PropertyTagTypeASCII;
				pOriginalDateItem->length = csDate.GetLength() + 1;
				pOriginalDateItem->value = csDate.GetBuffer( pOriginalDateItem->length );

				// smart pointer to the digitized date property item
				unique_ptr<Gdiplus::PropertyItem> pDigitizedDateItem =
					unique_ptr<Gdiplus::PropertyItem>( new Gdiplus::PropertyItem );
				pDigitizedDateItem->id = PropertyTagExifDTDigitized;
				pDigitizedDateItem->type = PropertyTagTypeASCII;
				pDigitizedDateItem->length = csDate.GetLength() + 1;
				pDigitizedDateItem->value = csDate.GetBuffer( pDigitizedDateItem->length );

				// if these properties exist they will be replaced
				// if these properties do not exist they will be created
				pImage->SetPropertyItem( pOriginalDateItem.get() );
				pImage->SetPropertyItem( pDigitizedDateItem.get() );

				// save the image to the new path
				Save( csPath, pImage.get() );

				// release the date buffer
				csDate.ReleaseBuffer();
			}
		}
	}

	finder.Close();

} // RecursePath

/////////////////////////////////////////////////////////////////////////////
// set the current file extension which will automatically lookup the
// related mime type and class ID and set their respective properties
void CExtension::SetFileExtension( CString value )
{
	m_csFileExtension = value;

	if ( m_mapExtensions.Exists[ value ] )
	{
		MimeType = *m_mapExtensions.find( value );

		// populate the mime type map the first time it is referenced
		if ( m_mapMimeTypes.Count == 0 )
		{
			UINT num = 0;
			UINT size = 0;

			// gets the number of available image encoders and 
			// the total size of the array
			Gdiplus::GetImageEncodersSize( &num, &size );
			if ( size == 0 )
			{
				return;
			}

			ImageCodecInfo* pImageCodecInfo =
				(ImageCodecInfo*)malloc( size );
			if ( pImageCodecInfo == nullptr )
			{
				return;
			}

			// Returns an array of ImageCodecInfo objects that contain 
			// information about the image encoders built into GDI+.
			Gdiplus::GetImageEncoders( num, size, pImageCodecInfo );

			// populate the map of mime types the first time it is 
			// needed
			for ( UINT nIndex = 0; nIndex < num; ++nIndex )
			{
				CString csKey;
				csKey = CW2A( pImageCodecInfo[ nIndex ].MimeType );
				CLSID classID = pImageCodecInfo[ nIndex ].Clsid;
				m_mapMimeTypes.add( csKey, new CLSID( classID ) );
			}

			// clean up
			free( pImageCodecInfo );
		}

		ClassID = *m_mapMimeTypes.find( MimeType );

	} else
	{
		MimeType = _T( "" );
	}
} // CExtension::SetFileExtension

/////////////////////////////////////////////////////////////////////////////
// a console application that can crawl through the file
// system and troll for image metadata properties
int _tmain( int argc, TCHAR* argv[], TCHAR* envp[] )
{
	HMODULE hModule = ::GetModuleHandle( NULL );
	if ( hModule == NULL )
	{
		_tprintf( _T( "Fatal Error: GetModuleHandle failed\n" ) );
		return 1;
	}

	// initialize MFC and error on failure
	if ( !AfxWinInit( hModule, NULL, ::GetCommandLine(), 0 ) )
	{
		_tprintf( _T( "Fatal Error: MFC initialization failed\n " ) );
		return 2;
	}

	CStdioFile fOut( stdout );

	// if the expected number of parameters are not found
	// give the user some usage information
	if ( argc != 3 && argc != 4 )
	{
		fOut.WriteString( _T( ".\n" ) );
		fOut.WriteString
		(
			_T( "FixDateTaken, Copyright (c) 2020, " )
			_T( "by W. T. Block.\n" )
		);
		fOut.WriteString( _T( ".\n" ) );
		fOut.WriteString( _T( "Usage:\n" ) );
		fOut.WriteString( _T( ".\n" ) );
		fOut.WriteString( _T( ".  OffsetHours pathname hour_offset [recurse_folders]\n" ) );
		fOut.WriteString( _T( ".\n" ) );
		fOut.WriteString( _T( "Where:\n" ) );
		fOut.WriteString( _T( ".\n" ) );
		fOut.WriteString
		(
			_T( ".  pathname is the root of the tree to be scanned.\n" )
		);
		fOut.WriteString
		(
			_T( ".  hour_offset is the number of hours to offset the date taken by.\n" )
		);
		fOut.WriteString
		(
			_T( ".  recurse_folders is optional true | false parameter\n" )
			_T( ".    to include sub-folders or not (default is false).\n" )
		);
		fOut.WriteString( _T( ".\n" ) );
		return 3;
	}

	// display the executable path
	CString csMessage;
	csMessage.Format( _T( "Executable pathname: %s\n" ), argv[ 0 ] );
	fOut.WriteString( _T( ".\n" ) );
	fOut.WriteString( csMessage );
	fOut.WriteString( _T( ".\n" ) );

	// retrieve the pathname and validate the pathname exists
	CString csPath = argv[ 1 ];
	if ( !::PathFileExists( csPath ) )
	{
		csMessage.Format( _T( "Invalid pathname: %s\n" ), csPath );
		fOut.WriteString( _T( ".\n" ) );
		fOut.WriteString( csMessage );
		fOut.WriteString( _T( ".\n" ) );
		return 4;

	} else
	{
		csMessage.Format( _T( "Given pathname: %s\n" ), csPath );
		fOut.WriteString( _T( ".\n" ) );
		fOut.WriteString( csMessage );
		fOut.WriteString( _T( ".\n" ) );
	}

	// get the number of hours to offset the date taken metadata
	m_nHourOffset = _tstol( argv[ 2 ] );
	if ( m_nHourOffset == 0 )
	{
		csMessage.Format( _T( "Invalid hour offset: %s\n" ), argv[ 2 ] );
		fOut.WriteString( _T( ".\n" ) );
		fOut.WriteString( csMessage );
		fOut.WriteString( _T( ".\n" ) );
		return 5;
	}

	// start up COM
	AfxOleInit();
	::CoInitialize( NULL );

	// reference to GDI+
	InitGdiplus();

	// crawl through directory tree defined by the command line
	// parameter trolling for image files
	RecursePath( csPath );

	// clean up references to GDI+
	TerminateGdiplus();

	// all is good
	return 0;

} // _tmain

/////////////////////////////////////////////////////////////////////////////
