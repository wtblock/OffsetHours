/////////////////////////////////////////////////////////////////////////////
// Copyright © by W. T. Block, all rights reserved
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "OffsetHours.h"
#include "CHelper.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

/////////////////////////////////////////////////////////////////////////////
// The one and only application object
CWinApp theApp;

/////////////////////////////////////////////////////////////////////////////
// given an image pointer and an ASCII property ID, return the property value
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
// The date and time when the original image data was generated.
// For a digital still camera, this is the date and time the picture 
// was taken or recorded. The format is "YYYY:MM:DD HH:MM:SS" with time 
// shown in 24-hour format, and the date and time separated by one blank 
// character (hex 20).
void CDate::SetDateTaken( CString csDate )
{
	// reset the m_Date to undefined state
	Year = -1;
	Month = -1;
	Day = -1;
	Hour = 0;
	Minute = 0;
	Second = 0;
	bool value = Okay;

	// parse the date into a vector of string tokens
	const CString csDelim( _T( ": " ) );
	int nStart = 0;
	vector<CString> tokens;

	do
	{
		const CString csToken = 
				csDate.Tokenize( csDelim, nStart ).MakeLower();
		if ( csToken.IsEmpty() )
		{
			break;
		}

		tokens.push_back( csToken );

	} while ( true );

	// there should be six tokens in the proper format of
	// "YYYY:MM:DD HH:MM:SS"
	const size_t tTokens = tokens.size();
	if ( tTokens != 6 )
	{
		return;
	}

	// populate the date and time members with the values
	// in the vector
	TOKEN_NAME eToken = tnYear;
	int nToken = 0;

	for ( CString csToken : tokens )
	{
		int nValue = _tstol( csToken );

		switch ( eToken )
		{
			case tnYear:
			{
				Year = nValue;
				break;
			}
			case tnMonth:
			{
				Month = nValue;
				break;
			}
			case tnDay:
			{
				Day = nValue;
				break;
			}
			case tnHour:
			{
				Hour = nValue;
				break;
			}
			case tnMinute:
			{
				Minute = nValue;
				break;
			}
			case tnSecond:
			{
				Second = nValue;
				break;
			}
		}

		nToken++;
		eToken = (TOKEN_NAME)nToken;
	}

	// this will be true if all of the values define a proper date and time
	value = Okay;

} // SetDateTaken

/////////////////////////////////////////////////////////////////////////////
// get the current date taken, if any, from the given filename
// which should be in the format "YYYY:MM:DD HH:MM:SS"
CString GetCurrentDateTaken( LPCTSTR lpszPathName )
{
	USES_CONVERSION;

	CString value;

	// smart pointer to the image representing this file
	// (smart pointer release their resources when they
	// go out of context)
	unique_ptr<Gdiplus::Image> pImage =
		unique_ptr<Gdiplus::Image>
		(
			Gdiplus::Image::FromFile( T2CW( lpszPathName ) )
		);

	// test the date properties stored in the given image
	const CString csOriginal =
		GetStringProperty( pImage.get(), PropertyTagExifDTOrig );
	const CString csDigitized =
		GetStringProperty( pImage.get(), PropertyTagExifDTDigitized );

	// officially the original property is the date taken in this
	// format: "YYYY:MM:DD HH:MM:SS"
	m_Date.DateTaken = csOriginal;
	if ( m_Date.Okay )
	{
		value = csOriginal;

	} else // alternately use the date digitized
	{
		m_Date.DateTaken = csDigitized;
		if ( m_Date.Okay )
		{
			value = csDigitized;
		}
	}

	return value;
} // GetCurrentDateTaken

/////////////////////////////////////////////////////////////////////////////
// Save the data inside pImage to the given lpszPathName
bool Save( LPCTSTR lpszPathName, Gdiplus::Image* pImage )
{
	USES_CONVERSION;

	CString csExt = CHelper::GetExtension( lpszPathName );

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
	const CString csFolder = CHelper::GetFolder( lpszPathName ) + csCorrected;
	if ( !::PathFileExists( csFolder ) )
	{
		if ( !CreatePath( csFolder ) )
		{
			return false;
		}
	}

	// filename plus extension
	const CString csData = CHelper::GetDataName( lpszPathName );
	const CString csPath = csFolder + _T( "\\" ) + csData;

	CLSID clsid = m_Extension.ClassID;
	Status status = pImage->Save( T2CW( csPath ), &clsid, &param );
	return status == Ok;
} // Save

/////////////////////////////////////////////////////////////////////////////
// crawl through the given directory tree which may include wild cards
void RecursePath( LPCTSTR path )
{
	USES_CONVERSION;

	// valid file extensions
	const CString csValidExt = _T( ".jpg;.jpeg;.png;.gif;.bmp;.tif;.tiff" );

	// the new folder under the image folder to contain the corrected images
	const CString csCorrected = GetCorrectedFolder();
	const int nCorrected = GetCorrectedFolderLength();

	// get the folder which will trim any wild card data
	CString csPathname = CHelper::GetFolder( path );

	// wild cards are in use if the pathname does not equal the given path
	const bool bWildCards = csPathname != path;
	csPathname.TrimRight( _T( "\\" ) );
	CString csData;

	// build a string with wild-cards
	CString strWildcard;
	if ( bWildCards )
	{
		csData = CHelper::GetDataName( path );
		strWildcard.Format( _T( "%s\\%s" ), csPathname, csData );

	} else // no wild cards, just a folder
	{
		strWildcard.Format( _T( "%s\\*.*" ), csPathname );
	}

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

			// if the user did not specify recursing into sub-folders
			// then we can ignore any directories found
			if ( m_bRecurse == false )
			{
				continue;
			}

			// do not recurse into the corrected folder
			const CString str = finder.GetFilePath().TrimRight( _T( "\\" ) );
			if ( str.Right( nCorrected ) == csCorrected )
			{
				continue;
			}

			// if wild cards in use, build a path with the wild cards
			if ( bWildCards )
			{
				CString csPath;
				csPath.Format( _T( "%s\\%s" ), str, csData );

				// recurse into the new directory with wild cards
				RecursePath( csPath );

			} else // recurse into the new directory
			{
				RecursePath( str + _T( "\\" ) );
			}

		} else // write the properties if it is a valid extension
		{
			const CString csPath = finder.GetFilePath();
			const CString csExt = CHelper::GetExtension( csPath ).MakeLower();
			const CString csFile = CHelper::GetFileName( csPath );

			if ( -1 != csValidExt.Find( csExt ) )
			{
				m_Extension.FileExtension = csExt;

				CStdioFile fout( stdout );
				fout.WriteString( csPath + _T( "\n" ) );

				// read the current date and time from the metadata
				// which should be in this format from the image
				// "YYYY:MM:DD HH:MM:SS"
				const CString csDateTaken = GetCurrentDateTaken( csPath );

				// if the date taken is empty, there is nothing for us
				// to do
				CString csOutput;
				if ( csDateTaken.IsEmpty() )
				{
					fout.WriteString( _T( ".\n") );
					fout.WriteString( _T( "Old Date Taken is missing.\n" ));
					fout.WriteString( _T( ".\n" ) );
					continue;
				}

				m_Date.DateTaken = csDateTaken;
				bool bValid = m_Date.Okay;
				if ( !bValid )
				{
					csOutput.Format
					( 
						_T( "Old Date Taken is invalid: %s.\n" ), csDateTaken 
					);
					fout.WriteString( _T( ".\n" ) );
					fout.WriteString( csOutput );
					fout.WriteString( _T( ".\n" ) );
					continue;

				} else
				{
					csOutput.Format
					(
						_T( "Old Date Taken is: %s.\n" ), csDateTaken
					);
					fout.WriteString( csOutput );
				}

				// get the date and time from the m_Date member class
				COleDateTime oDT = m_Date.DateAndTime;

				// convert the offset in hours to days which is the 
				// internal representation of the COleDateTime class
				const double dOffset = m_dHourOffset / 24.0;

				// modify the date and time by adding in the offset
				oDT.m_dt += dOffset; 

				// change the date
				m_Date.DateAndTime = oDT;

				CString csDate = m_Date.Date;
				bValid = m_Date.Okay;
				if ( !bValid )
				{
					csOutput.Format
					( 
						_T( "New Date Taken is invalid: %s.\n" ), csDate
					);
					fout.WriteString( _T( ".\n" ) );
					fout.WriteString( csOutput );
					fout.WriteString( _T( ".\n" ) );
					continue;
				}

				csOutput.Format( _T( "New Date Taken is: %s\n" ), csDate );
				fout.WriteString( csOutput );
				fout.WriteString( _T( ".\n" ) );

				// smart pointer to the image representing this element
				unique_ptr<Gdiplus::Image> pImage =
					unique_ptr<Gdiplus::Image>
					(
						Gdiplus::Image::FromFile( T2CW( csPath ) )
					);

				// smart pointer to the original date property item
				// (smart pointers automatically release their resources
				// when they go out of context)
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

	// clean up and go home
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

			// a smart pointer to the image codex information
			unique_ptr<ImageCodecInfo> pImageCodecInfo =
				unique_ptr<ImageCodecInfo>
				( 
					(ImageCodecInfo*)malloc( size ) 
				);
			if ( pImageCodecInfo == nullptr )
			{
				return;
			}

			// Returns an array of ImageCodecInfo objects that contain 
			// information about the image encoders built into GDI+.
			Gdiplus::GetImageEncoders( num, size, pImageCodecInfo.get() );

			// populate the map of mime types the first time it is 
			// needed
			for ( UINT nIndex = 0; nIndex < num; ++nIndex )
			{
				CString csKey;
				csKey = CW2A( pImageCodecInfo.get()[ nIndex ].MimeType );
				CLSID classID = pImageCodecInfo.get()[ nIndex ].Clsid;
				m_mapMimeTypes.add( csKey, new CLSID( classID ) );
			}
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

	// do some common command line argument corrections
	vector<CString> arrArgs = CHelper::CorrectedCommandLine( argc, argv );
	size_t nArgs = arrArgs.size();

	CStdioFile fOut( stdout );
	CString csMessage;

	// display the number of arguments if not 1 to help the user 
	// understand what went wrong if there is an error in the
	// command line syntax
	if ( nArgs != 1 )
	{
		fOut.WriteString( _T( ".\n" ) );
		csMessage.Format( _T( "The number of parameters are %d\n.\n" ), nArgs - 1 );
		fOut.WriteString( csMessage );

		// display the arguments
		for ( int i = 1; i < nArgs; i++ )
		{
			csMessage.Format( _T( "Parameter %d is %s\n.\n" ), i, arrArgs[ i ] );
			fOut.WriteString( csMessage );
		}
	}

	// if the expected number of parameters are not found
	// give the user some usage information
	if ( nArgs != 3 && nArgs != 4 )
	{
		fOut.WriteString( _T( ".\n" ) );
		fOut.WriteString
		(
			_T( "OffsetHours, Copyright (c) 2020, " )
			_T( "by W. T. Block.\n" )
		);

		fOut.WriteString
		( 
			_T( ".\n" ) 
			_T( "Usage:\n" )
			_T( ".\n" )
			_T( ".  OffsetHours pathname hour_offset [recurse_folders]\n" )
			_T( ".\n" )
			_T( "Where:\n" )
			_T( ".\n" )
		);

		fOut.WriteString
		(
			_T( ".  pathname is the root of the tree to be scanned, but\n" )
			_T( ".  may contain wild cards like the following:\n" )
			_T( ".    \"c:\\Picture\\DisneyWorldMary2 *.JPG\"\n" )
			_T( ".  will process all files with that pattern, or\n" )
			_T( ".    \"c:\\Picture\\DisneyWorldMary2 231.JPG\"\n" )
			_T( ".  will process a single defined image file.\n" )
			_T( ".  (NOTE: using wild cards will prevent recursion\n" )
			_T( ".    into sub-folders because the folders will likely\n" )
			_T( ".    not fall into the same pattern and therefore\n" )
			_T( ".    sub-folders will not be found by the search).\n" )
		);
		fOut.WriteString
		(
			_T( ".  hour_offset is the number of hours to offset the date taken by.\n" )
			_T( ".    where 12 will change AM to PM, and\n" )
			_T( ".    where -12 will change PM to AM, and\n" )
			_T( ".    where 24 will change to the next day, and\n" )
			_T( ".    where -24 will change to the previous day, and\n" )
			_T( ".    where fractional values are okay:\n" )
			_T( ".      A minute is 0.0166666666666667.\n" )
			_T( ".      A second is 0.0002777777777777.\n" )
		);
		fOut.WriteString
		(
			_T( ".  recurse_folders is optional true | false parameter\n" )
			_T( ".    to include sub-folders or not (default is false).\n" )
			_T( ".  (NOTE: using wild cards will prevent recursion\n" )
			_T( ".    into sub-folders because the folders will likely\n" )
			_T( ".    not fall into the same pattern and therefore\n" )
			_T( ".    sub-folders will not be found by the search).\n" )
		);
		fOut.WriteString( _T( ".\n" ) );
		return 3;
	}

	// display the executable path
	//csMessage.Format( _T( "Executable pathname: %s\n" ), arrArgs[ 0 ] );
	//fOut.WriteString( _T( ".\n" ) );
	//fOut.WriteString( csMessage );
	//fOut.WriteString( _T( ".\n" ) );

	// retrieve the pathname which may include wild cards
	CString csPathParameter = arrArgs[ 1 ];

	// trim off any wild card data
	const CString csFolder = CHelper::GetFolder( csPathParameter );

	// test for current folder character (a period)
	bool bExists = csPathParameter == _T( "." );


	// if it is a period, add a wild card of *.* to retrieve
	// all folders and files
	if ( bExists )
	{
		csPathParameter = _T( ".\\*.*" );

	// if it is not a period, test to see if the folder exists
	} else 
	{
		if ( ::PathFileExists( csFolder ))
		{
			bExists = true;
		}
	}

	// give feedback to the user
	if ( !bExists )
	{
		csMessage.Format( _T( "Invalid pathname: %s\n" ), csPathParameter );
		fOut.WriteString( _T( ".\n" ) );
		fOut.WriteString( csMessage );
		fOut.WriteString( _T( ".\n" ) );
		return 4;

	} else
	{
		csMessage.Format( _T( "Given pathname: %s\n" ), csPathParameter );
		fOut.WriteString( _T( ".\n" ) );
		fOut.WriteString( csMessage );
		fOut.WriteString( _T( ".\n" ) );
	}

	// get the number of hours to offset the date taken metadata
	m_dHourOffset = _tstof( arrArgs[ 2 ] );

	if ( NearlyEqual( m_dHourOffset, 0.0 ))
	{
		csMessage.Format( _T( "Invalid hour offset: %s\n" ), arrArgs[ 2 ] );
		fOut.WriteString( _T( ".\n" ) );
		fOut.WriteString( csMessage );
		fOut.WriteString( _T( ".\n" ) );
		return 5;
	}

	// default to no recursion through sub-folders
	m_bRecurse = false;

	// test for the recursion parameter
	if ( nArgs == 4 )
	{
		CString csRecurse = arrArgs[ 3 ];
		csRecurse.MakeLower();

		// if the text is "true" the turn on recursion
		if ( csRecurse == _T( "true" ))
		{
			m_bRecurse = true;
		}
	}

	// start up COM
	AfxOleInit();
	::CoInitialize( NULL );

	// reference to GDI+
	InitGdiplus();

	// crawl through directory tree defined by the command line
	// parameter trolling for image files
	RecursePath( csPathParameter );

	// clean up references to GDI+
	TerminateGdiplus();

	// all is good
	return 0;

} // _tmain

/////////////////////////////////////////////////////////////////////////////
