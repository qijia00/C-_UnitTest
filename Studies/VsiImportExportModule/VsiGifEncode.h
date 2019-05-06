/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiGifEncode.h
**
**	Description:
**		Code to enable animation of gif files
**
********************************************************************************/

#pragma once

/*********************************************************************************

GIF format

Byte Order: Little-endian

GIF Header
------------------------------------------------------------------------
	Offset   Length   Contents
	  0      3 bytes  "GIF"
	  3      3 bytes  "87a" or "89a"
	  6      2 bytes  <Logical Screen Width>
	  8      2 bytes  <Logical Screen Height>
	 10      1 byte   bit 0:    Global Color Table Flag (GCTF)
					  bit 1..3: Color Resolution
					  bit 4:    Sort Flag to Global Color Table
					  bit 5..7: Size of Global Color Table: 2^(1+n)
	 11      1 byte   <Background Color Index>
	 12      1 byte   <Pixel Aspect Ratio>
	 13      ? bytes  <Global Color Table(0..255 x 3 bytes) if GCTF is one>
			 ? bytes  <Blocks>
			 1 bytes  <Trailer> (0x3b)



Image Block
------------------------------------------------------------------------
	Offset   Length   Contents
	  0      1 byte   Image Separator (0x2c)
	  1      2 bytes  Image Left Position
	  3      2 bytes  Image Top Position
	  5      2 bytes  Image Width
	  7      2 bytes  Image Height
	  8      1 byte   bit 0:    Local Color Table Flag (LCTF)
					  bit 1:    Interlace Flag
					  bit 2:    Sort Flag
					  bit 2..3: Reserved
					  bit 4..7: Size of Local Color Table: 2^(1+n)
			 ? bytes  Local Color Table(0..255 x 3 bytes) if LCTF is one
			 1 byte   LZW Minimum Code Size
	[ // Blocks
			 1 byte   Block Size (s)
			(s)bytes  Image Data
	]*
			 1 byte   Block Terminator(0x00)



Graphic Control Extension Block
------------------------------------------------------------------------
	Offset   Length   Contents
	  0      1 byte   Extension Introducer (0x21)
	  1      1 byte   Graphic Control Label (0xf9)
	  2      1 byte   Block Size (0x04)
	  3      1 byte   bit 0..2: Reserved
					  bit 3..5: Disposal Method
					  bit 6:    User Input Flag
					  bit 7:    Transparent Color Flag
	  4      2 bytes  Delay Time (1/100ths of a second)
	  6      1 byte   Transparent Color Index
	  7      1 byte   Block Terminator(0x00)



Comment Extension Block
------------------------------------------------------------------------
	Offset   Length   Contents
	  0      1 byte   Extension Introducer (0x21)
	  1      1 byte   Comment Label (0xfe)
	[
			 1 byte   Block Size (s)
			(s)bytes  Comment Data
	]*
			 1 byte   Block Terminator(0x00)



Plain Text Extension Block
------------------------------------------------------------------------
	Offset   Length   Contents
	  0      1 byte   Extension Introducer (0x21)
	  1      1 byte   Plain Text Label (0x01)
	  2      1 byte   Block Size (0x0c)
	  3      2 bytes  Text Grid Left Position
	  5      2 bytes  Text Grid Top Position
	  7      2 bytes  Text Grid Width
	  9      2 bytes  Text Grid Height
	 10      1 byte   Character Cell Width(
	 11      1 byte   Character Cell Height
	 12      1 byte   Text Foreground Color Index(
	 13      1 byte   Text Background Color Index(
	[
			 1 byte   Block Size (s)
			(s)bytes  Plain Text Data
	]*
			 1 byte   Block Terminator(0x00)




Application Extension Block
------------------------------------------------------------------------
	Offset   Length   Contents
	  0      1 byte   Extension Introducer (0x21)
	  1      1 byte   Application Label (0xff)
	  2      1 byte   Block Size (0x0b)
	  3      8 bytes  Application Identifire
	[
			 1 byte   Block Size (s)
			(s)bytes  Application Data
	]*
			 1 byte   Block Terminator(0x00)



GIF87a:
------------------------------------------------------------------------
	GIF Header
	Image Block
	Trailer



GIF89a:
------------------------------------------------------------------------
	GIF Header
	Graphic Control Extension
	Image Block
	Trailer

GIF Animation
------------------------------------------------------------------------
	GIF Header
	Application Extension
	[
	  Graphic Control Extension
	  Image Block
	]*
	Trailer
*/

class CVsiGifEncode
{
private:
	HANDLE m_hOutput;
	DWORD m_dwFileSize;	

	USHORT m_usXRes;
	USHORT m_usYRes;
	USHORT m_usDelay100sSec;
	
public:
	CVsiGifEncode() :
		m_hOutput(NULL),
		m_dwFileSize(0),
		m_usXRes(0),
		m_usYRes(0),
		m_usDelay100sSec(0)
	{
	}
	~CVsiGifEncode()
	{
		_ASSERT(NULL == m_hOutput);
	}

	HRESULT Initialize(LPCWSTR pszOutput, USHORT usXRes, USHORT usYRes, USHORT usDelay100sSec)
	{		
		HRESULT hr = S_OK;

		try
		{
			VSI_VERIFY(m_hOutput == NULL, VSI_LOG_ERROR, NULL);

			m_hOutput = CreateFile(
				pszOutput, 
				GENERIC_WRITE, 
				FILE_SHARE_WRITE, 
				NULL,CREATE_ALWAYS, 
				0, 
				NULL);
			VSI_WIN32_VERIFY(m_hOutput != INVALID_HANDLE_VALUE && NULL != m_hOutput, VSI_LOG_ERROR, L"Failure opening gif image");
		
			m_usXRes = usXRes;
			m_usYRes = usYRes;
			m_usDelay100sSec = usDelay100sSec;
			m_dwFileSize = 0;

			hr = WriteGifHeader(usXRes, usYRes);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure writing gif header");

			hr = WriteNetscapeControlExt();
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure writing gif header");
		}
		VSI_CATCH(hr);
		
		return hr;
	}


	HRESULT AddImage(BYTE *pbtGifBuffer, DWORD dwDataSize)
	{
		HRESULT hr = S_OK;		

		try
		{
			m_dwFileSize = dwDataSize;

			BYTE pPalette[256 * 3];
			hr = ExtractImagePalette(pPalette, pbtGifBuffer);
			VSI_CHECK_HR(hr, VSI_LOG_WARNING, L"Failure extracting image palette");

			hr = WriteGraphicsControlExt(m_usDelay100sSec);
			VSI_CHECK_HR(hr, VSI_LOG_WARNING, L"Failure writing image descriptor");

			hr = WriteImageDescriptor(pPalette);
			VSI_CHECK_HR(hr, VSI_LOG_WARNING, L"Failure writing image descriptor");

			hr = WriteImageData(pbtGifBuffer);
			VSI_CHECK_HR(hr, VSI_LOG_WARNING, L"Failure writing image descriptor");			
		}
		VSI_CATCH(hr);
		
		m_dwFileSize = 0;
	
		return hr;
	}

	HRESULT Uninitialize()
	{
		DWORD dwWritten;
		HRESULT hr = S_OK;		

		try
		{
			BYTE btTrailer = 0x3B;

			// Write terminator
			if(NULL != m_hOutput)
			{
				BOOL bRet = WriteFile(m_hOutput, &btTrailer, sizeof(BYTE), &dwWritten, NULL);
				VSI_WIN32_VERIFY(bRet, VSI_LOG_WARNING, L"Failure writing file");	
			}
		}
		VSI_CATCH(hr);

		if(INVALID_HANDLE_VALUE != m_hOutput && NULL != m_hOutput)
		{
			CloseHandle(m_hOutput);
			m_hOutput = NULL;
		}

		m_dwFileSize = 0;

		return hr;
	}

private:	
	// Write primary gif header
	HRESULT WriteGifHeader(USHORT usXres, USHORT usYRes)
	{
		// don't pack this structure
		#pragma pack(push, 1)
		typedef struct tagVSI_GIF_HEADER
		{
			BYTE pbtGif89A[6];

			USHORT usXRes;
			USHORT usYRes;

			BYTE btPackedColor;				// Global Color Table Flag 1 bit
											// Color Resolution 3 bits
											// Sort flag 1 bit
											// size of global color table 3 bits
			BYTE btBackgroundColorIndex;
			BYTE btPixelAspectRatio;

			// no global color table

		} VSI_GIF_HEADER;
		#pragma pack(pop)

		DWORD dwWritten(0);
		VSI_GIF_HEADER header;
		HRESULT hr = S_OK;

		try
		{
			// structure header
			BYTE pbtGif89A[6] = {'G','I','F','8','9','a'};
			CopyMemory(header.pbtGif89A, pbtGif89A, 6);		// GIF89a

			header.usXRes = usXres;							// x resolution
			header.usYRes = usYRes;							// y resolution

			header.btPackedColor = 0x70;					// original colour resolution is 24 bits ??

			header.btBackgroundColorIndex = 0;				// no global colour table, not required
			header.btPixelAspectRatio = 0;					// no aspect ratio is given
								   							
			BOOL bRet = WriteFile(m_hOutput, &header, sizeof(header), &dwWritten, NULL);
			VSI_WIN32_VERIFY(bRet, VSI_LOG_WARNING, L"Failure writing file");			
		}
		VSI_CATCH(hr);

		return hr;
	}

	// Extract Image Descriptor from gif file and verify that source is reasonable
	HRESULT ExtractImagePalette(BYTE *pPalette, BYTE *pbtFile)
	{
		HRESULT hr = S_OK;	
		
		try
		{
			// Verify image is a gif file
			VSI_VERIFY(pbtFile[0]=='G' && pbtFile[1]=='I' && pbtFile[2]=='F' &&
					pbtFile[3]=='8' && pbtFile[4]=='9' && pbtFile[5]=='a', VSI_LOG_ERROR,
					L"Invalid file format");

			// Get resolution			
			USHORT usXRes = *(USHORT *)&pbtFile[6];
			USHORT usYRes = *(USHORT *)&pbtFile[6+2];

			// Verify that these resolutions match
			VSI_VERIFY(usXRes == m_usXRes, VSI_LOG_ERROR, L"Image resolutions don't match");
			VSI_VERIFY(usYRes == m_usYRes, VSI_LOG_ERROR, L"Image resolutions don't match");
		
			// Get packed fields from image
			BYTE btGlobalPackedColor = pbtFile[6 + 2 + 2];

			// Do we have a global colour table
			BOOL bGlobalColorTable = (btGlobalPackedColor & 0x80) > 0;
			BYTE btGlobalSizeColorTable = (btGlobalPackedColor & 0x7);		
			BOOL bGlobalSorted = (btGlobalPackedColor & 0x8) > 0;

			// Test that we are using an expected image type
			VSI_VERIFY(bGlobalColorTable, VSI_LOG_ERROR, L"No global color table");
			VSI_VERIFY(btGlobalSizeColorTable == 0x7, VSI_LOG_ERROR, L"Invalid colour table size");
			VSI_VERIFY(bGlobalSorted == FALSE, VSI_LOG_ERROR, L"Invalid table sort");
			
			DWORD dwColorTableSize = 3 * 256;
			ZeroMemory(pPalette, dwColorTableSize);

			CopyMemory(pPalette, &pbtFile[6+7], dwColorTableSize);
			
			// Verify what comes next is a valid image descriptor
			BYTE btImageSeparator = pbtFile[6+7+dwColorTableSize];
			VSI_VERIFY(btImageSeparator == 0x2C, VSI_LOG_ERROR, L"Invalid image separator");

			usXRes = *(USHORT *)&pbtFile[6+7+dwColorTableSize+1+2+2];
			usYRes = *(USHORT *)&pbtFile[6+7+dwColorTableSize+1+2+2+2];

			// Verify that these resolutions match
			VSI_VERIFY(usXRes == m_usXRes, VSI_LOG_ERROR, L"Image resolutions don't match");
			VSI_VERIFY(usYRes == m_usYRes, VSI_LOG_ERROR, L"Image resolutions don't match");

			BYTE btLocalPackedColor = pbtFile[6+7+dwColorTableSize+1+2+2+2+2];

			BOOL bLocalColorTable = (btLocalPackedColor & 0x80) > 0;
			BOOL bLocalInterlace = (btLocalPackedColor & 0x40) > 0;
			BOOL bLocalSorted = (btLocalPackedColor & 0x20) > 0;
			//BYTE btLocalSizeColorTable = (btLocalPackedColor & 0x7);

			VSI_VERIFY(!bLocalColorTable, VSI_LOG_ERROR, L"Local color table");
			VSI_VERIFY(!bLocalInterlace, VSI_LOG_ERROR, L"Local interlaced");			
			VSI_VERIFY(bLocalSorted == FALSE, VSI_LOG_ERROR, L"Invalid local sort");			
		}
		VSI_CATCH(hr);

		return hr;
	}

	// Write new Image Descriptor
	HRESULT WriteImageDescriptor(BYTE *pPalette)
	{	
		// don't pack this structure
		#pragma pack(push, 1)
		typedef struct tagVSI_IMAGE_DESCRIPTOR
		{
			BYTE btImageSeparator;			// 0x2c

			USHORT usImageLeft;				// 0
			USHORT usImageTop;				// 0
			USHORT usImageWidth;			// xres
			USHORT usImageHeight;			// yres
			BYTE btPackedColor;
			
			BYTE pPalette[256*3];			// local color table
				
		} VSI_IMAGE_DESCRIPTOR;
		#pragma pack(pop)

		DWORD dwWritten(0);
		VSI_IMAGE_DESCRIPTOR header;
		HRESULT hr = S_OK;
		
		try
		{
			header.btImageSeparator = 0x2C;
			header.usImageLeft = 0;
			header.usImageTop = 0;
			header.usImageWidth = m_usXRes;
			header.usImageHeight = m_usYRes;

			header.btPackedColor = 0x87;	// 1,0,0,0,0,0x7
											// local color table, not interlace, not sorted, 256 colors

			DWORD dwColorTableSize = 3 * 256;
			CopyMemory(header.pPalette, pPalette, dwColorTableSize);

			BOOL bRet = WriteFile(m_hOutput, &header, sizeof(header), &dwWritten, NULL);
			VSI_WIN32_VERIFY(bRet, VSI_LOG_WARNING, L"Failure writing file");			
		}
		VSI_CATCH(hr);

		return hr;
	}

	// Extract image data and write to end
	HRESULT WriteImageData(BYTE *pbtFile)
	{
		DWORD dwWritten;
		HRESULT hr = S_OK;

		try
		{			
			// Image data starts
			BYTE *pbtDataStart = &pbtFile[6 + 7 + 3 * 256 + 10];

			// Verify laste byte is trailer
			BYTE btBlockTerminator = pbtFile[m_dwFileSize - 2];
			VSI_VERIFY(btBlockTerminator == 0x00, VSI_LOG_ERROR, L"Invalid block terminator");

			BYTE btTrailer = pbtFile[m_dwFileSize - 1];
			VSI_VERIFY(btTrailer == 0x3B, VSI_LOG_ERROR, L"Invalid trailer");

			// Copy the rest of the image into the dest
			// Don't write trailer but do write block terminator
			BOOL bRet = WriteFile(m_hOutput, pbtDataStart, m_dwFileSize - (6 + 7 + 3 * 256 + 10 + 1), &dwWritten, NULL);
			VSI_WIN32_VERIFY(bRet, VSI_LOG_WARNING, L"Failure writing file");
		}
		VSI_CATCH(hr);
		
		return hr;
	}

	// Write graphics control extension
	HRESULT WriteGraphicsControlExt(USHORT usDelay100sSec)
	{	
		// don't pack this structure
		#pragma pack(push, 1)
		typedef struct tagVSI_GRAPHICS_CONTROL
		{
			BYTE btExtensionIntroducer;			// 0x21
			BYTE btGraphicControlLabel;			// 0xF9
			BYTE btBlockSize;					// 0x04

			BYTE btPackedFields;
			USHORT usDelayTime;
			BYTE btTransparentColor;
			BYTE btBlockTerminator;				// 0x00
				
		} VSI_GRAPHICS_CONTROL;
		#pragma pack(pop)

		DWORD dwWritten(0);
		VSI_GRAPHICS_CONTROL header;
		HRESULT hr = S_OK;
		
		try
		{
			header.btExtensionIntroducer = 0x21;
			header.btGraphicControlLabel = 0xF9;
			header.btBlockSize = 0x04;
			header.btPackedFields = 0x01 + 0x4;			// disposal previous, no user input, yes transparent color
														// disposal 0x4 = leave as is
														// disposal 0x8 = background
														// disposal
										
			header.usDelayTime = usDelay100sSec;	// in 100's of a second

			header.btTransparentColor = 255;
			header.btBlockTerminator = 0x00;

			BOOL bRet = WriteFile(m_hOutput, &header, sizeof(header), &dwWritten, NULL);
			VSI_WIN32_VERIFY(bRet, VSI_LOG_WARNING, L"Failure writing file");			
		}
		VSI_CATCH(hr);

		return hr;
	}

	// Write netscape control extension
	// byte   1       : 33 (hex 0x21) GIF Extension code
	// byte   2       : 255 (hex 0xFF) Application Extension Label
	// byte   3       : 11 (hex (0x0B) Length of Application Block 
	//				 (eleven bytes of data to follow)
	// bytes  4 to 11 : "NETSCAPE"
	// bytes 12 to 14 : "2.0"
	// byte  15       : 3 (hex 0x03) Length of Data Sub-Block 
	//				 (three bytes of data to follow)
	// byte  16       : 1 (hex 0x01)
	// bytes 17 to 18 : 0 to 65535, an unsigned integer in 
	//				 lo-hi byte format. This indicate the 
	//				 number of iterations the loop should 
	//				 be executed.
	// bytes 19       : 0 (hex 0x00) a Data Sub-block Terminator. 
	HRESULT WriteNetscapeControlExt()
	{	
		// don't pack this structure
		#pragma pack(push, 1)
		typedef struct tagVSI_NETSCAPE_EXTENSION
		{
			BYTE btGifExtensionCode;		// 0x21
			BYTE btAppExtensionlabel;		// 0xFF
			BYTE btLengthApplicationBlock;	// 0x0B
			BYTE pbtNETSCAPE[8];			// NETSCAPE
			BYTE pbtVersion[3];				// 2.0
			BYTE btLengthSubBlock;			// 0x03
			BYTE btData;					// 0x01
			USHORT usIterations;			// 0 for infinite (0xFF)
			BYTE btTerminator;				// 0x00
				
		} VSI_NETSCAPE_EXTENSION;
		#pragma pack(pop)

		DWORD dwWritten(0);
		VSI_NETSCAPE_EXTENSION header;
		HRESULT hr = S_OK;
		
		try
		{
			header.btGifExtensionCode = 0x21;
			header.btAppExtensionlabel = 0xFF;
			header.btLengthApplicationBlock = 0x0B;
			
			header.pbtNETSCAPE[0] = 'N';
			header.pbtNETSCAPE[1] = 'E';
			header.pbtNETSCAPE[2] = 'T';
			header.pbtNETSCAPE[3] = 'S';
			header.pbtNETSCAPE[4] = 'C';
			header.pbtNETSCAPE[5] = 'A';
			header.pbtNETSCAPE[6] = 'P';
			header.pbtNETSCAPE[7] = 'E';

			header.pbtVersion[0] = '2';
			header.pbtVersion[1] = '.';
			header.pbtVersion[2] = '0';

			header.btLengthSubBlock = 0x03;
			header.btData = 0x01;
			header.usIterations = 0;
			header.btTerminator = 0x00;

			BOOL bRet = WriteFile(m_hOutput, &header, sizeof(header), &dwWritten, NULL);
			VSI_WIN32_VERIFY(bRet, VSI_LOG_WARNING, L"Failure writing file");			
		}
		VSI_CATCH(hr);

		return hr;
	}
};
