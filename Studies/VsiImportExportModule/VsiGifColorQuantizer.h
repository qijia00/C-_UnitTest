/*******************************************************************************
**
**  Copyright (c) 1999-2011 VisualSonics Inc. All Rights Reserved.
**
**  VsiGifColorQuantizer.h
**
**	Description:
**		Implementation of CVsiGifColorQuantizer
**		Based on CVsiGifColorQuantizer from Jeff Prosise, Davide Pizzolato - www.xdp.it
**
********************************************************************************/

#pragma once

#include <hash_map>

template<class _Kty, class _Pr = _STD less<_Kty>>
class VsiQuantizeHashCompare : public stdext::hash_compare<_Kty, _Pr>
{
public:
	enum
	{	// parameters for hash table
		bucket_size = 1,			// 0 < bucket_size
		min_buckets = 256			// min_buckets = 2 ^^ N, 0 < N
	};
};

struct lessDWORD : public std::binary_function<DWORD, DWORD, bool>
{
	bool operator()(DWORD x, DWORD y) const
	{ return x < y; }
};

typedef stdext::hash_map<DWORD, BYTE, VsiQuantizeHashCompare<DWORD, lessDWORD>> CVsiQuantizeMap;
typedef CVsiQuantizeMap::iterator CVsiQuantizeMapIter;

class CVsiGifColorQuantizer
{
private:
	typedef struct _NODE 
	{
		BOOL bIsLeaf;               // TRUE if node has no children
		DWORD dwPixelCount;         // Number of pixels represented by this leaf
		DWORD dwRedSum;             // Sum of red components
		DWORD dwGreenSum;           // Sum of green components
		DWORD dwBlueSum;            // Sum of blue components
		DWORD dwAlphaSum;           // Sum of alpha components
		struct _NODE* pChild[8];    // Pointers to child nodes
		struct _NODE* pNext;        // Pointer to next reducible node
	} NODE;

    NODE* m_pTree;
    DWORD m_dwLeafCount;
    NODE* m_pReducibleNodes[9];
    DWORD m_dwMaxColors;  
    DWORD m_dwColorBits;

	RGBQUAD m_prgbPal[256];	

	CVsiQuantizeMap m_colorMap;

	int m_iWidth;
	int m_iHeight;
	int m_iPitch;

public:
    CVsiGifColorQuantizer();
    ~CVsiGifColorQuantizer ();

	HRESULT Initialize(
		int iWidth, 
		int iHeight, 
		int iPitch,
		DWORD dwBPP,
		BOOL bAllowTransparency);
    HRESULT ProcessImage(BYTE *pbBits);
	HRESULT Uninitialize();

    DWORD GetColorCount();
    void GetColorTable(RGBQUAD* prgb);
	CVsiQuantizeMap *GetColorMap();

private:
    void AddColor (NODE** ppNode, BYTE r, BYTE g, BYTE b, BYTE a, DWORD dwColorBits,
        DWORD nLevel, DWORD* pLeafCount, NODE** pReducibleNodes);

    void* CreateNode (DWORD nLevel, DWORD dwColorBits, DWORD* pLeafCount,
        NODE** pReducibleNodes);

    void ReduceTree (DWORD dwColorBits, DWORD* pLeafCount,
        NODE** pReducibleNodes);

    void DeleteTree (NODE** ppNode);

    void GetPaletteColors(NODE* pTree, RGBQUAD* prgb, DWORD* pIndex, DWORD* pSum);
	void CalculateColorTable(RGBQUAD* prgb);
	
	inline BYTE GetNearestIndex(RGBQUAD q)
	{	
		DWORD dwErrorBest(UINT_MAX);		
		DWORD dwError;
		BYTE btIndexBest(0);

		// Index zero is always 0, 0, 0 - most common colour most likely
		if(q.rgbBlue == 0 && q.rgbGreen == 0 && q.rgbRed == 0)
			return 0;
		
		for(int i=0;i<256;i++)
		{
			// Not sure what one is more efficient

			//dwError = (m_prgbPal[i].rgbBlue-q.rgbBlue)*(m_prgbPal[i].rgbBlue-q.rgbBlue)+
			//	(m_prgbPal[i].rgbGreen-q.rgbGreen)*(m_prgbPal[i].rgbGreen-q.rgbGreen)+
			//	(m_prgbPal[i].rgbRed-q.rgbRed)*(m_prgbPal[i].rgbRed-q.rgbRed);

			dwError = abs(m_prgbPal[i].rgbBlue-q.rgbBlue)+
				abs(m_prgbPal[i].rgbGreen-q.rgbGreen)+
				abs(m_prgbPal[i].rgbRed-q.rgbRed);

			if (dwError == 0)
			{
				return (BYTE)i;
			}
			if (dwError<dwErrorBest)
			{
				dwErrorBest = dwError;
				btIndexBest = (BYTE)i;
			}
		}

		return btIndexBest;
	}
};

