// DBFile.cpp : implementation file
//

#include "stdafx.h"
#include "GenMAPPDBDL.h"
#include "DBFile.h"
#include ".\dbfile.h"


// CDBFile

CDBFile::CDBFile()
: szSrcFileName(_T(""))
, bFTP(false)
, bFileIsCompressed(false)
, nDownloadProgress(0)
, nFileType(0)
, dwFileSize(0)
, bUpdateAvailable(false)
, bSelected(false)
, bIncludesGenBank(false)
, DBDate(0)
, nID(-1)
, nLVID(-1)
{
}

CDBFile::~CDBFile()
{
}


// Returns a formatted string version of the file's size. Automatically shows MB or GB,
// depending on size.
CString CDBFile::GetStringFileSize(DWORD FileSize)
{
	CString szFileSize;
	if (FileSize < 1073741824)
		szFileSize.Format("%.2f MB", (double)FileSize / 1048576);
	else
		szFileSize.Format("%.2f GB", (double)FileSize / 1073741824);

	return szFileSize;
}

CString CDBFile::GetFileTypeAsString(int nFileType)
{
	CString szFileType;
	switch (nFileType)
	{
		case GENE_DATABASE:
			szFileType = "Gene Database";
			break;

		case MAPP_DATABASE:
			szFileType = "MAPP";
			break;

		case MAPP_ARCHIVE:
			szFileType = "MAPP Archive";
			break;

		case GENEEXPRESS_DATABASE:
			szFileType = "Expression Dataset";
			break;

		default:
			szFileType = "Unknown";
	}

	return szFileType;
}
CString CDBFile::GetFileNameOnly(CString szSrcFileName)
{
	return GetFileNameOnly(szSrcFileName, false);
}

CString CDBFile::GetFileNameOnly(CString szSrcFileName, bool bIncludeExtension)
{
	CString szFileNameOnly = RemoveEscapeSequences(szSrcFileName.Right(szSrcFileName.GetLength() - szSrcFileName.ReverseFind('/') - 1));
	if (bIncludeExtension)
		return szFileNameOnly;
	else
		return szFileNameOnly.Left(szFileNameOnly.ReverseFind('.'));
}

CString CDBFile::RemoveEscapeSequences(CString szURI)
{
	int	nPercentIndex = 0;
	CString szFinalString = "";
	CString szURIPiece = szURI.Tokenize("%", nPercentIndex);
	while (szURIPiece != "")
	{
		szFinalString+= szURIPiece;
		if (nPercentIndex < szURI.GetLength())
		{
			if (szURI[nPercentIndex] != '%')
			{
				CString szTemp;
				szTemp.Format("%c", TwoDigitHexStringToInt(szURI.Mid(nPercentIndex, 2)));
				szFinalString += szTemp;
				nPercentIndex += 2;
			}
			else
			{
				szFinalString+= szURIPiece + "%";
				nPercentIndex++;
			}
		}

		szURIPiece = szURI.Tokenize("%", nPercentIndex);
	}

	szFinalString+= szURIPiece;
	return szFinalString;
}

int CDBFile::GetDecValFromHex(char cHexDigit)
{
	char	szHexNumerals[] = "0123456789ABCDEF", szHexDigit[2];

	szHexDigit[0] = cHexDigit;
	szHexDigit[1] = '\0';

	return (int)strcspn(szHexNumerals, szHexDigit);
}

int CDBFile::TwoDigitHexStringToInt(CString szHexString)
{
	int	nTheNumber = GetDecValFromHex(szHexString[0]);
	
	nTheNumber<<= 4;
    
	return nTheNumber+= GetDecValFromHex(szHexString[1]);
}