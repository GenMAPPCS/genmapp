#pragma once
#include "atltime.h"

// CDBFile command target

class CDBFile : public CObject
{
public:
	CDBFile();
	virtual ~CDBFile();
	// Name of the source file, including path. Can be CD-ROM or FTP based path.
	CString szSrcFileName;
	// Indicate whether source file exists on CD-ROM or FTP site
	bool bFTP;
	// Inidicates whether or not the source file is compressed
	bool bFileIsCompressed;
	// Percentage of download complete for this file
	int nDownloadProgress;
	// Describes the type of file. The value is on of GMDBType enum values.
	int nFileType;
	// Size of the database file, in KB.
	DWORD dwFileSize;
	// Indicates that an update is available on the FTP site.
	bool bUpdateAvailable;
	// Inidicates that file has been selected for download/copy
	bool bSelected;
	// If true, indicates that database contains GenBank data.
	bool bIncludesGenBank;
	CTime DBDate;
	// Returns a formatted string version of the file's size. Automatically shows MB or GB, depending on size.
	static CString GetStringFileSize(DWORD FileSize);
	static CString GetFileTypeAsString(int nFileType);
	// ID of the file. This matches the index in the Download Queue list box.
	int nID;
	// This file's index when displayed in the ListView control
	int nLVID;
	// Database server to which this file belongs. Must be cast to a CGMFTP object.
	CObList DBServerList;
	static CString GetFileNameOnly(CString szSrcFileName, bool bIncludeExtension);
	static CString GetFileNameOnly(CString szSrcFileName);
	static CString RemoveEscapeSequences(CString szURI);
	static int GetDecValFromHex(char cHexDigit);
	static int TwoDigitHexStringToInt(CString szHexString);
};


