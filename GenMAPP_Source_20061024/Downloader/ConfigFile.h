#pragma once
#include "afx.h"

#define		CFG_FILE_BUF		8192

class CConfigFile :
	public CStdioFile
{
public:
	CConfigFile(CString szBasePath);
	CConfigFile(CString szDLLPath, CString szFileName);
	~CConfigFile(void);
	// Read a string value from the configuration file
	CString ReadStringKey(CString szKeyName, CString szSectionName);
	CString ReadStringKey(CString szKeyName);
private:
	// Config file object
public:
	// Write a key value in the GenMAPP configuration file.
	bool WriteStringKey(CString szKeyName, CString szValue);
	bool WriteStringKey(CString szKeyName, CString szValue, CString szSection);
		// Returns a list of folders from DataLoacations.cfg that belong to the specified server.
	CStringList* RetreiveFoldersForServer(CString szServerName, bool bHTTP);
	CObList* RetreiveServerEntries();

private:
	CString m_szCFGPath;
	// Sets the file pointer to the line following the specified section
	BOOL FindSection(CString szSectionName);
public:
	// Returns the individual files associated with a server
	BOOL RetrieveFilesForServer(CString szServerName, CStringList* pFileList, CStringList* pFolderList, bool bHTTP);
	bool TestWriteInTheMiddle(CString szWhatToWrite);
private:
	// Takes a config file line and expands any replaceable paramters found in the line
	static CString ExpandReplaceableParam(CString szConfigLine);
public:
	// Provides a means to iterate through a config file. Returns false if there are no more lines matching the criteria.
	bool RetrieveNextKey(int nLineNumber, CString* pszKey, CString* pszValue, CString szSection);
private:
	// Either reads the [Servers] section or the [HTTP Servers] section of DataLocations.cfg. Returns an CObList containing CGMFTP objects representing the servers found.
	CObList* RetrieveServerEntriesByType(bool bHTTP);
};
