#pragma once



// CGMFTP command target
#ifndef cgmftp_h
#define cgmftp_h


class CGMFTP : public CInternetSession
{
public:
	CGMFTP(LPCTSTR pstrAgent = NULL,
		DWORD dwContext = 1,
		DWORD dwAccessType = PRE_CONFIG_INTERNET_ACCESS,
		LPCTSTR pstrProxyName = NULL,
		LPCTSTR pstrProxyBypass = NULL,
		DWORD dwFlags = 0);
	virtual ~CGMFTP();
	double SpeedTest(CString szAltSpeedTestDir);
	CFtpConnection* m_pConnect;
	// Connect to the specified GenMAPP server. All operations will take place on this server.
	UINT ConnectToGMServer();
	BOOL RetrieveFromRootFolder(CString szRootFolder);
		// Host name or IP address of FTP server
	CString szHost;
	// User name for read access to FTP server
	CString szUserName;
	// Password for read access to FTP server
	CString szPassword;
	// Text string describing physical location of server, for example, "San Francisco, CA, USA"
	CString szLocation;
	// Friendly name, suitable for display to user, that describes FTP server
	CString szAlias;

public:
private:
	// Adds a Database File object to the list for each database file contained in the specified direcory.
	bool BuildListFromDBDir(int nDBType, CObList* pList);
	bool PopulateListFromType(int nDBType, CObList* pList, CString szExt);
public:
	// Connection speed in kilobytes per second.
	double fConnSpeed;
	// Changes the current directory based on the file type.
	bool ChangeGMDir(int nDirType);
	// Logs out of the GenMAPP Server
	bool DisconnectFromGMServer(void);
	// Populates the variables necessary to log in to the server
	void SetLoginData(CString szHostInit, CString szUserNameInit, CString szPasswordInit, CString szLocationInit, CString szAliasInit);
	// Populates a DBFile object given the FTP file object and the handle to the folder it belongs under
	CDBFile* PopulateDBFile(CString szFileName, DWORD dwFileSize, bool bFTP);
	// The key name of the server in DataLocations.cfg
	CString szServerKeyName;
private:
	// Recursively drills down the folder structure, adding the folders and files found at each level
	BOOL RecurseGMFolder(CString szRootFolder, CString szStartingURL);
	// Parses the file name to determine the file type. Returns on of the predefined constants.
	int DetermineFileType(CString szFileName);
public:
	// Returns the string representation of a file type constant. If the the constant is unknown, returns an empty string.
	static CString RetriveFileTypeString(int nFileType);
	// Adds the specified folder to the database folder tree
	static HTREEITEM AddFolderToTree(HTREEITEM hParent, CString szFolderName);
public:
	// Verifies the existence of a data file found in DataLocations.cfg and adds to the folder tree and file list
	BOOL RetrieveIndividualFile(CString szPathAndFileName, CString szDataTreeFolder);
private:
	// Returns a pointer to the CDBFile object if it already exists. Returns NULL otherwise.
	CDBFile* DoesFileExist(CString szFileName);
public:
	// Folder containing the SpeedTester.bin file.
	CString szSpeedTestFolder;
	// Folder containing the SpeedTester.bin file.
	CString szHTTPSpeedTestFolder;
	// Adds a the folder name to the folder tree and folder tree collection parented to the specified item and returns a handle to the new folder object. If the folder already exists, simply returns a pointer to the folder.
	static CDBFolder* AddFolder(CString szFolderName, int nFileType, HTREEITEM hParentFolder);
	// Indicates this object has an FTP connection to its target server.
	bool bConnected;
	// If true, indicates that this object is connecting to an HTTP server. If false, the object is connecting to an FTP server.
	bool bHTTPServer;
	// When this class to connect to an http server, this is used as the CHttpConnection object.
	CHttpConnection* m_pHttpConnect;
private:
	// FTP version of the speed test
	double SpeedTestFTP(CString szAltSpeedTestDir);
	// Http version of the Speed Test
	double SpeedTestHttp(CString szAltSpeedTestDir);
public:
	// Given a URL, retrieves a http file and stores it on the local file system. Returns false if a failure occurs.
	BOOL GetHttpFile(CString szHttpFile, CString szDestFile);
	// Indicates that this server is a backup for an FTP server. This server should only be used if FTP connectivity is not available.
	bool bFTPBackup;
private:
	// Parses one line of a HTTP folder listing returns results. If the return value is -1, there are no more files on this line. If return value is a positive number, use this as the nNextFileIndex in subsequent calls to retrieve the next HTTP file. If a file is found, returns a populated CDBFile object. If a folder is found, returns the folder name.
	CDBFile* FindNextHTTPFile(int* pnNextBRIndex, CString* szFolder, int nNextFileIndex, CString* szLine, CString szPathToFile);
public:
	// Opens an HTTP connection, creates an HTTPFile object, adds the appropriate headers and then returns the HTTPFile object as an InternetFile object.
	CInternetFile* OpenHTTPFile(CString szURL);
	// Parses one line in an HTML folder listing and returns the size of the file or 0 if the file size could not be determined.
	DWORD GetHTMLFileSize(CString szHTMLLine);
	// Tier level of the server. 1 is the top tier and has priority over lower tier servers.
	int nTier;
	
	// Test the HTTP connection by getting a directory listing
	bool TestHTTPConnection();

};

#endif
