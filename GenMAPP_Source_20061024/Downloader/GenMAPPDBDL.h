// GenMAPPDBDL.h : main header file for the GenMAPPDBDL DLL
//

#pragma once

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols
#include "afxcoll.h"

//	Possible statuses of m_nInetConnected
#define		NOT_CONNECTED	-1
#define		CONNECTING		0
#define		CONNECT_FAILED	1
#define		CONNECTED		2
#define		SPEED_TEST		3
#define		READ_DIR		4

#define		UNKNOWN_DATABASE		0
#define		GENE_DATABASE			1
#define		MAPP_DATABASE			2
#define		GENEEXPRESS_DATABASE	3
#define		MAPP_ARCHIVE			4

#define		SPEED_SAMPLES			16
#define		SPEED_SAMPLE_SIZE		32

UINT ConnectToGenMAPP( LPVOID pParam );
UINT AttemptFTPConnection( LPVOID pParam );
UINT DownloadFiles( LPVOID pParam );
UINT ExtractFiles(LPVOID pParam);
CString GetDestPathFromDBFile(CDBFile* dbfile);
bool LaunchedFromInstaller(HWND hParentWnd);
DWORD64 GetTotalByteCount();
CObList* RetrieveServerInfo();
BOOL PopulateFoldersAndFiles();
bool RunProgram(CString szCommandLine, bool bWait);
bool IsSFXFile(CString szFileName);

// CGenMAPPDBDLApp
// See GenMAPPDBDL.cpp for the implementation of this class
//

class CGenMAPPDBDLApp : public CWinApp
{
public:
	CGenMAPPDBDLApp();
	int m_nInetConnected;
	HWND m_hParentWnd;
	void SelectFolder(CString* pszSelectedFolder);
	CString szBasePath;
	CString szMAPPPath;
	CString szGDBPath;
	CString szGEXPath;
	CString szOtrPath;
	CString m_szDLLPath;
	CString szPreferredServer;
	bool bDeleteSFX;
	bool bOverwriteDataFiles;
	CGMFTP* pGenMAPPServer;

// Overrides
public:
	virtual BOOL InitInstance();

	// Descirbes how to connect to the Internet. Any proxy info is set here.
	DECLARE_MESSAGE_MAP()
public:
	// Searchs all specified drives for he specified directory. If bCDROM is true, all CDROMs are searched. If false, all hard drives are searched.
	char* FindDirectory(char* szDirName, bool bCDROM);
	// Reads entries from GenMAPP.cfg and populates the corresponding member variables
	void ReadCfgEntries(void);
private:
	// Companion to FindDirectory, this function is called recusively for each new directory layer.
	char* RecurseDir(CString szPath, char* szTargetPath);
public:
	CObList* pFileList;
	double nSpeed;
	int nSelectedFileCount;
private:
	// Sometimes the mru entries in GenMAPP.cfg contain file names in addition to the path. The function will remove the file name.
	bool TrimFileName(CString* szFileName);
public:
	double scaleX;
	double scaleY;
private:
	// Determines the scalaing factor for Windows installations that use font sizes other than Small Fonts
	void InitScaling(void);
public:
	bool CalledFromInstaller;
	void FreeFileList(void);
	bool bAbortDownload;
	// List of FTP servers for database retrieval. Populated from DataLocations.cfg file.
	CObList* pServerList;
	// List of all of the folders retrieved from FTP servers
	CObList DBFolderList;
	// Returns a pointer to a CDBFolder object in the folder list that matches the specified folder name or NULL if not found
	CDBFolder* FindFolderByName(CString szFolderName, int nType);
	// Handle to the main Data Acquisition Tool dialog
	HWND hMainDlg;
	// Handle to the GenBank folder, which is treated differently from all other folders
	HTREEITEM hGenBankFolder;
	HTREEITEM hOtherSpeciesFolder;
	// Set during the reading of data files from the FTP server. This is the name of the server as specified in DataLocations.cfg
	CString szRetrievingFrom;
	// Indicates that a file download is in progress
	bool bDLInProgress;
	bool m_bUseHTTP;
	// Fallback HTTP server if FTP is not available on machine.
	CGMFTP* pGenMAPPHTTPServer;
	// Handles creating a new HTTPServer object for accessing the control server.
	CGMFTP* CreateHTTPControlConnection(void);
	// The number of pixels in the screen resolution
	int nScrResX;
	int nScrResY;
	// If true, the user has clicked the Pause button in the Progress dialog. The Download routine acts accordingly.
	bool bPauseDownload;
	// A list of possible data file server seeds
	CObList ServerSeeds;
	// List of Updater servers. See ConnectToServer() in UpdateDlg.cpp
	CObList UpdateServers;
	POSITION HTTPPos;
};
