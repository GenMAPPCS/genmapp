// GenMAPPDBDL.cpp : Defines the initialization routines for the DLL.
//

#include "stdafx.h"
#include <afxinet.h>
#include <shlobj.h> 
#include "GenMAPPDBDL.h"
#include "resource.h"
#include "ConfigFile.h"
#include "CDDBCopyDlg.h"
#include "GMFTP.h"
#include "ProgressDialog.h"
#include "UpdateDlg.h"
#include "DBFile.h"
#include "OverwriteDlg.h"
#include "GenMAPPDBDL_Definitions.h"
#include ".\genmappdbdl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

//
//	Note!
//
//		If this DLL is dynamically linked against the MFC
//		DLLs, any functions exported from this DLL which
//		call into MFC must have the AFX_MANAGE_STATE macro
//		added at the very beginning of the function.
//
//		For example:
//
//		extern "C" BOOL PASCAL EXPORT ExportedFunction()
//		{
//			AFX_MANAGE_STATE(AfxGetStaticModuleState());
//			// normal function body here
//		}
//
//		It is very important that this macro appear in each
//		function, prior to any calls into MFC.  This means that
//		it must appear as the first statement within the 
//		function, even before any object variable declarations
//		as their constructors may generate calls into the MFC
//		DLL.
//
//		Please see MFC Technical Notes 33 and 58 for additional
//		details.
//

// The one and only CGenMAPPDBDLApp object

CGenMAPPDBDLApp theApp;


// CGenMAPPDBDLApp

BEGIN_MESSAGE_MAP(CGenMAPPDBDLApp, CWinApp)
END_MESSAGE_MAP()

// CGenMAPPDBDLApp construction

CGenMAPPDBDLApp::CGenMAPPDBDLApp()
: m_nInetConnected(NOT_CONNECTED)
, szBasePath(_T(""))
, szMAPPPath(_T(""))
, szGDBPath(_T(""))
, szGEXPath(_T(""))
, m_szDLLPath(_T(""))
, pFileList(NULL)
, nSpeed(0)
, nSelectedFileCount(0)
, scaleX(0)
, scaleY(0)
, CalledFromInstaller(false)
, bAbortDownload(false)
, szRetrievingFrom(_T("genmapp.org"))
, bDLInProgress(false)
, m_bUseHTTP(false)
, pGenMAPPHTTPServer(NULL)
, nScrResX(0)
, nScrResY(0)
, bPauseDownload(false)
{
	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}



// CGenMAPPDBDLApp initialization

BOOL CGenMAPPDBDLApp::InitInstance()
{
	CWinApp::InitInstance();
	CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
	InitScaling();

	return TRUE;
}

extern "C" BOOL PASCAL EXPORT InvokeFullDBDL(HWND hParentWnd, char* szDLLPath)
{
	// CD Database Copy entry point
	int		nReturnCode;
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	theApp.m_hParentWnd = hParentWnd;

	if (szDLLPath != NULL)
	{
		theApp.m_szDLLPath.Format("%s", szDLLPath);
//		AfxMessageBox("Here's what you passed in: " + theApp.m_szDLLPath);
	}

	theApp.CalledFromInstaller = LaunchedFromInstaller(hParentWnd);

	theApp.ReadCfgEntries();

	CreateDirectory(theApp.szBasePath, NULL);
	CreateDirectory(theApp.szGDBPath, NULL);
	CreateDirectory(theApp.szGEXPath, NULL);
	CreateDirectory(theApp.szMAPPPath, NULL);
	CreateDirectory(theApp.szOtrPath, NULL);

	INITCOMMONCONTROLSEX ccex;
	ccex.dwICC = ICC_LISTVIEW_CLASSES;
	ccex.dwSize = sizeof(INITCOMMONCONTROLSEX);
	if (InitCommonControlsEx(&ccex) == false)
		AfxMessageBox("InitCommon from InvokeFull failed");

	
	// Populate server seeds
	// Make sure seeds are added FTP then HTTP
	// SpeedTestFolder is also the location of DataLocations.cfg
	CGMFTP* pSeed = new CGMFTP("GMDBDL");
	pSeed->SetLoginData("root.genmapp.org", "downloader", "fun4downloader", "", "");
	pSeed->szSpeedTestFolder = "ServerSeed";
	pSeed->szHTTPSpeedTestFolder = "/gmd/";
	theApp.ServerSeeds.AddHead(pSeed);

	pSeed = new CGMFTP("GMDBDL");
	pSeed->SetLoginData("root2.genmapp.org", "downloader", "fun4downloader", "", "");
	pSeed->szSpeedTestFolder = "ServerSeed";
	pSeed->szHTTPSpeedTestFolder = "/gmd/";
	theApp.ServerSeeds.AddTail(pSeed);

CDCopyDlg:
	CCDDBCopyDlg CDDlg;
	nReturnCode = (int)CDDlg.DoModal();
	if (nReturnCode != IDCANCEL && theApp.nSelectedFileCount > 0)
	{
		CProgressDialog PrgDlg;
		nReturnCode = (int)PrgDlg.DoModal();
	}
	else
		nReturnCode = nReturnCode == IDCANCEL ? 0 : IDOK;

	theApp.FreeFileList();

	if (nReturnCode == IDCANCEL)
		goto CDCopyDlg;

	return nReturnCode;
}

extern "C" BOOL PASCAL EXPORT InvokeCDDBCopier()
{
	// CD Database Copy entry point
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	
	CCDDBCopyDlg CDDlg;
	CDDlg.DoModal();
	return true;
}

extern "C" BOOL PASCAL EXPORT InvokeUpdate(HWND hParentWnd, char* szCfgPath)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (szCfgPath != NULL)
	{
		theApp.m_szDLLPath.Format("%s", szCfgPath);
		//AfxMessageBox("Here's what you passed in: " + theApp.m_szDLLPath);
	}

	theApp.CalledFromInstaller = LaunchedFromInstaller(hParentWnd);

	theApp.ReadCfgEntries();

	CreateDirectory(theApp.szBasePath, NULL);
	CreateDirectory(theApp.szGDBPath, NULL);
	CreateDirectory(theApp.szGEXPath, NULL);
	CreateDirectory(theApp.szMAPPPath, NULL);
	CreateDirectory(theApp.szOtrPath, NULL);

	CUpdateDlg UDDlg;
	UDDlg.szInstallDir = szCfgPath;
	UDDlg.DoModal();
	return TRUE;
}

extern "C" BOOL PASCAL EXPORT CheckForGMUpdates(HWND hParentWnd, char* szCfgPath)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (szCfgPath != NULL)
	{
		theApp.m_szDLLPath.Format("%s", szCfgPath);
		//AfxMessageBox("Here's what you passed in: " + theApp.m_szDLLPath);
	}

	theApp.CalledFromInstaller = LaunchedFromInstaller(hParentWnd);

	theApp.ReadCfgEntries();

	if (CheckForUpdates(NULL) == 1)
	{
		CreateDirectory(theApp.szBasePath, NULL);
		CreateDirectory(theApp.szGDBPath, NULL);
		CreateDirectory(theApp.szGEXPath, NULL);
		CreateDirectory(theApp.szMAPPPath, NULL);
		CreateDirectory(theApp.szOtrPath, NULL);

		CUpdateDlg UDDlg;
		UDDlg.szInstallDir = szCfgPath;
		UDDlg.DoModal();
		
		return TRUE;
	}

	return FALSE;
}

extern "C" BOOL PASCAL EXPORT WriteStringKey(char* szDLLPath, char* szKey, char* szValue)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	if (szDLLPath != NULL)
		theApp.m_szDLLPath.Format("%s", szDLLPath);

	::CConfigFile cf(theApp.m_szDLLPath);
	cf.WriteStringKey(szKey, szValue);
	
	return true;
}

UINT ConnectToGenMAPP( LPVOID pParam )
{
	CCDDBCopyDlg* pDlg = (CCDDBCopyDlg*)pParam;
	CString szFileSize, szSpeedString = "";
	DWORD	dwFTPTimeout = 0;
	POSITION pos;
	bool bConnected = false;

	//char	szCurDir[MAX_PATH];
	//GetCurrentDirectory(MAX_PATH, szCurDir);
	//AfxMessageBox(szCurDir);

	// InternetConnectionTest();
	theApp.bAbortDownload = false;

	theApp.hMainDlg = pDlg->GetSafeHwnd();
	theApp.m_nInetConnected = CONNECTING;
	pDlg->DBFolderTree.EnableWindow(FALSE);

	for( pos = theApp.ServerSeeds.GetHeadPosition(); pos != NULL && bConnected == false; )
	{
		theApp.pGenMAPPServer = (CGMFTP*)theApp.ServerSeeds.GetNext( pos );
		if (!theApp.m_bUseHTTP)
		{
			theApp.pGenMAPPServer->SetOption(INTERNET_OPTION_CONNECT_TIMEOUT, 1000);
			theApp.pGenMAPPServer->SetOption(INTERNET_OPTION_CONNECT_RETRIES, 0);
			
			DWORD dwTimeoutVal = 0;
			theApp.pGenMAPPServer->QueryOption(INTERNET_OPTION_CONNECT_TIMEOUT, dwTimeoutVal);
			TRACE("Timeout value: %d", dwTimeoutVal);
			theApp.pGenMAPPServer->QueryOption(INTERNET_OPTION_CONNECT_RETRIES, dwTimeoutVal);
			TRACE(", retries value: %d\n", dwTimeoutVal);

			if (theApp.pGenMAPPServer->ConnectToGMServer())
			{
				//theApp.m_nInetConnected = CONNECT_FAILED;
				theApp.bAbortDownload = false;
				//pDlg->btnRefresh.EnableWindow();
				//return 1;
			}
		}

		if (!theApp.pGenMAPPServer->bConnected)
		{
			//delete theApp.pGenMAPPServer;
			theApp.HTTPPos = pos;
			//theApp.pGenMAPPHTTPServer = (CGMFTP*)theApp.ServerSeeds.GetNext( pos );
			theApp.pGenMAPPHTTPServer = theApp.CreateHTTPControlConnection();
			if (theApp.pGenMAPPHTTPServer->ConnectToGMServer())
			{
				theApp.m_nInetConnected = CONNECT_FAILED;
				theApp.bAbortDownload = false;
				//pDlg->btnRefresh.EnableWindow();
			}
			else
			{
				//delete theApp.pGenMAPPHTTPServer;
				if (theApp.pGenMAPPHTTPServer->TestHTTPConnection())
				{
					if (!theApp.m_bUseHTTP)
					{
						CConfigFile cf(theApp.m_szDLLPath);
						cf.WriteStringKey("ForceHTTP", "True", "[Downloader]");

						theApp.m_bUseHTTP = true;
					}
				}
			}
		}

		if (theApp.m_bUseHTTP)
			bConnected = theApp.pGenMAPPHTTPServer->bConnected;
		else
			bConnected = theApp.pGenMAPPServer->bConnected;
	}

	if (!bConnected)
	{
		theApp.m_nInetConnected = CONNECT_FAILED;
		pDlg->DBFolderTree.EnableWindow(TRUE);
		return 1;
	}

	TRACE("About to run speed test.\n");

	if (theApp.m_bUseHTTP)
	{
		theApp.pGenMAPPHTTPServer = theApp.CreateHTTPControlConnection();
		theApp.nSpeed = theApp.pGenMAPPHTTPServer->SpeedTest("");
	}
	else
		theApp.nSpeed = theApp.pGenMAPPServer->SpeedTest("");

	szSpeedString.Format("Connection speed to GenMAPP server: %.2fKB/Sec.", theApp.nSpeed);
	pDlg->szConnSpeed = szSpeedString;

	theApp.m_nInetConnected = READ_DIR;	

	//  Read Server Info from DataLocations.cfg to populate
	//  FTP server objects
	if (!RetrieveServerInfo())
	{
		theApp.m_nInetConnected = CONNECT_FAILED;
		if (!theApp.m_bUseHTTP)
			theApp.pGenMAPPServer->DisconnectFromGMServer();

		theApp.bAbortDownload = false;
		pDlg->DBFolderTree.EnableWindow();

		return 1;
	}

	// Retrieve directory info from each server
	// and build folder tree and list of DBFiles
	// from DataLocations.cfg.
	if (!PopulateFoldersAndFiles())	// If false, user aborted
	{
		theApp.m_nInetConnected = CONNECT_FAILED;
		if (!theApp.m_bUseHTTP)
			theApp.pGenMAPPServer->DisconnectFromGMServer();

		theApp.bAbortDownload = false;
		pDlg->DBFolderTree.EnableWindow();

		return 1;
	}

	if (!theApp.m_bUseHTTP)
		theApp.pGenMAPPServer->DisconnectFromGMServer();

	theApp.m_nInetConnected = CONNECTED;

	UINT	nTimerID = IDC_DBDLLIST;
	pDlg->SetTimer((UINT_PTR)nTimerID, 100, NULL);

	pDlg->btnRefresh.EnableWindow();
	pDlg->DBFolderTree.EnableWindow();



    return 0;   // thread completed successfully
}

void CGenMAPPDBDLApp::SelectFolder(CString* pszSelectedFolder) 
//this function will browse for a folder 
//PARAMETERS: 
//message:a string displayed above tree control in browse dialog (IN parameter) 
//lpitidlist:pointer to a item identifiers list ( OUT parameter, NULL if dialog was canceled) 
{ 
	LPITEMIDLIST lpitidlist=NULL, lpInitialPathList=NULL;//pointer to items list 
//	CString ReturnPath;//will receive the folder complete path 
	LPMALLOC shallocator;//will point to IMalloc interface of shell memory allocator 
	BROWSEINFO bwinfo;//the structure which contains necessary informations 
	char buffer[MAX_PATH], folder_path[MAX_PATH];//will receive the folder name 
	FillMemory((void*)&bwinfo,sizeof(bwinfo),0);//memory cleaning 
	bwinfo.hwndOwner=NULL;//handle of owner window (NULL here) 
	bwinfo.pidlRoot=NULL;//starting point for browse (NULL means Computer Desktop) 
	bwinfo.pszDisplayName=(char*) buffer;//will receive folder name (as displayed) 
	bwinfo.lpszTitle="Please Select the Data Folder. Press F5 to Refresh the View.";//dialog title 
	bwinfo.ulFlags=BIF_USENEWUI | BIF_EDITBOX | BIF_SHAREABLE;//browsing flags (0 means no restrictions) 
	lpitidlist=SHBrowseForFolder(&bwinfo);// now browsing for folder...and keep the results 
	SHGetPathFromIDList(lpitidlist, folder_path);//get the folder path from SHBrowseForFolder result 
	SHGetMalloc(&shallocator);//get shell memory allocator (a pointer to his IMalloc interface) 
	shallocator->Free((void*) lpitidlist);//releasing memory via this allocator 
	shallocator->Release();//releasing allocator 
	pszSelectedFolder->Format("%s\\", folder_path);
}


// Reads keys from GenMAPP.cfg and populates
// the corresponding member variables. If the Keys do not exist in the
// config file (or the config file does not exist), default paths are
// created and stored in the config file (again, the latter assumes the
// config file exists).
void CGenMAPPDBDLApp::ReadCfgEntries(void)
{
	CConfigFile cf(m_szDLLPath);
	CString		szResult, szImportPath, szTempValue;

	szBasePath = cf.ReadStringKey("baseFolder");
	if (szBasePath == "")
	{
		szBasePath = "C:\\GenMAPP 2 Data\\";
		cf.WriteStringKey("baseFolder", szBasePath);
	}

	szGDBPath = cf.ReadStringKey("mruGeneDB");
	if (szGDBPath == "")
	{
		szGDBPath = szBasePath + "Gene Databases\\";
		cf.WriteStringKey("mruGeneDB", szGDBPath);
	}
	else
		TrimFileName(&szGDBPath);
		

	szMAPPPath = cf.ReadStringKey("mruMAPPPath");
	if (szMAPPPath == "")
	{
		szMAPPPath = szBasePath + "MAPPs\\";
		cf.WriteStringKey("mruMAPPPath", szMAPPPath);
	}
	else
		TrimFileName(&szMAPPPath);

	szOtrPath = cf.ReadStringKey("mruOtherInfo");
	if (szOtrPath == "")
	{
		szOtrPath = szBasePath + "Other Information\\";
		cf.WriteStringKey("mruOtherInfo", szOtrPath);
	}
	else
		TrimFileName(&szOtrPath);

	szGEXPath = cf.ReadStringKey("mruDataSet");
	if (szGEXPath == "")
	{
		szGEXPath = szBasePath + "Expression Datasets\\";
		cf.WriteStringKey("mruDataSet", szGEXPath);
	}
	else
		TrimFileName(&szGEXPath);

	szImportPath = cf.ReadStringKey("mruImportPath");
	if (szImportPath == "")
	{
		szImportPath = szBasePath + "Expression Datasets\\";
		cf.WriteStringKey("mruImportPath", szImportPath);
	}

	szTempValue = cf.ReadStringKey("ForceHTTP", "[Downloader]");
	if (szTempValue == "" || szTempValue.MakeUpper() == "FALSE")
        m_bUseHTTP = false;
	else
	{
		szTempValue.MakeUpper();
		m_bUseHTTP = (szTempValue == "TRUE");
	}

	szTempValue = cf.ReadStringKey("DeleteCompressedFiles", "[Downloader]");
	if (szTempValue == "")
        bDeleteSFX = true;
	else
	{
		szTempValue.MakeUpper();
		bDeleteSFX = (szTempValue == "TRUE");
	}

	szTempValue = cf.ReadStringKey("OverwriteDataFiles", "[Downloader]");
	if (szTempValue == "")
        bOverwriteDataFiles = true;
	else
	{
		szTempValue.MakeUpper();
		bOverwriteDataFiles = (szTempValue == "TRUE");
	}

	szPreferredServer = cf.ReadStringKey("PreferredServer", "[Downloader]");
}

// Searchs all specified drives for the specified directory.
// If bCDROM is true, all CDROMs are searched. If false, all
// hard drives are searched.
char* CGenMAPPDBDLApp::FindDirectory(char* szDirName, bool bCDROM)
{
	char		*szGMPath = NULL, szDrv[] = {"C:\\"};
	UINT		nDrvType = 0;

	for (szDrv[0]='C';szDrv[0]<='Z';szDrv[0]++)
	{
		nDrvType = GetDriveType(szDrv);
		if (nDrvType == (bCDROM ? DRIVE_CDROM : DRIVE_FIXED))
			szGMPath = RecurseDir(szDrv, szDirName);
			if (szGMPath != NULL)
				return szGMPath;
	}
	return NULL;
}

// Companion to FindDirectory, this function is called
// recusively for each new directory layer.
char* CGenMAPPDBDLApp::RecurseDir(CString szPath, char* szTargetPath)
{
	char			*szSearchBuf, **szLastOccur = NULL;
	WIN32_FIND_DATA	FindFileData;
	BOOL			bMore = true;

	szSearchBuf = new char[MAX_PATH * 4];
	if (SearchPath(szPath, szTargetPath, NULL, MAX_PATH * 4,
		szSearchBuf, szLastOccur))
		return szSearchBuf;
	else
		delete szSearchBuf;

	HANDLE hFind = FindFirstFile(szPath + "*.*", &FindFileData);
	while (bMore && hFind != INVALID_HANDLE_VALUE)
	{
		if (FindFileData.dwFileAttributes == FILE_ATTRIBUTE_DIRECTORY && 
			strcmp(FindFileData.cFileName, ".") &&
			strcmp(FindFileData.cFileName, ".."))
		{
			CString szNewPath;
			char	*szFoundPath;
			szNewPath.Format("%s%s\\", szPath, FindFileData.cFileName);
			szFoundPath = RecurseDir(szNewPath, szTargetPath);
			if (szFoundPath != NULL)
				return szFoundPath;
		}

		bMore = FindNextFile(hFind, &FindFileData);
	}
	
	if (hFind != INVALID_HANDLE_VALUE)
		FindClose(hFind);
	return NULL;
}

UINT DownloadFiles( LPVOID pParam )
{

	// For each server
	//			If selected
	//				Is server in list?
	//					yes - Is there a faster server in list?
	//						No - download
	//						yes - skip file
	
	CProgressDialog* pDlg = (CProgressDialog*)pParam;
	POSITION pos, pos2, pos3;
	CDBFile* dbfile;
	int	nCurrentFileNum = 0;
	DWORD64 nNumBytesDownloaded = 0, nCombinedFileSize = GetTotalByteCount();
	double			dNumBytesDownloaded, dCombinedFileSize = (double)nCombinedFileSize;
	UINT_PTR nTimerID =	111;

	pDlg->dwTotalBytesToDownload = (DWORD)GetTotalByteCount();

	theApp.bDLInProgress = true;
	pDlg->OverallProgress.SetRange32(0, 10000);
	pDlg->SetTimer(nTimerID, 500, NULL);

	try
	{
		// Speed test for all servers
		pDlg->lstQueuedFiles.ShowWindow(SW_HIDE);
		pDlg->m_lstSpeedTestProg.ShowWindow(SW_SHOW);
		
		pDlg->m_lstSpeedTestProg.AddString("FINDING FASTEST SERVER");
		pDlg->m_lstSpeedTestProg.AddString(" ");

		for( pos = theApp.pServerList->GetHeadPosition(); pos != NULL; )
		{
			//CGMFTP* pServer = (CGMFTP*)theApp.pFileList->GetNext( pos );
			CGMFTP* pServer = (CGMFTP*)theApp.pServerList->GetNext( pos );
			if (pServer->bHTTPServer == theApp.m_bUseHTTP)
			{
				if (!pServer->bHTTPServer)
				{
					if (pServer->ConnectToGMServer())
						continue;
				}

				pDlg->m_lstSpeedTestProg.AddString("Testing: " + pServer->szAlias + " in " + pServer->szLocation);
				pServer->SpeedTest("");

				pServer->DisconnectFromGMServer();
			}
		}

		pDlg->lstQueuedFiles.ShowWindow(SW_SHOW);
		pDlg->m_lstSpeedTestProg.ShowWindow(SW_HIDE);

		for( pos = theApp.pServerList->GetHeadPosition(); pos != NULL; )
		{
			BOOL bAlreadyConnected = FALSE;
			// Go through entire file list
//			CGMFTP* pServer = (CGMFTP*)theApp.pFileList->GetNext( pos );
			CGMFTP* pServer = (CGMFTP*)theApp.pServerList->GetNext( pos );
			for( pos2 = theApp.pFileList->GetHeadPosition(); pos2 != NULL; )
			{
				dbfile = (CDBFile*)theApp.pFileList->GetNext( pos2 );
				if (dbfile->bSelected)
				{
					double	fFastest = 0;
					CGMFTP* pFastestServer = NULL;
					int nHighestTier = 32767;
					for( pos3 = dbfile->DBServerList.GetHeadPosition(); pos3 != NULL; )
					{
						CGMFTP* pFileServer = (CGMFTP*)dbfile->DBServerList.GetNext( pos3 );
						if (pFileServer->nTier < nHighestTier)
							nHighestTier = pFileServer->nTier;
					}

					for( pos3 = dbfile->DBServerList.GetHeadPosition(); pos3 != NULL; )
					{
						CGMFTP* pFileServer = (CGMFTP*)dbfile->DBServerList.GetNext( pos3 );
						if (pFileServer->fConnSpeed > fFastest && pFileServer->nTier == nHighestTier)
						{
							fFastest = pFileServer->fConnSpeed;
							pFastestServer = pFileServer;
						}
					}

					if (pFastestServer == pServer)
					{
						if (!pServer->bHTTPServer && !bAlreadyConnected)
						{
							if (pServer->ConnectToGMServer())
								continue;
							else
								bAlreadyConnected = TRUE;
						}

						if (theApp.bAbortDownload)
						{
							if (!pServer->bHTTPServer)
								pServer->DisconnectFromGMServer();
							theApp.bDLInProgress = false;
							theApp.bAbortDownload = false;
							AfxEndThread(0, true);
							return 0;
						}

						if (!pServer->bHTTPServer)
							pServer->m_pConnect->SetCurrentDirectory(pServer->szSpeedTestFolder + dbfile->szSrcFileName.Left(dbfile->szSrcFileName.ReverseFind('/')));

						CString szStatic, szDirName = GetDestPathFromDBFile(dbfile);
						CInternetFile*	pSrcFTPFile = NULL;
						int				nBytesRead = 0, nCheckpoint = 1024;
						BYTE*			pReadBuf = new BYTE[99328]; // 97KB
						DWORD BeginTime, EndTime, dwBytesReadThisFile = 1024;

						szStatic.Format("%i of %i", ++nCurrentFileNum, theApp.nSelectedFileCount);
						pDlg->labFileTally.SetWindowText(szStatic);

						szStatic.Format("File Currently Downloading: %s.", dbfile->GetFileNameOnly(dbfile->szSrcFileName));
						pDlg->labFileInDL.SetWindowText(szStatic);

						// Set "Current File" information
						pDlg->lblFileName.SetWindowText(dbfile->GetFileNameOnly(dbfile->szSrcFileName));
						pDlg->lblServerName.SetWindowText(pServer->szAlias);
						pDlg->lblLocation.SetWindowText(pServer->szLocation);
						szStatic.Format("Downloaded 0 of %u bytes", dbfile->dwFileSize);
						pDlg->lblFileBytesDownloaded.SetWindowText(szStatic);
						pDlg->lblTransport.SetWindowText(pServer->bHTTPServer ? "HTTP" : "FTP");
						pDlg->lblFileType.SetWindowText(dbfile->GetFileTypeAsString(dbfile->nFileType));

						// Set icon of the item in the file queue list to a right arrow
						pDlg->UpdateFileQueueItem(1, 0);

						if (!CreateDirectory(szDirName, NULL))
							TRACE("Create Directory failed: %s\n", szDirName);
						
						if (!pServer->bHTTPServer)
							pSrcFTPFile = pServer->m_pConnect->OpenFile(dbfile->szSrcFileName.Right(dbfile->szSrcFileName.GetLength() - dbfile->szSrcFileName.ReverseFind('/') - 1));
						else
							pSrcFTPFile = pServer->OpenHTTPFile(pServer->szHTTPSpeedTestFolder + dbfile->szSrcFileName);


						if (pSrcFTPFile == NULL)
						{
							pDlg->UpdateFileQueueItem(2, 0);
							delete pReadBuf;
							continue;
						}

						szStatic.Format("%s%s", szDirName, dbfile->szSrcFileName.Right(dbfile->szSrcFileName.GetLength() - dbfile->szSrcFileName.ReverseFind('/') - 1));
						CFile DestFile(szStatic, CFile::modeCreate | CFile::modeWrite | CFile::shareExclusive);
						if (DestFile.m_hFile == CFile::hFileNull)
						{
							pDlg->UpdateFileQueueItem(2, 0);
							pSrcFTPFile->Close();
							delete pReadBuf;
							continue;
						}
						
						BeginTime = GetTickCount();
						nBytesRead = pSrcFTPFile->Read(pReadBuf, 1024);
						while (nBytesRead > 0)
						{
							CString szAllPurpose;

							if (theApp.bAbortDownload)
							{
								DestFile.Close();
								DeleteFile(szStatic);
								pSrcFTPFile->Close();
								if (!pServer->bHTTPServer)
									pServer->DisconnectFromGMServer();
								else
									delete pSrcFTPFile;
								theApp.bDLInProgress = false;
								theApp.bAbortDownload = false;
								pDlg->UpdateFileQueueItem(2, (int)(((double)dwBytesReadThisFile / (double)dbfile->dwFileSize) * 100));
								delete pReadBuf;
								AfxEndThread(0, true);
								return 0;
							}



							DestFile.Write(pReadBuf, nBytesRead);
							nNumBytesDownloaded += nBytesRead;
							dNumBytesDownloaded = (double)nNumBytesDownloaded;
							pDlg->OverallProgress.SetPos((int)((dNumBytesDownloaded / dCombinedFileSize) * 10000));

							// Update the "Downloaded x of x bytes" labels
							szAllPurpose.Format("%u of %u bytes", (DWORD)nNumBytesDownloaded, (DWORD)nCombinedFileSize);
							pDlg->lblByteTotals.SetWindowText(szAllPurpose);

							szAllPurpose.Format("%u of %u bytes", dwBytesReadThisFile, dbfile->dwFileSize);
							pDlg->lblFileBytesDownloaded.SetWindowText(szAllPurpose);


							// Set the percentage complete in the file queue
							pDlg->UpdateFileQueueItem(1, (int)(((double)dwBytesReadThisFile / (double)dbfile->dwFileSize) * 100));

							if (theApp.bPauseDownload)
							{
								bool	bIconOn = true;

								szStatic.Format("File Currently Downloading: %s (Paused).", dbfile->GetFileNameOnly(dbfile->szSrcFileName));
								pDlg->labFileInDL.SetWindowText(szStatic);
								pDlg->lblDownloadSpeed.SetWindowText("PAUSED");
								pDlg->UpdateFileQueueItem(4, (int)(((double)dwBytesReadThisFile / (double)dbfile->dwFileSize) * 100));

								while (theApp.bPauseDownload)
								{
									Sleep(500);
									bIconOn = !bIconOn;
									pDlg->UpdateFileQueueItem(bIconOn ? 4 : 5, (int)(((double)dwBytesReadThisFile / (double)dbfile->dwFileSize) * 100));
								}


								nBytesRead = pSrcFTPFile->Read(pReadBuf, 99328);
								if (nBytesRead < 99328)
								{
									TRACE("Pause Timeout\n");
									pDlg->labFileInDL.SetWindowText("Lost connection to server. Restarting download of this file.");
									bAlreadyConnected = FALSE;
									Sleep(1500);
									nCurrentFileNum--;
									pDlg->m_nCurrentQueueIndex--;
									
									// It occurs to me that I have too many variable keeping track of the bytes downloaded. That's spaghetti code for you!
									nNumBytesDownloaded -= dwBytesReadThisFile;
									dNumBytesDownloaded = (double)nNumBytesDownloaded;
									pDlg->dwBytesDownloaded -= dwBytesReadThisFile;
									pDlg->nHistoricalSpeedRecordCount = 0;

									theApp.pFileList->GetPrev( pos2 );
									if (pos2 == NULL)
										pos2 = theApp.pFileList->GetHeadPosition();

									break;
								}

								szStatic.Format("File Currently Downloading: %s.", dbfile->GetFileNameOnly(dbfile->szSrcFileName));
								pDlg->labFileInDL.SetWindowText(szStatic);

							}
							else
								nBytesRead = pSrcFTPFile->Read(pReadBuf, 1024);

							pDlg->dwBytesDownloaded+= nBytesRead;
							nCheckpoint+= nBytesRead;
							dwBytesReadThisFile+= nBytesRead;
							if (nCheckpoint >= (1024 * SPEED_SAMPLE_SIZE) - 1)
							{
								CString szSpeedString;
								EndTime = GetTickCount();
								pDlg->AddHistoricalSpeedRecord((nCheckpoint / 1024) / (((double)EndTime - (double)BeginTime) / 1000));
								pDlg->dCurrentSpeed = pDlg->GetAverageDLSpeed();
								szSpeedString.Format("%.2fKB/Sec.", pDlg->dCurrentSpeed);
								pDlg->lblDownloadSpeed.SetWindowText(szSpeedString);
								
								nCheckpoint = 0;
								BeginTime = GetTickCount();
							}
						}

						DestFile.Flush();
						DestFile.Close();
						pSrcFTPFile->Close();
						if (pServer->bHTTPServer)
							delete pSrcFTPFile;


						// Account for "fuzzy" file sizes gleemed from HTTP by
						// setting file D/L percentage and byte counts know that
						// they're known.
						pDlg->UpdateFileQueueItem(0, 100);
						szStatic.Format("%u of %u bytes", dwBytesReadThisFile, dwBytesReadThisFile);
						pDlg->lblFileBytesDownloaded.SetWindowText(szStatic);

						delete pReadBuf;
					}
				}
			}
			
			if (!pServer->bHTTPServer)
				pServer->DisconnectFromGMServer();
		}
	}

	catch (CInternetException* pEx)
	{
		TCHAR sz[1024];
		pEx->GetErrorMessage(sz, 1024);
		printf("ERROR!  %s\n", sz);
		pEx->Delete();
		pDlg->labFileTally.SetWindowText("Error downloading file. Please try again at a later time.");
		pDlg->labFileInDL.SetWindowText(" ");
		pDlg->btnOK.EnableWindow(true);
		theApp.bDLInProgress = false;
		pDlg->KillTimer(nTimerID);
		return 1;
	}

	CString	szAllPurpose;

	// Make sure progress bar is set to max. HTTP files sizes
	// are not always precise down to the byte.
	pDlg->OverallProgress.SetPos(10000);

	// Update the "Downloaded x of x bytes" label for the same reason,
	// using total bytes donloaded, as opposed to estimated combined file size.
	szAllPurpose.Format("%u of %u bytes", (DWORD)nNumBytesDownloaded, (DWORD)nNumBytesDownloaded);
	pDlg->lblByteTotals.SetWindowText(szAllPurpose);

	pDlg->labFileTally.SetWindowText("Uncompressing files.");
	pDlg->labFileInDL.SetWindowText(" ");
	pDlg->KillTimer(nTimerID);
	pDlg->lblRemainingTime.SetWindowText("00:00:00");
	ExtractFiles(pParam);


	pDlg->labFileTally.SetWindowText("Download Completed Successfully.");
	pDlg->btnOK.EnableWindow(true);
	theApp.bDLInProgress = false;

	return 0;
}

UINT ExtractFiles(LPVOID pParam)
{
	CProgressDialog* pDlg = (CProgressDialog*)pParam;
	POSITION pos;
	bool bYesToAll = theApp.bOverwriteDataFiles, bNoToAll = false;
	char szTempDir[_MAX_PATH];

	GetTempPath(_MAX_PATH, szTempDir);

	for( pos = theApp.pFileList->GetHeadPosition(); pos != NULL; )
	{
		CDBFile* dbfile = (CDBFile*)theApp.pFileList->GetNext( pos );

		if (dbfile->bSelected && (dbfile->szSrcFileName.Right(4).MakeUpper() == ".EXE" || 
			dbfile->szSrcFileName.Right(4).MakeUpper() == ".BAT"))
		{
			// Update file queue list icon with a right arrow
			pDlg->UpdateFileQueueItem(1, 0);

			if (IsSFXFile(GetDestPathFromDBFile(dbfile) + "\\" + dbfile->GetFileNameOnly(dbfile->szSrcFileName, true)))
			{
				bool bYesThisPass = false;
				if (!bYesToAll && !bNoToAll)
				{
					char szFullCmdLine[MAX_PATH];
					DWORD	dwBytesWritten;
					CString szTempFilePath, szLine;
					CStringList	FilesInSFX;
					bool	bFileExists = false;

					// Create a batch file that will dump the console
					// output of the PKSFX test command to a text file.
					// The text file will contain the names of all of the files
					// in the SFX.
					szTempFilePath.Format("%s~GetSFXContents.bat", szTempDir);
					HANDLE	hStdOutFile = CreateFile(szTempFilePath, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS,
						FILE_ATTRIBUTE_NORMAL, NULL);

					if (hStdOutFile == INVALID_HANDLE_VALUE)
					{
						pDlg->UpdateFileQueueItem(2, 0);
						continue;
					}

					wsprintf(szFullCmdLine, "\"%s%s\" -test=all > ZipContents.txt", GetDestPathFromDBFile(dbfile), dbfile->GetFileNameOnly(dbfile->szSrcFileName, true));
					WriteFile(hStdOutFile, szFullCmdLine, (DWORD)strlen(szFullCmdLine), &dwBytesWritten, NULL);
					CloseHandle(hStdOutFile);

					RunProgram(szTempFilePath, true);
					DeleteFile(szTempFilePath);


					// Now parse the text file to determine the names
					// of the files in the SFX.
					szTempFilePath.Format("%sZipContents.txt", szTempDir);
					CStdioFile ZipContents;
					CFileException fe;
					if (!ZipContents.Open(szTempFilePath, CFile::modeRead | CFile::shareDenyWrite, &fe))
					{
						// Couldn't open ZipContents, so just run the SFX without overwriting.
						char	szError[1024];
						fe.GetErrorMessage(szError, 1024);
						TRACE("File Exception reading ZipContents: %s", szError);
						DeleteFile(szTempFilePath);
						RunProgram("\"" + GetDestPathFromDBFile(dbfile) + dbfile->GetFileNameOnly(dbfile->szSrcFileName, true) + "\" -d -overwrite=never", true);
					}
	

					// The line in the text file we are interested in looks like this:
					// Testing: Dm-Std_20040411.gdb   OK
					// Keep reading the file until we find the first occurence.
					while (ZipContents.ReadString(szLine) && szLine.Left(8) != "Testing:");

					while (szLine.Left(8) == "Testing:")
					{
						// Pull out file name from szLine: it ends 3 spaces to the left of "OK"
						// and begins 1 space after Testing:
						FilesInSFX.AddTail(szLine.Mid(9, szLine.ReverseFind('O') - 12));
						
						if (!bFileExists)
						{
							HANDLE hExists = CreateFile(GetDestPathFromDBFile(dbfile) + szLine.Mid(9, szLine.ReverseFind('O') - 12),
								0, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
			                
							if (hExists != INVALID_HANDLE_VALUE)
								bFileExists = true;
						}

						ZipContents.ReadString(szLine);
					}

					ZipContents.Close();
					DeleteFile(szTempFilePath);

					if (bFileExists)
					{
						UINT nOverwriteReturn = 255;

						COverwriteDlg ORDlg;

						ORDlg.pFileNameList = &FilesInSFX; 
						nOverwriteReturn = (UINT)ORDlg.DoModal();

						switch (nOverwriteReturn)
						{
							case IDOK:
								bYesThisPass = true;
								break;

							case 10: // Yes to all
								bYesToAll = true;
								break;

							case 100:  // No to all
								bNoToAll = true;
						}
					}
				}

                if (bYesToAll || bYesThisPass)
					RunProgram("\"" + GetDestPathFromDBFile(dbfile) + dbfile->GetFileNameOnly(dbfile->szSrcFileName, true) + "\" -d -overwrite=all", true);
				else
					RunProgram("\"" + GetDestPathFromDBFile(dbfile) + dbfile->GetFileNameOnly(dbfile->szSrcFileName, true) + "\" -d -overwrite=never", true);
			}
			else
				RunProgram("\"" + GetDestPathFromDBFile(dbfile) + dbfile->GetFileNameOnly(dbfile->szSrcFileName, true) + "\"", true);

			if (theApp.bDeleteSFX)
				DeleteFile(GetDestPathFromDBFile(dbfile) + dbfile->GetFileNameOnly(dbfile->szSrcFileName, true));

			// Set the icon of the item in the file queue to a check mark
			pDlg->UpdateFileQueueItem(0, 100);

		}
	}

	return 1;
}

CString GetDestPathFromDBFile(CDBFile* dbfile)
{
	switch (dbfile->nFileType)
	{
		case GENE_DATABASE:
			return theApp.szGDBPath;

		case MAPP_DATABASE:
		case MAPP_ARCHIVE:
			return theApp.szMAPPPath;

		case GENEEXPRESS_DATABASE:
			return theApp.szGEXPath;

		default:
			return theApp.szBasePath;
	}
}

bool IsSFXFile(CString szFileName)
{
	BYTE bytBuffer[536];
	UINT nBytesRead = 0, nBufIndex = 0;
	CFileException fe;

	CFile ExecutableFile;
	if (!ExecutableFile.Open(szFileName, CFile::modeRead | CFile::shareDenyWrite, &fe))
	{
		fe.GetErrorMessage((char*)bytBuffer, 535);
		TRACE("File Exception: %s", bytBuffer);
		return false;
	}

	if (ExecutableFile.GetLength() < 535)
	{
		ExecutableFile.Close();
		return false;
	}

	while (nBufIndex < 534)
	{
		nBytesRead = ExecutableFile.Read(&bytBuffer, 535 - nBufIndex);
		nBufIndex+= nBytesRead;
	}

	ExecutableFile.Close();

	return (memcmp(&bytBuffer[526], "PKSFX CLI", 9) == 0);
}

// Generic function to run a program and optionally wait for it
// to complete. Working directory is the same directory as that
// of the executable. Returns false if something goes wrong.
bool RunProgram(CString szCommandLine, bool bWait)
{
	STARTUPINFO si;
	PROCESS_INFORMATION pi;
	DWORD dwExitCode = STILL_ACTIVE;
	char  szCPCmdLine[_MAX_PATH], szWorkingDir[_MAX_PATH];

    strcpy(szCPCmdLine, szCommandLine);

	// If there is a quote surrounding the path, remove it.
	if (szCommandLine.Left(1) == "\"")
		strcpy(szWorkingDir, szCommandLine.Mid(1, szCommandLine.ReverseFind('\\')));
	else
		strcpy(szWorkingDir, szCommandLine.Left(szCommandLine.ReverseFind('\\')));


	ZeroMemory( &si, sizeof(si) );
	si.cb = sizeof(si);
	si.wShowWindow = SW_HIDE;
	si.dwFlags = STARTF_USESHOWWINDOW;
	ZeroMemory( &pi, sizeof(pi) );
	if (!CreateProcess(NULL, szCPCmdLine, NULL, NULL, true, NORMAL_PRIORITY_CLASS,
		NULL, szWorkingDir, &si, &pi ))
	{
		LPVOID lpMsgBuf;
		FormatMessage( 
			FORMAT_MESSAGE_ALLOCATE_BUFFER | 
			FORMAT_MESSAGE_FROM_SYSTEM | 
			FORMAT_MESSAGE_IGNORE_INSERTS,
			NULL,
			GetLastError(),
			MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
			(LPTSTR) &lpMsgBuf,
			0,
			NULL );

		TRACE("Error running %s in working direcory %s. Error: %s\n", szCPCmdLine, szWorkingDir, (LPCTSTR)lpMsgBuf);
		LocalFree( lpMsgBuf );
		return false;
	}

	if (bWait)
	{
		while (dwExitCode == STILL_ACTIVE)
		{
			Sleep(1000);
			if (!GetExitCodeProcess(pi.hProcess, &dwExitCode))
				break;
		}
	}

	CloseHandle( pi.hProcess );
	CloseHandle( pi.hThread );

	return true;
}

// Sometimes the mru entries in GenMAPP.cfg contain file names in addition
// to the path. The function will remove the file name.
bool CGenMAPPDBDLApp::TrimFileName(CString* szFileName)
{
	if (szFileName->Right(4) == ".gdb" || szFileName->Right(4) == ".gex" ||
		szFileName->Right(5) == ".mapp")
		*szFileName = szFileName->Left(szFileName->ReverseFind('\\') + 1);
	return true;
}

// Determines the scalaing factor for Windows installations that
// use font sizes other than Small Fonts
void CGenMAPPDBDLApp::InitScaling(void)
{
   HDC screen = GetDC(0);
   scaleX = GetDeviceCaps(screen, LOGPIXELSX) / 96.0;
   scaleY = GetDeviceCaps(screen, LOGPIXELSY) / 96.0;
   nScrResX = GetDeviceCaps(screen, HORZRES);
   nScrResY = GetDeviceCaps(screen, VERTRES);
  ReleaseDC(0, screen);

}

bool LaunchedFromInstaller(HWND hParentWnd)
{
	char szWndText[255];
	bool bLFI = false;
	::GetWindowText(hParentWnd, szWndText, 255);
//	if (szWndText[0] == '\0')
	{
		HWND hNextWnd = NULL;
		hNextWnd = ::FindWindowEx(HWND_DESKTOP, hNextWnd, NULL, NULL);
			//= ::GetDesktopWindow();

		while (hNextWnd != NULL)
		{
			GetWindowText(hNextWnd, szWndText, 255);
			if (strncmp(szWndText, "InstallShield", 12) == 0)
			{
				bLFI = true;
				break;
			}
			hNextWnd = GetNextWindow(hNextWnd, GW_HWNDNEXT);
		}
	}

	return bLFI;
}

void CGenMAPPDBDLApp::FreeFileList(void)
{
	// Give time to abort FTP connection.
	while (bAbortDownload)
	{	TRACE("Waiting for abort...\r\n");
		Sleep(100);
	}

	if (pFileList != NULL)
	{
		POSITION pos;
		CDBFile* dbfile;
		for( pos = pFileList->GetHeadPosition(); pos != NULL; )
		{
			dbfile = (CDBFile*)pFileList->GetNext( pos );
			delete dbfile;
		}

		pFileList->RemoveAll();

		if (pFileList != NULL)
			delete pFileList;

		pFileList = NULL;
	}

	if (!DBFolderList.IsEmpty())
	{
		POSITION pos;
		CDBFolder* dbfolder;
		for( pos = DBFolderList.GetHeadPosition(); pos != NULL; )
		{
			dbfolder = (CDBFolder*)DBFolderList.GetNext( pos );
			delete dbfolder;
		}

		DBFolderList.RemoveAll();
	}

	if (pServerList != NULL)
	{
		POSITION pos;
		CGMFTP* pTempServer;
		for( pos = pServerList->GetHeadPosition(); pos != NULL; )
		{
			pTempServer = (CGMFTP*)pServerList->GetNext( pos );
			delete pTempServer;
		}

		pServerList->RemoveAll();

		if (pServerList != NULL)
			delete pServerList;

		pServerList = NULL;
	}
}

DWORD64 GetTotalByteCount()
{
	POSITION pos;
	CDBFile* dbfile;
	DWORD64 nTotalFileSize = 0;

	for( pos = theApp.pFileList->GetHeadPosition(); pos != NULL; )
	{
		dbfile = (CDBFile*)theApp.pFileList->GetNext( pos );
		if (dbfile->bSelected)
			nTotalFileSize += dbfile->dwFileSize;
	}

	// Top is 10000, (DL'ed / Total) * 10000

	return nTotalFileSize;
}

CObList* RetrieveServerInfo()
{
	char*		szTempDir;
	CString		szPathAndFile;

	szTempDir = getenv("TEMP");
	if (szTempDir == NULL)
		szTempDir = getenv("TMP");

	szPathAndFile.Format("%s\\DataLocations.cfg", szTempDir);

	try
	{
#ifdef DEBUG		
		if (!theApp.m_bUseHTTP)
		{
			CString szCurrDir = "";
			theApp.pGenMAPPServer->m_pConnect->GetCurrentDirectory(szCurrDir);
			TRACE("Current Directory is %s reading DataLocations.cfg\n", szCurrDir);
		}
#endif
		if (theApp.m_bUseHTTP)
		{
			//delete theApp.pGenMAPPHTTPServer;
			theApp.pGenMAPPHTTPServer = theApp.CreateHTTPControlConnection();
			if (!theApp.pGenMAPPHTTPServer->GetHttpFile(theApp.pGenMAPPServer->szHTTPSpeedTestFolder + "/DataLocations.cfg", szPathAndFile))
				return NULL;

			//delete theApp.pGenMAPPHTTPServer;
		}
		else
		{
			if (!theApp.pGenMAPPServer->m_pConnect->GetFile("DataLocations.cfg", szPathAndFile, false, FILE_ATTRIBUTE_NORMAL, FTP_TRANSFER_TYPE_ASCII))
				return NULL;
		}
	}
	catch (CInternetException* pEx)
	{
		TCHAR sz[1024];
		pEx->GetErrorMessage(sz, 1024);
		printf("ERROR!  %s\n", sz);
		pEx->Delete();
		return NULL;
	}

	CConfigFile cfgReader(szTempDir, "\\DataLocations.cfg");
	theApp.pServerList = cfgReader.RetreiveServerEntries();
	return theApp.pServerList;
}

BOOL PopulateFoldersAndFiles()
{
	POSITION	pos, pos2, pos3;
	CGMFTP*		pServer;
	char*		szTempDir;
	CString		szPathAndFile;

	if (theApp.pServerList == NULL)
		return FALSE;

	szTempDir = getenv("TEMP");
	if (szTempDir == NULL)
		szTempDir = getenv("TMP");

	szPathAndFile.Format("%s\\DataLocations.cfg", szTempDir);

	theApp.pFileList = new CObList();

	for( pos = theApp.pServerList->GetHeadPosition(); pos != NULL; )
	{
		CStringList*		pServerFolders, *pServerFiles = new CStringList(),
			*pServerFilesFolders = new CStringList();

		pServer = (CGMFTP*)theApp.pServerList->GetNext( pos );

		if (pServer->bFTPBackup && !theApp.m_bUseHTTP)
			continue;
		
		theApp.szRetrievingFrom = pServer->szAlias;
		
		CConfigFile cfgReader(szTempDir, "\\DataLocations.cfg");
		pServerFolders = cfgReader.RetreiveFoldersForServer(pServer->szServerKeyName, pServer->bHTTPServer);
		cfgReader.RetrieveFilesForServer(pServer->szServerKeyName, pServerFiles, pServerFilesFolders, pServer->bHTTPServer);

		if (pServerFolders->IsEmpty())
			continue;

		// Failure to connect is not an error. We'll just have
		// fewer files available.
		if (pServer->ConnectToGMServer())
//			return FALSE;
			continue;

		for( pos2 = pServerFolders->GetHeadPosition(); pos2 != NULL; )
		{
			CString		szFolder;
			szFolder = pServerFolders->GetNext( pos2 );
			if (!pServer->RetrieveFromRootFolder(szFolder))
			{
				pServer->DisconnectFromGMServer();

				delete pServerFiles;
				delete pServerFilesFolders;
				::DeleteFile(szPathAndFile);
				return FALSE;
			}
		}

		for( pos2 = pServerFiles->GetHeadPosition(), pos3 = pServerFilesFolders->GetHeadPosition(); pos2 != NULL; )
		{
			CString		szFolder, szFile;
			szFile = pServerFiles->GetNext( pos2 );
			szFolder = pServerFilesFolders->GetNext( pos3 );
			if (!pServer->RetrieveIndividualFile(szFile, szFolder))
			{
				pServer->DisconnectFromGMServer();

				delete pServerFiles;
				delete pServerFilesFolders;
				::DeleteFile(szPathAndFile);
				return FALSE;
			}
		}


		pServer->DisconnectFromGMServer();

		delete pServerFiles;
		delete pServerFilesFolders;

	}
	
		CDBFolder*	dbfolder;

		dbfolder = theApp.FindFolderByName("Gene Databases", GENE_DATABASE);
		if (dbfolder != NULL)
			theApp.hOtherSpeciesFolder = pServer->AddFolderToTree(dbfolder->hFolder, "Other Species");

	::DeleteFile(szPathAndFile);
	return TRUE;
}

// Returns a pointer to a CDBFolder object in the folder list that
// matches the specified folder name or NULL if not found
CDBFolder* CGenMAPPDBDLApp::FindFolderByName(CString szFolderName, int nType)
{
	POSITION pos;
	for( pos = DBFolderList.GetHeadPosition(); pos != NULL; )
	{
		CDBFolder*		pFolder;
		pFolder = (CDBFolder*)DBFolderList.GetNext( pos );

		if (pFolder->szFolderName == szFolderName && pFolder->nType == nType)
			return pFolder;
	}

	return NULL;
}
// Return the pointer to the FTP server and set the HTTPServer
// flag. Don't ask. The code got this way "orgranically."
CGMFTP* CGenMAPPDBDLApp::CreateHTTPControlConnection(void)
{
	theApp.pGenMAPPServer->bHTTPServer = true;
	return theApp.pGenMAPPServer;
}
