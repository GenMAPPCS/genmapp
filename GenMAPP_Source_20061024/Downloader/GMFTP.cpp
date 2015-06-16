// GMFTP.cpp : implementation file
//
#include "stdafx.h"
#include "DBFile.h"
#include "DBFolder.h"
#include "GenMAPPDBDL.h"
#include ".\gmftp.h"

// CGMFTP

CGMFTP::CGMFTP(LPCTSTR pstrAgent /*= NULL*/,
		DWORD dwContext /*= 1*/,
		DWORD dwAccessType /*= PRE_CONFIG_INTERNET_ACCESS*/,
		LPCTSTR pstrProxyName /*= NULL*/,
		LPCTSTR pstrProxyBypass /*= NULL*/,
		DWORD dwFlags /*= 0*/)
		: CInternetSession (pstrAgent, dwContext, dwAccessType, pstrProxyName, 
							pstrProxyBypass, dwFlags)
							, m_pConnect(NULL)
							, fConnSpeed(0)
							, szServerKeyName(_T(""))
							, szSpeedTestFolder(_T(""))
							, bConnected(false)
							, bHTTPServer(false)
							, m_pHttpConnect(NULL)
							, bFTPBackup(false)
							, nTier(1)
{
	pstrAgent = "GenMAPPDBDL";
}


CGMFTP::~CGMFTP()
{
	// if the connection is open, close it
	if (m_pConnect != NULL)
		m_pConnect->Close();
	bConnected = false;
	delete m_pConnect;
}


// CGMFTP member functions

// Determines how fast users connection is to GenMAPP FTP site
// Test involves downloading a 128KB, randomly genrerated file.
double CGMFTP::SpeedTest(CString szAltSpeedTestDir)
{
	CGenMAPPDBDLApp* m_pApp = (CGenMAPPDBDLApp*)AfxGetApp();
	m_pApp->m_nInetConnected = SPEED_TEST;
	try
	{
		if (!bHTTPServer)
			return SpeedTestFTP(szAltSpeedTestDir);
		else
			return SpeedTestHttp(szAltSpeedTestDir);
	}
	catch (CInternetException* pEx)
	{
		TCHAR sz[1024];
		pEx->GetErrorMessage(sz, 1024);
		printf("ERROR!  %s\n", sz);
		pEx->Delete();
	}

	return 0;
    	
}

// FTP version of the speed test
double CGMFTP::SpeedTestFTP(CString szAltSpeedTestDir)
{
	CGenMAPPDBDLApp* m_pApp = (CGenMAPPDBDLApp*)AfxGetApp();
#ifdef DEBUG		
	CString szCurrDir = "";
	m_pConnect->GetCurrentDirectory(szCurrDir);
	TRACE("Current Directory is %s running SpeedTest\n", szCurrDir);
	TRACE("Setting Directory to %s for SpeedTest\n", szAltSpeedTestDir == "" ? 
		szSpeedTestFolder : szAltSpeedTestDir);
#endif
	if (!m_pConnect->SetCurrentDirectory(szAltSpeedTestDir == "" ? 
		szSpeedTestFolder : szAltSpeedTestDir))
	{
		LPVOID lpMsgBuf;
		DWORD  dwLastError = GetLastError();
		FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | 
			FORMAT_MESSAGE_IGNORE_INSERTS,
			NULL,
			dwLastError,
			MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
			(LPTSTR) &lpMsgBuf,
			0,
			NULL );
		LocalFree(lpMsgBuf);
		//TRACE(szError);
	}
	DWORD BeginTime, EndTime;

	BeginTime = GetTickCount();
	if (!m_pConnect->GetFile("SpeedTester.bin", m_pApp->szBasePath + "\\GenMAPPSpeedTest.tmp", false))
	{
		LPVOID lpMsgBuf;
		DWORD  dwLastError = GetLastError();
		FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | 
			FORMAT_MESSAGE_IGNORE_INSERTS,
			NULL,
			dwLastError,
			MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
			(LPTSTR) &lpMsgBuf,
			0,
			NULL );
		LocalFree(lpMsgBuf);
	}
	EndTime = GetTickCount();

	::DeleteFile(m_pApp->szBasePath + "\\GenMAPPSpeedTest.tmp");
	double Elapsed = ((double)EndTime - (double)BeginTime) / 1000;

	// x seconds to dl 128k, how much can be dl in 1 sec.
	return fConnSpeed = 128 / Elapsed;
}

// Http version of the Speed Test
double CGMFTP::SpeedTestHttp(CString szAltSpeedTestDir)
{
	char*		szTempDir;
	CString		szPathAndFile, 
		szSpeedPath = szAltSpeedTestDir == "" ? szHTTPSpeedTestFolder : szAltSpeedTestDir;

	szTempDir = getenv("TEMP");
	if (szTempDir == NULL)
		szTempDir = getenv("TMP");

	szPathAndFile.Format("%s\\SpeedTester.bin", szTempDir);

	DWORD BeginTime, EndTime;
	BeginTime = GetTickCount();
	GetHttpFile(szSpeedPath + "SpeedTester.bin", szPathAndFile);

	EndTime = GetTickCount();

	::DeleteFile(szPathAndFile);

	double Elapsed = ((double)EndTime - (double)BeginTime) / 1000;

	// x seconds to dl 128k, how much can be dl in 1 sec.
	return fConnSpeed = 128 / Elapsed;
}

// Connect to the specified GenMAPP server. All operations
// will take place on this server.
UINT CGMFTP::ConnectToGMServer()
{
	DWORD	nRet = 0;

	if (!bHTTPServer && (szHost == "" || szUserName == "" || szPassword == ""))
		return 1;

	try
	{
		if (!bHTTPServer)
		{	m_pConnect = 	GetFtpConnection(_T(szHost), szUserName, szPassword);
			bConnected = (m_pConnect != NULL);
		}
		else
		{
			if (bConnected)
				DisconnectFromGMServer();

			m_pHttpConnect = GetHttpConnection(_T(szHost), INTERNET_FLAG_DONT_CACHE, 80, szUserName, szPassword);
			//bConnected = (m_pHttpConnect != NULL);
		}
	}
	catch (CInternetException* pEx)
	{
		char szTheError[255];
		nRet = (UINT)pEx->m_dwError;
		pEx->GetErrorMessage(szTheError, 255);
		TRACE("Error connecting to GMServer: %d, %s\n", pEx->m_dwError, szTheError);
		pEx->Delete();
		bConnected = false;
	}

	return nRet;
}

BOOL CGMFTP::RetrieveFromRootFolder(CString szRootFolder)
{
	try
	{
		if (RecurseGMFolder(szRootFolder.TrimRight("/"), szRootFolder) == FALSE)
			return FALSE;

		CGenMAPPDBDLApp* m_pApp = (CGenMAPPDBDLApp*)AfxGetApp();
		CDBFolder*	dbfolder = NULL;

		HWND	hTreeView = ::FindWindowEx(m_pApp->hMainDlg, NULL, "SysTreeView32", NULL);

		dbfolder = m_pApp->FindFolderByName(RetriveFileTypeString(GENE_DATABASE), GENE_DATABASE);
		//if (dbfolder != NULL)
		//	TreeView_Expand(hTreeView, dbfolder->hFolder, TVE_COLLAPSE);

		dbfolder = m_pApp->FindFolderByName(RetriveFileTypeString(MAPP_DATABASE), MAPP_DATABASE);
		//if (dbfolder != NULL)
		//	TreeView_Expand(hTreeView, dbfolder->hFolder, TVE_COLLAPSE);

		dbfolder = m_pApp->FindFolderByName(RetriveFileTypeString(GENEEXPRESS_DATABASE), GENEEXPRESS_DATABASE);
		//if (dbfolder != NULL)
		//	TreeView_Expand(hTreeView, dbfolder->hFolder, TVE_COLLAPSE);

		dbfolder = m_pApp->FindFolderByName(RetriveFileTypeString(MAPP_ARCHIVE), MAPP_ARCHIVE);
		//if (dbfolder != NULL)
		//	TreeView_Expand(hTreeView, dbfolder->hFolder, TVE_COLLAPSE);

	}
	catch (CInternetException* pEx)
	{
		UINT nRet = (UINT)pEx->m_dwError;
		pEx->Delete();
	}

	return TRUE;
}

// Adds a Database File object to the list for each database
// file contained in the specified direcory.
bool CGMFTP::BuildListFromDBDir(int nDBType, CObList* pList)
{
	CGenMAPPDBDLApp* m_pApp = (CGenMAPPDBDLApp*)AfxGetApp();
	if (!ChangeGMDir(nDBType))
		return false;

	try
	{
		switch (nDBType)
		{
			case GENE_DATABASE:
				PopulateListFromType(nDBType, pList, "*.gdb");
				break;

			case MAPP_DATABASE:
			case MAPP_ARCHIVE:
				PopulateListFromType(nDBType, pList, "*.mapp");
				break;

			case GENEEXPRESS_DATABASE:
				PopulateListFromType(nDBType, pList, "*.gex");
				break;

			default:
				return false;
		}

		if (m_pApp->bAbortDownload)
			return false;

		PopulateListFromType(nDBType, pList, "*.exe");
		if (m_pApp->bAbortDownload)
			return false;
		
		PopulateListFromType(nDBType, pList, "*.bat");
		if (m_pApp->bAbortDownload)
			return false;

	}
	catch (CInternetException* pEx)
	{
		TCHAR sz[1024];
		pEx->GetErrorMessage(sz, 1024);
		printf("ERROR!  %s\n", sz);
		pEx->Delete();
	}

	return true;
}

// *********** DEPRECIATED  ************
// Changes the current directory based on the file type.
bool CGMFTP::ChangeGMDir(int nDirType)
{
	CString	szDirName;
	

	
	try
	{
		if (!m_pConnect->SetCurrentDirectory(_T("/GenMAPP2Data")))
			return false;

		if (!m_pConnect->SetCurrentDirectory(_T(szDirName)))
			return false;
	}
	catch (CInternetException* pEx)
	{
		TCHAR sz[1024];
		pEx->GetErrorMessage(sz, 1024);
		printf("ERROR!  %s\n", sz);
		pEx->Delete();
	}

	return true;
}

// Populates CDBFile objects for each file found in the directory matching the extension
bool CGMFTP::PopulateListFromType(int nDBType, CObList* pList, CString szExt)
{
	CGenMAPPDBDLApp* m_pApp = (CGenMAPPDBDLApp*)AfxGetApp();
	try
	{
		int		nDateStart;
		CFtpFileFind finder(m_pConnect);

		// start looping
		BOOL bWorking = finder.FindFile(szExt, INTERNET_FLAG_EXISTING_CONNECT | INTERNET_FLAG_RELOAD);

		while (bWorking)
		{
			if (m_pApp->bAbortDownload == true)
			{
				finder.Close();
				DisconnectFromGMServer();
				return false;
			}
			CDBFile* dbfile = new CDBFile();

			bWorking = finder.FindNextFile();
			
			// Get the file name, chop off the path
			dbfile->szSrcFileName = finder.GetFileURL();
			dbfile->szSrcFileName = dbfile->szSrcFileName.Right(dbfile->szSrcFileName.GetLength() - dbfile->szSrcFileName.ReverseFind('/') - 1);
			
			dbfile->bFileIsCompressed = false;
			dbfile->bFTP = true;
			if (nDBType != GENE_DATABASE)
				dbfile->bIncludesGenBank = false;
			else
				dbfile->bIncludesGenBank = (dbfile->szSrcFileName.Find("-GB_") > 0); //|| 
					//dbfile->szSrcFileName.Find("-gb_") || dbfile->szSrcFileName.Find("-Gb_"));
			
			dbfile->dwFileSize = (DWORD)finder.GetLength();
			dbfile->nFileType = nDBType;
			nDateStart = dbfile->szSrcFileName.Find("_2") + 1;
			if (nDateStart)
				dbfile->DBDate = CTime(atoi(dbfile->szSrcFileName.Mid(nDateStart, 4)),
					atoi(dbfile->szSrcFileName.Mid(nDateStart + 4, 2)),
					atoi(dbfile->szSrcFileName.Mid(nDateStart + 6, 2)), 0, 0, 0);
			else
				dbfile->DBDate = CTime(1980, 01, 01, 0, 0, 0);
			
			pList->AddTail(dbfile);
		}

		finder.Close();
	}
	catch (CInternetException* pEx)
	{
		TCHAR sz[1024];
		pEx->GetErrorMessage(sz, 1024);
		printf("ERROR!  %s\n", sz);
		pEx->Delete();
	}
	return true;
}

// Logs out of the GenMAPP Server
bool CGMFTP::DisconnectFromGMServer(void)
{
	if (!bConnected)
		return false;

	bConnected = false;
	if (!bHTTPServer)
	{
		if (m_pConnect != NULL)
			m_pConnect->Close();
		else
			return false;
	}
	//else
	//{
	//	if (m_pHttpConnect != NULL)
	//	{
	//		//m_pHttpConnect->Close();
	//		delete m_pHttpConnect;
	//	}
	//	else
	//		return false;
	//}

	return true;
}

// Populates the variables necessary to log in to the server
void CGMFTP::SetLoginData(CString szHostInit, CString szUserNameInit, CString szPasswordInit, CString szLocationInit, CString szAliasInit)
{
	szHost = szHostInit;
	szUserName = szUserNameInit;
	szPassword = szPasswordInit;
	szLocation = szLocationInit;
	if (szAliasInit != "")
		szAlias = szAliasInit;
	else
		szAlias = szHostInit;
}

// Recursively drills down the folder structure, adding the folders 
// and files found at each level
BOOL CGMFTP::RecurseGMFolder(CString szRootFolder, CString szStartingURL)
{
	// Start a file finder
	// for each file:
	// is it a folder?
	//		yes - recurse
	// is it a file?
	//		yes - determine type
	//      Create folder for file type if not already created
	//		from root folder, tokenize each directory level. for each level:
	//			Does folder object already exist?
	//				No- create folder object, saving HTREEItem and ptr. Add folder to TreeView control.
	//				Yes - obtain HTREEItem and ptr for next iteration
	//      
	//		Add file to folder object, parenting to saved htreeitem
	CGenMAPPDBDLApp* m_pApp = (CGenMAPPDBDLApp*)AfxGetApp();
	// Did user click Cancel/Back button?
	if (m_pApp->bAbortDownload == true)
		return FALSE;


	CStringList		FolderList;
	POSITION		pos;
	BOOL			bAddtlLines = TRUE;
	CHttpFile*		pFile;
	CString			szLine;
	CFtpFileFind*	pFinder;

	if (!bHTTPServer)
		pFinder = new CFtpFileFind(m_pConnect);

	if (bHTTPServer)
	{
		if (ConnectToGMServer() != 0)
			return TRUE;

		pFile = m_pHttpConnect->OpenRequest(CHttpConnection::HTTP_VERB_GET, szStartingURL, NULL, 1, NULL, NULL, INTERNET_FLAG_DONT_CACHE);
		pFile->AddRequestHeaders("Content-Type: */*");
		if (!pFile->SendRequest())
			return TRUE;
	}
	else
		if (!m_pConnect->SetCurrentDirectory(_T(szStartingURL)))
			return TRUE;



	// Read a line. In IIS, there are no CRLFs, only <br>'s
	// Apache has no <br>'s and one CRLF per file/folder.
	if (bHTTPServer)
		bAddtlLines = pFile->ReadString(szLine);

	// outside HTTP loop here, always false for FTP
	while (bAddtlLines)
	{
		int		nNextFileIndex = 0;
		BOOL	bWorking = TRUE;
		// start inner loop
		if (!bHTTPServer)
		{
			bWorking = pFinder->FindFile("*", INTERNET_FLAG_EXISTING_CONNECT | INTERNET_FLAG_RELOAD);
			bAddtlLines = FALSE;
		}
		else
			bAddtlLines = pFile->ReadString(szLine);

		while (bWorking && nNextFileIndex != -1)
		{
			int			nFileType, nIndex = 0, nIndex2 = 0;
			CDBFolder*	pLastFolder = NULL;
			CDBFile*	dbfile = NULL;
			CString		szTemp, szFinderFile, szFolder;
			
			// Did user click Cancel/Back button?
			if (m_pApp->bAbortDownload == true)
			{
				if (!bHTTPServer)
					pFinder->Close();

				DisconnectFromGMServer();
				return FALSE;
			}

			if (bHTTPServer)
			{
				dbfile = FindNextHTTPFile(&nNextFileIndex, &szFolder, nNextFileIndex, &szLine, szStartingURL);
				if (dbfile == NULL && szFolder == "")
					continue;
			}
			else
			{
				bWorking = pFinder->FindNextFile();
				szFinderFile = pFinder->GetFileName();
			}

			// If it's a directory, save its name and search
			// after this directory has been processed.
			if ((!bHTTPServer && pFinder->IsDirectory()) || (bHTTPServer && szFolder != ""))
			{
				if (!bHTTPServer)
					szFolder = szFinderFile;

				if (szFolder.Right(5).MakeUpper() != "SPLIT")
					FolderList.AddHead(szFolder);

				continue;
			}
			
			// Filter out any non-database files
			nFileType = DetermineFileType(bHTTPServer ? dbfile->szSrcFileName : szFinderFile);
			if (nFileType == UNKNOWN_DATABASE)
			{
				if (bHTTPServer)
					delete dbfile;

				continue;
			}

			// Create root file type folder if it does not exist
			pLastFolder = AddFolder(RetriveFileTypeString(nFileType), nFileType, TVI_ROOT);

			// Find position of first subfolder after root directory of
			// the path we are currently processing. This folder is
			// the first we will add to the folder tree
			szTemp = szRootFolder.Tokenize("/", nIndex2);
			while (szTemp != "")
			{
				szTemp = szRootFolder.Tokenize("/", nIndex2);
				szStartingURL.Tokenize("/", nIndex);
			}

			// Add a folder for each folder in path of the file
			szTemp = szStartingURL.Tokenize("/", nIndex);
			while (szTemp != "")
			{
				CDBFolder* pNewFolder = AddFolder(szTemp, nFileType, pLastFolder->hFolder);
				pLastFolder = pNewFolder;
				szTemp = szStartingURL.Tokenize("/", nIndex);
			}

			// Add the file
			CDBFile* pDBFile;
			if (bHTTPServer)
				pDBFile = dbfile;
			else
				pDBFile = PopulateDBFile(pFinder->GetFilePath(), (DWORD)pFinder->GetLength(), true);

			if (pDBFile != NULL)
			{
				if (pDBFile->DBServerList.GetCount() < 2)
				{
					m_pApp->pFileList->AddHead(pDBFile);
					pLastFolder->DBFileList.AddHead(pDBFile);
				}
			}
		}
	}

	if (!bHTTPServer)
	{
		pFinder->Close();
		delete pFinder;
	}
	else
		DisconnectFromGMServer();

	// Now process all folders found in the current folder
	for( pos = FolderList.GetHeadPosition(); pos != NULL; )
	{
		CString szQueuedFolder = FolderList.GetNext( pos );
		szStartingURL.TrimRight("/");
		if (!RecurseGMFolder(szRootFolder, szStartingURL + "/" + szQueuedFolder + "/"))
			return FALSE;
	}

	return TRUE;
}


// Parses one line of a HTTP folder listing returns results. If the return value is -1,
// there are no more files on this line. If return value is a positive number, use this 
// as the nNextFileIndex in subsequent calls to retrieve the next HTTP file. 
// If a file is found, returns a populated CDBFile object. If a folder is found, returns 
// the folder name.
CDBFile* CGMFTP::FindNextHTTPFile(int* pnNextBRIndex, CString* szFolder, int nNextFileIndex, CString* szLine, CString szPathToFile)
{
	CString szLineToBR, szCapsLine = *szLine;
	CDBFile	staticDBFile;
	CDBFile* dbfile = NULL;

	TRACE("Line length %d\n", szLine->GetLength());

	szCapsLine.MakeUpper();
	while (nNextFileIndex != -1)
	{
		int nAIndex = 0;
		CString	szCapsLineToBR;
		*pnNextBRIndex = szCapsLine.Find("<BR", nNextFileIndex + 3);
		//nNextFileIndex = szCapsLine.Find("<BR", nNextFileIndex);
		szLineToBR = szLine->Mid(nNextFileIndex + 3, *pnNextBRIndex == -1 ? 
			szLine->GetLength() : *pnNextBRIndex);
		szCapsLineToBR = szLineToBR;
		szCapsLineToBR.MakeUpper();

		// Look for A HREF
		while (nAIndex != -1)
		{
			nAIndex = szCapsLineToBR.Find("<A HREF=\"", nAIndex);
			//if (nAIndex == -1)
			//	nAIndex = szCapsLineToBR.Find("<A HREF='", nAIndex);

			if (nAIndex != -1)
			{
				CString	szHREFURL, szDisplayedText;
				int nEndQuoteIndex = szCapsLineToBR.Find("\"", nAIndex + 9), 
					nCloseTagIndex = 0, nLastSlashIndex = -1;
				if (nEndQuoteIndex == -1)
					continue;

				szHREFURL = szCapsLineToBR.Mid(nAIndex + 9, nEndQuoteIndex - (nAIndex + 9));

				szHREFURL = staticDBFile.RemoveEscapeSequences(szHREFURL);
				if (szHREFURL.Right(1) == "/")
					szHREFURL = szHREFURL.TrimRight('/');
				

				nCloseTagIndex = szCapsLineToBR.Find("</A", nEndQuoteIndex);
				szDisplayedText = szCapsLineToBR.Mid(nEndQuoteIndex + 2, (nCloseTagIndex) - (nEndQuoteIndex + 2));
				if (szDisplayedText.Right(1) == "/")
					szDisplayedText = szDisplayedText.TrimRight('/');

				// MS: trim path in URL to target file or folder
				nLastSlashIndex = szHREFURL.ReverseFind('/');
				if (nLastSlashIndex != -1)
					szHREFURL = szHREFURL.Mid(nLastSlashIndex + 1, szHREFURL.GetLength() - nLastSlashIndex);

				TRACE("Right of URL Contains a slash: %s\n", (szLineToBR.Mid(nAIndex + 9 + nLastSlashIndex + 1, nEndQuoteIndex - (nAIndex + 9 + nLastSlashIndex + 1)).Right(1) == "/") ? "True" : "False");
				if (szDisplayedText == szHREFURL)
				{
					if (szCapsLineToBR.Find("DIR") != -1)
						// We have a folder
						*szFolder = szLineToBR.Mid(nEndQuoteIndex + 2, nCloseTagIndex - 
							((szLineToBR.Mid(nEndQuoteIndex + 2, nCloseTagIndex - (nEndQuoteIndex + 2)).Right(1) == "/") ? 1 : 0) - (nEndQuoteIndex + 2));
					else
						dbfile = PopulateDBFile(szPathToFile.TrimRight("/") + "/" + szLineToBR.Mid(nEndQuoteIndex + 2, nCloseTagIndex - 
							((szLineToBR.Mid(nEndQuoteIndex + 2, nCloseTagIndex - (nEndQuoteIndex + 2)).Right(1) == "/") ? 1 : 0) - (nEndQuoteIndex + 2)),
							GetHTMLFileSize(szLineToBR), false);
				}

				nAIndex = nEndQuoteIndex;
			}
		}

		nNextFileIndex = *pnNextBRIndex;
	}
	
	return dbfile;
}

// Parses the file name to determine the file type. Returns on of the predefined constants.
int CGMFTP::DetermineFileType(CString szFileName)
{
	// Gene Databaes:  xx-Std or gb
	// MAPP Archives:  xx-MAPP_Archive
	// MAPPs:
	// Expression Datasets:

	szFileName = szFileName.MakeUpper();
	if (szFileName.Find("-STD") != -1 || szFileName.Find("-GB") != -1 || szFileName.Find("CONVERTER") != -1)
		return GENE_DATABASE;

	if (szFileName.Find("MAPP_ARCHIVE") != -1)
		return MAPP_ARCHIVE;

	if (szFileName.Find(".GEX") != -1 || szFileName.Find("ED_") != -1)
		return GENEEXPRESS_DATABASE;

	if (szFileName.Find(".MAPP") != -1)
		return MAPP_DATABASE;

	return UNKNOWN_DATABASE;
}

// Returns the string representation of a file type constant. If the the constant is unknown, returns an empty string.
CString CGMFTP::RetriveFileTypeString(int nFileType)
{
	switch (nFileType)
	{
		case GENE_DATABASE:
			return "Gene Databases";

		case MAPP_DATABASE:
			return "MAPPs";

		case MAPP_ARCHIVE:
			return "MAPP Archives";

		case GENEEXPRESS_DATABASE:
			return "Expression Datasets";

		default:
			return "";
	}
}

// Populates a DBFile object given the FTP file object and the handle
// to the folder it belongs under
CDBFile* CGMFTP::PopulateDBFile(CString szFileName, DWORD dwFileSize, bool bFTP)
{
	CDBFile* dbfile = DoesFileExist(szFileName);

	if (dbfile == NULL)
	{
		dbfile = new CDBFile();
		if (dbfile == NULL)
			return NULL;

		int		nDateStart = -1;
		
		if (bFTP)
			dbfile->szSrcFileName = szFileName.Mid(szFileName.Find(this->szSpeedTestFolder) + this->szSpeedTestFolder.GetLength());
		else
			dbfile->szSrcFileName = szFileName.Mid(szFileName.Find(this->szHTTPSpeedTestFolder) + this->szHTTPSpeedTestFolder.GetLength());

		dbfile->bFileIsCompressed = false;
		dbfile->bFTP = bFTP;
		dbfile->nFileType = DetermineFileType(dbfile->szSrcFileName);

		if (dbfile->nFileType != GENE_DATABASE)
			dbfile->bIncludesGenBank = false;
		else
			dbfile->bIncludesGenBank = (dbfile->szSrcFileName.Find("-GB") > 0); 

		dbfile->bSelected = false;
		
		dbfile->dwFileSize = dwFileSize;
		nDateStart = dbfile->szSrcFileName.Find("_2") + 1;
		if (nDateStart)
			dbfile->DBDate = CTime(atoi(dbfile->szSrcFileName.Mid(nDateStart, 4)),
				atoi(dbfile->szSrcFileName.Mid(nDateStart + 4, 2)),
				atoi(dbfile->szSrcFileName.Mid(nDateStart + 6, 2)), 0, 0, 0);
	}

	dbfile->DBServerList.AddHead(this);
	
	return dbfile;
}

// Adds the specified folder to the database folder tree
HTREEITEM CGMFTP::AddFolderToTree(HTREEITEM hParent, CString szFolderName)
{
	CGenMAPPDBDLApp* pApp = (CGenMAPPDBDLApp*)AfxGetApp();
	HWND	hTreeView = ::FindWindowEx(pApp->hMainDlg, NULL, "SysTreeView32", NULL);

	if (hTreeView != NULL && szFolderName.GetLength() < 255)
	{
		TVINSERTSTRUCT tvins;
		HTREEITEM hItem = NULL;
		char	szFolderNameChar[255];

		sprintf(szFolderNameChar, "%s", szFolderName);
		tvins.hParent = hParent;
		tvins.hInsertAfter = TVI_SORT;
		tvins.item.mask = TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_TEXT;
		tvins.item.pszText = szFolderNameChar;
		tvins.item.cchTextMax = (int)strlen(szFolderNameChar);
		tvins.item.iImage = 0;
		tvins.item.iSelectedImage = 1;
		hItem = TreeView_InsertItem(hTreeView, &tvins);
		//TreeView_Expand(hTreeView, hItem, TVE_EXPAND);
		return hItem;
	}
	else
		return NULL;
}

// Verifies the existence of a data file found in DataLocations.cfg and adds to the folder tree and file list
BOOL CGMFTP::RetrieveIndividualFile(CString szPathAndFileName, CString szDataTreeFolder)
{
	int		nFileType = UNKNOWN_DATABASE, nIndex = 0;
	CDBFolder*	pLastFolder = NULL;
	CString		szTemp, 
		szFileNameOnly = szPathAndFileName.Right(szPathAndFileName.GetLength() - szPathAndFileName.ReverseFind('/') - 1),
		szPathOnly = szPathAndFileName.Left(szPathAndFileName.ReverseFind('/'));
	CGenMAPPDBDLApp* pApp = (CGenMAPPDBDLApp*)AfxGetApp();

	// Does file exist?
	if (!m_pConnect->SetCurrentDirectory(szPathOnly))
		return FALSE;

	CFtpFileFind pInetFile(m_pConnect);
	if (!pInetFile.FindFile(szFileNameOnly, INTERNET_FLAG_EXISTING_CONNECT | INTERNET_FLAG_RELOAD))
		return FALSE;
		
	pInetFile.FindNextFile();  // Actually, find the same file. Thanks MS for all the time it took me to figure that one out!

	nFileType = DetermineFileType(szFileNameOnly);
	if (nFileType == UNKNOWN_DATABASE)
		return FALSE;

	// Create root file type folder if it does not exist
	pLastFolder = AddFolder(RetriveFileTypeString(nFileType), nFileType, TVI_ROOT);

	// Add a folder for each folder in path of the file
	szTemp = szDataTreeFolder.Tokenize("/", nIndex);
	while (szTemp != "")
	{
		CDBFolder* pNewFolder = AddFolder(szTemp, nFileType, pLastFolder->hFolder);
		pLastFolder = pNewFolder;
		szTemp = szDataTreeFolder.Tokenize("/", nIndex);
	}

	// Add the file
	CDBFile* pDBFile = PopulateDBFile(pInetFile.GetFilePath(), (DWORD)pInetFile.GetLength(), true);
	if (pDBFile != NULL)
	{
		if (pDBFile->DBServerList.GetCount() < 2)
			pApp->pFileList->AddHead(pDBFile);
		pLastFolder->DBFileList.AddHead(pDBFile);
	}
	else
		return FALSE;

	pInetFile.Close();

	return TRUE;
}

// Returns a pointer to the CDBFile object if it already exists. Returns NULL otherwise.
CDBFile* CGMFTP::DoesFileExist(CString szFileName)
{
	CGenMAPPDBDLApp* pApp = (CGenMAPPDBDLApp*)AfxGetApp();
	if (pApp->pFileList != NULL)
	{
		CString szSrcFileNameOnly, szDestFileNameOnly;

		// Trim path and extension from file name
		szSrcFileNameOnly = szFileName.Right(szFileName.GetLength() - szFileName.ReverseFind('/') - 1);
		szSrcFileNameOnly = szSrcFileNameOnly.Left(szSrcFileNameOnly.ReverseFind('.'));

		POSITION	pos;
		CDBFile* dbfile;
		
		for( pos = pApp->pFileList->GetHeadPosition(); pos != NULL; )
		{
			dbfile = (CDBFile*)pApp->pFileList->GetNext( pos );
			szDestFileNameOnly = dbfile->GetFileNameOnly(dbfile->szSrcFileName);
			if (szSrcFileNameOnly == szDestFileNameOnly)
				return dbfile;
		}
	}

	return NULL;
}
// Adds a the folder name to the folder tree and folder tree collection parented 
// to the specified item and returns a handle to the new folder object. If the 
// folder already exists, simply returns a pointer to the folder.
CDBFolder* CGMFTP::AddFolder(CString szFolderName, int nFileType, HTREEITEM hParentFolder)
{
	CGenMAPPDBDLApp* pApp = (CGenMAPPDBDLApp*)AfxGetApp();
	CDBFolder*	pNewFolder = pApp->FindFolderByName(szFolderName, nFileType);
	if (pNewFolder == NULL)
	{
		pNewFolder = new CDBFolder();
		pNewFolder->szFolderName = szFolderName;
		pNewFolder->nType = nFileType;
		pNewFolder->hFolder = AddFolderToTree(hParentFolder, szFolderName);
		pNewFolder->hParentFolder = hParentFolder;
		pApp->DBFolderList.AddHead(pNewFolder);
		HWND	hTreeView = ::FindWindowEx(pApp->hMainDlg, NULL, "SysTreeView32", NULL);
		//TreeView_Expand(hTreeView, pNewFolder->hParentFolder, TVE_EXPAND);
	}

	return pNewFolder;
}


// Given a URL, retrieves a http file and stores it on the local file system. Returns false if a failure occurs.
BOOL CGMFTP::GetHttpFile(CString szHttpFile, CString szDestFile)
{
	szHttpFile.OemToAnsi();
	
	if (ConnectToGMServer())
		return FALSE;
	
	try
	{
		CFile DestFile(szDestFile, CFile::modeCreate | CFile::modeWrite | CFile::shareExclusive);
		if (DestFile.m_hFile == CFile::hFileNull)
			return FALSE;
		
		UINT	nBytesRead = 0;
		DWORD	dwStatusCode = 0;

		//szHttpFile = "/gmu/GenMAPP2/GenMAPPv2.exe";
		//szHttpFile = "/gmu/GenMAPP2/GenMAPPv2.exe";

		CHttpFile* pFile = m_pHttpConnect->OpenRequest(CHttpConnection::HTTP_VERB_GET, szHttpFile, NULL, GetTickCount(), NULL, NULL, INTERNET_FLAG_DONT_CACHE);

        if (pFile == NULL)
			return FALSE;
		
		pFile->QueryInfoStatusCode(dwStatusCode);
		TRACE("FullURL: %s, Status code: %i, ServerName: %s\n", pFile->GetFileURL(), dwStatusCode, m_pHttpConnect->GetServerName());

		pFile->AddRequestHeaders("Content-Type: */*");
		pFile->SendRequest();

		if (!pFile->QueryInfoStatusCode(dwStatusCode))
		{
			delete pFile;
			return FALSE;
		}
		else if (dwStatusCode < 200 || dwStatusCode > 299)
		{
			delete pFile;
			return FALSE;
		}

		BYTE*	pTempBuf = new BYTE[1024];
		
		nBytesRead = pFile->Read(pTempBuf, 1024);
		while (nBytesRead > 0)
		{
			DestFile.Write(pTempBuf, nBytesRead);
			nBytesRead = pFile->Read(pTempBuf, 1024);
		}

		DestFile.Close();
		pFile->Close();

		DisconnectFromGMServer();

		delete pFile;
		delete pTempBuf;

		return TRUE;
	}
	catch (CException* pEx)
	{
		TCHAR sz[1024];
		pEx->GetErrorMessage(sz, 1024);
		printf("ERROR!  %s\n", sz);
		pEx->Delete();
		return FALSE;
	}
}
// Opens an HTTP connection, creates an HTTPFile object, adds the appropriate
// headers and then returns the HTTPFile object as an InternetFile object.
CInternetFile* CGMFTP::OpenHTTPFile(CString szURL)
{
	try
	{
		ConnectToGMServer(); 
		CHttpFile* pFile = m_pHttpConnect->OpenRequest(CHttpConnection::HTTP_VERB_GET, szURL, NULL, 1, NULL, NULL, INTERNET_FLAG_DONT_CACHE);

		pFile->AddRequestHeaders("Content-Type: */*");
		pFile->SendRequest();

		return (CInternetFile*)pFile;
	}
	catch (CException* pEx)
	{
		TCHAR sz[1024];
		pEx->GetErrorMessage(sz, 1024);
		printf("ERROR!  %s\n", sz);
		pEx->Delete();
		return NULL;
	}
}

// Parses one line in an HTML folder listing and returns the size
// of the file or 0 if the file size could not be determined.
DWORD CGMFTP::GetHTMLFileSize(CString szHTMLLine)
{
	// Sample IIS line:
	// Tuesday, May 18, 2004 10:00 AM       148992 <A HREF="/WLNPDocs/End%20User%20Service%20Request.doc">End User Service Request.doc</A>

	// Sample Apache Line:
	// <img src="/icons/binary.gif" alt="[   ]" /> <a href="Dr-GB_20040411.exe">Dr-GB_20040411.exe</a>           06-Jul-2004 10:37   15M

	// The size always follows the time, so find the time first
	int i, nColonIndex = szHTMLLine.Find(':');
	double dFileSize = 0;

	// No colon, no time, let's get out of here!
	if (nColonIndex == -1)
		return 0;

	TRACE("This should be a colon- %c\n", szHTMLLine[nColonIndex]);

	// Make sure the character to the left is a number. If it's any other character, this
	// isn't the file time.
	if (szHTMLLine[nColonIndex - 1] < 48 || szHTMLLine[nColonIndex - 1] > 57)
		return 0;

	// Find the first numeric character after the time
	for (i=nColonIndex + 3;i < szHTMLLine.GetLength() - 1;i++)
	{
		if (szHTMLLine[i] >= 48 && szHTMLLine[i] <= 57)
		{
			// Found it, convert the number to a DWORD value
			dFileSize = atof(szHTMLLine.Mid(i, szHTMLLine.GetLength() - i));

			// In Apache, if the file size exceeds 1k, the file size is
			// abbreviated with a K, M or G designation. Find that designation
			// and do appropriate math to convert to a precise number.
			int j;
			TRACE("Size suffix search: ");
			for (j=i;j < szHTMLLine.GetLength() - 1;j++)
			{
				TRACE("%c", szHTMLLine[j]);
				if ((szHTMLLine[j] < 48 || szHTMLLine[j] > 57) && szHTMLLine[j] != 46)
				{
					switch (szHTMLLine[j])
					{
						case 'K':
							dFileSize*= 1024;
							break;

						case 'M':
							dFileSize*= 1048576;
							break;

						case 'G':
							dFileSize*= 1073741824;
							break;
					}

					TRACE("  Done\n");
					return (DWORD)dFileSize;
				}
			}
		}
	}

	return (DWORD)dFileSize;
}

bool CGMFTP::TestHTTPConnection()
{
	try
	{
		CHttpFile*	pFile = m_pHttpConnect->OpenRequest(CHttpConnection::HTTP_VERB_GET, szHTTPSpeedTestFolder, NULL, 1, NULL, NULL, INTERNET_FLAG_DONT_CACHE);
		if (pFile == NULL)
			return false;

		pFile->AddRequestHeaders("Content-Type: */*");
		bConnected = (pFile->SendRequest());
		return bConnected;
	}
	catch (CInternetException* pEx)
	{
		return false;
	}
}

