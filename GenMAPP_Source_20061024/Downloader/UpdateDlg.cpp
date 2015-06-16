// UpdateDlg.cpp : implementation file
//

#include "stdafx.h"
#include "GenMAPPDBDL.h"
#include "ConfigFile.h"
#include "UpdateDlg.h"
#include ".\updatedlg.h"
#include "PromptForUpdate.h"

#define SCALEX(argX) ((int) ((argX) * m_pApp->scaleX))
#define SCALEY(argY) ((int) ((argY) * m_pApp->scaleY))

// CUpdateDlg dialog

IMPLEMENT_DYNAMIC(CUpdateDlg, CDialog)
CUpdateDlg::CUpdateDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CUpdateDlg::IDD, pParent)
	, m_pApp(NULL)

	, szInstallDir(_T(""))
	, nLastLogIndex(0)
{
	m_pApp = (CGenMAPPDBDLApp*)AfxGetApp();
}

CUpdateDlg::~CUpdateDlg()
{
}

void CUpdateDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LSTUPDATELOG, lstUpdateLog);
	DDX_Control(pDX, IDC_LBLUPDATELOG, m_labUpdateLog);
	DDX_Control(pDX, IDC_GENMAPPSPLASH, m_picGenMAPPSplash);
	DDX_Control(pDX, IDC_LABUPDATEPRG, m_labUpdatePrg);
	DDX_Control(pDX, IDC_GRPUPDATEINFO, m_grpUpdateInfo);
	DDX_Control(pDX, IDC_LABUPDATESUM, m_labUpdateSummary);
	DDX_Control(pDX, IDC_OVERALLPRG, prgUpdateProgress);
}


BEGIN_MESSAGE_MAP(CUpdateDlg, CDialog)
	ON_BN_CLICKED(IDOK, OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, OnBnClickedCancel)
	ON_BN_CLICKED(IDC_BUTTON1, OnBnClickedButton1)
	ON_WM_SIZE()
END_MESSAGE_MAP()


// CUpdateDlg message handlers

void CUpdateDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
	OnOK();
}

void CUpdateDlg::OnBnClickedCancel()
{
	// TODO: Add your control notification handler code here
	OnCancel();
}

BOOL CUpdateDlg::OnInitDialog()
{
	CDialog::OnInitDialog();
	CGenMAPPDBDLApp* pApp = (CGenMAPPDBDLApp*)AfxGetApp();

	SetIcon(LoadIcon(pApp->m_hInstance, MAKEINTRESOURCE(IDI_GENMAPP)), TRUE);

	DoScaling();
	CreateUpdateLog();
	prgUpdateProgress.SetRange(0, 3);

	::SetForegroundWindow(this->GetSafeHwnd());

	CString	szConfirmMessage;
	szConfirmMessage.Format("This operation will update your installation of GenMAPP 2, located at %s. Update requires an Internet connection and will download any new GenMAPP program and data files to your computer. Do you wish to continue?", pApp->m_szDLLPath);
	if (AfxMessageBox(szConfirmMessage, MB_YESNO | MB_APPLMODAL) == IDNO)
	{
		EndDialog(0);
		return TRUE;
	}

	AfxBeginThread(::BeginUpdate, this);

	UpdateData(FALSE);
	UpdateWindow();

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}
void CUpdateDlg::OnBnClickedButton1()
{
	CConfigFile	cf("");
//	cf.TestWriteInTheMiddle("Here is the test line");
	cf.WriteStringKey("LookForMe", "Something else");//, "[TestSect]");
}

// Creates the Update Log listview control
void CUpdateDlg::CreateUpdateLog(void)
{
	LVCOLUMN iconcolumn;

	// Create image list of folder icons
	UpdateLogImageList.Create(16, 16, ILC_COLOR8, 0, 4);

	// Add folder icons.
	UpdateLogImageList.Add(AfxGetApp()->LoadIcon(IDI_CHECKMARK));
	UpdateLogImageList.Add(AfxGetApp()->LoadIcon(IDI_RIGHTARROW));
	UpdateLogImageList.Add(AfxGetApp()->LoadIcon(IDI_REDX));
	lstUpdateLog.SetImageList(&UpdateLogImageList, LVSIL_SMALL);

	iconcolumn.pszText = "";
	iconcolumn.cchTextMax = 1;
	iconcolumn.cx = SCALEX(22);
	iconcolumn.fmt = LVCFMT_CENTER;
	iconcolumn.iImage = 0;
	iconcolumn.iOrder = 0;
	iconcolumn.iSubItem = 0;
	iconcolumn.mask = LVCF_TEXT| LVCF_WIDTH| LVCF_SUBITEM | LVCF_FMT;

	lstUpdateLog.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

	lstUpdateLog.InsertColumn(0, &iconcolumn);
	lstUpdateLog.InsertColumn(1, NULL, LVCFMT_LEFT, SCALEX(355), 1);

}

UINT BeginUpdate( LPVOID pParam )
{
	CGenMAPPDBDLApp* pApp = (CGenMAPPDBDLApp*)AfxGetApp();
	CUpdateDlg* pUpdateDlg = (CUpdateDlg*)pParam;
	bool	bMoreKeys = true;
	int		i;
	char	szTempDir[_MAX_PATH];
	CString	szKey = "", szValue = "", WhatDir;
	GetTempPath(_MAX_PATH, szTempDir);
	CString	szUpdateFilePath;
	CStringList	KeysForBatchFile;

	//CString	szConfirmMessage;
	//szConfirmMessage.Format("This operation will update your installation of GenMAPP 2, located at %s. Update requires an Internet connection and will download any new GenMAPP program and data files to your computer. Do you wish to continue?", pApp->m_szDLLPath);
	//if (AfxMessageBox(szConfirmMessage, MB_YESNO | MB_APPLMODAL) == IDNO)
	//{
	//	//pUpdateDlg->EndDialog(0);
	//	pUpdateDlg->PostMessage(WM_CLOSE);
	//	return TRUE;
	//}

	szUpdateFilePath.Format("%s%s", szTempDir, "UpdateInfo.cfg");

	pUpdateDlg->AddItemToUpdateLog("Checking for Updates");
	pUpdateDlg->prgUpdateProgress.SetPos(1);

	if (!pApp->m_bUseHTTP)
	{
		if (!ConnectToServer())
		{
			AfxMessageBox("Unable to connect to GenMAPP Update server. Please try again later. If this continues to be a problem, please contact GenMAPP support at support@genmapp.org");
			return 1;
		}

		pApp->pGenMAPPServer->m_pConnect->SetCurrentDirectory("UpdateData");
		pApp->pGenMAPPServer->m_pConnect->GetCurrentDirectory(WhatDir);
		pApp->pGenMAPPServer->m_pConnect->GetFile("UpdateInfo.cfg", szUpdateFilePath, FALSE);
	}
/*	else
	{
		//pApp->pGenMAPPHTTPServer = pApp->CreateHTTPControlConnection();

		if (!pApp->pGenMAPPHTTPServer->GetHttpFile("/gmu/UpdateInfo.cfg", szUpdateFilePath))
			return 1000;
	}

*/
	PopulateInUseList(pParam, szTempDir);

	CConfigFile cf(szTempDir, "UpdateInfo.cfg");

	for (i=1;bMoreKeys;i++)
	{
		bMoreKeys = cf.RetrieveNextKey(i, &szKey, &szValue, "[Files]");
		if (szValue != "")
		{
			//MessageBox(NULL, szKey, szValue, 0);
			CFile	LocalFile;
			CString	szWindowTitle, szFTPFileName;
			CFileStatus LocalFileStatus;

			TRACE("FTPFile- Year: %d, Month: %d, Day: %d\n", atoi(szValue.Left(4)), atoi(szValue.Mid(4, 2)), atoi(szValue.Mid(6, 2)));
			CTime	FTPFileTime(atoi(szValue.Left(4)), atoi(szValue.Mid(4, 2)), atoi(szValue.Mid(6, 2)),
				0, 0, 0);

			HANDLE	hLocalFile = CreateFile(szKey, 0, FILE_SHARE_READ, NULL, OPEN_EXISTING,
				FILE_ATTRIBUTE_NORMAL, NULL);

			// File exists. Is older?
			if (hLocalFile != INVALID_HANDLE_VALUE)
			{
				CloseHandle(hLocalFile);
				LocalFile.Open(szKey, CFile::modeRead | CFile::shareDenyNone);


				LocalFile.GetStatus(LocalFileStatus);
				CString szLocalFileTime;
				szLocalFileTime.Format("LocalFile %s - Year: %d, Month: %d, Day: %d\n", szKey, LocalFileStatus.m_mtime.GetYear(), LocalFileStatus.m_mtime.GetMonth(), LocalFileStatus.m_mtime.GetDay());
				TRACE("LocalFile %s - Year: %d, Month: %d, Day: %d\n", szKey, LocalFileStatus.m_mtime.GetYear(), LocalFileStatus.m_mtime.GetMonth(), LocalFileStatus.m_mtime.GetDay());
				LocalFile.Close();
				if (LocalFileStatus.m_mtime >= FTPFileTime) 
					continue;  // Not older. Skip it.
			}
			else
				CloseHandle(hLocalFile);


			szFTPFileName = szValue.Right(szValue.GetLength() - szValue.ReverseFind('|') - 1) + "/" + 
				szKey.Right(szKey.GetLength() - szKey.ReverseFind('\\') - 1);

			pUpdateDlg->UpdatePreviousLogItem(true, pUpdateDlg->nLastLogIndex);
			pUpdateDlg->AddItemToUpdateLog("Updating file: " + szFTPFileName.Right(szFTPFileName.GetLength() - szFTPFileName.ReverseFind('/') - 1));
			pUpdateDlg->prgUpdateProgress.SetPos(2);

			if (pUpdateDlg->CouldBeInUseList.Lookup(szKey, szWindowTitle))
			{
				KeysForBatchFile.AddHead(szKey);
				szKey += ".new";
			}

			TRACE("FTPPath: %s, LocalPath: %s\n", szFTPFileName, szKey);
			if (!pApp->m_bUseHTTP)
				pApp->pGenMAPPServer->m_pConnect->GetFile(szFTPFileName, szKey, FALSE);
			else
				pApp->pGenMAPPHTTPServer->GetHttpFile("/gmu/" + szFTPFileName, szKey);

			LocalFile.Open(szKey, CFile::modeReadWrite | CFile::shareExclusive, NULL);
			LocalFile.GetStatus(LocalFileStatus);
			LocalFile.Close();
			LocalFileStatus.m_mtime = FTPFileTime;
			LocalFile.SetStatus(szKey, LocalFileStatus);
		}
	}

	if (!KeysForBatchFile.IsEmpty())
	{
		BuildAndRunBatchFile(pParam, &KeysForBatchFile);
		pUpdateDlg->m_labUpdateSummary.SetWindowText("Update process complete. Update will take effect the next time GenMAPP is run.");
	}
	else
		pUpdateDlg->m_labUpdateSummary.SetWindowText("Update process complete.");

	pUpdateDlg->UpdatePreviousLogItem(true, pUpdateDlg->nLastLogIndex);

	pUpdateDlg->prgUpdateProgress.SetPos(3);
	if (pUpdateDlg->nLastLogIndex > 0)
		pUpdateDlg->AddItemToUpdateLog("Update Complete");
	else
		pUpdateDlg->AddItemToUpdateLog("No Updates are Available at this Time");

	pUpdateDlg->UpdatePreviousLogItem(true, pUpdateDlg->nLastLogIndex);

	DeleteFile(szUpdateFilePath);

	return 1;
}

UINT CheckForUpdates( LPVOID pParam )
{
	CGenMAPPDBDLApp* pApp = (CGenMAPPDBDLApp*)AfxGetApp();
	bool	bMoreKeys = true;
	int		i;
	char	szTempDir[_MAX_PATH];
	CString	szKey = "", szValue = "", WhatDir;
	GetTempPath(_MAX_PATH, szTempDir);
	CString	szUpdateFilePath;

	szUpdateFilePath.Format("%s%s", szTempDir, "UpdateInfo.cfg");

	if (!ConnectToServer())
	{
		AfxMessageBox("Unable to connect to GenMAPP Update server. Please try again later. If this continues to be a problem, please contact GenMAPP support at support@genmapp.org");
		return 0;
	}

	if (!pApp->m_bUseHTTP)
	{
		pApp->pGenMAPPServer->m_pConnect->SetCurrentDirectory("UpdateData");
		pApp->pGenMAPPServer->m_pConnect->GetFile("UpdateInfo.cfg", szUpdateFilePath, FALSE);
	}
	else
	{
		//pApp->pGenMAPPHTTPServer = pApp->CreateHTTPControlConnection();

		if (!pApp->pGenMAPPHTTPServer->GetHttpFile("/gmu/UpdateInfo.cfg", szUpdateFilePath))
			return 1000;
	}

	CConfigFile cf(szTempDir, "UpdateInfo.cfg");

	for (i=1;bMoreKeys;i++)
	{
		bMoreKeys = cf.RetrieveNextKey(i, &szKey, &szValue, "[Files]");
		if (szValue != "")
		{
			CFile	LocalFile;
			CString	szWindowTitle, szFTPFileName;
			CFileStatus LocalFileStatus;

			TRACE("FTPFile- Year: %d, Month: %d, Day: %d\n", atoi(szValue.Left(4)), atoi(szValue.Mid(4, 2)), atoi(szValue.Mid(6, 2)));
			CTime	FTPFileTime(atoi(szValue.Left(4)), atoi(szValue.Mid(4, 2)), atoi(szValue.Mid(6, 2)),
				0, 0, 0);

			HANDLE	hLocalFile = CreateFile(szKey, 0, FILE_SHARE_READ, NULL, OPEN_EXISTING,
				FILE_ATTRIBUTE_NORMAL, NULL);

			// File exists. Is older?
			if (hLocalFile != INVALID_HANDLE_VALUE)
			{
				CloseHandle(hLocalFile);
				LocalFile.Open(szKey, CFile::modeRead | CFile::shareDenyNone);


				LocalFile.GetStatus(LocalFileStatus);
				CString szLocalFileTime;
				szLocalFileTime.Format("LocalFile %s - Year: %d, Month: %d, Day: %d\n", szKey, LocalFileStatus.m_mtime.GetYear(), LocalFileStatus.m_mtime.GetMonth(), LocalFileStatus.m_mtime.GetDay());
				TRACE("LocalFile %s - Year: %d, Month: %d, Day: %d\n", szKey, LocalFileStatus.m_mtime.GetYear(), LocalFileStatus.m_mtime.GetMonth(), LocalFileStatus.m_mtime.GetDay());
				LocalFile.Close();
				if (LocalFileStatus.m_mtime >= FTPFileTime) 
					continue;  // Not older. Skip it.
			}
			else
				CloseHandle(hLocalFile);

			CPromptForUpdate pfu;
			int nReturnCode = (int)pfu.DoModal();

			if (nReturnCode == IDOK)
				return 1;
			else
				return 0;
			//if (AfxMessageBox("There is an update to GenMAPP available. Would you like to download and apply the update now?", MB_YESNO) == IDYES)
			//	return 1;
			//else
			//	return 0;
		}
	}

	return 0;
}

bool ConnectToServer()
{
	CGenMAPPDBDLApp* pApp = (CGenMAPPDBDLApp*)AfxGetApp();
	POSITION pos;
	bool bConnected = false;
	
	// Populate Update server addresses
	// Make servers are added FTP then HTTP
	// Speed test folders are used as the Update folders
	CGMFTP* pSeed = new CGMFTP("GMDBDL");
	pSeed->SetLoginData("root.genmapp.org", "downloader", "fun4downloader", "", "");
	pSeed->szSpeedTestFolder = "UpdateData";
	pSeed->szHTTPSpeedTestFolder = "/gmu/";
	pApp->UpdateServers.AddHead(pSeed);

	pSeed = new CGMFTP("GMDBDL");
	pSeed->SetLoginData("root2.genmapp.org", "downloader", "fun4downloader", "", "");
	pSeed->szSpeedTestFolder = "UpdateData";
	pSeed->szHTTPSpeedTestFolder = "/gmu/";
	pApp->UpdateServers.AddTail(pSeed);


	for( pos = pApp->UpdateServers.GetHeadPosition(); pos != NULL && bConnected == false; )
	{
		pApp->pGenMAPPServer = (CGMFTP*)pApp->UpdateServers.GetNext( pos );
		if (!pApp->m_bUseHTTP)
		{
			pApp->pGenMAPPServer->SetOption(INTERNET_OPTION_CONNECT_TIMEOUT, 1000);
			pApp->pGenMAPPServer->SetOption(INTERNET_OPTION_CONNECT_RETRIES, 0);
			
			DWORD dwTimeoutVal = 0;
			pApp->pGenMAPPServer->QueryOption(INTERNET_OPTION_CONNECT_TIMEOUT, dwTimeoutVal);
			TRACE("Timeout value: %d", dwTimeoutVal);
			pApp->pGenMAPPServer->QueryOption(INTERNET_OPTION_CONNECT_RETRIES, dwTimeoutVal);
			TRACE(", retries value: %d\n", dwTimeoutVal);

			if (pApp->pGenMAPPServer->ConnectToGMServer())
				pApp->bAbortDownload = false;
		}

		if (!pApp->pGenMAPPServer->bConnected)
		{
			pApp->HTTPPos = pos;
			pApp->pGenMAPPHTTPServer = pApp->CreateHTTPControlConnection();
			if (pApp->pGenMAPPHTTPServer->ConnectToGMServer())
			{
				pApp->m_nInetConnected = CONNECT_FAILED;
				pApp->bAbortDownload = false;
			}
			else
			{
				if (pApp->pGenMAPPHTTPServer->TestHTTPConnection())
				{
					if (!pApp->m_bUseHTTP)
					{
						CConfigFile cf(pApp->m_szDLLPath);
						cf.WriteStringKey("ForceHTTP", "True", "[Downloader]");

						pApp->m_bUseHTTP = true;
					}
				}
			}
		}

		if (pApp->m_bUseHTTP)
			bConnected = pApp->pGenMAPPHTTPServer->bConnected;
		else
			bConnected = pApp->pGenMAPPServer->bConnected;
	}

	if (!bConnected)
	{
		pApp->m_nInetConnected = CONNECT_FAILED;
		return false;
	}

	return true;





/*
	if (!pApp->m_bUseHTTP)
	{
		pApp->pGenMAPPServer = new CGMFTP("GMDBDL");
		pApp->pGenMAPPServer->SetOption(INTERNET_OPTION_CONNECT_TIMEOUT, 1000);
		pApp->pGenMAPPServer->SetOption(INTERNET_OPTION_CONNECT_RETRIES, 0);
		
		pApp->pGenMAPPServer->SetLoginData("root.genmapp.org", "gmdownloader", "fun4downloader", "", "");
		pApp->pGenMAPPServer->ConnectToGMServer();
	}

	if (pApp->pGenMAPPServer == NULL || !pApp->pGenMAPPServer->bConnected)
	{
		if (pApp->pGenMAPPServer != NULL)
			delete pApp->pGenMAPPServer;

		pApp->pGenMAPPHTTPServer = pApp->CreateHTTPControlConnection();
		if (pApp->pGenMAPPHTTPServer->ConnectToGMServer())
		{
			pApp->m_nInetConnected = CONNECT_FAILED;
			pApp->bAbortDownload = false;
			return 1;
		}
		else
		{
			delete pApp->pGenMAPPHTTPServer;
			if (!pApp->m_bUseHTTP)
			{
				CConfigFile cf(pApp->m_szDLLPath);
				cf.WriteStringKey("ForceHTTP", "True", "[Downloader]");

				pApp->m_bUseHTTP = true;
			}
		}
	}
*/
}

// Performs scaling of the dialog controls based on font scaling factor
bool CUpdateDlg::DoScaling(void)
{
	CRect ListRect, ListLabelRect, SplashRect, DialogRect, UpdatePrgRect, UpdateInfRect, UpdateSum;
	GetWindowRect(&DialogRect);
	if (lstUpdateLog.m_hWnd == NULL)
		return false;

	lstUpdateLog.GetWindowRect(&ListRect);
	m_labUpdateLog.GetWindowRect(&ListLabelRect);
	m_picGenMAPPSplash.GetWindowRect(SplashRect);
	m_labUpdatePrg.GetWindowRect(&UpdatePrgRect);
	m_grpUpdateInfo.GetWindowRect(&UpdateInfRect);
	m_labUpdateSummary.GetWindowRect(&UpdateSum);

	ListLabelRect.bottom = 7 + ListLabelRect.bottom - ListLabelRect.top;
	ListLabelRect.top = 7;
	ListLabelRect.right = 7 + ListLabelRect.right - ListLabelRect.left;
	ListLabelRect.left = 7;
	m_labUpdateLog.MoveWindow(&ListLabelRect);

	SplashRect.bottom = SplashRect.bottom - SplashRect.top;
	SplashRect.top = 0;
	SplashRect.left = (DialogRect.right - DialogRect.left) - (SplashRect.right - SplashRect.left) - (SCALEX(3));
	SplashRect.right = DialogRect.right - DialogRect.left - (SCALEX(3));
	m_picGenMAPPSplash.MoveWindow(&SplashRect);

	ListRect.top = ListLabelRect.bottom + SCALEY(2);
	ListRect.left = ListLabelRect.left;
	ListRect.right = SplashRect.left - SCALEX(8);
	ListRect.bottom = (UpdatePrgRect.top - DialogRect.top) - ListRect.top - (SCALEY(14));
	lstUpdateLog.MoveWindow(&ListRect);

	UpdateInfRect.top = SplashRect.bottom + SCALEY(4);
	UpdateInfRect.right = DialogRect.right - DialogRect.left - (SCALEX(14));
	UpdateInfRect.bottom = ListRect.bottom;
	UpdateInfRect.left = SplashRect.left;
	m_grpUpdateInfo.MoveWindow(&UpdateInfRect);

	UpdateSum.bottom = (UpdateSum.bottom - UpdateSum.top) + SCALEY(30) + UpdateInfRect.top;
	UpdateSum.top = UpdateInfRect.top + SCALEY(30);
	UpdateSum.right = UpdateInfRect.right - (SCALEX(4));
	UpdateSum.left = UpdateInfRect.left + (SCALEX(22));
	m_labUpdateSummary.MoveWindow(&UpdateSum);

	return true;
}

void PopulateInUseList( LPVOID pParam, char* szTempDir )
{
	int i;
	BOOL bMoreKeys = true;
	CString szKey = "", szValue = "";
	CUpdateDlg* pUpdateDlg = (CUpdateDlg*)pParam;
	if (!pUpdateDlg->CouldBeInUseList.IsEmpty())
		pUpdateDlg->CouldBeInUseList.RemoveAll();

	CConfigFile cf(szTempDir, "UpdateInfo.cfg");

	for (i=1;bMoreKeys;i++)
	{
		bMoreKeys = cf.RetrieveNextKey(i, &szKey, &szValue, "[CouldBeInUse]");
		if (szValue != "")
			pUpdateDlg->CouldBeInUseList.SetAt(szKey, szValue);
	}
}

void BuildAndRunBatchFile(LPVOID pParam, CStringList* FilesToWaitOn)
{
	CUpdateDlg* pUpdateDlg = (CUpdateDlg*)pParam;
	POSITION	pos;
	CStdioFile	BatchFile(pUpdateDlg->szInstallDir + "~GMLastUpdateStep.bat", CFile::modeCreate |
		CFile::modeReadWrite | CFile::shareExclusive);

	for( pos = FilesToWaitOn->GetHeadPosition(); pos != NULL; )
	{
		int		nTokenPos = 0;
		CString szWindowTitleCombined, szWindowTitleParsed, 
			szWaitFile = FilesToWaitOn->GetNext(pos), szPID;

		szPID.Format("%u", GetCurrentProcessId());

//		pUpdateDlg->CouldBeInUseList.Lookup(szWaitFile, szWindowTitleCombined);
//		szWindowTitleParsed = szWindowTitleCombined.Tokenize("|", nTokenPos);
//		while (szWindowTitleParsed != "")
//		{
			BatchFile.WriteString("call \"" + pUpdateDlg->szInstallDir + "GenMAPPWindowWatcher.exe\" " + szPID + "\r\n");
//			szWindowTitleParsed = szWindowTitleCombined.Tokenize("|", nTokenPos);
//		}
	}

	for( pos = FilesToWaitOn->GetHeadPosition(); pos != NULL; )
	{
		CString szWindowTitleCombined, szWindowTitleParsed, 
			szWaitFile = FilesToWaitOn->GetNext(pos);
		
		BatchFile.WriteString("del \"" + szWaitFile + "\"\r\n");
		BatchFile.WriteString("ren \"" + szWaitFile + ".new\"" + " \"" + szWaitFile.Right(szWaitFile.GetLength() - szWaitFile.ReverseFind('\\') - 1) + "\"\r\n");
	}

	BatchFile.WriteString("del \"" + pUpdateDlg->szInstallDir + "~GMLastUpdateStep.bat\"");
	BatchFile.Flush();
	BatchFile.Close();

	ShellExecute(NULL, "open", pUpdateDlg->szInstallDir + "~GMLastUpdateStep.bat", NULL, pUpdateDlg->szInstallDir, SW_HIDE);
}

// Add a string to the Update Log. Returns the index for updating later.
int CUpdateDlg::AddItemToUpdateLog(CString szItem)
{
	LVITEM  itemdata;

	itemdata.mask =  LVIF_IMAGE | LVIF_STATE;
	itemdata.iImage = 1;
	itemdata.state  = 0;
	itemdata.stateMask = 0;
	itemdata.iSubItem = 0;
	itemdata.iItem = lstUpdateLog.GetItemCount();

	nLastLogIndex = lstUpdateLog.InsertItem(&itemdata);
	lstUpdateLog.SetItem(nLastLogIndex, 1, LVIF_TEXT, szItem, 0, 0, 0, 0, 0);

	return nLastLogIndex;
}

// Sets the icon of a previous Update Log item to a success or failure icon.
void CUpdateDlg::UpdatePreviousLogItem(bool bSuccess, int nIndex)
{
	LVITEM  itemdata;

	itemdata.mask =  LVIF_IMAGE | LVIF_STATE;
	itemdata.iImage = bSuccess ? 0 : 2;
	itemdata.state  = 0;
	itemdata.stateMask = 0;
	itemdata.iSubItem = 0;
	itemdata.iItem = nIndex;

	lstUpdateLog.SetItem(&itemdata);
}

void CUpdateDlg::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	DoScaling();

}
