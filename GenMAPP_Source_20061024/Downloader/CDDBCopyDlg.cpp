// CDDBCopyDlg.cpp : implementation file
//
#include "stdafx.h"
#include <afxinet.h>
#include "ConfigFile.h"
#include "GMFTP.h"
#include "GenMAPPDBDL.h"
#include "GenBankWarnDlg.h"
#include "AdvancedOptsDlg.h"
#include "CDDBCopyDlg.h"
#include "DBFile.h"
#include ".\cddbcopydlg.h"

#define SCALEX(argX) ((int) ((argX) * m_pApp->scaleX))
#define SCALEY(argY) ((int) ((argY) * m_pApp->scaleY))


// CCDDBCopyDlg dialog

IMPLEMENT_DYNAMIC(CCDDBCopyDlg, CDialog)
CCDDBCopyDlg::CCDDBCopyDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CCDDBCopyDlg::IDD, pParent)
	, m_DotCount(0)
	, m_labInetStat(_T(""))
	, m_pApp(NULL)
	, m_txtGDB(_T(""))
	, m_txtMAPP(_T(""))
	, m_txtGEX(_T(""))
	, szConnSpeed(_T(""))
	, m_nLastListCount(0)
	, m_bFirstActivate(false)
	, m_txtOtr(_T(""))
	, m_bOKToSize(false)
	, nMinDialogHeight(0)
	, nMinDialogWidth(0)
	, nOrigSelSumHeight(0)
	, m_bCapturingMouse(FALSE)
	, m_bInResizeZone(FALSE)
{
	m_pApp = (CGenMAPPDBDLApp*)AfxGetApp();
}

CCDDBCopyDlg::~CCDDBCopyDlg()
{
}

void CCDDBCopyDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_INETSTAT, m_labInetStat);
	DDX_Control(pDX, IDC_INETSTAT, m_ctlInetStat);
	DDX_Text(pDX, IDC_GDB, m_txtGDB);
	DDX_Text(pDX, IDC_MAPP, m_txtMAPP);
	DDX_Text(pDX, IDC_GEX, m_txtGEX);
	DDX_Control(pDX, IDOK, btnNext);
	DDX_Text(pDX, IDC_CONNSPEED, szConnSpeed);
	DDX_Control(pDX, IDC_SELCOUNT, SelectCount);
	DDX_Control(pDX, IDC_DLTIME, DLTime);
	DDX_Control(pDX, IDC_TOTALDLSIZE, m_labTotalDLSize);
	DDX_Control(pDX, IDCANCEL, btnBack);
	DDX_Control(pDX, IDC_DBDLLIST, DBList);
	DDX_Control(pDX, IDC_DBFOLDERTREE, DBFolderTree);
	DDX_Control(pDX, IDC_DLQUEUE, DLQueueList);
	DDX_Text(pDX, IDC_OTR, m_txtOtr);
	DDX_Control(pDX, IDC_SHOWREADME, m_btnShowReadMe);
	DDX_Control(pDX, IDC_READMETEXT, m_labReadMeText);
	DDX_Control(pDX, IDC_SELECTSUM, m_frmSelectSum);
	DDX_Control(pDX, IDC_CONNSTATFRM, m_frmConnStat);
	DDX_Control(pDX, IDC_REMOVEDB, m_btnRemoveDB);
	DDX_Control(pDX, IDC_STATUS_IND, m_icoStatusIndicator);
	DDX_Control(pDX, IDC_REFRESH, btnRefresh);
	DDX_Control(pDX, IDC_FREESPACELST, m_lstFreeSpace);
	DDX_Control(pDX, IDC_LBLTOPMESSAGE, lblTopMessage);
	DDX_Control(pDX, IDC_GRPDLDESTINATIONS, grpDLDestinations);
	DDX_Control(pDX, IDC_GRPOPTIONS, grpOptions);
	DDX_Control(pDX, IDC_BTNADVDLG, btnAdvanced);
	DDX_Control(pDX, IDC_CONNSPEED, lblConnSpeed);
	DDX_Control(pDX, IDC_LBLGENEDB, lblGeneDB);
	DDX_Control(pDX, IDC_GDB, txtGeneDB);
	DDX_Control(pDX, IDC_GENEDIR, btnGeneDBBrowse);
	DDX_Control(pDX, IDC_LBLEXPDATA, lblExpData);
	DDX_Control(pDX, IDC_GEX, txtExpData);
	DDX_Control(pDX, IDC_MAPPDIR, btnMappArchBrowse);
	DDX_Control(pDX, IDC_LBLMAPPARC, lblMAPPArch);
	DDX_Control(pDX, IDC_MAPP, txtMAPPArch);
	DDX_Control(pDX, IDC_GEXDIR, btnExpDataBrowse);
	DDX_Control(pDX, IDC_LBLOTRINFO, lblOtrInfo);
	DDX_Control(pDX, IDC_OTR, txtOtrInfo);
	DDX_Control(pDX, IDC_OTRDIR, btnOtrInfoBrowse);
	DDX_Control(pDX, IDC_LBLSELCOUNT, lblSelCountKey);
	DDX_Control(pDX, IDC_LBLESTDLTIME, lblDLTimeKey);
	DDX_Control(pDX, IDC_LBLDISKSPACE, lblReqDiskSpaceKey);
	DDX_Control(pDX, IDC_LBLFREESPACE, lblFreeSpaceKey);
	DDX_Control(pDX, IDC_LBLDLQUEUE, lblDLQueueKey);
	DDX_Control(pDX, IDC_LBLDLSIZEKEY, lblDLSizeKey);
	DDX_Control(pDX, IDC_LBLDLSIZEVAL, lblDLSizeVal);
}

afx_msg void CCDDBCopyDlg::OnTimer(UINT_PTR nIDEvent)
{
	UINT nTimerID = 123;

	switch (nIDEvent)
	{
		case 123: // Internet Connectivity UI Element
			switch (m_pApp->m_nInetConnected)
			{

				case NOT_CONNECTED:
				case CONNECTING:
					m_icoStatusIndicator.SetIcon( ::LoadIcon(m_pApp->m_hInstance, MAKEINTRESOURCE(IDI_YELLOWLIGHT)));
					m_icoStatusIndicator.SetWindowPos(&wndBottom, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
					if (m_DotCount < 4)
					{
						char	szElipsis[] = {"...\0"};
						
						m_labInetStat.Format("Testing for Internet Connectivity%s", &szElipsis[3 - m_DotCount]);
						UpdateData(false);
					}
					
					m_DotCount != 6 ? m_DotCount++ : m_DotCount = 0;
					break;

				case SPEED_TEST:
					m_icoStatusIndicator.SetIcon( ::LoadIcon(m_pApp->m_hInstance, MAKEINTRESOURCE(IDI_YELLOWLIGHT)));
					m_labInetStat = "Determining Internet connection speed.";
					UpdateData(false);
					break;

				case READ_DIR:
					if (m_DotCount < 4)
					{
						char	szElipsis[] = {"...\0"};
						
						m_labInetStat.Format("Retrieving data available for download from %s%s", m_pApp->szRetrievingFrom, &szElipsis[3 - m_DotCount]);
						UpdateData(false);
					}
					
					m_DotCount != 6 ? m_DotCount++ : m_DotCount = 0;
					break;

				case CONNECTED:
					m_icoStatusIndicator.SetIcon( ::LoadIcon(m_pApp->m_hInstance, MAKEINTRESOURCE(IDI_GREENLIGHT)));
					m_icoStatusIndicator.SetWindowPos(&wndBottom, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
					KillTimer((UINT_PTR)nTimerID);
					m_labInetStat = "Connected to GenMAPP server.";
					DBFolderTree.Expand(TVI_ROOT, TVE_COLLAPSE);
					UpdateData(false);
					break;

				case CONNECT_FAILED:
					m_icoStatusIndicator.SetIcon( ::LoadIcon(m_pApp->m_hInstance, MAKEINTRESOURCE(IDI_REDLIGHT)));
					m_icoStatusIndicator.SetWindowPos(&wndBottom, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
					KillTimer((UINT_PTR)nTimerID);
					m_labInetStat = "Unable to connect to GenMAPP server.";
					UpdateData(false);
					break;
			}

		case IDC_DBDLLIST:
			int		nItemCount = DBList.GetItemCount();
			POSITION pos;
			CDBFile* dbfile;

			if (m_nLastListCount != nItemCount)
			{
				UpdateData(false);
				m_nLastListCount = nItemCount;
			}

			if (m_pApp->pFileList != NULL)
			{
				DWORD	dwTotalSize = 0, dwLargestExe = 0, dwTotalDiskSpace = 0;
				int		nTotalSelected = 0;
				for( pos = m_pApp->pFileList->GetHeadPosition(); pos != NULL; )
				{
					dbfile = (CDBFile*)m_pApp->pFileList->GetNext( pos );
					if (dbfile->bSelected)
					{
						// Add to download size
						dwTotalSize += dbfile->dwFileSize;
						nTotalSelected++;

						// Calculate disk space requirements
						dwTotalDiskSpace += (dbfile->dwFileSize * 6);
						if (dbfile->szSrcFileName.Right(4).MakeUpper() == ".EXE")
						{
							// If we're deleting the self-extracting .exe after decompressing,
							// only count the largest of the SFXs in the disk space requirement.
							if (m_pApp->bDeleteSFX)
							{
								if (dbfile->dwFileSize > dwLargestExe)
									dwLargestExe = dbfile->dwFileSize;
							}
							else
								dwTotalDiskSpace += dbfile->dwFileSize;
						}			
					}

				}

				if (m_pApp->bDeleteSFX)
					dwTotalDiskSpace += dwLargestExe;

				m_pApp->nSelectedFileCount = nTotalSelected;
				CString szStatic;

				szStatic.Format("%d", nTotalSelected);
				SelectCount.SetWindowText(szStatic);

				if (dwTotalSize < 1073741824)
					szStatic.Format("%.2f MB", (double)dwTotalSize / 1048576);
				else
					szStatic.Format("%.2f GB", (double)dwTotalSize / 1073741824);

				lblDLSizeVal.SetWindowText(szStatic);

				// Set disk space label
				if (dwTotalDiskSpace < 1073741824)
					szStatic.Format("%.2f MB", (double)dwTotalDiskSpace / 1048576);
				else
					szStatic.Format("%.2f GB", (double)dwTotalDiskSpace / 1073741824);

				m_labTotalDLSize.SetWindowText(szStatic);

				// Total # of bytes / speed ( bytes * 1024) = num of seconds
                
				DWORD dwNumberOfSeconds;
				UINT nTotalDays, nTotalHours, nTotalMinutes, nTotalSeconds;
				dwNumberOfSeconds = (DWORD)(dwTotalSize / 1024) / (double)m_pApp->nSpeed;

				nTotalDays = dwNumberOfSeconds / 86400;
				nTotalHours = (dwNumberOfSeconds / 3600) % 24;
				nTotalMinutes = (dwNumberOfSeconds / 60) % 60;
				nTotalSeconds = dwNumberOfSeconds % 60;

				if (nTotalDays)
					szStatic.Format("%d Day%s %002i:%002i:%002i", nTotalDays, nTotalDays > 1 ? "s" : "",
						nTotalHours, nTotalMinutes, nTotalSeconds);
				else
					szStatic.Format("%002i:%002i:%002i", nTotalHours, nTotalMinutes, nTotalSeconds);

				DLTime.SetWindowText(szStatic);
			}
	}
}
BOOL CCDDBCopyDlg::OnInitDialog()
{
	UINT nTimerID = 123;
	UpdateWindow();

//	SetParent(FromHandle(m_pApp->m_hParentWnd));
//	::SetParent(this->GetSafeHwnd(), m_pApp->m_hParentWnd);

	::SetForegroundWindow(this->GetSafeHwnd());
//	SetActiveWindow();
//	BringWindowToTop();

	SetIcon(LoadIcon(m_pApp->m_hInstance, MAKEINTRESOURCE(IDI_GENMAPP)), TRUE);
	
	if (m_pApp->CalledFromInstaller)
	{
//		::EnableWindow(::FindWindowEx(this->GetSafeHwnd(), NULL, NULL, "Cancel"), false);
		::SetWindowText(::FindWindowEx(this->GetSafeHwnd(), NULL, NULL, "Start"), "Next");
		::SetWindowText(::FindWindowEx(this->GetSafeHwnd(), NULL, NULL, "Cancel"), "Back");
	}

	// Set Data Folder Paths
	m_txtGDB = m_pApp->szGDBPath;
	m_txtMAPP = m_pApp->szMAPPPath;
	m_txtGEX = m_pApp->szGEXPath;
	m_txtOtr = m_pApp->szOtrPath;

	UpdateData(false);

	// Fill in Free Space label
//	::SetWindowText(::FindWindowEx(this->GetSafeHwnd(), NULL, NULL, "  "), GetFreeSpaceString());
	while (m_lstFreeSpace.DeleteString(0) != LB_ERR);
	GetFreeSpaceString();

	if (!CreateDBListCtrl())
		return false;

	SizeDialogToScreen();
	m_bOKToSize = true;
	DoControlScaling();

	UpdateData(false);
	// Begin the Internet Connectivity process
	// Start the thread that:
	//		1.	Connects to the GenMAPP FTP server
	//		2.	Performs the speed test
	//		3.  Read the contents of the data directories and populate the list box
	SetTimer((UINT_PTR)nTimerID, 250, NULL);
	AfxBeginThread(::ConnectToGenMAPP, this);
	return true;
}

BOOL CCDDBCopyDlg::AddDatabasesToList(CString szDir, CString szExt)
{
	HANDLE hExists;
	CString szFullPath = "GenMAPP v2 Data\\" + szDir;
	
	//  Ensure the data directory exists
	hExists = CreateFile(szFullPath, 0, FILE_SHARE_READ, NULL, OPEN_EXISTING,
		FILE_ATTRIBUTE_NORMAL, NULL);

	if (hExists == INVALID_HANDLE_VALUE)
	{
		EndDialog(0);
		return false;
	}
	else
		CloseHandle(hExists);

	return true;
}

BEGIN_MESSAGE_MAP(CCDDBCopyDlg, CDialog)
    ON_WM_TIMER()
	ON_BN_CLICKED(IDC_GENEDIR, OnBnClickedGenedir)
	ON_BN_CLICKED(IDC_MAPPDIR, OnBnClickedMappdir)
	ON_BN_CLICKED(IDC_GEXDIR, OnBnClickedGexdir)
	ON_WM_PAINT()
	ON_BN_CLICKED(IDOK, OnBnClickedOk)
	ON_WM_SHOWWINDOW()
	ON_WM_ACTIVATE()
	ON_NOTIFY(NM_CUSTOMDRAW, IDC_DBDLLIST, OnNMCustomdrawDbdllist)
	ON_BN_CLICKED(IDCANCEL, OnBnClickedCancel)
	ON_NOTIFY(TVN_SELCHANGED, IDC_DBFOLDERTREE, OnTvnSelchangedDbfoldertree)
	ON_NOTIFY(NM_CLICK, IDC_DBDLLIST, OnNMClickDbdllist)
	ON_BN_CLICKED(IDC_OTRDIR, OnBnClickedOtrdir)
	ON_BN_CLICKED(IDC_SHOWREADME, OnBnClickedShowreadme)
	ON_BN_CLICKED(IDC_REMOVEDB, OnBnClickedRemovedb)
	ON_LBN_SELCHANGE(IDC_DLQUEUE, OnLbnSelchangeDlqueue)
	ON_BN_CLICKED(IDC_REFRESH, OnBnClickedRefresh)
	ON_BN_CLICKED(IDC_BTNADVDLG, OnBnClickedBtnadvdlg)
	ON_WM_SIZE()
	ON_WM_GETMINMAXINFO()
	ON_WM_LBUTTONDOWN()
	ON_WM_MOUSEMOVE()
	ON_WM_LBUTTONUP()
	ON_WM_SETCURSOR()
END_MESSAGE_MAP()

void CCDDBCopyDlg::OnBnClickedGenedir()
{
	CString*		szFolderName = new CString();
	m_pApp->SelectFolder(szFolderName);

	if (*szFolderName != "\\")
	{
		CConfigFile cf(m_pApp->m_szDLLPath);
		cf.WriteStringKey("mruGeneDB", *szFolderName);

		m_txtGDB = *szFolderName;
		UpdateData(FALSE);
	}

	delete szFolderName;
}

void CCDDBCopyDlg::OnBnClickedMappdir()
{
	CString*		szFolderName = new CString();
	m_pApp->SelectFolder(szFolderName);

	if (*szFolderName != "\\")
	{
		CConfigFile cf(m_pApp->m_szDLLPath);
		cf.WriteStringKey("mruMAPPPath", *szFolderName);

		m_txtMAPP = *szFolderName;
		UpdateData(FALSE);
	}

	delete szFolderName;
}

void CCDDBCopyDlg::OnBnClickedGexdir()
{
	CString*		szFolderName = new CString();
	m_pApp->SelectFolder(szFolderName);

	if (*szFolderName != "\\")
	{
		CConfigFile cf(m_pApp->m_szDLLPath);
		cf.WriteStringKey("mruDataSet", *szFolderName);

		m_txtGEX = *szFolderName;
		UpdateData(FALSE);
	}

	delete szFolderName;
}

void CCDDBCopyDlg::OnBnClickedOtrdir()
{
	CString*		szFolderName = new CString();
	m_pApp->SelectFolder(szFolderName);

	if (*szFolderName != "\\")
	{
		CConfigFile cf(m_pApp->m_szDLLPath);
		cf.WriteStringKey("mruOtherInfo", *szFolderName);

		m_txtOtr = *szFolderName;
		UpdateData(FALSE);
	}

	delete szFolderName;
}

//HBRUSH CCDDBCopyDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
//{
//	HBRUSH hbr = CDialog::OnCtlColor(pDC, pWnd, nCtlColor);

	//if (pWnd->m_hWnd == m_ctlInetStat.m_hWnd)
	//{
	//	switch (m_pApp->m_nInetConnected)
	//	{

	//		case CONNECT_FAILED:
	//		case NOT_CONNECTED:
	//			pDC->SetTextColor(RGB(255,0,0));
	//			break;

	//		case CONNECTING:
	//		case SPEED_TEST:
	//		case READ_DIR:
	//			pDC->SetTextColor(RGB(160,160,16 )); // need darker color
	//			break;

	//		case CONNECTED:
	//			pDC->SetTextColor(RGB(0,128,0));
	//			break;
	//	}
	//}

	//if (pWnd->m_hWnd == m_frmConnStat.m_hWnd)
	//	pDC->SetBkColor(RGB(255, 255, 255));

	// TODO:  Change any attributes of the DC here

	// TODO:  Return a different brush if the default is not desired
//	return hbr;
//}

void CCDDBCopyDlg::OnPaint()
{
	CPaintDC dc(this); // device context for painting
//	dc.SetBkColor(RGB(255, 0, 0));
	//// Show GenBank warning dialog box, if applicable.
	//if (!m_bFirstActivate)
	//{
	//	m_bFirstActivate = true;
	//	ShowGenBankWarning();
	//}
	// TODO: Add your message handler code here
	// Do not call CDialog::OnPaint() for painting messages
}


BOOL CCDDBCopyDlg::UpdateDlgControls(bool bSave)
{
	return UpdateData(bSave);
}

// Creates and sizes the two List View controls 
// that comprise the "Explorer."
bool CCDDBCopyDlg::CreateDBListCtrl(void)
{
	INITCOMMONCONTROLSEX ccex;
	ccex.dwICC = ICC_LISTVIEW_CLASSES;
	ccex.dwSize = sizeof(INITCOMMONCONTROLSEX);
	if (InitCommonControlsEx(&ccex) == false)
		return false;

	CRect ListRect, TreeRect, DBSelectRect, DialogRect;
	GetClientRect(&ListRect);
	DBFolderTree.GetClientRect(&TreeRect);
	m_frmSelectSum.GetWindowRect(&DBSelectRect);
	GetWindowRect(&DialogRect);

	ListRect.top += 51;
	ListRect.bottom = (DBSelectRect.top - DialogRect.top) - SCALEY(33);//ListRect.top + 145;
	ListRect.left = TreeRect.right + 5;//10;
	ListRect.right -= 10;

	TreeRect.top = ListRect.top;
	TreeRect.bottom = ListRect.bottom;
	TreeRect.left = 10;

	DBList.SetExtendedStyle(LVS_EX_FULLROWSELECT|LVS_EX_CHECKBOXES);

	DBList.MoveWindow(&ListRect);
	DBFolderTree.MoveWindow(&TreeRect);

	int x = SCALEX(20);
	// set up columns
	DBList.InsertColumn(0, NULL, LVCFMT_LEFT, SCALEX(21), -1);
	DBList.InsertColumn(1, "Database File Name", LVCFMT_LEFT, SCALEX(130), 1);
	DBList.InsertColumn(2, "Date", LVCFMT_LEFT, SCALEX(60), 2);
	DBList.InsertColumn(3, "Size", LVCFMT_LEFT, SCALEX(65), 4);
	DBList.InsertColumn(4, "Location", LVCFMT_LEFT, SCALEX(103), 4);

	// Create image list of folder icons
	DBFolderTreeIcons.Create(16, 16, ILC_COLOR8, 0, 4);

	// Add folder icons.
	DBFolderTreeIcons.Add(AfxGetApp()->LoadIcon(IDI_FOLDERCLOSED));
	DBFolderTreeIcons.Add(AfxGetApp()->LoadIcon(IDI_FOLDEROPEN));
	DBFolderTree.SetImageList(&DBFolderTreeIcons, TVSIL_NORMAL);

	return true;
}

void CCDDBCopyDlg::OnBnClickedOk()
{
	// Make sure paths in folder text boxes are valid. Only check if there's
	// actually something to download.
	if (DLQueueList.GetCount() > 0 && !AreGenMAPPFoldersValid())
		return;

	UINT nTimerID = IDC_DBDLLIST;
	KillTimer((UINT_PTR)nTimerID);

	if (m_pApp->m_nInetConnected != CONNECTED && m_pApp->m_nInetConnected != CONNECT_FAILED)
		m_pApp->bAbortDownload = true;

	DBFolderTree.EnableWindow(TRUE);

	OnOK();
}
// Shows the GenBank ID warning dialog box, if the user has not indicated he or she does not want to see it.
//bool CCDDBCopyDlg::ShowGenBankWarning(void)
//{
//	CConfigFile cf(m_pApp->m_szDLLPath);
//	if (cf.ReadStringKey("ShowGenBankWarning") != "False")
//	{
//		CGenBankWarnDlg gbWarn;
//		gbWarn.DoModal();
//	}
//	return true;
//}

void CCDDBCopyDlg::OnActivate(UINT nState, CWnd* pWndOther, BOOL bMinimized)
{
	CDialog::OnActivate(nState, pWndOther, bMinimized);
}

// Retrieve a string suitable for display that shows the
// free hard drive space for all of the fixed drives on 
// the system.
CString CCDDBCopyDlg::GetFreeSpaceString()
{
	char			szDrv[] = {"C:\\"};
	UINT			nDrvType = 0;
	CString			szFreeSpaceString = "";
	ULARGE_INTEGER	nFreeBytesAvailable;

	nFreeBytesAvailable.QuadPart = 0;
	for (szDrv[0]='C';szDrv[0]<='Z';szDrv[0]++)
	{
		nDrvType = GetDriveType(szDrv);
		if (nDrvType == DRIVE_FIXED)
		{
			GetDiskFreeSpaceEx(szDrv, &nFreeBytesAvailable, NULL, NULL);
			double nDisplayedSize = nFreeBytesAvailable.QuadPart > 1073741824 ? (double)nFreeBytesAvailable.QuadPart / 1073741824 :
				(double)nFreeBytesAvailable.QuadPart / 1048576;
			szFreeSpaceString.Format("%s has %.2f %s free", szDrv, 
				nFreeBytesAvailable.QuadPart > 1073741824 ? (double)nFreeBytesAvailable.QuadPart / 1073741824 : 
				(double)nFreeBytesAvailable.QuadPart / 1048576, nFreeBytesAvailable.QuadPart > 1073741824 ? "GB" : "MB");
			m_lstFreeSpace.InsertString(m_lstFreeSpace.GetCount(), szFreeSpaceString);
		}
	}
	return szFreeSpaceString;

}

void CCDDBCopyDlg::OnNMCustomdrawDbdllist(NMHDR *pNMHDR, LRESULT *pResult)
{
	*pResult = 0;
	return;
	LPNMCUSTOMDRAW pNMCD = reinterpret_cast<LPNMCUSTOMDRAW>(pNMHDR);
	LPNMLVCUSTOMDRAW lplvcd = reinterpret_cast<LPNMLVCUSTOMDRAW>(pNMHDR);
	// TODO: Add your control notification handler code here
	if (pNMCD->dwDrawStage == CDDS_PREPAINT)
	{
		if (pNMCD->dwItemSpec == 15)
		{
			*pResult = CDRF_NOTIFYITEMDRAW;
			return;
		}
	}

	if (pNMCD->dwDrawStage == CDDS_ITEMPREPAINT && pNMCD->dwItemSpec == 3)
	{
		lplvcd->clrText = RGB(255, 0, 0);
		*pResult = CDRF_NEWFONT;
		return;
	}

	*pResult = 0;
}

void CCDDBCopyDlg::OnBnClickedCancel()
{
	UINT nTimerID = IDC_DBDLLIST;
	KillTimer((UINT_PTR)nTimerID);

	if (m_pApp->m_nInetConnected != CONNECTED && m_pApp->m_nInetConnected != CONNECT_FAILED)
		m_pApp->bAbortDownload = true;

	DBFolderTree.EnableWindow(TRUE);
	OnCancel();
}

void CCDDBCopyDlg::OnTvnSelchangedDbfoldertree(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMTREEVIEW pNMTreeView = reinterpret_cast<LPNMTREEVIEW>(pNMHDR);
	
	if (!m_pApp->DBFolderList.IsEmpty())
	{
		if (pNMTreeView->itemNew.hItem == m_pApp->hGenBankFolder || 
			pNMTreeView->itemNew.hItem == m_pApp->hOtherSpeciesFolder)
		{
			DBList.ShowWindow(SW_HIDE);
			m_btnShowReadMe.ShowWindow(SW_SHOW);
			m_labReadMeText.ShowWindow(SW_SHOW);
		}
		else
		{
			POSITION	pos;
			int			nListIndex = 0;
			CDBFolder*	dbfolder;
			
			DBList.ShowWindow(SW_SHOW);
			m_btnShowReadMe.ShowWindow(SW_HIDE);
			m_labReadMeText.ShowWindow(SW_HIDE);
			DBList.DeleteAllItems();
			for( pos = m_pApp->pFileList->GetHeadPosition(); pos != NULL; )
			{
				CDBFile* dbfile;
				dbfile = (CDBFile*)m_pApp->pFileList->GetNext( pos );
				dbfile->nLVID = -1;
			}
			
			for( pos = m_pApp->DBFolderList.GetHeadPosition(); pos != NULL; )
			{
				dbfolder = (CDBFolder*)m_pApp->DBFolderList.GetNext( pos );
				if (dbfolder->hFolder == pNMTreeView->itemNew.hItem)
					break;
			}

			if (dbfolder == NULL)
				return;

			for( pos = dbfolder->DBFileList.GetHeadPosition(); pos != NULL; )
			{
				CDBFile* dbfile;
				CString	szFileNameOnly;

				dbfile = (CDBFile*)dbfolder->DBFileList.GetNext( pos );
				nListIndex = DBList.InsertItem(DBList.GetItemCount(), NULL);
				dbfile->nLVID = nListIndex;
				DBList.SetCheck(nListIndex, dbfile->bSelected);
				DBList.SetItem(nListIndex, 1, LVIF_TEXT, dbfile->GetFileNameOnly(dbfile->szSrcFileName), 0, 0, 0, 0, 0);
				DBList.SetItem(nListIndex, 2, LVIF_TEXT, dbfile->DBDate.Format("%x"), 0, 0, 0, 0, 0);  // %x - date for locale
				DBList.SetItem(nListIndex, 3, LVIF_TEXT, dbfile->GetStringFileSize(dbfile->dwFileSize), 0, 0, 0, 0, 0);
				
				if (dbfile->DBServerList.GetCount() > 1)
					DBList.SetItem(nListIndex, 4, LVIF_TEXT, "< Multiple Locations >", 0, 0, 0, 0, 0);
				else
				{
					CGMFTP* pServer = (CGMFTP*)dbfile->DBServerList.GetHead();
					DBList.SetItem(nListIndex, 4, LVIF_TEXT, pServer->szLocation, 0, 0, 0, 0, 0);
				}
			}
		}
	}

	*pResult = 0;
}

void CCDDBCopyDlg::OnNMClickDbdllist(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMIA = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	int			nSelectedIndex = pNMIA->iItem;
	POSITION	pos;
	CDBFile* dbfile;
	
	if (m_pApp->pFileList != NULL)
	{
		for( pos = m_pApp->pFileList->GetHeadPosition(); pos != NULL; )
		{
			dbfile = (CDBFile*)m_pApp->pFileList->GetNext( pos );
			if (dbfile->nLVID == nSelectedIndex)
				break;
		}
	
		dbfile->bSelected = !dbfile->bSelected;
		// Only set the checkmark if the user clicked outside of check area
		// Windows will check automatically if the check area itself is clicked.
		TRACE("x= %d, y=%d\n", pNMIA->ptAction.x, pNMIA->ptAction.y);
		if (pNMIA->ptAction.x >= SCALEX(16))
            DBList.SetCheck(nSelectedIndex, dbfile->bSelected);

		UpdateDLQueueList();

	}
	
	*pResult = 0;
}

// Builds the Database Queue List when an item has been selected or removed
BOOL CCDDBCopyDlg::UpdateDLQueueList(void)
{
	if (m_pApp->pFileList != NULL)
	{
		POSITION	pos;
		CDBFile* dbfile;

		while (DLQueueList.DeleteString(0) != LB_ERR);
			
		for( pos = m_pApp->pFileList->GetHeadPosition(); pos != NULL; )
		{
			dbfile = (CDBFile*)m_pApp->pFileList->GetNext( pos );
			if (dbfile->bSelected)
				dbfile->nID = DLQueueList.InsertString(-1, dbfile->GetFileNameOnly(dbfile->szSrcFileName));
		}
	}
	return TRUE;
}

void CCDDBCopyDlg::OnBnClickedShowreadme()
{
	CString		szReadMeType = ".rtf";
	CString		szReadMeFileName = "";
	if (DBFolderTree.GetSelectedItem() == m_pApp->hOtherSpeciesFolder)
		szReadMeFileName = "Non-supportedSpeciesDBs";
	else
		szReadMeFileName = "ObtainingGBFiles";

	try
	{
		HKEY hKey = NULL;

		if (RegOpenKeyEx( HKEY_CLASSES_ROOT,
               ".pdf",
               0, KEY_QUERY_VALUE, &hKey ) == ERROR_SUCCESS)
		{
			szReadMeType = ".pdf";
			RegCloseKey(hKey);
		}
		
		if (m_pApp->pGenMAPPServer->ConnectToGMServer())
			return;

		if (!m_pApp->pGenMAPPServer->bHTTPServer)
		{
			if (!m_pApp->pGenMAPPServer->m_pConnect->SetCurrentDirectory("/home2/ServerSeed/"))
				return;
		}
		else
			return;

		// Make sure file exists on server. If looking for .pdf file and it
		// doesn't exist, try the .rtf. If neither exist, abort.
		CFtpFileFind pInetFile(m_pApp->pGenMAPPServer->m_pConnect);
		if (!pInetFile.FindFile(szReadMeFileName + szReadMeType, INTERNET_FLAG_EXISTING_CONNECT | INTERNET_FLAG_RELOAD))
		{
			if (szReadMeType == ".pdf")
			{
				pInetFile.Close();
				szReadMeType = ".rtf";
				if (!pInetFile.FindFile(szReadMeFileName + szReadMeType, INTERNET_FLAG_EXISTING_CONNECT | INTERNET_FLAG_RELOAD))
				{
					pInetFile.Close();
					return;
				}
			}
			else
			{
				pInetFile.Close();
				return;
			}
		}
		
		pInetFile.Close();

		CreateDirectory(m_pApp->szOtrPath, NULL);

		if (!m_pApp->pGenMAPPServer->m_pConnect->GetFile(szReadMeFileName + szReadMeType, m_pApp->szOtrPath + szReadMeFileName + szReadMeType, false))
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

		m_pApp->pGenMAPPServer->DisconnectFromGMServer();
	}
	catch (CInternetException* pEx)
	{
		TCHAR sz[1024];
		pEx->GetErrorMessage(sz, 1024);
		printf("ERROR!  %s\n", sz);
		pEx->Delete();
	}

	ShellExecute(NULL, "open", m_pApp->szOtrPath + szReadMeFileName + szReadMeType, NULL, m_pApp->szOtrPath, SW_SHOW);
}

void CCDDBCopyDlg::OnBnClickedRemovedb()
{
	int nSelection = DLQueueList.GetCurSel();
	POSITION	pos;
	CDBFile* dbfile;

	for( pos = m_pApp->pFileList->GetHeadPosition(); pos != NULL; )
	{
		dbfile = (CDBFile*)m_pApp->pFileList->GetNext( pos );
		if (dbfile->nID == nSelection)
		{
			HTREEITEM	hSelectedFolder = DBFolderTree.GetSelectedItem();
			
			dbfile->bSelected = FALSE;
			dbfile->nID = -1;
			if (dbfile->nLVID != -1)
				DBList.SetCheck(dbfile->nLVID, dbfile->bSelected);
			break;

		}
	}

	UpdateDLQueueList();
	m_btnRemoveDB.EnableWindow(FALSE);
}

void CCDDBCopyDlg::OnLbnSelchangeDlqueue()
{
	m_btnRemoveDB.EnableWindow((DLQueueList.GetCurSel() != LB_ERR));
}

void CCDDBCopyDlg::OnBnClickedRefresh()
{
	UINT nTimerID = 123;

	// Disable the Refresh button while refreshing
	btnRefresh.EnableWindow(FALSE);
	
	// Clear all lists
	DBList.DeleteAllItems();
	DBFolderTree.DeleteAllItems();
	while (DLQueueList.GetCount() != 0)
		DLQueueList.DeleteString(0);

	UpdateData(FALSE);

	// Free memory, including servers, files and folders
	m_pApp->FreeFileList();
	UpdateData(FALSE);

	// Start from scratch
	SetTimer((UINT_PTR)nTimerID, 250, NULL);
	AfxBeginThread(::ConnectToGenMAPP, this);
}
// Verifies that paths specifed in the four folder text boxes are valid. Sets
// the folder variables to the paths in the text boxes if valid.
bool CCDDBCopyDlg::AreGenMAPPFoldersValid(void)
{
	DWORD	dwResult;

	UpdateData(TRUE);
	dwResult = ::GetFileAttributes(m_txtGDB);
	if (dwResult == INVALID_FILE_ATTRIBUTES || !(dwResult | FILE_ATTRIBUTE_DIRECTORY))
	{
		AfxMessageBox("The Gene Databases folder specified in Download Destination Locations is invalid. Please click the \"Change\" button to select a new folder.");
		return false;
	}
	m_pApp->szGDBPath = m_txtGDB;

	dwResult = ::GetFileAttributes(m_txtGEX);
	if (dwResult == INVALID_FILE_ATTRIBUTES || !(dwResult | FILE_ATTRIBUTE_DIRECTORY))
	{
		AfxMessageBox("The Expression Datasets folder specified in Download Destination Locations is invalid. Please click the \"Change\" button to select a new folder.");
		return false;
	}
	m_pApp->szGEXPath = m_txtGEX;

	dwResult = ::GetFileAttributes(m_txtMAPP);
	if (dwResult == INVALID_FILE_ATTRIBUTES || !(dwResult | FILE_ATTRIBUTE_DIRECTORY))
	{
		AfxMessageBox("The MAPP Archives folder specified in Download Destination Locations is invalid. Please click the \"Change\" button to select a new folder.");
		return false;
	}
	m_pApp->szMAPPPath = m_txtMAPP;

	dwResult = ::GetFileAttributes(m_txtOtr);
	if (dwResult == INVALID_FILE_ATTRIBUTES || !(dwResult | FILE_ATTRIBUTE_DIRECTORY))
	{
		AfxMessageBox("The Other Information folder specified in Download Destination Locations is invalid. Please click the \"Change\" button to select a new folder.");
		return false;
	}
	m_pApp->szOtrPath = m_txtOtr;

	return true;

}

// Launch the Advanced Options dialog
void CCDDBCopyDlg::OnBnClickedBtnadvdlg()
{
   	CAdvancedOptsDlg AODlg;
	AODlg.DoModal();
}

// Adjusts the dialog size to fit the screen resoultion. Also stores the width and height when finished.
void CCDDBCopyDlg::SizeDialogToScreen(void)
{
	RECT DialogRect, SelectSumGrpRect;
	
	GetWindowRect(&DialogRect);
	nMinDialogWidth = DialogRect.right - DialogRect.left;
	nMinDialogHeight = DialogRect.bottom - DialogRect.top;

	m_frmSelectSum.GetClientRect(&SelectSumGrpRect);
	nOrigSelSumHeight = SelectSumGrpRect.bottom;

	if (m_pApp->nScrResX >= 1024 && m_pApp->nScrResY >= 768)
	{
		//1000 x 684
		DialogRect.left = (m_pApp->nScrResX / 2) - (1000 / 2);
		DialogRect.right = DialogRect.left + 1000;
		DialogRect.top = (m_pApp->nScrResY / 2) - (684 / 2);
		DialogRect.bottom = DialogRect.top + 684;
		MoveWindow(&DialogRect);
	}
}

// Code to scale the sizes and positions of the controls based on the dialog box size
void CCDDBCopyDlg::DoControlScaling(void)
{
	RECT DialogRect, MessageRect, MessageDlgRect, FolderTreeRect, FolderTreeDlgRect, DBListRect,
		ConnectionStatRect, DLDestinationsRect, SelectSumGrpRect, NextBtnRect, NextBtnDlgRect,
		BackBtnRect, BackBtnDlgRect, OptionsGrpDlgRect, AdvBtnRect, RefreshBtnRect, ConnStatGrpRect,
		ConnStatGrpDlgRect, ConnStatLblRect, StatIconRect, DLDestGrpRect, DLDestGrpDlgRect, GeneDBLblRect,
		GeneDBBrowseRect, GeneDBTxtRect, DialogWindowRect, SelCountKeyRect, DLSelSumDlgRect, SelCountValRect,
		FreeSpaceKeyRect, DLQueueLstDlgRect, ReadMeLblRect, ShowReadMeBtnRect, FileListDlgRect;

	int nNonAdjControlHeight;

	//int	nLabelWidths = 100, nFromTop = 20, nValuePaddingFromRight = 28;

	GetClientRect(&DialogRect);
	GetWindowRect(&DialogWindowRect);
	lblTopMessage.GetClientRect(&MessageRect);
	DBFolderTree.GetClientRect(&FolderTreeRect);
	DBList.GetClientRect(&DBListRect);
	m_frmConnStat.GetClientRect(&ConnectionStatRect);
	grpDLDestinations.GetClientRect(&DLDestinationsRect);
	m_frmSelectSum.GetClientRect(&SelectSumGrpRect);
	btnNext.GetClientRect(&NextBtnRect);
	btnBack.GetClientRect(&BackBtnRect);
	btnAdvanced.GetClientRect(&AdvBtnRect);
	btnRefresh.GetClientRect(&RefreshBtnRect);
	m_frmConnStat.GetClientRect(&ConnStatGrpRect);
	m_ctlInetStat.GetClientRect(&ConnStatLblRect);
	m_icoStatusIndicator.GetClientRect(&StatIconRect);
	grpDLDestinations.GetClientRect(&DLDestGrpRect);
	lblGeneDB.GetClientRect(&GeneDBLblRect);
	btnGeneDBBrowse.GetClientRect(&GeneDBBrowseRect);
	txtGeneDB.GetClientRect(&GeneDBTxtRect);
	lblSelCountKey.GetClientRect(&SelCountKeyRect);
	SelectCount.GetClientRect(&SelCountValRect);
	lblFreeSpaceKey.GetClientRect(&FreeSpaceKeyRect);
	m_labReadMeText.GetClientRect(&ReadMeLblRect);
	m_btnShowReadMe.GetClientRect(&ShowReadMeBtnRect);

	// OK and Cancel
	::MoveWindow(btnNext.GetSafeHwnd(), DialogRect.right - NextBtnRect.right - SCALEX(10), DialogRect.bottom - NextBtnRect.bottom - SCALEY(10), NextBtnRect.right, NextBtnRect.bottom, true);
	NextBtnDlgRect = AdjustToDlgCoordinates(&btnNext);
	::MoveWindow(btnBack.GetSafeHwnd(), NextBtnDlgRect.left - BackBtnRect.right - SCALEX(14), DialogRect.bottom - BackBtnRect.bottom - SCALEY(10), BackBtnRect.right, BackBtnRect.bottom, true);

	
	// Options group
	BackBtnDlgRect = AdjustToDlgCoordinates(&btnBack);
	::MoveWindow(grpOptions.GetSafeHwnd(), BackBtnDlgRect.left - (int)(AdvBtnRect.right * 2.4) - SCALEX(8), BackBtnDlgRect.bottom - (int)(AdvBtnRect.bottom * 2.5) - SCALEY(7), (int)(AdvBtnRect.right * 2.4), (int)(AdvBtnRect.bottom * 2.5) + SCALEY(4), true);
	OptionsGrpDlgRect = AdjustToDlgCoordinates(&grpOptions);
	::MoveWindow(btnRefresh.GetSafeHwnd(), OptionsGrpDlgRect.left + (int)(AdvBtnRect.right * .1), OptionsGrpDlgRect.top + SCALEY(20), RefreshBtnRect.right, RefreshBtnRect.bottom, true);
	::MoveWindow(btnAdvanced.GetSafeHwnd(), OptionsGrpDlgRect.left + AdvBtnRect.right + SCALEX(15), OptionsGrpDlgRect.top + SCALEY(20), AdvBtnRect.right, AdvBtnRect.bottom, true);

	
	// Internet Connection Status group
	::MoveWindow(m_frmConnStat.GetSafeHwnd(), SCALEX(10), BackBtnDlgRect.bottom - (int)(AdvBtnRect.bottom * 2.5) - SCALEY(7), OptionsGrpDlgRect.left - SCALEX(8) - SCALEX(10), OptionsGrpDlgRect.bottom - OptionsGrpDlgRect.top, true);
	ConnStatGrpDlgRect = AdjustToDlgCoordinates(&m_frmConnStat);
	::MoveWindow(m_ctlInetStat.GetSafeHwnd(), ConnStatGrpDlgRect.left + StatIconRect.right + SCALEX(3), ConnStatGrpDlgRect.top + SCALEY(15), ConnStatGrpDlgRect.right - ConnStatGrpDlgRect.left - StatIconRect.right - SCALEX(3) - SCALEX(10) - SCALEX(5), ConnStatLblRect.bottom, true);
	::MoveWindow(lblConnSpeed.GetSafeHwnd(), ConnStatGrpDlgRect.left + StatIconRect.right + SCALEX(3), ConnStatGrpDlgRect.top + SCALEY(15) + (int)(ConnStatLblRect.bottom * 1.5), ConnStatGrpDlgRect.right - ConnStatGrpDlgRect.left - StatIconRect.right - SCALEX(3) - SCALEX(10) - SCALEX(5), ConnStatLblRect.bottom, true);
	::MoveWindow(m_icoStatusIndicator.GetSafeHwnd(), ConnStatGrpDlgRect.left + SCALEX(2), ConnStatGrpDlgRect.top + SCALEY(15) + (int)(ConnStatLblRect.bottom * .5), StatIconRect.right, StatIconRect.bottom, true);


	// Download Destinations group
	::MoveWindow(grpDLDestinations.GetSafeHwnd(), SCALEX(10), ConnStatGrpDlgRect.top - DLDestGrpRect.bottom - SCALEY(7), DialogRect.right - DialogRect.left - SCALEX(20), DLDestGrpRect.bottom, true);
	DLDestGrpDlgRect = AdjustToDlgCoordinates(&grpDLDestinations);
	::MoveWindow(lblGeneDB.GetSafeHwnd(), DLDestGrpDlgRect.left + SCALEX(7), DLDestGrpDlgRect.top + SCALEY(15), GeneDBLblRect.right, GeneDBLblRect.bottom, true);
	::MoveWindow(btnGeneDBBrowse.GetSafeHwnd(), DLDestGrpDlgRect.right - DLDestGrpDlgRect.left - GeneDBBrowseRect.right - SCALEX(4), DLDestGrpDlgRect.top + SCALEY(15), GeneDBBrowseRect.right + SCALEX(2), GeneDBBrowseRect.bottom + SCALEX(2), true);
	::MoveWindow(txtGeneDB.GetSafeHwnd(), DLDestGrpDlgRect.left + SCALEX(7) + GeneDBLblRect.right + SCALEX(7), DLDestGrpDlgRect.top + SCALEY(15), DLDestGrpDlgRect.right - DLDestGrpDlgRect.left - GeneDBBrowseRect.right - GeneDBLblRect.right - SCALEX(21) - SCALEX(17), GeneDBTxtRect.bottom, true);
	::MoveWindow(lblExpData.GetSafeHwnd(), DLDestGrpDlgRect.left + SCALEX(7), DLDestGrpDlgRect.top + SCALEY(13) + GeneDBTxtRect.bottom, GeneDBLblRect.right, GeneDBLblRect.bottom, true);
	::MoveWindow(btnExpDataBrowse.GetSafeHwnd(), DLDestGrpDlgRect.right - DLDestGrpDlgRect.left - GeneDBBrowseRect.right - SCALEX(4), DLDestGrpDlgRect.top + SCALEY(13) + GeneDBTxtRect.bottom, GeneDBBrowseRect.right + SCALEX(2), GeneDBBrowseRect.bottom + SCALEX(2), true);
	::MoveWindow(txtExpData.GetSafeHwnd(), DLDestGrpDlgRect.left + SCALEX(7) + GeneDBLblRect.right + SCALEX(7), DLDestGrpDlgRect.top + SCALEY(13) + GeneDBTxtRect.bottom, DLDestGrpDlgRect.right - DLDestGrpDlgRect.left - GeneDBBrowseRect.right - GeneDBLblRect.right - SCALEX(21) - SCALEX(17), GeneDBTxtRect.bottom, true);
	::MoveWindow(lblMAPPArch.GetSafeHwnd(), DLDestGrpDlgRect.left + SCALEX(7), DLDestGrpDlgRect.top + SCALEY(11) + (GeneDBTxtRect.bottom * 2), GeneDBLblRect.right, GeneDBLblRect.bottom, true);
	::MoveWindow(btnMappArchBrowse.GetSafeHwnd(), DLDestGrpDlgRect.right - DLDestGrpDlgRect.left - GeneDBBrowseRect.right - SCALEX(4), DLDestGrpDlgRect.top + SCALEY(11) + (GeneDBTxtRect.bottom * 2), GeneDBBrowseRect.right + SCALEX(2), GeneDBBrowseRect.bottom + SCALEX(2), true);
	::MoveWindow(txtMAPPArch.GetSafeHwnd(), DLDestGrpDlgRect.left + SCALEX(7) + GeneDBLblRect.right + SCALEX(7), DLDestGrpDlgRect.top + SCALEY(11) + (GeneDBTxtRect.bottom * 2), DLDestGrpDlgRect.right - DLDestGrpDlgRect.left - GeneDBBrowseRect.right - GeneDBLblRect.right - SCALEX(21) - SCALEX(17), GeneDBTxtRect.bottom, true);
	::MoveWindow(lblOtrInfo.GetSafeHwnd(), DLDestGrpDlgRect.left + SCALEX(7), DLDestGrpDlgRect.top + SCALEY(9) + (GeneDBTxtRect.bottom * 3), GeneDBLblRect.right, GeneDBLblRect.bottom, true);
	::MoveWindow(btnOtrInfoBrowse.GetSafeHwnd(), DLDestGrpDlgRect.right - DLDestGrpDlgRect.left - GeneDBBrowseRect.right - SCALEX(4), DLDestGrpDlgRect.top + SCALEY(9) + (GeneDBTxtRect.bottom * 3), GeneDBBrowseRect.right + SCALEX(2), GeneDBBrowseRect.bottom + SCALEX(2), true);
	::MoveWindow(txtOtrInfo.GetSafeHwnd(), DLDestGrpDlgRect.left + SCALEX(7) + GeneDBLblRect.right + SCALEX(7), DLDestGrpDlgRect.top + SCALEY(9) + (GeneDBTxtRect.bottom * 3), DLDestGrpDlgRect.right - DLDestGrpDlgRect.left - GeneDBBrowseRect.right - GeneDBLblRect.right - SCALEX(21) - SCALEX(17), GeneDBTxtRect.bottom, true);


	// Add height of all controls whose height wont change + spaces. Remainder is split between
	// folder list and selection summary.
	m_frmConnStat.GetClientRect(&ConnectionStatRect);
	grpDLDestinations.GetClientRect(&DLDestinationsRect);
	MessageDlgRect = AdjustToDlgCoordinates(&lblTopMessage);
	nNonAdjControlHeight = ConnectionStatRect.bottom + DLDestinationsRect.bottom + MessageDlgRect.bottom + SCALEY(10 + 7 + 7 + 12);
	int nSelSumHeight;
	if (DialogWindowRect.bottom - DialogWindowRect.top <= nMinDialogHeight + (int)(nMinDialogHeight * .17))
		nSelSumHeight = nOrigSelSumHeight + (int)((DialogWindowRect.bottom - DialogWindowRect.top - nMinDialogHeight)  * .33);//nOrigSelSumHeight + (int)((DialogRect.bottom - nNonAdjControlHeight)  * .2);
	else
		nSelSumHeight = nOrigSelSumHeight + (int)(nOrigSelSumHeight * .17);

	int nFileAreaHeight = (DialogWindowRect.bottom - DialogWindowRect.top) - (nNonAdjControlHeight + nSelSumHeight + SCALEY(22));

	// Selection Summary Group
	::MoveWindow(m_frmSelectSum.GetSafeHwnd(), SCALEX(10), DLDestGrpDlgRect.top - nSelSumHeight - SCALEY(7), DialogRect.right - DialogRect.left - SCALEX(20), nSelSumHeight, true);
	DLSelSumDlgRect = AdjustToDlgCoordinates(&m_frmSelectSum);
	::MoveWindow(lblSelCountKey.GetSafeHwnd(), DLSelSumDlgRect.left + SCALEX(7), DLSelSumDlgRect.top + SCALEY(15), SelCountKeyRect.right, SelCountKeyRect.bottom, true);
	::MoveWindow(SelectCount.GetSafeHwnd(), (int)((DLSelSumDlgRect.right - DLSelSumDlgRect.left) / 2) - SCALEX(15) - SelCountValRect.right, DLSelSumDlgRect.top + SCALEY(15), SelCountValRect.right, SelCountValRect.bottom, true);
	::MoveWindow(lblDLTimeKey.GetSafeHwnd(), DLSelSumDlgRect.left + SCALEX(7), DLSelSumDlgRect.top + SCALEY(15) + (int)(SelCountValRect.bottom * 1.5), SelCountKeyRect.right, SelCountKeyRect.bottom, true);
	::MoveWindow(DLTime.GetSafeHwnd(), (int)((DLSelSumDlgRect.right - DLSelSumDlgRect.left) / 2) - SCALEX(15) - SelCountValRect.right, DLSelSumDlgRect.top + SCALEY(15) + (int)(SelCountValRect.bottom * 1.5), SelCountValRect.right, SelCountValRect.bottom, true);
	::MoveWindow(lblDLSizeKey.GetSafeHwnd(), DLSelSumDlgRect.left + SCALEX(7), DLSelSumDlgRect.top + SCALEY(15) + (int)(SelCountValRect.bottom * 3), SelCountKeyRect.right, SelCountKeyRect.bottom, true);
	::MoveWindow(lblDLSizeVal.GetSafeHwnd(), (int)((DLSelSumDlgRect.right - DLSelSumDlgRect.left) / 2) - SCALEX(15) - SelCountValRect.right, DLSelSumDlgRect.top + SCALEY(15) + (int)(SelCountValRect.bottom * 3), SelCountValRect.right, SelCountValRect.bottom, true);
	::MoveWindow(lblReqDiskSpaceKey.GetSafeHwnd(), DLSelSumDlgRect.left + SCALEX(7), DLSelSumDlgRect.top + SCALEY(15) + (int)(SelCountValRect.bottom * 4.5), SelCountKeyRect.right, SelCountKeyRect.bottom, true);
	::MoveWindow(m_labTotalDLSize.GetSafeHwnd(), (int)((DLSelSumDlgRect.right - DLSelSumDlgRect.left) / 2) - SCALEX(15) - SelCountValRect.right, DLSelSumDlgRect.top + SCALEY(15) + (int)(SelCountValRect.bottom * 4.5), SelCountValRect.right, SelCountValRect.bottom, true);
	::MoveWindow(lblFreeSpaceKey.GetSafeHwnd(), DLSelSumDlgRect.left + SCALEX(7), DLSelSumDlgRect.top + SCALEY(15) + (int)(SelCountValRect.bottom * 6), FreeSpaceKeyRect.right, FreeSpaceKeyRect.bottom, true);
	int nFreeSpaceListWidth = (((int)(DLSelSumDlgRect.right - DLSelSumDlgRect.left) / 2) - SCALEX(15) - (DLSelSumDlgRect.left + SCALEX(7) + FreeSpaceKeyRect.right)) - SCALEX(7);
	::MoveWindow(m_lstFreeSpace.GetSafeHwnd(), nFreeSpaceListWidth > SCALEX(145) ? (int)((DLSelSumDlgRect.right - DLSelSumDlgRect.left) / 2) - SCALEX(15) - SCALEX(145) : DLSelSumDlgRect.left + SCALEX(7) + FreeSpaceKeyRect.right + SCALEX(7), DLSelSumDlgRect.top + SCALEY(15) + (int)(SelCountValRect.bottom * 6), nFreeSpaceListWidth > SCALEX(145) ? SCALEX(145) : nFreeSpaceListWidth, DLSelSumDlgRect.bottom - (DLSelSumDlgRect.top + SCALEY(15) + (int)(SelCountValRect.bottom * 6)) - SCALEY(8), true);
	::MoveWindow(lblDLQueueKey.GetSafeHwnd(), (int)((DLSelSumDlgRect.right - DLSelSumDlgRect.left) / 2) + SCALEX(15), DLSelSumDlgRect.top + SCALEY(15), SelCountKeyRect.right, SelCountKeyRect.bottom, true);
	::MoveWindow(DLQueueList.GetSafeHwnd(), (int)((DLSelSumDlgRect.right - DLSelSumDlgRect.left) / 2) + SCALEX(15), DLSelSumDlgRect.top + SCALEY(15) + SelCountKeyRect.bottom, (int)((DLSelSumDlgRect.right - DLSelSumDlgRect.left) / 2) - SCALEX(22), (DLSelSumDlgRect.bottom - DLSelSumDlgRect.top) - SCALEY(16) - SelCountKeyRect.bottom - GeneDBBrowseRect.bottom - SCALEY(12), true);
	DLQueueLstDlgRect = AdjustToDlgCoordinates(&DLQueueList);
	::MoveWindow(m_btnRemoveDB.GetSafeHwnd(), DLQueueLstDlgRect.right - GeneDBBrowseRect.right, DLQueueLstDlgRect.bottom, GeneDBBrowseRect.right, GeneDBBrowseRect.bottom + SCALEY(2), true);


	// Folder Tree and File List Box
	::MoveWindow(DBFolderTree.GetSafeHwnd(), SCALEX(10), MessageDlgRect.bottom + SCALEY(7), DialogRect.right - (LONG)(DialogRect.right * .65), nFileAreaHeight, true);
	FolderTreeDlgRect = AdjustToDlgCoordinates(&DBFolderTree);
	::MoveWindow(DBList.GetSafeHwnd(), FolderTreeDlgRect.right, MessageDlgRect.bottom + SCALEY(7), DialogRect.right - FolderTreeDlgRect.right - SCALEX(10), nFileAreaHeight, true);
	FileListDlgRect	= AdjustToDlgCoordinates(&DBList);
	::MoveWindow(m_labReadMeText.GetSafeHwnd(), FileListDlgRect.left + (int)((FileListDlgRect.right - FileListDlgRect.left) / 2) - (int)((ReadMeLblRect.right - ReadMeLblRect.left) / 2), FileListDlgRect.top + (int)((FileListDlgRect.bottom - FileListDlgRect.top) / 2) - (int)((ReadMeLblRect.bottom + ShowReadMeBtnRect.bottom + SCALEY(15)) / 2), ReadMeLblRect.right, ReadMeLblRect.bottom, true);
	::MoveWindow(m_btnShowReadMe.GetSafeHwnd(), FileListDlgRect.left + (int)((FileListDlgRect.right - FileListDlgRect.left) / 2) - (int)((ShowReadMeBtnRect.right - ShowReadMeBtnRect.left) / 2), FileListDlgRect.top + (int)((FileListDlgRect.bottom - FileListDlgRect.top) / 2) - (int)((ReadMeLblRect.bottom + ShowReadMeBtnRect.bottom + SCALEY(15)) / 2) + ReadMeLblRect.bottom + SCALEY(15), ShowReadMeBtnRect.right, ShowReadMeBtnRect.bottom, true);

	// Maximize width of the columns using list box width: Total width of columns: 379. File Name 130 or 34%, Location 103 or 27%. File Name is 56% of remaining width, Location 44%
	DBList.GetClientRect(&DBListRect);
	int nOtherColumnsWidth = SCALEX(21) + SCALEX(60) + SCALEX(65) + SCALEX(22); // Last 22 is for scroll bar
	DBList.SetColumnWidth(1, (int)((DBListRect.right - nOtherColumnsWidth) * .56));
	DBList.SetColumnWidth(4, (int)((DBListRect.right - nOtherColumnsWidth) * .44));

	InvalidateRect(NULL);
	UpdateWindow();
}

void CCDDBCopyDlg::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);

	if (m_bOKToSize)
		DoControlScaling();
}

// Returns a CRect object with the coordinates adjusted to be relative to the dialog.
CRect CCDDBCopyDlg::AdjustToDlgCoordinates(CWnd* pWindowToAdjust)
{
	CRect DialogScreenRect, ControlRect;
	int nCaptionHeight = GetSystemMetrics(SM_CYCAPTION);
	
	GetWindowRect(&DialogScreenRect);

	// Now adjust for caption bar.
	DialogScreenRect.top+= nCaptionHeight;

	pWindowToAdjust->GetWindowRect(&ControlRect);

	ControlRect.bottom-= DialogScreenRect.top;
	ControlRect.top-= DialogScreenRect.top;
	ControlRect.left-= DialogScreenRect.left;
	ControlRect.right-= DialogScreenRect.left;

	return ControlRect;
}
void CCDDBCopyDlg::OnGetMinMaxInfo(MINMAXINFO* lpMMI)
{
	lpMMI->ptMinTrackSize.x = nMinDialogWidth;
	lpMMI->ptMinTrackSize.y = nMinDialogHeight;

	CDialog::OnGetMinMaxInfo(lpMMI);
}

void CCDDBCopyDlg::OnLButtonDown(UINT nFlags, CPoint point)
{
	RECT FolderRect, FileListRect, DialogRect;
	CString szMessageText;
	int	nXBorderWidth = GetSystemMetrics(SM_CXSIZEFRAME),
		nYBorderHeight = GetSystemMetrics(SM_CYSIZEFRAME) + GetSystemMetrics(SM_CYCAPTION);

	DBFolderTree.GetWindowRect(&FolderRect);
	DBList.GetWindowRect(&FileListRect);
	GetWindowRect(&DialogRect);

	if (point.x > FolderRect.right - DialogRect.left - nXBorderWidth && point.x < FileListRect.left - DialogRect.left - nXBorderWidth - SCALEX(1) &&  
		point.y >= FolderRect.top - DialogRect.top - nYBorderHeight && point.y <= FolderRect.bottom - DialogRect.top - nYBorderHeight)
	{
		m_bCapturingMouse = TRUE;
		SetCapture();
	}

	CDialog::OnLButtonDown(nFlags, point);
}

void CCDDBCopyDlg::OnMouseMove(UINT nFlags, CPoint point)
{
	RECT FolderRect, FileListRect, DialogRect;
	CString szMessageText;
	int	nXBorderWidth = GetSystemMetrics(SM_CXSIZEFRAME),
		nYBorderHeight = GetSystemMetrics(SM_CYSIZEFRAME) + GetSystemMetrics(SM_CYCAPTION);

	DBFolderTree.GetWindowRect(&FolderRect);
	DBList.GetWindowRect(&FileListRect);
	GetWindowRect(&DialogRect);

	if (m_bCapturingMouse)
	{
		::MoveWindow(DBFolderTree.GetSafeHwnd(), FolderRect.left - DialogRect.left - nXBorderWidth, FolderRect.top - DialogRect.top - nYBorderHeight,
			point.x - SCALEX(8), FolderRect.bottom - FolderRect.top, true);

		DBFolderTree.GetWindowRect(&FolderRect);

		::MoveWindow(DBList.GetSafeHwnd(), FolderRect.right - DialogRect.left, FileListRect.top - DialogRect.top - nYBorderHeight,
			DialogRect.right - FolderRect.right - SCALEX(18), FolderRect.bottom - FolderRect.top, true);
	}

	if (point.x > FolderRect.right - DialogRect.left - nXBorderWidth && point.x < FileListRect.left - DialogRect.left - nXBorderWidth - SCALEX(1) &&  
		point.y >= FolderRect.top - DialogRect.top - nYBorderHeight && point.y <= FolderRect.bottom - DialogRect.top - nYBorderHeight)
	{
		TRACE("In the Zone!\n");
		m_bInResizeZone = TRUE;
	}
	else
		m_bInResizeZone = FALSE;


	CDialog::OnMouseMove(nFlags, point);
}

void CCDDBCopyDlg::OnLButtonUp(UINT nFlags, CPoint point)
{
	if (m_bCapturingMouse || m_bInResizeZone)
	{
		ReleaseCapture();
		m_bCapturingMouse = FALSE;
	}

	CDialog::OnLButtonUp(nFlags, point);
}

BOOL CCDDBCopyDlg::OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message)
{
//	WINDOWINFO wi;
    if (m_bCapturingMouse || m_bInResizeZone)
    {
		TRACE("Size Cursor\n");
        ::SetCursor(AfxGetApp()->LoadStandardCursor(IDC_SIZEWE));
        return TRUE;
    }

/*	DBFolderTree.GetWindowInfo(&wi);
	
	if (wi.dwExStyle & WS_DISABLED)
	{
        ::SetCursor(AfxGetApp()->LoadStandardCursor(IDC_WAIT));
		return TRUE;
	}
*/
	return CDialog::OnSetCursor(pWnd, nHitTest, message);
}
