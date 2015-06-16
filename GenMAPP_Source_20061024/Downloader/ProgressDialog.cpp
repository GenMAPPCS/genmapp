// ProgressDialog.cpp : implementation file
//

#include "stdafx.h"
#include "GMFTP.h"
#include "GenMAPPDBDL.h"
#include "DBFile.h"
#include "ProgressDialog.h"
#include ".\progressdialog.h"

#define SCALEX(argX) ((int) ((argX) * m_pApp->scaleX))
#define SCALEY(argY) ((int) ((argY) * m_pApp->scaleY))


// CProgressDialog dialog

IMPLEMENT_DYNAMIC(CProgressDialog, CDialog)
CProgressDialog::CProgressDialog(CWnd* pParent /*=NULL*/)
	: CDialog(CProgressDialog::IDD, pParent)
	, m_pApp(NULL)
	, dwBytesDownloaded(0)
	, dwTotalBytesToDownload(0)
	, dCurrentSpeed(0)
	, nHistoricalSpeedRecordCount(-1)
	, m_bOKToSize(false)
	, nMinDialogHeight(0)
	, nMinDialogWidth(0)
{
	m_pApp = (CGenMAPPDBDLApp*)AfxGetApp();
}

CProgressDialog::~CProgressDialog()
{
}

void CProgressDialog::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_OVERALLPRG, OverallProgress);
	DDX_Control(pDX, IDC_FILEINDL, labFileInDL);
	DDX_Control(pDX, IDC_FILETALLY, labFileTally);
	DDX_Control(pDX, IDOK, btnOK);
	DDX_Control(pDX, IDC_GENMAPPSPLASH, picGenMAPPSplash);
	DDX_Control(pDX, IDC_ABORT, btnAbort);
	DDX_Control(pDX, IDC_GRPOPTIONS, grpOptions);
	DDX_Control(pDX, IDC_PAUSE, btnPause);
	DDX_Control(pDX, IDCANCEL, btnCancel);
	DDX_Control(pDX, IDC_LSTQUEUEDFILES, lstQueuedFiles);
	DDX_Control(pDX, IDC_DLSPEED, lblDownloadSpeed);
	DDX_Control(pDX, IDC_REMAININGTIME, lblRemainingTime);
	DDX_Control(pDX, IDC_LBLDLTOTALS, lblByteTotals);
	DDX_Control(pDX, IDC_FILENAME, lblFileName);
	DDX_Control(pDX, IDC_SERVERNAME, lblServerName);
	DDX_Control(pDX, IDC_LOCATION, lblLocation);
	DDX_Control(pDX, IDC_BYTESDOWNLOADED, lblFileBytesDownloaded);
	DDX_Control(pDX, IDC_TRANSPORT, lblTransport);
	DDX_Control(pDX, IDC_FILETYPE, lblFileType);
	DDX_Control(pDX, IDC_GRPDLSTATISTICS, grpDownloadStatistics);
	DDX_Control(pDX, IDC_GRPCURRENTFILE, grpCurrentFile);
	DDX_Control(pDX, IDC_LBLDLSPEED, lblDLSpeedKey);
	DDX_Control(pDX, IDC_LBLREMAININGTIME, lblRemainingTimeKey);
	DDX_Control(pDX, IDC_LBLFILENAME, lblFileNameKey);
	DDX_Control(pDX, IDC_LBLSERVERNAME, lblServerNameKey);
	DDX_Control(pDX, IDC_LBLLOCATION, lblLocationKey);
	DDX_Control(pDX, IDC_LBLBYTEDOWNLOADED, lblFileBytesDLKey);
	DDX_Control(pDX, IDC_LBLTRANSPORT, lblNetTransKey);
	DDX_Control(pDX, IDC_LBLFILETYPE, lblFileTypeKey);
	DDX_Control(pDX, IDC_DLFILE, lblFileTallyKey);
	DDX_Control(pDX, IDC_TOTALBYTESDL, lblTotalBytesDLKey);
	DDX_Control(pDX, IDC_SPEEDPROG, m_lstSpeedTestProg);
}


BEGIN_MESSAGE_MAP(CProgressDialog, CDialog)
	ON_WM_SYSCOLORCHANGE()
	ON_WM_CTLCOLOR()
	ON_BN_CLICKED(IDCANCEL, OnBnClickedCancel)
	ON_BN_CLICKED(IDC_ABORT, OnBnClickedAbort)
	ON_WM_TIMER()
	ON_BN_CLICKED(IDOK, OnBnClickedOk)
	ON_WM_SIZE()
	ON_BN_CLICKED(IDC_PAUSE, OnBnClickedPause)
	ON_WM_GETMINMAXINFO()
END_MESSAGE_MAP()


// CProgressDialog message handlers


HBRUSH CProgressDialog::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialog::OnCtlColor(pDC, pWnd, nCtlColor);


//	pDC->SetBkColor(RGB(220, 220, 220));
	//red
	//pDC->SetTextColor(RGB(255, 0, 0));

	
	// TODO:  Change any attributes of the DC here

	// TODO:  Return a different brush if the default is not desired
	return hbr;
}

BOOL CProgressDialog::OnInitDialog()
{
	CDialog::OnInitDialog();

	SetIcon(LoadIcon(m_pApp->m_hInstance, MAKEINTRESOURCE(IDI_GENMAPP)), TRUE);

	SizeDialogToScreen();
	
	m_bOKToSize = true;
	MoveAndSizeControls();

	if (m_pApp->CalledFromInstaller)
		::SetWindowText(::FindWindowEx(this->GetSafeHwnd(), NULL, NULL, "OK"), "Next");

	CreateDirectory(m_pApp->szGDBPath, NULL);
	CreateDirectory(m_pApp->szGEXPath, NULL);
	CreateDirectory(m_pApp->szMAPPPath, NULL);

	CreateFileOpsList();

    AfxBeginThread(::DownloadFiles, this);


	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

void CProgressDialog::OnBnClickedCancel()
{
	if (m_pApp->bDLInProgress)
	{
		if (AfxMessageBox("This will abort the download in progress. Are you sure?", MB_YESNO) == IDYES)
		{
			m_pApp->bAbortDownload = true;
			m_FileOpImageList.Detach();
			m_FileOpImageList.DeleteImageList();
			//Sleep(20);
			OnCancel();
		}
	}
	else
	{
		m_FileOpImageList.Detach();
		m_FileOpImageList.DeleteImageList();
		//Sleep(20);
		OnCancel();
	}
}

void CProgressDialog::OnBnClickedOk()
{
	m_FileOpImageList.Detach();
	m_FileOpImageList.DeleteImageList();
	//Sleep(2500);
	OnOK();
}

void CProgressDialog::OnBnClickedAbort()
{
	if (AfxMessageBox("This will abort the download in progress. Are you sure?", MB_YESNO) == IDYES)
	{
		m_pApp->bAbortDownload = true;
		if (m_pApp->CalledFromInstaller)
			labFileInDL.SetWindowText("Download aborted. Click Back to select new files for download or click Next to complete the install of GenMAPP 2.");
		else
			labFileInDL.SetWindowText("Download aborted. Click Back to select new files for download or click OK to exit.");

		labFileTally.SetWindowText(" ");

		btnAbort.EnableWindow(false);
		UpdateData();
		btnOK.EnableWindow();
	}
}

void CProgressDialog::OnTimer(UINT nIDEvent)
{
	// Total # of bytes / speed ( bytes * 1024) = num of seconds
	CString szStatic;
	DWORD dwNumberOfSeconds;
	UINT nTotalDays, nTotalHours, nTotalMinutes, nTotalSeconds;
	dwNumberOfSeconds = (DWORD)((((dwTotalBytesToDownload - dwBytesDownloaded) + 1) / 1024) / (double)dCurrentSpeed);

	nTotalDays = dwNumberOfSeconds / 86400;
	nTotalHours = (dwNumberOfSeconds / 3600) % 24;
	nTotalMinutes = (dwNumberOfSeconds / 60) % 60;
	nTotalSeconds = dwNumberOfSeconds % 60;

	// shows max time after download
	if (nTotalDays)
		szStatic.Format("%d Day%s %002i:%002i:%002i", nTotalDays, nTotalDays > 1 ? "s" : "",
			nTotalHours, nTotalMinutes, nTotalSeconds);
	else
		szStatic.Format("%002i:%002i:%002i", nTotalHours, nTotalMinutes, nTotalSeconds);

	lblRemainingTime.SetWindowText(szStatic);

	CDialog::OnTimer(nIDEvent);
}

// Based on the user's font size setting, move and size the controls 
// relative to the splashscreen which is not proportional
void CProgressDialog::MoveAndSizeControls(void)
{
	CRect DialogRect, SplashRect, FileTallyRect, OptionsRect, PauseRect, AbortRect,
		CancelDlgRect, FileQueueRect, FileQueueDialogRect, SplashDlgRect, 
		CurrentFileRect, CurrentFileDialogRect, DownloadStatisticsRect, DownloadStatsDlgRect, CancelRect,
		OKRect, OKDlgRect, OptionsGrpDlgRect, FileInDLRect, ProgressRect;

	int	nLabelWidths = 100, nFromTop = 20, nValuePaddingFromRight = 28;

	GetClientRect(&DialogRect);
	picGenMAPPSplash.GetClientRect(&SplashRect);
	labFileTally.GetClientRect(&FileTallyRect);
	grpOptions.GetClientRect(&OptionsRect);
	btnPause.GetClientRect(&PauseRect);
	btnAbort.GetClientRect(&AbortRect);
	lstQueuedFiles.GetClientRect(&FileQueueRect);
	grpCurrentFile.GetClientRect(&CurrentFileRect);
	grpDownloadStatistics.GetClientRect(&DownloadStatisticsRect);
	btnCancel.GetClientRect(&CancelRect);
	labFileInDL.GetClientRect(&FileInDLRect);
	OverallProgress.GetClientRect(&ProgressRect);
	btnOK.GetClientRect(&OKRect);
	

	// Splashscreen
	::MoveWindow(picGenMAPPSplash.GetSafeHwnd(), DialogRect.right  - SplashRect.right - SCALEX(7), SCALEY(7), SplashRect.right, SplashRect.bottom, true);


	// File Queue List Box
	SplashDlgRect = AdjustToDlgCoordinates(&picGenMAPPSplash);
	::MoveWindow(lstQueuedFiles.GetSafeHwnd(), SCALEX(7), SCALEY(10), SplashDlgRect.left - SCALEX(17), SplashDlgRect.bottom - SCALEY(15), true);
	
	// Maximize width of the "Operation" column using list box width - Percent column width - icon colum width - vert. scroll bar witdh
	lstQueuedFiles.GetClientRect(&FileQueueRect);
	lstQueuedFiles.SetColumnWidth(1, FileQueueRect.right - SCALEX(70) - SCALEX(22) - SCALEX(20));

	
	// Current File group
	// Note that FileTally.bottom represents the label height of all labels
	FileQueueDialogRect = AdjustToDlgCoordinates(&lstQueuedFiles);
	::MoveWindow(grpCurrentFile.GetSafeHwnd(), DialogRect.right - (LONG)(DialogRect.right * .50), FileQueueDialogRect.bottom + SCALEY(7), (int)(DialogRect.right * .50) - SCALEX(7), SCALEY(150), true);
	CurrentFileDialogRect = AdjustToDlgCoordinates(&grpCurrentFile);
	::MoveWindow(lblFileNameKey.GetSafeHwnd(), CurrentFileDialogRect.left + SCALEX(15), CurrentFileDialogRect.top + SCALEY(nFromTop), SCALEX(nLabelWidths), FileTallyRect.bottom, true);
	::MoveWindow(lblFileName.GetSafeHwnd(), CurrentFileDialogRect.left + SCALEX(15) + SCALEX(nLabelWidths), CurrentFileDialogRect.top + SCALEY(nFromTop), CurrentFileDialogRect.right - CurrentFileDialogRect.left - SCALEX(nValuePaddingFromRight) - SCALEX(nLabelWidths), FileTallyRect.bottom, true);
	::MoveWindow(lblServerNameKey.GetSafeHwnd(), CurrentFileDialogRect.left + SCALEX(15), CurrentFileDialogRect.top + SCALEY(nFromTop) + FileTallyRect.bottom, SCALEX(nLabelWidths), FileTallyRect.bottom, true);
	::MoveWindow(lblServerName.GetSafeHwnd(), CurrentFileDialogRect.left + SCALEX(15) + SCALEX(nLabelWidths), CurrentFileDialogRect.top + SCALEY(nFromTop) + FileTallyRect.bottom, CurrentFileDialogRect.right - CurrentFileDialogRect.left - SCALEX(nValuePaddingFromRight) - SCALEX(nLabelWidths), FileTallyRect.bottom, true);
	::MoveWindow(lblLocationKey.GetSafeHwnd(), CurrentFileDialogRect.left + SCALEX(15), CurrentFileDialogRect.top + SCALEY(nFromTop) + (FileTallyRect.bottom * 2), SCALEX(nLabelWidths), FileTallyRect.bottom, true);
	::MoveWindow(lblLocation.GetSafeHwnd(), CurrentFileDialogRect.left + SCALEX(15) + SCALEX(nLabelWidths), CurrentFileDialogRect.top + SCALEY(nFromTop) + (FileTallyRect.bottom * 2), CurrentFileDialogRect.right - CurrentFileDialogRect.left - SCALEX(nValuePaddingFromRight) - SCALEX(nLabelWidths), FileTallyRect.bottom, true);
	::MoveWindow(lblFileBytesDLKey.GetSafeHwnd(), CurrentFileDialogRect.left + SCALEX(15), CurrentFileDialogRect.top + SCALEY(nFromTop) + (FileTallyRect.bottom * 3), SCALEX(nLabelWidths), FileTallyRect.bottom, true);
	::MoveWindow(lblFileBytesDownloaded.GetSafeHwnd(), CurrentFileDialogRect.left + SCALEX(15) + SCALEX(nLabelWidths), CurrentFileDialogRect.top + SCALEY(nFromTop) + (FileTallyRect.bottom * 3), CurrentFileDialogRect.right - CurrentFileDialogRect.left - SCALEX(nValuePaddingFromRight) - SCALEX(nLabelWidths), FileTallyRect.bottom, true);
	::MoveWindow(lblNetTransKey.GetSafeHwnd(), CurrentFileDialogRect.left + SCALEX(15), CurrentFileDialogRect.top + SCALEY(nFromTop) + (FileTallyRect.bottom * 4), SCALEX(nLabelWidths), FileTallyRect.bottom, true);
	::MoveWindow(lblTransport.GetSafeHwnd(), CurrentFileDialogRect.left + SCALEX(15) + SCALEX(nLabelWidths), CurrentFileDialogRect.top + SCALEY(nFromTop) + (FileTallyRect.bottom * 4), CurrentFileDialogRect.right - CurrentFileDialogRect.left - SCALEX(nValuePaddingFromRight) - SCALEX(nLabelWidths), FileTallyRect.bottom, true);
	::MoveWindow(lblFileTypeKey.GetSafeHwnd(), CurrentFileDialogRect.left + SCALEX(15), CurrentFileDialogRect.top + SCALEY(nFromTop) + (FileTallyRect.bottom * 5), SCALEX(nLabelWidths), FileTallyRect.bottom, true);
	::MoveWindow(lblFileType.GetSafeHwnd(), CurrentFileDialogRect.left + SCALEX(15) + SCALEX(nLabelWidths), CurrentFileDialogRect.top + SCALEY(nFromTop) + (FileTallyRect.bottom * 5), CurrentFileDialogRect.right - CurrentFileDialogRect.left - SCALEX(nValuePaddingFromRight) - SCALEX(nLabelWidths), FileTallyRect.bottom, true);


	// Download statistics group
	::MoveWindow(grpDownloadStatistics.GetSafeHwnd(), SCALEX(7), FileQueueDialogRect.bottom + SCALEY(7), CurrentFileDialogRect.left - SCALEX(17), SCALEY(150), true);
	DownloadStatsDlgRect = AdjustToDlgCoordinates(&grpDownloadStatistics);
	::MoveWindow(lblFileTallyKey.GetSafeHwnd(), DownloadStatsDlgRect.left + SCALEX(15), DownloadStatsDlgRect.top + SCALEY(nFromTop), SCALEX(nLabelWidths), FileTallyRect.bottom, true);
	::MoveWindow(labFileTally.GetSafeHwnd(), DownloadStatsDlgRect.left + SCALEX(15) + SCALEX(nLabelWidths), DownloadStatsDlgRect.top + SCALEY(nFromTop), DownloadStatsDlgRect.right - DownloadStatsDlgRect.left - SCALEX(nValuePaddingFromRight) - SCALEX(nLabelWidths), FileTallyRect.bottom, true);
	::MoveWindow(lblTotalBytesDLKey.GetSafeHwnd(), DownloadStatsDlgRect.left + SCALEX(15), DownloadStatsDlgRect.top + SCALEY(nFromTop) + FileTallyRect.bottom, SCALEX(nLabelWidths), FileTallyRect.bottom, true);
	::MoveWindow(lblByteTotals.GetSafeHwnd(), DownloadStatsDlgRect.left + SCALEX(15) + SCALEX(nLabelWidths), DownloadStatsDlgRect.top + SCALEY(nFromTop) + FileTallyRect.bottom, DownloadStatsDlgRect.right - DownloadStatsDlgRect.left - SCALEX(nValuePaddingFromRight) - SCALEX(nLabelWidths), FileTallyRect.bottom, true);
	::MoveWindow(lblDLSpeedKey.GetSafeHwnd(), DownloadStatsDlgRect.left + SCALEX(15), DownloadStatsDlgRect.top + SCALEY(nFromTop) + (FileTallyRect.bottom * 2), SCALEX(nLabelWidths), FileTallyRect.bottom, true);
	::MoveWindow(lblDownloadSpeed.GetSafeHwnd(), DownloadStatsDlgRect.left + SCALEX(15) + SCALEX(nLabelWidths), DownloadStatsDlgRect.top + SCALEY(nFromTop) + (FileTallyRect.bottom * 2), DownloadStatsDlgRect.right - DownloadStatsDlgRect.left - SCALEX(nValuePaddingFromRight) - SCALEX(nLabelWidths), FileTallyRect.bottom, true);
	::MoveWindow(lblRemainingTimeKey.GetSafeHwnd(), DownloadStatsDlgRect.left + SCALEX(15), DownloadStatsDlgRect.top + SCALEY(nFromTop) + (FileTallyRect.bottom * 3), SCALEX(nLabelWidths), FileTallyRect.bottom, true);
	::MoveWindow(lblRemainingTime.GetSafeHwnd(), DownloadStatsDlgRect.left + SCALEX(15) + SCALEX(nLabelWidths), DownloadStatsDlgRect.top + SCALEY(nFromTop) + (FileTallyRect.bottom * 3), DownloadStatsDlgRect.right - DownloadStatsDlgRect.left - SCALEX(nValuePaddingFromRight) - SCALEX(nLabelWidths), FileTallyRect.bottom, true);


	// OK and Cancel
	::MoveWindow(btnOK.GetSafeHwnd(), DialogRect.right - OKRect.right - SCALEX(7), DialogRect.bottom - OKRect.bottom - SCALEY(7), OKRect.right, OKRect.bottom, true);
	OKDlgRect = AdjustToDlgCoordinates(&btnOK);
	::MoveWindow(btnCancel.GetSafeHwnd(), OKDlgRect.left - CancelRect.right - SCALEX(15), DialogRect.bottom - CancelRect.bottom - SCALEY(7), CancelRect.right, CancelRect.bottom, true);

	
	// Options group
	CancelDlgRect = AdjustToDlgCoordinates(&btnCancel);
	::MoveWindow(grpOptions.GetSafeHwnd(), CancelDlgRect.left - (int)(PauseRect.right * 2.7) - SCALEX(15), CancelDlgRect.bottom - (PauseRect.bottom * 2) - SCALEY(7), (int)(PauseRect.right * 2.7), (PauseRect.bottom * 2) + SCALEY(4), true);
	OptionsGrpDlgRect = AdjustToDlgCoordinates(&grpOptions);
	::MoveWindow(btnAbort.GetSafeHwnd(), OptionsGrpDlgRect.left + (int)(PauseRect.right * .2), OptionsGrpDlgRect.top + SCALEY(15), AbortRect.right, AbortRect.bottom, true);
	::MoveWindow(btnPause.GetSafeHwnd(), OptionsGrpDlgRect.left + PauseRect.right + SCALEX(22), OptionsGrpDlgRect.top + SCALEY(15), PauseRect.right, PauseRect.bottom, true);

	// Progress bar and status message
	// Bottom of statistics + (Height between statistics and options buttons / 2) - (Height of both controls / 2)
	::MoveWindow(labFileInDL.GetSafeHwnd(), (DialogRect.right / 2) - (int)(DialogRect.right * .45),
		DownloadStatsDlgRect.bottom + ((OptionsGrpDlgRect.top - DownloadStatsDlgRect.bottom) / 2) - ((FileInDLRect.bottom + ProgressRect.bottom) / 2), 
		(int)(DialogRect.right * .9), FileInDLRect.bottom, true);
	::MoveWindow(OverallProgress.GetSafeHwnd(), (DialogRect.right / 2) - (int)(DialogRect.right * .45),
		DownloadStatsDlgRect.bottom + ((OptionsGrpDlgRect.top - DownloadStatsDlgRect.bottom) / 2) - ((FileInDLRect.bottom + ProgressRect.bottom) / 2) + FileInDLRect.bottom, 
		(int)(DialogRect.right * .9), SCALEY(25), true);

	InvalidateRect(NULL);
	UpdateWindow();

}

// Returns a CRect object with the coordinates adjusted to be relative to the dialog.
CRect CProgressDialog::AdjustToDlgCoordinates(CWnd* pWindowToAdjust)
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

// Manages the Historical Speed Records array. Adds the speed record passed
// in to the array and bumps the oldest speed record out.
void CProgressDialog::AddHistoricalSpeedRecord(double fDLSpeed)
{
	int i;

	if (nHistoricalSpeedRecordCount >= SPEED_SAMPLES - 1)
	{
		for (i=1;i<=SPEED_SAMPLES - 1;i++)
			fHistoricalSpeedRecords[i - 1] = fHistoricalSpeedRecords[i];
	}
	else
		nHistoricalSpeedRecordCount++;

	fHistoricalSpeedRecords[nHistoricalSpeedRecordCount] = fDLSpeed;
}

// Calculates the average download speed based on saved speed records.
double CProgressDialog::GetAverageDLSpeed(void)
{
	if (nHistoricalSpeedRecordCount == 0)
		return 0;

	int i;
	double fAvgSpeed = 0;

	for (i=0;i<nHistoricalSpeedRecordCount;i++)
		fAvgSpeed+= fHistoricalSpeedRecords[i];

	return fAvgSpeed / nHistoricalSpeedRecordCount;

}

// Creates and populates the file operations list.
bool CProgressDialog::CreateFileOpsList(void)
{
	LVCOLUMN iconcolumn;
	POSITION pos;
	LVITEM  itemdata;
	RECT FileQueueRect;

	itemdata.mask =  LVIF_IMAGE | LVIF_STATE;
	itemdata.iImage = 3;
	itemdata.state  = 0;
	itemdata.stateMask = 0;
	itemdata.iSubItem = 0;
	itemdata.iItem = 0;

	// Create image list of folder icons
	m_FileOpImageList.Create(16, 16, ILC_COLOR8, 6, 5);

	// Add folder icons.
	m_FileOpImageList.Add(AfxGetApp()->LoadIcon(IDI_CHECKMARK));
	m_FileOpImageList.Add(AfxGetApp()->LoadIcon(IDI_RIGHTARROW));
	m_FileOpImageList.Add(AfxGetApp()->LoadIcon(IDI_REDX));
	m_FileOpImageList.Add(AfxGetApp()->LoadIcon(IDI_DOT));
	m_FileOpImageList.Add(AfxGetApp()->LoadIcon(IDI_PAUSE));
	m_FileOpImageList.Add(AfxGetApp()->LoadIcon(IDI_BLANK));
	lstQueuedFiles.SetImageList(&m_FileOpImageList, LVSIL_SMALL);

	iconcolumn.pszText = "";
	iconcolumn.cchTextMax = 1;
	iconcolumn.cx = SCALEX(22);
	iconcolumn.fmt = LVCFMT_CENTER;
	iconcolumn.iImage = 0;
	iconcolumn.iOrder = 0;
	iconcolumn.iSubItem = 0;
	iconcolumn.mask = LVCF_TEXT| LVCF_WIDTH| LVCF_SUBITEM | LVCF_FMT;

	lstQueuedFiles.GetClientRect(&FileQueueRect);

	lstQueuedFiles.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
	lstQueuedFiles.InsertColumn(0, &iconcolumn);
	lstQueuedFiles.InsertColumn(1, "Operation", LVCFMT_LEFT, FileQueueRect.right - SCALEX(70) - SCALEX(22) - SCALEX(20), 1);
	lstQueuedFiles.InsertColumn(2, "Complete", LVCFMT_CENTER, SCALEX(70), 1);

	for( pos = m_pApp->pFileList->GetHeadPosition(); pos != NULL; )
	{
		CDBFile* dbfile = (CDBFile*)m_pApp->pFileList->GetNext( pos );

		if (dbfile->bSelected)
		{
			itemdata.iItem = lstQueuedFiles.GetItemCount();
			itemdata.iItem = lstQueuedFiles.InsertItem(&itemdata);
			lstQueuedFiles.SetItem(itemdata.iItem, 1, LVIF_TEXT, "Download " + dbfile->GetFileNameOnly(dbfile->szSrcFileName), 0, 0, 0, 0, 0);
			lstQueuedFiles.SetItem(itemdata.iItem, 2, LVIF_TEXT, "0%", 0, 0, 0, 0, 0);
		}

	}

	for( pos = m_pApp->pFileList->GetHeadPosition(); pos != NULL; )
	{
		CDBFile* dbfile = (CDBFile*)m_pApp->pFileList->GetNext( pos );

		if (dbfile->bSelected&& (dbfile->szSrcFileName.Right(4).MakeUpper() == ".EXE" || 
			dbfile->szSrcFileName.Right(4).MakeUpper() == ".BAT"))
		{
			itemdata.iItem = lstQueuedFiles.GetItemCount();
			itemdata.iItem = lstQueuedFiles.InsertItem(&itemdata);
			lstQueuedFiles.SetItem(itemdata.iItem, 1, LVIF_TEXT, "Extract " + dbfile->GetFileNameOnly(dbfile->szSrcFileName), 0, 0, 0, 0, 0);
			lstQueuedFiles.SetItem(itemdata.iItem, 2, LVIF_TEXT, "0%", 0, 0, 0, 0, 0);
		}
	}

	m_nCurrentQueueIndex = 0;

	return true;
}

// Updates the current item in the File Queue List. 
// An icon and percentage can be specified. Increments the 
// current item when a check or an X icon is set.
void CProgressDialog::UpdateFileQueueItem(int nIcon, int nPercent)
{
	if (lstQueuedFiles.GetSafeHwnd() != NULL)
	{
		CString szPercent;
		LVITEM  itemdata;

		itemdata.mask =  LVIF_IMAGE | LVIF_STATE;
		itemdata.iImage = nIcon;
		itemdata.state  = 0;
		itemdata.stateMask = 0;
		itemdata.iSubItem = 0;
		itemdata.iItem = m_nCurrentQueueIndex;

		lstQueuedFiles.SetItem(&itemdata);

		if (nPercent > 100)
			nPercent = 100;

		szPercent.Format("%d%%", nPercent);

		lstQueuedFiles.SetItem(m_nCurrentQueueIndex, 2, LVIF_TEXT, szPercent, 0, 0, 0, 0, 0);

		if (nIcon == 2 || nIcon == 0)
			m_nCurrentQueueIndex++;
	}
}

void CProgressDialog::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);

	if (m_bOKToSize)
		MoveAndSizeControls();
}

// Adjusts the dialog size to fit the screen resoultion. Also stores the width and height when finished.
void CProgressDialog::SizeDialogToScreen(void)
{
	RECT DialogRect;
	
	GetWindowRect(&DialogRect);
	nMinDialogWidth = DialogRect.right - DialogRect.left;
	nMinDialogHeight = DialogRect.bottom - DialogRect.top;

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

void CProgressDialog::OnBnClickedPause()
{
	m_pApp->bPauseDownload = !m_pApp->bPauseDownload;
	btnPause.SetWindowText(m_pApp->bPauseDownload ? "Resume" : "Pause");
}

void CProgressDialog::OnGetMinMaxInfo(MINMAXINFO* lpMMI)
{
	lpMMI->ptMinTrackSize.x = nMinDialogWidth;
	lpMMI->ptMinTrackSize.y = nMinDialogHeight;

	CDialog::OnGetMinMaxInfo(lpMMI);
}
