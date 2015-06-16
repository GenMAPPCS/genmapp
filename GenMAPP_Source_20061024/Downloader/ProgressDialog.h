#pragma once
#include "afxcmn.h"
#include "afxwin.h"



// CProgressDialog dialog

class CProgressDialog : public CDialog
{
	DECLARE_DYNAMIC(CProgressDialog)

public:
	CProgressDialog(CWnd* pParent = NULL);   // standard constructor
	virtual ~CProgressDialog();

// Dialog Data
	enum { IDD = IDD_PROGDLG };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	CProgressCtrl OverallProgress;
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedCancel();
	CGenMAPPDBDLApp* m_pApp;
	CStatic labFileInDL;
	CStatic labFileTally;
	CButton btnOK;
	afx_msg void OnBnClickedAbort();
	afx_msg void OnTimer(UINT nIDEvent);
	CStatic picGenMAPPSplash;
	CButton btnAbort;
	// Options Group Box
	CStatic grpOptions;
	// Based on the user's font size setting, move and size the controls relative to the splashscreen which is not proportional
	void MoveAndSizeControls(void);
	CButton btnPause;
	CButton btnCancel;
private:
	// Returns a CRect object with the coordinates adjusted to be relative to the dialog.
	CRect AdjustToDlgCoordinates(CWnd* pWindowToAdjust);
public:
	CListCtrl lstQueuedFiles;
	CStatic lblDownloadSpeed;
	CStatic lblRemainingTime;
	DWORD dwBytesDownloaded;
	DWORD dwTotalBytesToDownload;
	double dCurrentSpeed;
public:
	CStatic lblByteTotals;
	int nHistoricalSpeedRecordCount;
	double fHistoricalSpeedRecords[SPEED_SAMPLES + 1];
	// Manages the Historical Speed Records array. Adds the speed record passed in to the array and bumps the oldest speed record out.
	void AddHistoricalSpeedRecord(double fDLSpeed);
	// Calculates the average download speed based on saved speed records.
	double GetAverageDLSpeed(void);
	// Creates and populates the file operations list.
	bool CreateFileOpsList(void);
private:
	CImageList m_FileOpImageList;
public:
	// Updates the current item in the File Queue List. An icon and percentage can be specified. Increments the current item when a check or an X icon is set.
	void UpdateFileQueueItem(int nIcon, int nPercent);
	int	m_nCurrentQueueIndex;
	CStatic lblFileName;
	CStatic lblServerName;
	CStatic lblLocation;
	CStatic lblFileBytesDownloaded;
	CStatic lblTransport;
	CStatic lblFileType;
	afx_msg void OnBnClickedOk();
	CStatic grpDownloadStatistics;
	CStatic grpCurrentFile;
	afx_msg void OnSize(UINT nType, int cx, int cy);
private:
	bool m_bOKToSize;
public:
	CStatic lblDLSpeedKey;
	CStatic lblRemainingTimeKey;
	CStatic lblFileNameKey;
	CStatic lblServerNameKey;
	CStatic lblLocationKey;
	CStatic lblFileBytesDLKey;
	CStatic lblNetTransKey;
	CStatic lblFileTypeKey;
	int nMinDialogHeight;
	int nMinDialogWidth;
	// Adjusts the dialog size to fit the screen resoultion. Also stores the width and height when finished.
	void SizeDialogToScreen(void);
	CStatic lblFileTallyKey;
	CStatic lblTotalBytesDLKey;
	afx_msg void OnBnClickedPause();
	afx_msg void OnGetMinMaxInfo(MINMAXINFO* lpMMI);
	// List box that shows at the start of download and displays the progress of speed testing
	CListBox m_lstSpeedTestProg;
};
