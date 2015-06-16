#pragma once
#include "afxcmn.h"
#include "afxwin.h"
#include "afxcoll.h"

UINT BeginUpdate( LPVOID pParam );
UINT CheckForUpdates( LPVOID pParam );
bool ConnectToServer();
void PopulateInUseList( LPVOID pParam, char* szTempDir );
void BuildAndRunBatchFile(LPVOID pParam, CStringList* FilesToWaitOn);

// CUpdateDlg dialog

class CUpdateDlg : public CDialog
{
	DECLARE_DYNAMIC(CUpdateDlg)

public:
	CUpdateDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CUpdateDlg();

// Dialog Data
	enum { IDD = IDD_UPDATEDLG };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	// Update log list control
	CListCtrl lstUpdateLog;
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	virtual BOOL OnInitDialog();
private:
	CGenMAPPDBDLApp* m_pApp;
public:
	// Holds the arrow, check and X images
	CImageList UpdateLogImageList;
	afx_msg void OnBnClickedButton1();
private:
	// Creates the Update Log listview control
	void CreateUpdateLog(void);
	// Performs scaling of the dialog controls based on font scaling factor
	bool DoScaling(void);
public:
	// The Update Log Label above the list view control
	CStatic m_labUpdateLog;
	// The splashscreen
	CStatic m_picGenMAPPSplash;
	CStatic m_labUpdatePrg;
	CStatic m_grpUpdateInfo;
	CStatic m_labUpdateSummary;
	CMapStringToString CouldBeInUseList;
	// Installation directory, set by external function call InvokeUpdate
	CString szInstallDir;
	// Add a string to the Update Log. Returns the index for updating later.
	int AddItemToUpdateLog(CString szItem);
	// The most recent Update Log index. Useful for setting the result of the current item.
	int nLastLogIndex;
	// Sets the icon of a previous Update Log item to a success or failure icon.
	void UpdatePreviousLogItem(bool bSuccess, int nIndex);
	// Update Progress control
	CProgressCtrl prgUpdateProgress;
	afx_msg void OnSize(UINT nType, int cx, int cy);
};
