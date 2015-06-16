#pragma once
#ifndef ccddbcopydlg_h
#define ccddbcopydlg_h


#include "afxcmn.h"
#include "afxwin.h"
#include "afxcoll.h"


// CCDDBCopyDlg dialog

class CCDDBCopyDlg : public CDialog
{
	DECLARE_DYNAMIC(CCDDBCopyDlg)

public:
	CCDDBCopyDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CCDDBCopyDlg();

// Dialog Data
	enum { IDD = IDD_CDDBCOPY };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

private:
	BOOL AddDatabasesToList(CString szDir, CString szExt);
public:
	// Label that is set to the status of the user's Internet connectivity
	CStatic m_labInetStatus;
	afx_msg void CCDDBCopyDlg::OnTimer(UINT_PTR nIDEvent);
	// The number of dots in the elipsis after the Checking for Internet Connectivity message
	int m_DotCount;
	afx_msg void OnStnClickedInetstat();
	CString m_labInetStat;
	// Pointer to the App class
	CGenMAPPDBDLApp* m_pApp;
	// Call to indicate that we were unable to connect to GenMAPP.org
	CStatic m_ctlInetStat;
	afx_msg void OnBnClickedGenedir();
protected:
	CString m_txtGDB;
	CString m_txtMAPP;
	CString m_txtGEX;
public:
	afx_msg void OnBnClickedMappdir();
	afx_msg void OnBnClickedGexdir();
	CButton btnNext;
	CButton myButton;
	CString szConnSpeed;
//	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg void OnPaint();
	BOOL UpdateDlgControls(bool bSave);
private:
	// The number of items in the list box the last time the IDC_DBDLLIST timer was fired.
	int m_nLastListCount;
public:
	// Create the main List View control which displays the database files.
	bool CreateDBListCtrl(void);
	afx_msg void OnBnClickedOk();
	CStatic SelectCount;
	CStatic DLTime;
	CStatic m_labTotalDLSize;
	CButton btnBack;
private:
public:
	afx_msg void OnActivate(UINT nState, CWnd* pWndOther, BOOL bMinimized);
private:
	bool m_bFirstActivate;
public:
	// Retrieve a string suitable for display that shows the free hard drive space for all of the fixed drives on the system.
	CString GetFreeSpaceString();
	CListCtrl DBList;
	afx_msg void OnNMCustomdrawDbdllist(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedCancel();
	// TreeView control that displays database folders
	CTreeCtrl DBFolderTree;
	// Icons used in the database folder tree control
	CImageList DBFolderTreeIcons;
	CListBox DLQueueList;
	afx_msg void OnTvnSelchangedDbfoldertree(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMClickDbdllist(NMHDR *pNMHDR, LRESULT *pResult);
	// Builds the Database Queue List when an item has been selected or removed
	BOOL UpdateDLQueueList(void);
//	afx_msg void OnBnClickedGexdir2();
	afx_msg void OnBnClickedOtrdir();
	CString m_txtOtr;
	afx_msg void OnBnClickedShowreadme();
	CButton m_btnShowReadMe;
	CStatic m_labReadMeText;
	CStatic m_frmSelectSum;
	CStatic m_frmConnStat;
	CButton m_btnRemoveDB;
	afx_msg void OnBnClickedRemovedb();
	afx_msg void OnLbnSelchangeDlqueue();
	CStatic m_icoStatusIndicator;
	afx_msg void OnBnClickedRefresh();
	// Refresh button for re-reading folder structure and files from data providers
	CButton btnRefresh;
	CListBox m_lstFreeSpace;
private:
	// Verifies that paths specifed in the four folder text boxes are valid
	bool AreGenMAPPFoldersValid(void);
public:
	afx_msg void OnBnClickedBtnadvdlg();
	// Adjusts the dialog size to fit the screen resoultion. Also stores the width and height when finished.
	void SizeDialogToScreen(void);
	// Code to scale the sizes and positions of the controls based on the dialog box size
	void DoControlScaling(void);
	CRect AdjustToDlgCoordinates(CWnd* pWindowToAdjust);
	CStatic lblTopMessage;
	afx_msg void OnSize(UINT nType, int cx, int cy);
private:
	bool m_bOKToSize;
public:
	CStatic grpDLDestinations;
	CStatic grpOptions;
	CButton btnAdvanced;
	CStatic lblConnSpeed;
	CStatic lblGeneDB;
	CEdit txtGeneDB;
	CButton btnGeneDBBrowse;
	CStatic lblExpData;
	CEdit txtExpData;
	CButton btnExpDataBrowse;
	CStatic lblMAPPArch;
	CEdit txtMAPPArch;
	CButton btnMappArchBrowse;
	CStatic lblOtrInfo;
	CEdit txtOtrInfo;
	CButton btnOtrInfoBrowse;
	int nMinDialogHeight;
	int nMinDialogWidth;
	int nOrigSelSumHeight;
	CStatic lblSelCountKey;
	CStatic lblDLTimeKey;
	CStatic lblReqDiskSpaceKey;
	CStatic lblFreeSpaceKey;
	CStatic lblDLQueueKey;
	CStatic lblDLSizeKey;
	CStatic lblDLSizeVal;
	afx_msg void OnGetMinMaxInfo(MINMAXINFO* lpMMI);
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
private:
	BOOL m_bInResizeZone;
	BOOL m_bCapturingMouse;
public:
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg BOOL OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message);
};

#endif