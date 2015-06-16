#pragma once
#include "afxwin.h"


// CAdvancedOptsDlg dialog

class CAdvancedOptsDlg : public CDialog
{
	DECLARE_DYNAMIC(CAdvancedOptsDlg)

public:
	CAdvancedOptsDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CAdvancedOptsDlg();

// Dialog Data
	enum { IDD = IDD_ADVOPTIONS };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	BOOL chkDeleteCompFiles;
	BOOL chkOverwriteDataFiles;
	virtual BOOL OnInitDialog();
	afx_msg void OnActivate(UINT nState, CWnd* pWndOther, BOOL bMinimized);
	CButton ctlDeleteCompressed;
	CButton ctlOverwriteData;
};
