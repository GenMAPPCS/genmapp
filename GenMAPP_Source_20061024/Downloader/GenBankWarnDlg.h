#pragma once
#include "afxcmn.h"
#include "afxwin.h"


// CGenBankWarnDlg dialog

class CGenBankWarnDlg : public CDialog
{
	DECLARE_DYNAMIC(CGenBankWarnDlg)

public:
	CGenBankWarnDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CGenBankWarnDlg();

// Dialog Data
	enum { IDD = IDD_GENBANKWARNING };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	CButton m_chkDontShow;
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedOk();
};
