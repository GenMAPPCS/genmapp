#pragma once
#include "afxwin.h"


// COverwriteDlg dialog

class COverwriteDlg : public CDialog
{
	DECLARE_DYNAMIC(COverwriteDlg)

public:
	COverwriteDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~COverwriteDlg();

// Dialog Data
	enum { IDD = IDD_FILEOVERWRITE };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	CListBox lstOverwriteFiles;
	virtual BOOL OnInitDialog();
	CGenMAPPDBDLApp* m_pApp;

	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedBtnyestoall();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnBnClickedBtnnotoall();
	// This is a list of file names that will appear in the overwrite dialog's list box.
	CStringList* pFileNameList;
};
