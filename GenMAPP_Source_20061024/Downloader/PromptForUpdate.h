#pragma once


// CPromptForUpdate dialog

class CPromptForUpdate : public CDialog
{
	DECLARE_DYNAMIC(CPromptForUpdate)

public:
	CPromptForUpdate(CWnd* pParent = NULL);   // standard constructor
	virtual ~CPromptForUpdate();

// Dialog Data
	enum { IDD = IDD_PROMPTFORUPDATE };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
};
