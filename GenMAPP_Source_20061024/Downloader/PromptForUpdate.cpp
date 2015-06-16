// PromptForUpdate.cpp : implementation file
//

#include "stdafx.h"
#include "GenMAPPDBDL.h"
#include "PromptForUpdate.h"
#include ".\promptforupdate.h"


// CPromptForUpdate dialog

IMPLEMENT_DYNAMIC(CPromptForUpdate, CDialog)
CPromptForUpdate::CPromptForUpdate(CWnd* pParent /*=NULL*/)
	: CDialog(CPromptForUpdate::IDD, pParent)
{
}

CPromptForUpdate::~CPromptForUpdate()
{
}

void CPromptForUpdate::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CPromptForUpdate, CDialog)
	ON_BN_CLICKED(IDOK, OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, OnBnClickedCancel)
END_MESSAGE_MAP()


// CPromptForUpdate message handlers

BOOL CPromptForUpdate::OnInitDialog()
{
	CDialog::OnInitDialog();

	CGenMAPPDBDLApp* pApp = (CGenMAPPDBDLApp*)AfxGetApp();
	SetIcon(LoadIcon(pApp->m_hInstance, MAKEINTRESOURCE(IDI_GENMAPP)), TRUE);

	::SetForegroundWindow(this->GetSafeHwnd());

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

void CPromptForUpdate::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
	OnOK();
}

void CPromptForUpdate::OnBnClickedCancel()
{
	// TODO: Add your control notification handler code here
	OnCancel();
}
