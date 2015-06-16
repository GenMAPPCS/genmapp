// GenBankWarnDlg.cpp : implementation file
//

#include "stdafx.h"
#include "GMFTP.h"
#include "GenMAPPDBDL.h"
#include "ConfigFile.h"
#include "GenBankWarnDlg.h"
#include ".\genbankwarndlg.h"


// CGenBankWarnDlg dialog

IMPLEMENT_DYNAMIC(CGenBankWarnDlg, CDialog)
CGenBankWarnDlg::CGenBankWarnDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CGenBankWarnDlg::IDD, pParent)
{
}

CGenBankWarnDlg::~CGenBankWarnDlg()
{
}

void CGenBankWarnDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_CHKDONTSHOW, m_chkDontShow);
}


BEGIN_MESSAGE_MAP(CGenBankWarnDlg, CDialog)
	ON_BN_CLICKED(IDOK, OnBnClickedOk)
END_MESSAGE_MAP()


// CGenBankWarnDlg message handlers
BOOL CGenBankWarnDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	return TRUE;
}
void CGenBankWarnDlg::OnBnClickedOk()
{
	CGenMAPPDBDLApp*	pApp = (CGenMAPPDBDLApp*)AfxGetApp();
	CConfigFile cf(pApp->m_szDLLPath);
	cf.WriteStringKey("ShowGenBankWarning", m_chkDontShow.GetCheck() ? "False" : "True");

	OnOK();
}
