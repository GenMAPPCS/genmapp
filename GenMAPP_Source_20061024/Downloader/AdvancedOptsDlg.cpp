// AdvancedOptsDlg.cpp : implementation file
//

#include "stdafx.h"
#include "GenMAPPDBDL.h"
#include "ConfigFile.h"
#include "AdvancedOptsDlg.h"
#include ".\advancedoptsdlg.h"


// CAdvancedOptsDlg dialog

IMPLEMENT_DYNAMIC(CAdvancedOptsDlg, CDialog)
CAdvancedOptsDlg::CAdvancedOptsDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CAdvancedOptsDlg::IDD, pParent)
	, chkDeleteCompFiles(TRUE)
	, chkOverwriteDataFiles(FALSE)
{
}

CAdvancedOptsDlg::~CAdvancedOptsDlg()
{
}

void CAdvancedOptsDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Check(pDX, IDC_CHKDELCOMPFILES, chkDeleteCompFiles);
	DDX_Check(pDX, IDC_CHKOVRWRITEDATA, chkOverwriteDataFiles);
	DDX_Control(pDX, IDC_CHKDELCOMPFILES, ctlDeleteCompressed);
	DDX_Control(pDX, IDC_CHKOVRWRITEDATA, ctlOverwriteData);
}


BEGIN_MESSAGE_MAP(CAdvancedOptsDlg, CDialog)
	ON_BN_CLICKED(IDOK, OnBnClickedOk)
	ON_WM_ACTIVATE()
END_MESSAGE_MAP()


// CAdvancedOptsDlg message handlers

void CAdvancedOptsDlg::OnBnClickedOk()
{
	CGenMAPPDBDLApp* pApp = (CGenMAPPDBDLApp*)AfxGetApp();
	CConfigFile cf(pApp->m_szDLLPath);

	UpdateData();

	pApp->bDeleteSFX = (bool)chkDeleteCompFiles;
	pApp->bOverwriteDataFiles = (bool)chkOverwriteDataFiles;
	cf.WriteStringKey("DeleteCompressedFiles", chkDeleteCompFiles ? "True" : "False", "[Downloader]");
	cf.WriteStringKey("OverwriteDataFiles", chkOverwriteDataFiles ? "True" : "False", "[Downloader]");

	OnOK();
}

BOOL CAdvancedOptsDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	CString szResult = "";
	CGenMAPPDBDLApp* pApp = (CGenMAPPDBDLApp*)AfxGetApp();
	CConfigFile cf(pApp->m_szDLLPath);

	szResult = cf.ReadStringKey("DeleteCompressedFiles", "[Downloader]");
	ctlDeleteCompressed.SetCheck((szResult.MakeUpper() != "FALSE"));

	szResult = cf.ReadStringKey("OverwriteDataFiles", "[Downloader]");
	ctlOverwriteData.SetCheck(!(szResult.MakeUpper() != "TRUE"));

	UpdateData();

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

void CAdvancedOptsDlg::OnActivate(UINT nState, CWnd* pWndOther, BOOL bMinimized)
{
	CDialog::OnActivate(nState, pWndOther, bMinimized);
}
