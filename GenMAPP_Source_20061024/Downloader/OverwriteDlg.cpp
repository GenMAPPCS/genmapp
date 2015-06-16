// OverwriteDlg.cpp : implementation file
//

#include "stdafx.h"
#include "GenMAPPDBDL.h"
#include "OverwriteDlg.h"
#include ".\overwritedlg.h"


// COverwriteDlg dialog

IMPLEMENT_DYNAMIC(COverwriteDlg, CDialog)
COverwriteDlg::COverwriteDlg(CWnd* pParent /*=NULL*/)
	: CDialog(COverwriteDlg::IDD, pParent)
	, pFileNameList(NULL)
{
	m_pApp = (CGenMAPPDBDLApp*)AfxGetApp();
}

COverwriteDlg::~COverwriteDlg()
{
}

void COverwriteDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LSTOVERWRITE, lstOverwriteFiles);
}


BEGIN_MESSAGE_MAP(COverwriteDlg, CDialog)
	ON_BN_CLICKED(IDOK, OnBnClickedOk)
	ON_BN_CLICKED(IDC_BTNYESTOALL, OnBnClickedBtnyestoall)
	ON_BN_CLICKED(IDCANCEL, OnBnClickedCancel)
	ON_BN_CLICKED(IDC_BTNNOTOALL, OnBnClickedBtnnotoall)
END_MESSAGE_MAP()


// COverwriteDlg message handlers

BOOL COverwriteDlg::OnInitDialog()
{
	POSITION pos;
	CDialog::OnInitDialog();
	SetIcon(LoadIcon(m_pApp->m_hInstance, MAKEINTRESOURCE(IDI_GENMAPP)), TRUE);
	for( pos = pFileNameList->GetHeadPosition(); pos != NULL; )
		lstOverwriteFiles.AddString(pFileNameList->GetNext(pos));

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

void COverwriteDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
	OnOK();
}

void COverwriteDlg::OnBnClickedBtnyestoall()
{
	EndDialog(10);
}

void COverwriteDlg::OnBnClickedCancel()
{
	// TODO: Add your control notification handler code here
	OnCancel();
}

void COverwriteDlg::OnBnClickedBtnnotoall()
{
	EndDialog(100);
}
