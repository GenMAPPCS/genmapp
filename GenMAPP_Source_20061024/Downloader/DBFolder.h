#pragma once
#include "afxcoll.h"



// CDBFolder command target

class CDBFolder : public CObject
{
public:
	CDBFolder();
	virtual ~CDBFolder();
	// Handle to folder
	HTREEITEM hFolder;
	// Handle to parent folder
	HTREEITEM hParentFolder;
	// Name of folder, as read from the FTP server
	CString szFolderName;
	// List of DBFiles displayed when user clicks on folder
	CObList DBFileList;
	// Indicates which types of databases the folder holds. DB types are identified in GenMAPPDBDL.h.
	int nType;
};


