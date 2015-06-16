
#include "StdAfx.h"
#include "GMFTP.h"
#include "GenMAPPDBDL.h"
#include ".\configfile.h"

CConfigFile::CConfigFile(CString szDLLPath)
: m_szCFGPath(_T(""))
{

	if (szDLLPath == "")
		m_szCFGPath = ".\\GenMAPP.cfg";
	else
		m_szCFGPath.Format("%sGenMAPP.cfg", szDLLPath);
}

// Call this constructor to read a config file other than GenMAPP.cfg
CConfigFile::CConfigFile(CString szDLLPath, CString szFileName)
: m_szCFGPath(_T(""))
{

	if (szDLLPath == "")
		m_szCFGPath = ".\\" + szFileName;
	else
		m_szCFGPath.Format("%s%s", szDLLPath, szFileName);
}

CConfigFile::~CConfigFile(void)
{
	if (m_hFile != CFile::hFileNull)
		Close();
}


/*
		public static string ReadCfgLine(string strKey)
		{
			if (PreviouslyReadValues.ContainsKey(strKey))
				return (string)PreviouslyReadValues[strKey];
			
			using (StreamReader sr = new StreamReader("PicSuck.ini")) 
			{
				string strLine;
				while ((strLine = sr.ReadLine()) != null) 
				{
					if (strLine.IndexOf("=") != -1 && 
						strLine.Substring(0, strLine.IndexOf("=")) == strKey)
					{
						string strValue = strLine.Substring(strLine.IndexOf("=") + 1);
						Logger.Log("ReadCfg-\tRead " + strKey + " from Ini File. Value is " + strValue);
						PreviouslyReadValues.Add(strKey, strValue);
						return strValue;
					}
				}
			}

			Logger.Log("ReadCfg-\tCouldn't find " + strKey);

			return "";
		}

		public static bool WriteCfgLine(string strKey, string strValue)
		{
			if (PreviouslyReadValues.ContainsKey(strKey))
				PreviouslyReadValues.Remove(strKey);

			bool bFoundLine = false;
			ArrayList IniFile = new ArrayList();
			using (StreamReader sr = new StreamReader("PicSuck.ini")) 
			{
				string strLine;
				while ((strLine = sr.ReadLine()) != null)
				{
					if (strLine.IndexOf("=") != -1 && 
						strLine.Substring(0, strLine.IndexOf("=")) == strKey)
					{
						IniFile.Add(string.Format("{0}={1}", strKey, strValue));
						bFoundLine = true;
					}
					else
						IniFile.Add(strLine);
				}
			}

			if (!bFoundLine)
				IniFile.Add(string.Format("{0}={1}", strKey, strValue));
				
			using (StreamWriter sw = new StreamWriter("PicSuck.ini", false)) 
			{
				foreach (object strLine in IniFile)
					sw.WriteLine((string)strLine);
			}

			return true;
		}
	}

*/


// Read a string value from the configuration file,
// without regard to section name
CString CConfigFile::ReadStringKey(CString szKeyName)
{
	return ReadStringKey(szKeyName, "");
}

// Read a string value from the configuration file
CString CConfigFile::ReadStringKey(CString szKeyName, CString szSectionName)
{
	CFileException ex;
	if (!Open(m_szCFGPath, CFile::modeRead | CFile::shareDenyWrite, &ex))
	{
		TCHAR szError[1024];
		ex.GetErrorMessage(szError, 1024);
		TRACE("Couldn't open source file: %s\n", szError);
		return "";
	}

	if (m_hFile != CFile::hFileNull)
	{
		bool	bFoundLine = false;
		CString	szCfgLine;

		if (szSectionName != "")
			FindSection(szSectionName);
		else
			SeekToBegin();

		while (!bFoundLine)
		{
			if (!ReadString(szCfgLine))
				break;

			if (szSectionName != "" && szCfgLine.Find('[') != -1)
				break;

			if (szCfgLine.Find('%') != -1)
				szCfgLine = ExpandReplaceableParam(szCfgLine);

			if (szCfgLine.Left(szCfgLine.Find(":")) == szKeyName)
			{
				bFoundLine = true;
				Close();
				return szCfgLine.Right(szCfgLine.GetLength() - (szCfgLine.Find(":") + 2));
			}
		}
	}

	Close();
	return "";
}




// Write a key value in the GenMAPP configuration file (without regard to section name).
bool CConfigFile::WriteStringKey(CString szKeyName, CString szValue)
{
	return WriteStringKey(szKeyName, szValue, "");
}

bool CConfigFile::WriteStringKey(CString szKeyName, CString szValue, CString szSection)
{
	CFileException ex;
	if (!Open(m_szCFGPath, CFile::modeReadWrite | CFile::shareDenyWrite, &ex))
	{
		TCHAR szError[1024];
		ex.GetErrorMessage(szError, 1024);
		TRACE("Couldn't open source file: %s\n", szError);
		return false;
	}

	// Walk through file
	// Are we in section?
	//    Yes- Is this the line?
	//		Yes- Change it
	//		No- Is this end of section?
	//			Yes- Add new key value
	// Add read-in key value
	CStringList CfgFileContents;
	if (m_hFile != CFile::hFileNull)
	{
		CString	szCfgLine;
		bool bInSection = false, bWroteLine = false, bFoundSection = false;
		while (ReadString(szCfgLine))
		{
			if (!bInSection && szSection != "")
				bFoundSection = bInSection = (szCfgLine.Left(szSection.GetLength()) == szSection);

			if (bInSection || szSection == "")
			{
				if (szCfgLine.Left(szKeyName.GetLength() + 1) == szKeyName + ":")
				{
					szCfgLine = szKeyName + ": " + szValue;
					bWroteLine = true;
				}

				if (!bWroteLine && szCfgLine.Left(1) == "[" && !(szCfgLine.Left(szSection.GetLength()) == szSection))
				{
					CfgFileContents.AddTail(szKeyName + ": " + szValue);
					bWroteLine = true;
					bInSection = false;
				}
			}

			CfgFileContents.AddTail(szCfgLine);
		}

		if (!bWroteLine)
		{
			if (szSection != "" && !bFoundSection)
				CfgFileContents.AddTail(szSection);

			CfgFileContents.AddTail(szKeyName + ": " + szValue);
		}
	}

	Close();

	if (!Open(m_szCFGPath, CFile::modeCreate | CFile::modeReadWrite | CFile::shareDenyRead | CFile::shareDenyWrite, &ex))
	{
		TCHAR szError[1024];
		ex.GetErrorMessage(szError, 1024);
		TRACE("Couldn't open source file: %s\n", szError);
		return false;
	}

	POSITION pos;
	for( pos = CfgFileContents.GetHeadPosition(); pos != NULL; )
	{
		CString		szCfgLine;
		szCfgLine = CfgFileContents.GetNext(pos);
		WriteString(szCfgLine + "\n");
	}
	Flush();

	Close();

	return true;
}

/*
// Write a key value in the GenMAPP configuration file under specified section.
// If section does not exist, it is created at the end of the file.
bool CConfigFile::WriteStringKey(CString szKeyName, CString szValue, CString szSection)
{
	CFileException ex;
	if (!Open(m_szCFGPath, CFile::modeReadWrite | CFile::shareDenyRead | CFile::shareDenyWrite, &ex))
	{
		TCHAR szError[1024];
		ex.GetErrorMessage(szError, 1024);
		TRACE("Couldn't open source file: %s\n", szError);
		return false;
	}

	if (m_hFile != CFile::hFileNull)
	{
		BYTE*			pSrcBuffer = new BYTE[(size_t)GetLength()], *pDestBuffer = new BYTE[(size_t)GetLength() + 1024];
		UINT			nBytesRead = 0, nBufPos = 0, nDestBufPos = 0;
		bool			bFoundKey = false, bFoundSect = false;
		CFileException	ex;

		if (szSection != "")
		{
			if (!FindSection(szSection))
			{
				SeekToBegin();
				nBytesRead = Read(pSrcBuffer, (UINT)GetLength());
				nBufPos = nBytesRead;//(UINT)GetPosition();
				nDestBufPos = nBytesRead;//(UINT)GetPosition();
			}
			else
			{
				UINT nAlmostSectionPos = (UINT)GetPosition();
				SeekToBegin();
				nBytesRead = Read(pSrcBuffer, (UINT)GetLength());
				nBufPos = nAlmostSectionPos;
				for (;pSrcBuffer[nBufPos] != '\n';nBufPos--);
				nDestBufPos = nBufPos;
				bFoundSect = true;
			}

			memcpy(pDestBuffer, pSrcBuffer, nBufPos);
		}
		else
			nBytesRead = Read(pSrcBuffer, (UINT)GetLength());

		if (nBytesRead > 0)
		{
			// Look for key
			for (;nBufPos < nBytesRead;nBufPos++, nDestBufPos++)
			{
				// If end of line, look for the key in that line
				if (pSrcBuffer[nBufPos] == '\n')
				{
					// Did we encounter the next section?
					if (pSrcBuffer[nBufPos + 1] == '[')
					{
						// Yes: We must not have found the key,
						// so write it out at the end of the section.
						char	szNewLine[256];

						bFoundKey = true;

						wsprintf(szNewLine, "%s: %s\n\n", szKeyName, szValue);
						memcpy(&pDestBuffer[nDestBufPos], szNewLine, strlen(szNewLine) + 1);
						nDestBufPos += (UINT)strlen(szNewLine) - 1;

						memcpy(&pDestBuffer[nDestBufPos], &pSrcBuffer[nBufPos], nBytesRead - nBufPos);
						nDestBufPos += nBytesRead - nBufPos;
						break;
					}

					// If the length of the line is shorter than the key we're looking
					// for, the line doesn't contain the key
					if (szKeyName.GetLength() + nBufPos + 1 < nBytesRead)
					{
						if (!memcmp(&pSrcBuffer[nBufPos + 1], szKeyName, szKeyName.GetLength()))
						{
							// Key has been found, modify the value
							char	szNewLine[256];
							BYTE	*pSrcNextLine;

							bFoundKey = true;

							wsprintf(szNewLine, "\n%s: %s\n", szKeyName, szValue);
							memcpy(&pDestBuffer[nDestBufPos], szNewLine, strlen(szNewLine) + 1);
							nDestBufPos += (UINT)strlen(szNewLine) - 1;

							pSrcNextLine = (BYTE*)memchr(&pSrcBuffer[nBufPos + 1], '\n', 256);
							nBufPos += (UINT)(pSrcNextLine - &pSrcBuffer[nBufPos]);
							memcpy(&pDestBuffer[nDestBufPos], &pSrcBuffer[nBufPos], nBytesRead - nBufPos);
							nDestBufPos += nBytesRead - nBufPos;
							break;
						}
						else
							pDestBuffer[nDestBufPos] = pSrcBuffer[nBufPos];
					}
				}
				else
					pDestBuffer[nDestBufPos] = pSrcBuffer[nBufPos]; // EOL not reached, just copy byte to dest.
			}
		}

		if (!bFoundKey)
		{
			// Key was not found, so add it to the config file
			char	szNewLine[256];

			pDestBuffer[nDestBufPos - 1] = pSrcBuffer[nBufPos - 1];

			if (szSection != "" && !bFoundSect)
			{
				//wsprintf(szNewLine, "%s\n", pDestBuffer[nDestBufPos - 1] == '\n' ? "" : "\n",
				//	szKeyName, szValue);
				wsprintf(szNewLine, "\n%s\n", szSection);
				memcpy(&pDestBuffer[nDestBufPos], szNewLine, strlen(szNewLine) + 1);
				nDestBufPos += (UINT)strlen(szNewLine) - 1;
			}

			wsprintf(szNewLine, "%s%s: %s\n", pDestBuffer[nDestBufPos - 1] == '\n' ? "" : "\n",
				szKeyName, szValue);
			memcpy(&pDestBuffer[nDestBufPos], szNewLine, strlen(szNewLine) + 1);
			nDestBufPos += (UINT)strlen(szNewLine) - 1;
		}

		Close();

		// Now write out the modified file currently in memory
		DeleteFile(m_szCFGPath);
		if (!Open(m_szCFGPath, CFile::modeReadWrite | CFile::shareDenyWrite | CFile::modeCreate, &ex))
		{
			TCHAR szError[1024];
			ex.GetErrorMessage(szError, 1024);
			TRACE("Couldn't open source file: %s\n", szError);
		}
		else
		{
			SeekToBegin();
			Write(pDestBuffer, nDestBufPos);
			Flush();
		}

		delete pSrcBuffer;
		delete pDestBuffer;
	}
	
	if (m_hFile != CFile::hFileNull)
	{
		Close();
		return true;
	}
	else
		return false;
}
*/
// Sets the file pointer to the line following the specified section.
// If section is not found, pointer will be set to the end of the file.
// This function requires that the section name be enclosed in brackets []
// Returns TRUE if found or FALSE if not.
BOOL CConfigFile::FindSection(CString szSectionName)
{
	CString	szSectLine;
	if (GetPosition() > 0)
		SeekToBegin();

	while (ReadString(szSectLine))
	{
		if (szSectLine.Left(szSectLine.Find("]") + 1) == szSectionName)
			return TRUE;
	}

	SeekToEnd();
	return FALSE;
}

// Returns a list of folders from DataLoacations.cfg that belong to
// the specified server. Make sure to free the memory associated
// with the strings in the list when you're finished using them.
CObList* CConfigFile::RetreiveServerEntries()
{
	CGenMAPPDBDLApp* pApp = (CGenMAPPDBDLApp*)AfxGetApp();
	CObList*	pServerList = NULL, *pReturnedServerList = NULL;
	CFileException ex;

	if (!Open(m_szCFGPath, CFile::modeRead | CFile::shareDenyWrite, &ex))
	{
		TCHAR szError[1024];
		ex.GetErrorMessage(szError, 1024);
		TRACE("Couldn't open source file: %s\n", szError);
		return NULL;
	}

	pServerList = new CObList();

	if (!pApp->m_bUseHTTP)
	{
		pReturnedServerList = RetrieveServerEntriesByType(false);
		if (pReturnedServerList != NULL)
		{
			pServerList->AddHead(pReturnedServerList);
			delete pReturnedServerList;
		}
	}

	pReturnedServerList = RetrieveServerEntriesByType(true);
	if (pReturnedServerList != NULL)
	{
		pServerList->AddTail(pReturnedServerList);
		delete pReturnedServerList;
	}

	Close();

	return pServerList;
}

// Either reads the [Servers] section or the [HTTP Servers] section of 
// DataLocations.cfg. Returns an CObList containing CGMFTP objects
// representing the servers found.
CObList* CConfigFile::RetrieveServerEntriesByType(bool bHTTP)
{
	CString		szCfgLine;
	CGMFTP*		pServer;
	CObList*	pServerList = NULL;
	CGenMAPPDBDLApp* pApp = (CGenMAPPDBDLApp*)AfxGetApp();

	FindSection(bHTTP ? "[HTTP Servers]" : "[Servers]");
	while (ReadString(szCfgLine))
	{
		int		nColon = -1;
		CString szTemp;

		if (szCfgLine.Find('[') != -1)
			break;

		nColon = szCfgLine.Find(':');
		if (nColon == -1)
			continue;

		if (szCfgLine.Find('|') == -1)
			continue;

		// Check for presence of undocumented "PreferredServer" GenMAPP.cfg entry. If found, only
		// retrieve info for this server.
		if (pApp->szPreferredServer != "" && pApp->szPreferredServer != szCfgLine.Left(nColon))
			continue;

		pServer = new CGMFTP();

		pServer->szServerKeyName = szCfgLine.Left(nColon);
		nColon+= 2;
		pServer->szHost = szCfgLine.Tokenize("|",nColon);
		pServer->szUserName = szCfgLine.Tokenize("|",nColon);
		pServer->szPassword = szCfgLine.Tokenize("|",nColon);
		pServer->szAlias = szCfgLine.Tokenize("|",nColon);
		pServer->szLocation = szCfgLine.Tokenize("|",nColon);
		pServer->szSpeedTestFolder = szCfgLine.Tokenize("|",nColon);
		pServer->szHTTPSpeedTestFolder = pServer->szSpeedTestFolder;
		szTemp = szCfgLine.Tokenize("|", nColon);
		if (bHTTP && szTemp != "" && szTemp.MakeUpper() != "NONE")
			pServer->bFTPBackup = true;
		pServer->bHTTPServer = bHTTP;

		pServer->nTier = atoi(szCfgLine.Tokenize("|",nColon));
		if (pServer->nTier == 0)
			pServer->nTier = 1;
		
		if (pServerList == NULL)
			pServerList = new CObList();
		pServerList->AddHead(pServer);

		
	}

	return pServerList;
}

// Returns a list of folders from DataLoacations.cfg that belong to
// the specified server. Make sure to free the memory associated
// with the strings in the list when you're finished using them.
CStringList* CConfigFile::RetreiveFoldersForServer(CString szServerName, bool bHTTP)
{
	CStringList*	pFolderList = new CStringList();
	CString			szCfgLine;
	CFileException	ex;

	if (!Open(m_szCFGPath, CFile::modeRead | CFile::shareDenyWrite, &ex))
	{
		TCHAR szError[1024];
		ex.GetErrorMessage(szError, 1024);
		TRACE("Couldn't open source file: %s\n", szError);
		return NULL;
	}

	FindSection(bHTTP ? "[HTTP Data Folders]" : "[Data Folders]");
	while (ReadString(szCfgLine))
	{
		int		nColon = -1, nPipe = -1;
		if (szCfgLine.Find('[') != -1)
			break;

		nColon = szCfgLine.Find(':');
		if (nColon == -1)
			continue;
		nColon+= 2;


		nPipe = szCfgLine.Find('|');
		if (nPipe == -1)
			continue;

		if (szCfgLine.Mid(nColon, nPipe - nColon) == szServerName)
		{
			CString szFolder;
			szFolder = szCfgLine.Right(szCfgLine.GetLength() - nPipe - 1);
			pFolderList->AddHead(szFolder);
		}
	}

	Close();

	return pFolderList;
}

// Returns the individual files associated with a server
BOOL CConfigFile::RetrieveFilesForServer(CString szServerName, CStringList* pFileList, CStringList* pFolderList, bool bHTTP)
{
	CString			szCfgLine;
	CFileException	ex;

	if (!Open(m_szCFGPath, CFile::modeRead | CFile::shareDenyWrite, &ex))
	{
		TCHAR szError[1024];
		ex.GetErrorMessage(szError, 1024);
		TRACE("Couldn't open source file: %s\n", szError);
		return FALSE;
	}

	FindSection(bHTTP ? "[HTTP Data Files]" : "[Data Files]");
	while (ReadString(szCfgLine))
	{
		int		nColon = -1, nPipe = -1;
		if (szCfgLine.Find('[') != -1)
			break;

		nColon = szCfgLine.Find(':');
		if (nColon == -1)
			continue;
		nColon+= 2;

		nPipe = szCfgLine.Find('|');
		if (nPipe == -1)
			continue;

		if (szCfgLine.Mid(nColon, nPipe - nColon) == szServerName)
		{
			CString szFile, szFolder;
			
			szFile = szCfgLine.Tokenize("|", nPipe);
			pFileList->AddHead(szFile);

			szFolder = szCfgLine.Tokenize("|", nPipe);
			pFolderList->AddHead(szFolder);
		}
	}

	Close();

	return TRUE;
}

// A Test...
bool CConfigFile::TestWriteInTheMiddle(CString szWhatToWrite)
{
	CFileException ex;
	char szCurrPath[_MAX_PATH];
	if (!Open(m_szCFGPath, CFile::modeReadWrite | CFile::shareDenyRead | CFile::shareDenyWrite, &ex))
	{
		TCHAR szError[1024];
		ex.GetErrorMessage(szError, 1024);
		TRACE("Couldn't open source file: %s\n", szError);
		return false;
	}
	
	GetCurrentDirectory(_MAX_PATH, szCurrPath);
	TRACE(szCurrPath);
	TRACE("\n");

	if (m_hFile != CFile::hFileNull)
	{
		CString AString;

		SeekToBegin();
		WriteString(szWhatToWrite);
		Flush();
		ReadString(AString);
		ReadString(AString);
		ReadString(AString);
		ReadString(AString);
		ReadString(AString);
		ReadString(AString);
		Close();
	}

	return true;
}

// Takes a config file line and expands any replaceable paramters found in the line
CString CConfigFile::ExpandReplaceableParam(CString szConfigLine)
{
	/*
	
	Here is the current list of replaceable paramters:
		%InstallDir%
		%BaseDataFolder%
		%OtherInfo%
		%GeneDatabases%
		%MAPPs%
		%ExpressionDatasets%
		%TempDir%
	*/

	CGenMAPPDBDLApp* pApp = (CGenMAPPDBDLApp*)AfxGetApp();
    int		nClosingPercentIndex = -1, nLastPercentIndex = szConfigLine.Find('%');
	CString	szExpandedLine = nLastPercentIndex == -1 ? 
		szConfigLine : szConfigLine.Left(nLastPercentIndex - 1);

	while (nLastPercentIndex != -1)
	{
		nClosingPercentIndex = szConfigLine.Find('%', nLastPercentIndex + 1);
		if (nClosingPercentIndex == -1)
		{
			nLastPercentIndex = -1;
			break;
		}

		CString		szParam = szConfigLine.Mid(nLastPercentIndex + 1, nClosingPercentIndex - nLastPercentIndex - 1);

		if (szParam == "InstallDir")
			szExpandedLine += pApp->m_szDLLPath;

		if (szParam == "BaseDataFolder")
			szExpandedLine += pApp->szBasePath;

		if (szParam == "OtherInfo")
			szExpandedLine += pApp->szOtrPath;

		if (szParam == "GeneDatabases")
			szExpandedLine += pApp->szGDBPath;

		if (szParam == "MAPPs")
			szExpandedLine += pApp->szMAPPPath;

		if (szParam == "ExpressionDatasets")
			szExpandedLine += pApp->szGEXPath;

		if (szParam == "TempDir")
		{
			char	szTempDir[_MAX_PATH];			
			CString	szCStringTempDir;
			GetTempPath(_MAX_PATH, szTempDir);
			szCStringTempDir.Format("%s", szTempDir);
			szExpandedLine += szCStringTempDir + "\\";
		}

		nLastPercentIndex = szConfigLine.Find('%', nClosingPercentIndex + 1);
	}

	return szExpandedLine += szConfigLine.Right(szConfigLine.GetLength() - nClosingPercentIndex - 1);
}

// Provides a means to iterate through a config file. Returns false if
// there are no more lines matching the criteria. Pass in an empty string for
// szSection to read all sections. Line numbers begin at 1.
bool CConfigFile::RetrieveNextKey(int nLineNumber, CString* pszKey, CString* pszValue, CString szSection)
{
	CFileException ex;
	if (!Open(m_szCFGPath, CFile::modeRead | CFile::shareDenyWrite, &ex))
	{
		TCHAR szError[1024];
		ex.GetErrorMessage(szError, 1024);
		TRACE("Couldn't open source file: %s\n", szError);
		return false;
	}

	if (m_hFile != CFile::hFileNull)
	{
		CString	szCfgLine;
		int		i;

		if (szSection != "")
			FindSection(szSection);
		else
			SeekToBegin();

		for (i=1;i<nLineNumber;i++)
		{
			if (!ReadString(szCfgLine))
			{
				Close();
				return false;
			}

			if (szSection != "")
			{
				if (szCfgLine.Left(1) == "[")
				{
					Close();
					return false;
				}
			}
		}

		if (!ReadString(szCfgLine) || szCfgLine.Left(1) == "[")
		{
			Close();
			return false;
		}


		if (szCfgLine.Find('%') != -1)
			szCfgLine = ExpandReplaceableParam(szCfgLine);

		if (szCfgLine.Left(szCfgLine.Find(": ")) != -1)
		{
			*pszKey = szCfgLine.Left(szCfgLine.Find(": "));
			*pszValue = szCfgLine.Right(szCfgLine.GetLength() - (szCfgLine.Find(": ") + 2));

			if (!ReadString(szCfgLine) || szCfgLine.Left(1) == "[")
			{
				Close();
				return false;
			}
			else
			{
				Close();
				return true;
			}
		}
		else
		{
			Close();		
			return false;
		}
	}

	return false;
}

