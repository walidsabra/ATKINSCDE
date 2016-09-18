// PWSDKMethods.cpp : Defines the exported functions for the DLL application.
//

#include "stdafx.h"
#include "PWSDKMethods.h"

#include <stdio.h>
#include <aawindms.h>
#include <aaapi.h>
#include <aadmsapi.h>
#include <aawinapi.h>
#include <tchar.h>
#include <stdlib.h>
#include <string>
#include "aaodsapi.h"
#include <algorithm>
#include <sstream>
#include <iostream>

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define DLLEXPORT __declspec( dllexport )
#define WINAPI __stdcall

using namespace std;

//aaApi_Login
//aaApi_CheckInDocument
extern "C" __declspec(dllexport) BOOL CheckIn(char* UserName, char* UserPassword, char* DSName, int ProjectID, int DocumentID)
{
	LONG validLogin = 0;
	BOOL valid = false;
	BOOL init = aaApi_Initialize(AAMODULE_ALL);

	// User Name
	WCHAR  userName[200];
	MultiByteToWideChar(0, 0, UserName, 200, userName, 200);
	LPCWSTR  User_Name = userName;

	// User Password
	WCHAR  userPassword[200];
	MultiByteToWideChar(0, 0, UserPassword, 200, userPassword, 200);
	LPCWSTR  User_Password = userPassword;

	// Current DataSource
	WCHAR  dsName[200];
	MultiByteToWideChar(0, 0, DSName, 200, dsName, 200);
	LPCWSTR  DS_Name = dsName;

	validLogin = aaApi_Login(AAAPIDB_UNKNOWN, DS_Name, User_Name, User_Password, NULL);

	if (validLogin == 0)
		MessageBoxA(NULL, "User Login Failed", "ATKINS CDE", MB_OK);
	else if (validLogin == 1)
		valid = aaApi_CheckInDocument(ProjectID, DocumentID);

	return valid;

	//return true;
}

