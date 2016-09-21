/*--------------------------------------------------------------------------------------+
|
|
|  $Copyright: (c) 2010 Bentley Systems, Incorporated. All rights reserved. $
|
|  Limited permission is hereby granted to reproduce and modify this copyrighted
|  material provided that the resulting code is used only in conjunction with
|  Bentley Systems products under the terms of the license agreement provided
|  therein, and that this notice is retained in its entirety in any such
|  reproduction or modification.
|
+--------------------------------------------------------------------------------------*/

// stdafx.h : include file for standard system include files,
//  or project specific include files that are used frequently, but
//      are changed infrequently
//

#if !defined(AFX_STDAFX_H__473CE3B1_AC5D_4142_A6B9_C95BC810A8E5__INCLUDED_)
#define AFX_STDAFX_H__473CE3B1_AC5D_4142_A6B9_C95BC810A8E5__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0501
#endif

#define VC_EXTRALEAN            // Exclude rarely-used stuff from Windows headers

#include <afxwin.h>         // MFC core and standard components
#include <afxext.h>         // MFC extensions

#ifndef _AFX_NO_OLE_SUPPORT
#include <afxole.h>         // MFC OLE classes
#include <afxodlgs.h>       // MFC OLE dialog classes
#include <afxdisp.h>        // MFC Automation classes
#endif // _AFX_NO_OLE_SUPPORT


#ifndef _AFX_NO_DB_SUPPORT
#include <afxdb.h>                      // MFC ODBC database classes
#endif // _AFX_NO_DB_SUPPORT

#ifndef _AFX_NO_DAO_SUPPORT
#include <afxdao.h>                     // MFC DAO database classes
#endif // _AFX_NO_DAO_SUPPORT

#include <afxdtctl.h>           // MFC support for Internet Explorer 4 Common Controls
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>                     // MFC support for Windows Common Controls
#endif // _AFX_NO_AFXCMN_SUPPORT


//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__473CE3B1_AC5D_4142_A6B9_C95BC810A8E5__INCLUDED_)
