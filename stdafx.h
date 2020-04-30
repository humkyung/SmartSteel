// This MFC Samples source code demonstrates using MFC Microsoft Office Fluent User Interface 
// (the "Fluent UI") and is provided only as referential material to supplement the 
// Microsoft Foundation Classes Reference and related electronic documentation 
// included with the MFC C++ library software.  
// License terms to copy, use or distribute the Fluent UI are available separately.  
// To learn more about our Fluent UI licensing program, please visit 
// http://msdn.microsoft.com/officeui.
//
// Copyright (C) Microsoft Corporation
// All rights reserved.

// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently,
// but are changed infrequently

#pragma once

#ifndef _SECURE_ATL
#define _SECURE_ATL 1
#endif

#ifndef VC_EXTRALEAN
#define VC_EXTRALEAN            // Exclude rarely-used stuff from Windows headers
#endif

#include "targetver.h"

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS      // some CString constructors will be explicit

// turns off MFC's hiding of some common and often safely ignored warning messages
#define _AFX_ALL_WARNINGS

#include <afxwin.h>         // MFC core and standard components
#include <afxext.h>         // MFC extensions


#include <afxdisp.h>        // MFC Automation classes



#ifndef _AFX_NO_OLE_SUPPORT
#include <afxdtctl.h>           // MFC support for Internet Explorer 4 Common Controls
#endif
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>             // MFC support for Windows Common Controls
#endif // _AFX_NO_AFXCMN_SUPPORT

#include <afxcontrolbars.h>     // MFC support for ribbons and control bars

#define _CRTDBG_MAP_ALLOC
#include "crtdbg.h"
//#ifdef _DEBUG
//	#ifndef DBG_NEW      
//		#define DBG_NEW new ( _NORMAL_BLOCK , __FILE__ , __LINE__ )
//		#define new DBG_NEW   
//	#endif
//#endif  // _DEBUG

//#include <tbb\tbbmalloc_proxy.h>

#include <afxsock.h>	/// for CSocket

#ifdef	SMART_STEEL
#include <IsGuiInf.h>
#include <Splash/SplashScreenFx.h>
//#include "Socket/ClientSocket.h"

#define PRODUCT_PUBLISHER	_T("TechSun")
#define PRODUCT_NAME		_T("SmartSteel")
#else
#define PRODUCT_PUBLISHER	_T("Daelim")
#define PRODUCT_NAME		_T("ezGussetPlate")
#endif

/// unit is mm
#define	CONN_PNT_TOLER		((UNIT::M == docData.m_oPlateCfg.unit_) ? 0.001 : 1)
#define	EXTEND_OFFSET		((UNIT::M == docData.m_oPlateCfg.unit_) ? 1 : 1000)
#define	MAXIMUM_EDGE_LEGNTH	((UNIT::M == docData.m_oPlateCfg.unit_) ? docData.m_oPlateCfg.max_edge_length_to_merge_*0.001 : docData.m_oPlateCfg.max_edge_length_to_merge_)
#define	SHAPE_LEN_OFFSET	((UNIT::M == docData.m_oPlateCfg.unit_) ? 0.015 : 15)
#define	SHAPE_START_LEN		((UNIT::M == docData.m_oPlateCfg.unit_) ? 0.05 : 50)
/// up to here

#define	DISPLAY_PROGRESSBAR	(WM_USER + 102)
#define	DISPLAY_MESSAGE		(WM_USER + 103)
#define	DISPLAY_STATUSBAR	(WM_USER + 104)

typedef enum
{
	MESSAGE_WARNING	= 0x01,
	MESSAGE_INFO	= 0x02
}MessageType;

#include <IsTools.h>
///#include <BugTrap.h>

#include "SmartSteelDoc.h"

#define	BEAM_ICON					4
#define	HBRACE_ICON					5
#define	GUSSET_PLATE_ICON			6
#define	END_PLATE_ICON				7
#define	GUSSET_PLATE_DELETED_ICON	8
#define	END_PLATE_DELETED_ICON		9

typedef enum
{
	M = 0x01,
	MM = 0x02
}UNIT;

typedef struct
{
	int class_;
	bool generate_for_web_type_brace;
	bool generate_endplate_depend_on_beam_length;	/// 2013.10.24 added by humkyung
	STRING_T grade_;
	STRING_T gusset_plate_display_color_;
	STRING_T end_plate_display_color_;
	UNIT unit_;	/// 2014.02.08 added by humkyung
	long max_edge_length_to_merge_;	/// 2014.02.14 added by humkyung
}PlateCfg;

extern CSmartSteelDoc* GetSDIActiveDocument();
extern CString GetExecPath();

#ifdef _UNICODE
#if defined _M_IX86
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_IA64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='ia64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif
#else
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif


