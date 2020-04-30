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

// SmartSteel.h : main header file for the SmartSteel application
//
#pragma once

#ifndef __AFXWIN_H__
	#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"       // main symbols
#include "singleinstance.h"
#include <OCC_3dAppEx.h>

// CSmartSteelApp:
// See SmartSteel.cpp for the implementation of this class
//

class CSmartSteelApp : public OCC_3dAppEx , public CSingleInstance
{
public:
	CSmartSteelApp();


// Overrides
public:
	virtual BOOL InitInstance();

// Implementation
	BOOL  m_bHiColorIcons;

	virtual void PreLoadState();
	virtual void LoadCustomState();
	virtual void SaveCustomState();

	afx_msg void OnFileOpen();
	afx_msg void OnAppAbout();
	DECLARE_MESSAGE_MAP()
private:
#ifdef	SMART_STEEL
	int AutoUpdate(void);
#endif
};

extern CSmartSteelApp theApp;
