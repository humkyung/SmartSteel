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

// stdafx.cpp : source file that includes just the standard includes
// SmartSteel.pch will be the pre-compiled header
// stdafx.obj will contain the pre-compiled type information

#include "stdafx.h"


/******************************************************************************
    @author     humkyung
    @date       2011-09-07
    @class
    @function   GetSDIActiveDocument
    @return     CSmartSteelDoc*
    @brief
******************************************************************************/
CSmartSteelDoc* GetSDIActiveDocument()
{
	CSmartSteelDoc* pDoc = NULL;

	CWnd* pWndMain = AfxGetMainWnd();
	ASSERT(pWndMain);
	ASSERT(pWndMain->IsKindOf(RUNTIME_CLASS(CFrameWndEx)) && 
		!pWndMain->IsKindOf(RUNTIME_CLASS(CMDIFrameWnd))); /// Not an MDI app.

	pDoc = (CSmartSteelDoc*)(((CFrameWnd*)pWndMain)->GetActiveDocument());

	return pDoc;
}