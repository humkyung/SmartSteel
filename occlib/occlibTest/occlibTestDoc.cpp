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

// occlibTestDoc.cpp : implementation of the CocclibTestDoc class
//

#include "stdafx.h"
#include "occlibTest.h"

#include "occlibTestDoc.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CocclibTestDoc

IMPLEMENT_DYNCREATE(CocclibTestDoc, OCC_3dBaseDoc)

BEGIN_MESSAGE_MAP(CocclibTestDoc, OCC_3dBaseDoc)
END_MESSAGE_MAP()


// CocclibTestDoc construction/destruction

CocclibTestDoc::CocclibTestDoc()
{
	// TODO: add one-time construction code here

}

CocclibTestDoc::~CocclibTestDoc()
{
}

BOOL CocclibTestDoc::OnNewDocument()
{
	if (!OCC_3dBaseDoc::OnNewDocument())
		return FALSE;

	// TODO: add reinitialization code here
	// (SDI documents will reuse this document)

	return TRUE;
}




// CocclibTestDoc serialization

void CocclibTestDoc::Serialize(CArchive& ar)
{
	if (ar.IsStoring())
	{
		// TODO: add storing code here
	}
	else
	{
		// TODO: add loading code here
	}
}


// CocclibTestDoc diagnostics

#ifdef _DEBUG
void CocclibTestDoc::AssertValid() const
{
	OCC_3dBaseDoc::AssertValid();
}

void CocclibTestDoc::Dump(CDumpContext& dc) const
{
	OCC_3dBaseDoc::Dump(dc);
}
#endif //_DEBUG


// CocclibTestDoc commands
