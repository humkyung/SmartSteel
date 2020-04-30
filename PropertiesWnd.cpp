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

#include "stdafx.h"
#include <assert.h>
#include "PropertiesWnd.h"
#include "Resource.h"
#include "MainFrm.h"
#include "SmartSteel.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

/////////////////////////////////////////////////////////////////////////////
// CResourceViewBar

CPropertiesWnd::CPropertiesWnd()
{
}

CPropertiesWnd::~CPropertiesWnd()
{
}

BEGIN_MESSAGE_MAP(CPropertiesWnd, CDockablePane)
	ON_WM_CREATE()
	ON_WM_SIZE()
	ON_COMMAND(ID_EXPAND_ALL, OnExpandAllProperties)
	ON_UPDATE_COMMAND_UI(ID_EXPAND_ALL, OnUpdateExpandAllProperties)
	ON_COMMAND(ID_SORTPROPERTIES, OnSortProperties)
	ON_UPDATE_COMMAND_UI(ID_SORTPROPERTIES, OnUpdateSortProperties)
	ON_COMMAND(ID_PROPERTIES1, OnProperties1)
	ON_UPDATE_COMMAND_UI(ID_PROPERTIES1, OnUpdateProperties1)
	ON_COMMAND(ID_PROPERTIES2, OnProperties2)
	ON_UPDATE_COMMAND_UI(ID_PROPERTIES2, OnUpdateProperties2)
	ON_WM_SETFOCUS()
	ON_WM_SETTINGCHANGE()
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CResourceViewBar message handlers

void CPropertiesWnd::AdjustLayout()
{
	if (GetSafeHwnd() == NULL)
	{
		return;
	}

	CRect rectClient,rectCombo;
	GetClientRect(rectClient);

	int cyCmb = 1;
	int cyTlb = m_wndToolBar.CalcFixedLayout(FALSE, TRUE).cy;

	m_wndToolBar.SetWindowPos(NULL, rectClient.left, rectClient.top + cyCmb, rectClient.Width(), cyTlb, SWP_NOACTIVATE | SWP_NOZORDER);
	m_wndPropList.SetWindowPos(NULL, rectClient.left, rectClient.top + cyCmb + cyTlb, rectClient.Width(), rectClient.Height() -(cyCmb+cyTlb), SWP_NOACTIVATE | SWP_NOZORDER);
	m_wndPlatePropList.SetWindowPos(NULL, rectClient.left, rectClient.top + cyCmb + cyTlb, rectClient.Width(), rectClient.Height() -(cyCmb+cyTlb), SWP_NOACTIVATE | SWP_NOZORDER);
}

int CPropertiesWnd::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CDockablePane::OnCreate(lpCreateStruct) == -1)
		return -1;

	CRect rectDummy;
	rectDummy.SetRectEmpty();

	if (!m_wndPropList.Create(WS_VISIBLE | WS_CHILD, rectDummy, this, 2))
	{
		TRACE0("Failed to create Properties Grid \n");
		return -1;      // fail to create
	}
	if (!m_wndPlatePropList.Create(WS_CHILD, rectDummy, this, 3))
	{
		TRACE0("Failed to create Properties Grid \n");
		return -1;      // fail to create
	}

	InitPropList();

	m_wndToolBar.Create(this, AFX_DEFAULT_TOOLBAR_STYLE, IDR_PROPERTIES);
	m_wndToolBar.LoadToolBar(IDR_PROPERTIES, 0, 0, TRUE /* Is locked */);
	m_wndToolBar.CleanUpLockedImages();
	m_wndToolBar.LoadBitmap(theApp.m_bHiColorIcons ? IDB_PROPERTIES_HC : IDR_PROPERTIES, 0, 0, TRUE /* Locked */);

	m_wndToolBar.SetPaneStyle(m_wndToolBar.GetPaneStyle() | CBRS_TOOLTIPS | CBRS_FLYBY);
	m_wndToolBar.SetPaneStyle(m_wndToolBar.GetPaneStyle() & ~(CBRS_GRIPPER | CBRS_SIZE_DYNAMIC | CBRS_BORDER_TOP | CBRS_BORDER_BOTTOM | CBRS_BORDER_LEFT | CBRS_BORDER_RIGHT));
	m_wndToolBar.SetOwner(this);

	// All commands will be routed via this control , not via the parent frame:
	m_wndToolBar.SetRouteCommandsViaFrame(FALSE);

	AdjustLayout();
	return 0;
}

void CPropertiesWnd::OnSize(UINT nType, int cx, int cy)
{
	CDockablePane::OnSize(nType, cx, cy);
	AdjustLayout();
}

void CPropertiesWnd::OnExpandAllProperties()
{
	m_wndPropList.ExpandAll();
}

void CPropertiesWnd::OnUpdateExpandAllProperties(CCmdUI* pCmdUI)
{
}

void CPropertiesWnd::OnSortProperties()
{
	m_wndPropList.SetAlphabeticMode(!m_wndPropList.IsAlphabeticMode());
}

void CPropertiesWnd::OnUpdateSortProperties(CCmdUI* pCmdUI)
{
	pCmdUI->SetCheck(m_wndPropList.IsAlphabeticMode());
}

void CPropertiesWnd::OnProperties1()
{
	// TODO: Add your command handler code here
}

void CPropertiesWnd::OnUpdateProperties1(CCmdUI* /*pCmdUI*/)
{
	// TODO: Add your command update UI handler code here
}

void CPropertiesWnd::OnProperties2()
{
	// TODO: Add your command handler code here
}

void CPropertiesWnd::OnUpdateProperties2(CCmdUI* /*pCmdUI*/)
{
	// TODO: Add your command update UI handler code here
}

void CPropertiesWnd::InitPropList()
{
	SetPropListFont();

	{
		m_wndPropList.EnableHeaderCtrl(FALSE);
		m_wndPropList.EnableDescriptionArea();
		m_wndPropList.SetVSDotNetLook();
		m_wndPropList.MarkModifiedProperties();

		CMFCPropertyGridProperty* pGroup1 = new CMFCPropertyGridProperty(_T("SDNF Element"));

		pGroup1->AddSubItem(new CMFCPropertyGridProperty(_T("ID"), _T("") , _T("Member ID")));
		pGroup1->AddSubItem(new CMFCPropertyGridProperty(_T("Type"), _T("") , _T("SDNF Element Type")));
		pGroup1->AddSubItem(new CMFCPropertyGridProperty(_T("Section"), _T("") , _T("Section value")));
		pGroup1->AddSubItem(new CMFCPropertyGridProperty(_T("Grade"), _T("") , _T("Grade value")));
		pGroup1->AddSubItem(new CMFCPropertyGridProperty(_T("Cardinal"), _T("") , _T("Cardinal Point")));
		pGroup1->AddSubItem(new CMFCPropertyGridProperty(_T("Start"), _T("") , _T("start")));
		pGroup1->AddSubItem(new CMFCPropertyGridProperty(_T("End"), _T("") , _T("end")));
		pGroup1->AddSubItem(new CMFCPropertyGridProperty(_T("Rotation"), _T("") , _T("Rotation Angle in degree")));

		m_wndPropList.AddProperty(pGroup1);
	}

	{
		m_wndPlatePropList.EnableHeaderCtrl(FALSE);
		m_wndPlatePropList.EnableDescriptionArea();
		m_wndPlatePropList.SetVSDotNetLook();
		m_wndPlatePropList.MarkModifiedProperties();

		CMFCPropertyGridProperty* pGroup = new CMFCPropertyGridProperty(_T("Plate Element"));

		pGroup->AddSubItem(new CMFCPropertyGridProperty(_T("Type"), _T("") , _T("Plate Type(Gusset Plate/End Plate)")));
		pGroup->AddSubItem(new CMFCPropertyGridProperty(_T("NO"), _T("") , _T("number")));

		m_wndPlatePropList.AddProperty(pGroup);
	}
}

void CPropertiesWnd::OnSetFocus(CWnd* pOldWnd)
{
	CDockablePane::OnSetFocus(pOldWnd);
	m_wndPropList.SetFocus();
}

void CPropertiesWnd::OnSettingChange(UINT uFlags, LPCTSTR lpszSection)
{
	CDockablePane::OnSettingChange(uFlags, lpszSection);
	SetPropListFont();
}

void CPropertiesWnd::SetPropListFont()
{
	::DeleteObject(m_fntPropList.Detach());

	LOGFONT lf;
	afxGlobalData.fontRegular.GetLogFont(&lf);

	NONCLIENTMETRICS info;
	info.cbSize = sizeof(info);

	afxGlobalData.GetNonClientMetrics(info);

	lf.lfHeight = info.lfMenuFont.lfHeight;
	lf.lfWeight = info.lfMenuFont.lfWeight;
	lf.lfItalic = info.lfMenuFont.lfItalic;

	m_fntPropList.CreateFontIndirect(&lf);

	m_wndPropList.SetFont(&m_fntPropList);
}

/**
	@brief	fill properties of sdnf element attribute
	@author	humkyung
	@date	2013.06.11
*/
int CPropertiesWnd::FillPropertiesOf(CSDNFAttribute* pAttr)
{
	assert(pAttr && "pAttr is NULL");

	if(pAttr)
	{
		m_wndPlatePropList.ShowWindow(SW_HIDE);
		m_wndPropList.ShowWindow(SW_SHOW);

		CMFCPropertyGridProperty* pGroup = m_wndPropList.GetProperty(0);
		if(NULL != pGroup)
		{
			pGroup->GetSubItem(0)->SetValue(CString(pAttr->id().c_str()));
			pGroup->GetSubItem(1)->SetValue(CString(pAttr->type().c_str()));
			pGroup->GetSubItem(2)->SetValue(CString(pAttr->section().c_str()));
			pGroup->GetSubItem(3)->SetValue(CString(pAttr->grade().c_str()));
			
			CString str;
			str.Format(_T("%d") , pAttr->CardinalPnt());
			pGroup->GetSubItem(4)->SetValue(str);

			pGroup->GetSubItem(5)->SetValue(CString(pAttr->start().c_str()));
			pGroup->GetSubItem(6)->SetValue(CString(pAttr->end().c_str()));

			str.Format(_T("%lf") , pAttr->Rotation());
			pGroup->GetSubItem(7)->SetValue(str);
		}

		return ERROR_SUCCESS;
	}

	return ERROR_INVALID_PARAMETER;
}

/**
	@brief	fill properties of steelplate attribute
	@author	humkyung
	@date	2013.08.04
*/
int CPropertiesWnd::FillPropertiesOf(CSteelPlate* pPlate)
{
	assert(pPlate && "pPlate is NULL");

	if(pPlate)
	{
		m_wndPropList.ShowWindow(SW_HIDE);
		m_wndPlatePropList.ShowWindow(SW_SHOW);

		CMFCPropertyGridProperty* pGroup = m_wndPlatePropList.GetProperty(0);
		if(NULL != pGroup)
		{
			pGroup->GetSubItem(0)->SetValue(pPlate->GetTypeString().c_str());
			CString str;
			str.Format(_T("%d") , pPlate->id());
			pGroup->GetSubItem(1)->SetValue(str);
		}

		return ERROR_SUCCESS;
	}

	return ERROR_INVALID_PARAMETER;
}