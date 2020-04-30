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

// SmartSteelDoc.h : interface of the CSmartSteelDoc class
//


#pragma once

#include "stdafx.h"
#include <IsVolume.h>
#include <occlib.h>
#include <SDNFFile.h>
#include <SDNFLinearMember.h>
#include <SDNFLinearMemberPacket.h>
#include <OCC_3dBaseDoc.h>
#include <ComplexShapeEntity.h>

#include "SteelConnPoint.h"
#include "ConnPntGenerator.h"
#include "SteelSectionBuilder.h"

#include "GussetPlate.h"
#include "EndPlate.h"
#include "SDNFAttribute.h"

#include <map>
#include <unordered_map>
using namespace std;

class CSmartSteelDoc : public OCC_3dBaseDoc
{
	enum
	{
		NONE_OP			= 0x00,
		FILE_OPEN		= 0x01,
		APP_GENERATE	= 0x02
	};

protected: // create from serialization only
	CSmartSteelDoc();
	DECLARE_DYNCREATE(CSmartSteelDoc)

// Attributes
public:
	int m_iAppOperation;

	CSDNFFile m_sdnf;
	tr1::unordered_map<STRING_T,OCC::CComplexShapeEntity* > m_oLinearShapeMap;
	list<OCC::CComplexShapeEntity* > m_oSteelShapeList;
	static map<STRING_T , CGussetPlate::Param* > m_oGussetPlateParamMap;
	static map<STRING_T , CEndPlate::Param* > m_oEndPlateParamMap;
	static map<STRING_T , CSteelSectionBuilder::ShapeParam* > m_oShapeParamMap;
private:
	CString m_sIniFilePath;

	friend class CConnPntGenerator;
	friend class CGussetPlateGenerator;
// Operations
public:
	/// zoom given entity
	int ZoomOCCEntity(CSteelPlate*);
	int ZoomOCCEntity(CSDNFLinearMember*);

	/**
	@brief	delete doc items
	*/
	void DeleteContents();

	/**
	@brief	return end parameter corresponding to given section name
	*/
	CEndPlate::Param* GetEndPlateParamOf(const STRING_T& sSectionName) const;

	/**
	@brief	return shape parameter corresponding to given section name
	*/
	CSteelSectionBuilder::ShapeParam* GetShapeParamOf(const STRING_T& sSectionName) const;

	/// 2013.06.26 added by humkyung
	double GetPlateThicknessOf(const STRING_T& sSectionName) const;
	/// up to here
	CGussetPlate::Param* GetSteelValue(const CSDNFLinearMember::ElmType& type , const STRING_T& section/* , const STRING_T& sKey*/);
	double GetSteelJointLength(const CSDNFLinearMember::ElmType& type , const STRING_T& section);
	CSteelPlate* GetPlateFrom(const TopoDS_Shape& selectedShape);
	OCC::CComplexShapeEntity* GetOCCEntityFrom(const TopoDS_Shape& selectedShape);
	/// @brief	remove occ entity associated to given shape
	int RemoveOCCEntity(const TopoDS_Shape& selectedShape);

	int UpdateAllViews(const bool& bFitView);
// Overrides
public:
	virtual BOOL OnNewDocument();
	virtual void Serialize(CArchive& ar);

// Implementation
public:
	virtual ~CSmartSteelDoc();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

private:
	/// load gusset/end plate from saved file - 2014.09.03 added by humkyung
	int LoadPlate(const CString& sFilePath);

	int CheckLinearMember();

	/// delete gusset,end plate
	void DeletePlateContents();

	/// do post process - check maximum edge length and end plate overlapping
	int DoPostProcess();
	/// get volume of linear member
	CIsVolume GetVolumeOf(CSDNFLinearMember* pLinearMem);
	static UINT StatusThreadEntry(LPVOID pParam);
	UINT ThreadEntry(LPVOID pParam );

	///@brief	load setting data
	int LoadSettingData();

	/// create a new connection point - 2014.08.15 added by humkyung
	CSteelConnPoint* CreateConnPoint(const CSteelConnPoint::ConnType& iType , const CIsPoint3d&);

	/// generate connection point
	int GenerateConnPoint();
	/// 2013.06.27 added by humkyung
	STRING_T GetColorStringOf(const CSDNFLinearMember::ElmType&) const;
	/// up to here
	/// 2013.06.26 added by humkyung
	void SendKeyIn(const STRING_T& keyin, HWND fwpWnd = NULL, CString findCaption = _T(""));
	bool SendCommand(const STRING_T& sDatFilePath , const STRING_T& sConfigFilePath);
	/// up to here
	bool CheckDuplicate(const int& iType , const CIsPoint3d& pt , CSDNFLinearMember* pMem1 = NULL , CSDNFLinearMember* pMem2 = NULL);	/// 2013.06.19 added by humkyung
// Generated message map functions
protected:
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnFileOpen();
	afx_msg void OnUpdateFileSave(CCmdUI* pCmdUI);	/// 2014.09.03 added by humkyung
	afx_msg void OnFileSave();						/// 2014.09.03 added by humkyung
	afx_msg void OnAppGenerate();
	afx_msg void OnUpdateAppGenerate(CCmdUI* pCmdUI);
#ifdef SMART_STEEL
	afx_msg void OnExportBREP();
#endif
	afx_msg void OnAppExport();		/// 2013.06.25 added by humkyung
	afx_msg void OnUpdateAppExport(CCmdUI* pCmdUI);
	afx_msg void OnAppOptions();	/// 2013.07.01 added by humkyung
	afx_msg void OnUpdateOCCViewShaded(CCmdUI* pCmdUI);
	afx_msg void OnOCCViewShaded();
	afx_msg void OnUpdateOCCViewWireFrame(CCmdUI* pCmdUI);
	afx_msg void OnOCCViewWireframe();
private:
	AIS_DisplayMode m_eDisplayMode;	/// 2014.08.18 added by humkyung

	map<CSDNFLinearMember::ElmType , list<CSDNFLinearMember*>* > m_oSDNFElmMap;
	list<CSteelConnPoint*> m_oConnPntList;
	list<CGussetPlate*> m_oGussetPlateList;
	list<CEndPlate*> m_oEndPlateList;
	list<CSDNFAttribute*> m_oAttrList;
public:
	virtual BOOL OnOpenDocument(LPCTSTR lpszPathName);
};


