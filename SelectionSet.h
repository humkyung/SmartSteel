#pragma once

#include "stdafx.h"
#include <occlib.h>

#include <vector>
using namespace std;

class Handle_AIS_InteractiveContext;
class Handle_V3d_View;
class Handle(AIS_InteractiveObject);
class CSelectionSet
{
public:
	CSelectionSet(Handle_AIS_InteractiveContext& hAisContext, Handle_V3d_View& h3DView);
	~CSelectionSet(void);

	void HilightSelected(bool bUpdate);
	AIS_StatusOfPick Select(long x, long y);
	void DeSelectAll(bool bUpdate);
	void DynamicSense(long x, long y);

	void SetSelectMode(TopAbs_ShapeEnum nMode);
	void RemoveSelectMode(TopAbs_ShapeEnum nMode);

	void OpenLocalContext();
	void CloseLocalContext();

	void AddOrRemoveSelected(const Handle(AIS_InteractiveObject)& aniobj,const Standard_Boolean updateviewer = Standard_True);

	//
	bool GetSelected(vector<TopoDS_Shape>& shapeList);
	int SelectedCount();
	bool GetSelected(Handle(AIS_InteractiveObject)& aisObject);
private:
	Handle_AIS_InteractiveContext& m_hAisContext;
	Handle_V3d_View& m_OccView;
	int m_nLocalContexId;
};


class CAutoLocalContext
{
public:
	CAutoLocalContext(CSelectionSet& selectSet)
		:m_selectSet(selectSet)
	{
		m_selectSet.OpenLocalContext();
	}
	~CAutoLocalContext()
	{
		m_selectSet.CloseLocalContext();
	}
private:
	CSelectionSet& m_selectSet;
};

class CAutoSelectMode
{
public:
	CAutoSelectMode(CSelectionSet& selectSet, TopAbs_ShapeEnum nMode)
		:m_selectSet(selectSet),
		m_nMode(nMode)
	{
		m_selectSet.SetSelectMode(m_nMode);
	}
	~CAutoSelectMode()
	{
		m_selectSet.RemoveSelectMode(m_nMode);
	}

private:
	CSelectionSet& m_selectSet;
	TopAbs_ShapeEnum m_nMode;
};