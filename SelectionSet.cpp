#include "StdAfx.h"
#include "SelectionSet.h"
#include "TColStd_ListOfInteger.hxx"
#include <StdSelect_EdgeFilter.hxx> 
CSelectionSet::CSelectionSet(Handle_AIS_InteractiveContext& hAisContext, Handle_V3d_View& h3DView)
:m_hAisContext(hAisContext),
m_OccView(h3DView),
m_nLocalContexId(-1)
{

}

CSelectionSet::~CSelectionSet(void)
{
	m_hAisContext->ClearCurrents();
	m_hAisContext->ClearSelected();
}

void CSelectionSet::HilightSelected(bool bUpdate)
{
	//m_hAisContext->HilightCurrents(bUpdate);
}

void CSelectionSet::OpenLocalContext()
{
	m_nLocalContexId = m_hAisContext->OpenLocalContext();
}

void CSelectionSet::CloseLocalContext()
{
	m_hAisContext->CloseLocalContext(m_hAisContext->IndexOfCurrentLocal());
}

AIS_StatusOfPick CSelectionSet::Select(long x, long y)
{
	return m_hAisContext->Select();
}

void CSelectionSet::AddOrRemoveSelected(const Handle(AIS_InteractiveObject)& aniobj,const Standard_Boolean updateviewer)
{
	m_hAisContext->AddOrRemoveSelected(aniobj , updateviewer);
}

void CSelectionSet::DeSelectAll(bool bUpdate)
{
	m_hAisContext->UnhilightCurrents(Standard_True);
	m_hAisContext->UnhilightSelected(Standard_True);
	m_hAisContext->ClearCurrents(Standard_True);
	m_hAisContext->ClearSelected(bUpdate);
	m_hAisContext->Select(bUpdate);
}

void CSelectionSet::SetSelectMode(TopAbs_ShapeEnum nMode)
{
	m_hAisContext->ActivateStandardMode(nMode);
}

void CSelectionSet::RemoveSelectMode(TopAbs_ShapeEnum nMode)
{
	m_hAisContext->DeactivateStandardMode(nMode);
}

void CSelectionSet::DynamicSense(long x, long y)
{
	m_hAisContext->MoveTo(x, y, m_OccView);
}

bool CSelectionSet::GetSelected(vector<TopoDS_Shape>& shapeList)
{
	shapeList.clear();

	m_hAisContext->InitSelected();
	while(m_hAisContext->MoreSelected())
	{
		shapeList.push_back( m_hAisContext->SelectedShape() );
		m_hAisContext->NextSelected();
	}

	return !shapeList.empty();
}

int CSelectionSet::SelectedCount()
{
	m_hAisContext->InitSelected();
	return m_hAisContext->NbSelected();
}

 bool CSelectionSet::GetSelected(Handle(AIS_InteractiveObject)& aisObject)
{
	m_hAisContext->InitSelected();
	if(m_hAisContext->MoreSelected())
	{
		aisObject = m_hAisContext->SelectedInteractive();
	}

	return aisObject.IsNull() == Standard_False;
}