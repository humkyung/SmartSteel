#include "StdAfx.h"
#include "DeleteCommand.h"
#include "../MainFrm.h"

using namespace Command;

CDeleteCommand::CDeleteCommand(const PlateSet& plateSet) : CAbstractCommand(plateSet)
{
}

CDeleteCommand::~CDeleteCommand(void)
{
}

/**
	@brief	delete given plate. it's not actually delete, just hide
	@author	humkyung
	@date	2013.07.27
*/
int CDeleteCommand::Do()
{
	CMainFrame* pFrame = CMainFrame::GetInstance();
	CFileView& oFileView = pFrame->GetFileView();
	CViewTree& oViewTree = oFileView.GetViewTree();

	CSmartSteelDoc* pDoc = GetSDIActiveDocument();
	for(vector<CSteelPlate*>::const_iterator itr = m_oPlateSet.begin();itr != m_oPlateSet.end();++itr)
	{
		HTREEITEM hItem = oFileView.FindItemWith(*itr);
		if(NULL == hItem) continue;

		int nImage = -1, nSelectedImage=-1;
		oViewTree.GetItemImage(hItem , nImage , nSelectedImage);
		if(GUSSET_PLATE_ICON == nImage)
		{
			oViewTree.SetItemImage(hItem , GUSSET_PLATE_DELETED_ICON , GUSSET_PLATE_DELETED_ICON);
		}
		else if(END_PLATE_ICON == nImage)
		{
			oViewTree.SetItemImage(hItem , END_PLATE_DELETED_ICON , END_PLATE_DELETED_ICON);
		}

		/*(*itr)->Show(pDoc->GetAISContext() , false);*/
		(*itr)->status() = CSteelPlate::DELETED;
	}
	pDoc->UpdateAllViews(false);

	return ERROR_SUCCESS;
}

/**
	@brief	undelete given plate
	@author	humkyung
	@date	2013.07.27
*/
int CDeleteCommand::Undo()
{
	CMainFrame* pFrame = CMainFrame::GetInstance();
	CFileView& oFileView = pFrame->GetFileView();
	CViewTree& oViewTree = oFileView.GetViewTree();

	CSmartSteelDoc* pDoc = GetSDIActiveDocument();
	for(vector<CSteelPlate*>::const_iterator itr = m_oPlateSet.begin();itr != m_oPlateSet.end();++itr)
	{
		HTREEITEM hItem = oFileView.FindItemWith(*itr);
		if(NULL == hItem) continue;

		int nImage = -1, nSelectedImage=-1;
		oViewTree.GetItemImage(hItem , nImage , nSelectedImage);
		if(GUSSET_PLATE_DELETED_ICON == nImage)
		{
			oViewTree.SetItemImage(hItem , GUSSET_PLATE_ICON , GUSSET_PLATE_ICON);
		}
		else if(END_PLATE_DELETED_ICON == nImage)
		{
			oViewTree.SetItemImage(hItem , END_PLATE_ICON , END_PLATE_ICON);
		}

		//(*itr)->Show(pDoc->GetAISContext() , true);
		(*itr)->status() = CSteelPlate::ALIVE;
	}
	pDoc->UpdateAllViews(false);

	return ERROR_SUCCESS;
}