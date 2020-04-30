#include "stdafx.h"
#include <assert.h>
#include <Tokenizer.h>
#include "OCCEntity.h"
#include <util/Guid.h>

using namespace OCC;

double COCCEntity::EPSILON = 0.0000001;

COCCEntity::COCCEntity(void) : m_type(_T("")) , m_sColor(_T("green")) , m_bSolid(true)
{
	m_dTransparency = 0.0;

	Reset(NULL);
}

COCCEntity::~COCCEntity(void)
{
	try
	{
		for(vector<Handle(AIS_Shape) , allocator_Handle_AIS_Shape>::iterator itr = m_oAISShapeList.begin();itr != m_oAISShapeList.end();++itr)
		{
			(*itr)->GetContext()->Remove(*itr , false);
			(*itr).Nullify();
		}
	}
	catch(...)
	{
	}
}

/******************************************************************************
    @author     humkyung
    @date       2011-09-05
    @class      COCCEntity
    @function   guid
    @return     CString
    @brief
******************************************************************************/
CString COCCEntity::guid() const
{
	return m_guid;
}

STRING_T COCCEntity::type() const
{
	return m_type;
}

int COCCEntity::SetColor(const CString& sColor)
{
	m_sColor = sColor;
	return ERROR_SUCCESS;
}

/******************************************************************************
    @author     humkyung
    @date       2011-??-??
    @class      COCCEntity
    @function   Reset
    @return     int
    @param      Handle_AIS_InteractiveContext   hAISContext
    @brief
******************************************************************************/
int COCCEntity::Reset(Handle_AIS_InteractiveContext hAISContext)
{
	m_ptOrigin = gp_Pnt(0.0 , 0.0 , 0.0);
	{
		CGUID guid;
		m_guid = guid.m_str;
	}
	m_sColor = _T("green");	/// default value is 'green'

	///if(!hAISContext.IsNull())
	{
		for(vector<Handle(AIS_Shape) , allocator_Handle_AIS_Shape>::iterator itr = m_oAISShapeList.begin();itr != m_oAISShapeList.end();++itr)
		{
			(*itr)->Delete();
			///hAISContext->Remove*itr);
		}
	}
	m_oAISShapeList.clear();

	return ERROR_SUCCESS;
}

/******************************************************************************
    @author     humkyung
    @date       2011-09-30
    @class      COCCEntity
    @function   Clone
    @return     COCCEntity*
    @brief
******************************************************************************/
COCCEntity* COCCEntity::Clone()
{
	return NULL;
}

void COCCEntity::Translate(Handle_AIS_InteractiveContext hContext , gp_Vec& V)
{
	m_ptOrigin.Translate(V);
	
	///if(!m_oAISShapeList.empty())
	///{
	CreateShape();
	///}
	//Redisplay(hContext);
}

void COCCEntity::Rotate( const double& angle )
{
	/// do nothing
}

void COCCEntity::Rotate( const gp_Ax1& axis , const double& angle )
{
	m_ptOrigin.Rotate( axis , angle*PI180 );
}

void COCCEntity::Update(vector<CString>& oObject , Handle_AIS_InteractiveContext hContext)
{
}

TopoDS_Shape COCCEntity::CreateShape(BRep_Builder* pBuilder , TopoDS_Compound* pCompound){ return TopoDS_Shape(); }

/******************************************************************************
    @author     humkyung
    @date       ????.??.??
    @class      COCCEntity
    @function   GetColorFrom
    @return     Quantity_NameOfColor
    @param      const   CString&
    @param      sColor
    @brief
******************************************************************************/
Quantity_NameOfColor COCCEntity::GetColorFrom(const CString& sColor)
{
	if(0 == sColor.CompareNoCase(_T("green")))
	{
		return Quantity_NOC_GREEN;
	}
	else if(0 == sColor.CompareNoCase(_T("red")))
	{
		return Quantity_NOC_RED;
	}
	else if(0 == sColor.CompareNoCase(_T("blue")))
	{
		return Quantity_NOC_MATRABLUE;
	}
	else if(0 == sColor.CompareNoCase(_T("yellow")))
	{
		return Quantity_NOC_YELLOW;
	}
	else if(0 == sColor.CompareNoCase(_T("gray")))
	{
		return Quantity_NOC_GRAY;
	}

	return Quantity_NOC_GREEN;
}

Handle_AIS_InteractiveContext COCCEntity::GetContext()
{
	if(!m_oAISShapeList.empty()) return m_oAISShapeList[0]->GetContext();

	return NULL;
}

gp_Pnt COCCEntity::GetOrigin() const
{
	return m_ptOrigin;
}

int COCCEntity::SetOrigin(gp_Pnt& origin)
{
	m_ptOrigin = origin;
	
	return ERROR_SUCCESS;
}

/******************************************************************************
    @author     humkyung
    @date       2011-11-21
    @class      COCCEntity
    @function   SetTransparency
    @return     int
    @param      const           double
    @param      dTransparency
    @brief
******************************************************************************/
int COCCEntity::SetTransparency( const double dTransparency )
{
	m_dTransparency = dTransparency;
	for(vector<Handle(AIS_Shape) , allocator_Handle_AIS_Shape>::iterator itr = m_oAISShapeList.begin();itr != m_oAISShapeList.end();++itr)
	{
		(*itr)->SetTransparency( dTransparency );
	}

	return ERROR_SUCCESS;
}

COCCAttribute* COCCEntity::GetAttributeAt( const int& at )
{
	if( at < int(m_oAttributeList.size()) )
	{
		return m_oAttributeList[at];
	}

	return NULL;
}

int COCCEntity::GetAttributeCount() const
{
	return m_oAttributeList.size();
}

int COCCEntity::AddAttribute(COCCAttribute* pAttr)
{
	assert(pAttr && "pAttr is NULL");

	if(pAttr)
	{
		m_oAttributeList.push_back( pAttr );
		return ERROR_SUCCESS;
	}

	return ERROR_INVALID_PARAMETER;
}

/******************************************************************************
    @author     humkyung
    @date       2011-12-04
    @class      COCCEntity
    @function   GetBoundingBox
    @return     int
    @param      Bnd_Box&    oBndBox
    @brief
******************************************************************************/
Bnd_Box COCCEntity::BoundingBox()
{
	Bnd_Box oBndBox;
	for(vector<Handle(AIS_Shape) , allocator_Handle_AIS_Shape>::iterator itr = m_oAISShapeList.begin();itr != m_oAISShapeList.end();++itr)
	{
		Bnd_Box _oBndBox = (*itr)->BoundingBox();
		oBndBox.Add( _oBndBox );
	}

	return oBndBox;
}

/******************************************************************************
    @author     humkyung
    @date       2011-12-07
    @class      COCCEntity
    @function   CreateWire
    @return     TopoDS_Wire
    @param      const       vector<gp_Pnt>&
    @param      oPointList
    @brief
******************************************************************************/
TopoDS_Wire COCCEntity::CreateWire(const vector<gp_Pnt>& oPointList)
{
	BRepBuilderAPI_MakeWire oMakeWire;
	for(vector<gp_Pnt>::const_iterator itr = oPointList.begin();(itr + 1) != oPointList.end();++itr)
	{
		gp_Pnt start(*itr);
		gp_Pnt end(*(itr + 1));
		if(!start.IsEqual(end , COCCEntity::EPSILON))
		{
			TopoDS_Edge E = BRepBuilderAPI_MakeEdge(start , end);
			oMakeWire.Add(E);
		}
	}

	return oMakeWire.Wire();
}

/******************************************************************************
    @author     humkyung
    @date       2011-12-08
    @class      COCCEntity
    @function   SetSolid
    @return     void
    @param      const   bool&
    @param      bSolid
    @brief
******************************************************************************/
void COCCEntity::SetSolid(const bool& bSolid)
{
	m_bSolid = bSolid;
}

/******************************************************************************
    @author     humkyung
    @date       2011-11-22
    @class      COCCEntity
    @function   Show
    @return     int
    @param      const   bool&
    @param      bShow
    @brief
******************************************************************************/
int COCCEntity::Show(Handle_AIS_InteractiveContext hContext , const bool& bShow)
{
	for(vector<Handle(AIS_Shape) , allocator_Handle_AIS_Shape>::iterator itr = m_oAISShapeList.begin();itr != m_oAISShapeList.end();++itr)
	{
		if(true == bShow)
		{
			hContext->Display(*itr , false);
		}
		else
		{
			hContext->Remove(*itr , false);
		}
	}

	return ERROR_SUCCESS;
}

/******************************************************************************
    @author     humkyung
    @date       2011-09-07
    @class      COCCEntity
    @function   Display
    @return     void
    @param      Handle_AIS_InteractiveContext   hContext
    @brief
******************************************************************************/
void COCCEntity::Display(Handle_AIS_InteractiveContext hContext , const Standard_Integer& aMode)
{
	assert(hContext && "hContext is NULL");

	if(hContext)
	{
		if(m_oAISShapeList.empty())
		{
			CreateShape();
		}

		for(vector<Handle(AIS_Shape) , allocator_Handle_AIS_Shape>::iterator itr = m_oAISShapeList.begin();itr != m_oAISShapeList.end();++itr)
		{
			if(-1 != m_sColor.Find(','))
			{
				vector<STRING_T> oResult;
				CTokenizer<CIsComma>::Tokenize(oResult , m_sColor.operator LPCTSTR() , CIsComma());
				if(3 == oResult.size())
				{
					Quantity_Color aColor = Quantity_Color (ATOF_T(oResult[0].c_str()) /255. , ATOF_T(oResult[1].c_str()) /255. , ATOF_T(oResult[2].c_str()) /255. , 
						Quantity_TOC_RGB);
					(*itr)->SetColor( aColor );
				}
			}
			else
			{
				(*itr)->SetColor(GetColorFrom(m_sColor));
			}
			(*itr)->SetMaterial(Graphic3d_NOM_DEFAULT);
			(*itr)->SetTransparency(m_dTransparency);
			
			try
			{
				hContext->SetDisplayMode((*itr) , aMode , Standard_False);
				hContext->Display((*itr) , Standard_False);
				hContext->SetCurrentObject((*itr) , Standard_False);
			}
			catch(...)
			{
			}
		}
	}
}

void COCCEntity::Redisplay(Handle_AIS_InteractiveContext hContext)
{
	assert(hContext && "hContext is NULL");

	if(hContext)
	{
		for(vector<Handle(AIS_Shape) , allocator_Handle_AIS_Shape>::iterator itr = m_oAISShapeList.begin();itr != m_oAISShapeList.end();++itr)
		{
			hContext->Redisplay(*itr);
		}
	}
}

/**
	@brief	selected AIS_Shape
	@author	humkyung
	@date	2013.07.27
*/
void COCCEntity::Select(Handle(AIS_InteractiveContext) hContext)
{
	if(hContext)
	{
		for(vector<Handle(AIS_Shape) , allocator_Handle_AIS_Shape>::iterator itr = m_oAISShapeList.begin();itr != m_oAISShapeList.end();++itr)
		{
			hContext->AddOrRemoveSelected((*itr) , false);
		}
	}
}