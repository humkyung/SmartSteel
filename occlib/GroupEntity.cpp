#include "stdafx.h"
#include <assert.h>
#include "OCCEntFactory.h"
#include "GroupEntity.h"
#include "RevolutionSurfEntity.h"

using namespace OCC;
CGroupEntity::CGroupEntity(void)
{
	m_type = CGroupEntity::TypeString();

	Reset(NULL);
}

CGroupEntity::~CGroupEntity(void)
{
	try
	{
		int d = 1;
	}
	catch(...)
	{
	}
}

/******************************************************************************
    @author     humkyung
    @date       2011-11-21
    @class      CGroupEntity
    @function   TypeString
    @return     STRING_T
    @brief
******************************************************************************/
STRING_T CGroupEntity::TypeString()
{
	static const STRING_T __type_str__(_T("group"));
	
	return __type_str__;
}

/******************************************************************************
    @author     humkyung
    @date       2011-11-21
    @class      CGroupEntity
    @function   Add
    @return     int
    @param      COCCEntity* pEnt
    @brief
******************************************************************************/
int CGroupEntity::Add(COCCEntity* pEnt)
{
	assert(pEnt && "pEnt is NULL");

	if(pEnt)
	{
		m_oOCCEntityList.push_back(pEnt);
		return ERROR_SUCCESS;
	}

	return ERROR_INVALID_PARAMETER;
}

/******************************************************************************
    @author     humkyung
    @date       2011-09-19
    @class      CGroupEntity
    @function   Reset
    @return     int
    @param      Handle_AIS_InteractiveContext   hAISContext
    @brief
******************************************************************************/
int CGroupEntity::Reset(Handle_AIS_InteractiveContext hAISContext)
{
	COCCEntity::Reset(hAISContext);
	for(list<OCC::COCCEntity*>::iterator itr = m_oOCCEntityList.begin();itr != m_oOCCEntityList.end();++itr)
	{
		(*itr)->Reset( hAISContext );
	}

	m_oOCCEntityList.clear();
	m_ptOrigin = gp_Pnt(0.0 , 0.0 , 0.0);

	return ERROR_SUCCESS;
}

/******************************************************************************
    @author     humkyung
    @date       2011-09-30
    @class      CGroupEntity
    @function   Clone
    @return     COCCEntity*
    @brief
******************************************************************************/
COCCEntity* CGroupEntity::Clone()
{
	COCCEntFactory factory = COCCEntFactory::Instance();
	CGroupEntity* pTemplate = (CGroupEntity*)(factory.GetEntity( m_type ));
	if(pTemplate)
	{
		pTemplate->m_oParamList.insert(pTemplate->m_oParamList.begin() , m_oParamList.begin() , m_oParamList.end());
		pTemplate->m_sScript = m_sScript;

		for(list<OCC::COCCEntity*>::iterator itr = m_oOCCEntityList.begin();itr != m_oOCCEntityList.end();++itr)
		{
			OCC::COCCEntity* pClone = (*itr)->Clone();
			pTemplate->m_oOCCEntityList.push_back( pClone );
		}
	}

	return pTemplate;
}

int CGroupEntity::SetParamList(vector<CString> oParamList)
{
	for(vector<CString>::iterator itr = oParamList.begin();itr != oParamList.end();++itr)
	{
		for(list<pair<STRING_T,STRING_T> >::iterator jtr = m_oParamList.begin();jtr != m_oParamList.end();++jtr)
		{
			if(jtr->first == STRING_T((*itr).operator LPCTSTR()))
			{
				jtr->second = (*(itr + 1)).operator LPCTSTR();
				++itr;
				break;
			}
		}
	}

	return ERROR_SUCCESS;
}

/******************************************************************************
    @author     humkyung
    @date       2011-09-30
    @class      CGroupEntity
    @function   CreateShape
    @return     int
    @brief
******************************************************************************/
TopoDS_Shape CGroupEntity::CreateShape(BRep_Builder* pBuilder , TopoDS_Compound* pCompound)
{
	/// clear
	if(!m_oOCCEntityList.empty())
	{
		OCC::COCCEntFactory& factory = OCC::COCCEntFactory::Instance();
		for(list<OCC::COCCEntity*>::iterator itr = m_oOCCEntityList.begin();itr != m_oOCCEntityList.end();++itr)
		{
			(*itr)->CreateShape(pBuilder , pCompound);
		}
	}
	/// up to here

	return (*pCompound);
}

/******************************************************************************
    @author     humkyung
    @date       2011-09-30
    @class      CGroupEntity
    @function   Rotate
    @return     void
    @param      const   gp_Ax1&
    @param      axis    const
    @param      double& angle
    @brief
******************************************************************************/
void CGroupEntity::Rotate( const gp_Ax1& axis , const double& angle )
{
	COCCEntity::Rotate( axis , angle );
	if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}
}

/******************************************************************************
    @author     humkyung
    @date       2011-09-30
    @class      CGroupEntity
    @function   Translate
    @return     void
    @param      gp_Vec& V
    @brief
******************************************************************************/
void CGroupEntity::Translate(Handle_AIS_InteractiveContext hContext , gp_Vec& V)
{
	m_ptOrigin.Translate(V);
	for(list<OCC::COCCEntity*>::iterator itr = m_oOCCEntityList.begin();itr != m_oOCCEntityList.end();++itr)
	{
		(*itr)->Translate( hContext , V );
	}
}

/******************************************************************************
    @author     humkyung
    @date       2011-12-04
    @class      CGroupEntity
    @function   BoundingBox
    @return     Bnd_Box
    @brief
******************************************************************************/
Bnd_Box CGroupEntity::BoundingBox()
{
	Bnd_Box oBndBox;
	for(list<OCC::COCCEntity*>::iterator itr = m_oOCCEntityList.begin();itr != m_oOCCEntityList.end();++itr)
	{
		Bnd_Box _oBndBox = (*itr)->BoundingBox( );
		oBndBox.Add(_oBndBox);
	}

	return oBndBox;
}

/******************************************************************************
    @author     humkyung
    @date       2011-09-30
    @class      CGroupEntity
    @function   Update
    @return     void
    @param      vector<CString>&                oObject
    @param      Handle_AIS_InteractiveContext   hContext
    @brief
******************************************************************************/
void CGroupEntity::Update(vector<CString>& oObject , Handle_AIS_InteractiveContext hContext)
{
	for(vector<CString>::iterator itr = oObject.begin();itr != oObject.end();++itr)
	{
		if(_T("origin") == (*itr))
		{
			m_ptOrigin.SetCoord( atof(*(itr + 1)) , atof(*(itr + 2)) , atof(*(itr + 3)) );
			itr+=3;
		}
		else
		{
			for(list<pair<STRING_T,STRING_T> >::iterator jtr = m_oParamList.begin();jtr != m_oParamList.end();++jtr)
			{
				if(jtr->first == STRING_T((*itr).operator LPCTSTR()))
				{
					jtr->second = (*(itr + 1)).operator LPCTSTR();
					++itr;
					break;
				}
			}
		}
		if(itr == oObject.end()) break;
	}

	gp_Vec V(m_ptOrigin.Coord());
	if(!CreateShape().IsNull())
	{
		Translate(hContext , V);
		Redisplay(hContext);
	}
}

/******************************************************************************
    @author     humkyung
    @date       2011-11-22
    @class      CGroupEntity
    @function   Show
    @return     int
    @param      Handle_AIS_InteractiveContext   hContext
    @param      const                           bool&
    @param      bShow
    @brief
******************************************************************************/
int CGroupEntity::Show(Handle_AIS_InteractiveContext hContext , const bool& bShow)
{
	if(hContext)
	{
		for(list<OCC::COCCEntity*>::iterator itr = m_oOCCEntityList.begin();itr != m_oOCCEntityList.end();++itr)
		{
			(*itr)->Show(hContext , bShow);
		}
	}
	
	return ERROR_SUCCESS;
}

/******************************************************************************
    @author     humkyung
    @date       2011-11-28
    @class      CGroupEntity
    @function   SetTransparency
    @return     int
    @param      const double
    @param      dTransparency
    @brief
******************************************************************************/
int CGroupEntity::SetTransparency( const double dTransparency )
{
	m_dTransparency = dTransparency;
	for(list<OCC::COCCEntity*>::iterator itr = m_oOCCEntityList.begin();itr != m_oOCCEntityList.end();++itr)
	{
		(*itr)->SetTransparency( m_dTransparency );
	}

	return ERROR_SUCCESS;
}

void CGroupEntity::Display(Handle_AIS_InteractiveContext hContext , const Standard_Integer& aMode)
{
	assert(hContext && "hContext is NULL");

	if(hContext)
	{
		for(list<OCC::COCCEntity*>::iterator itr = m_oOCCEntityList.begin();itr != m_oOCCEntityList.end();++itr)
		{
			(*itr)->SetTransparency( m_dTransparency );
			(*itr)->Display( hContext , aMode );
		}
	}
}