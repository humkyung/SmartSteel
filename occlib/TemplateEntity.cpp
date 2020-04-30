#include "stdafx.h"
#include <assert.h>
#include "OCCEntFactory.h"
#include "TemplateEntity.h"

using namespace OCC;
CTemplateEntity::CTemplateEntity(void)
{
	m_type = CTemplateEntity::TypeString();

	Reset(NULL);
}

CTemplateEntity::~CTemplateEntity(void)
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
    @class      CTemplateEntity
    @function   TypeString
    @return     STRING_T
    @brief
******************************************************************************/
STRING_T CTemplateEntity::TypeString()
{
	static const STRING_T __type_str__(_T("template"));
	
	return __type_str__;
}

/******************************************************************************
    @author     humkyung
    @date       2011-09-19
    @class      CTemplateEntity
    @function   Reset
    @return     int
    @param      Handle_AIS_InteractiveContext   hAISContext
    @brief
******************************************************************************/
int CTemplateEntity::Reset(Handle_AIS_InteractiveContext hAISContext)
{
	COCCEntity::Reset(hAISContext);
	for(map<STRING_T , OCC::COCCEntity*>::iterator itr = m_oOCCEntityMap.begin();itr != m_oOCCEntityMap.end();++itr)
	{
		itr->second->Reset( hAISContext );
	}

	m_oOCCEntityMap.clear();
	m_ptOrigin = gp_Pnt(0.0 , 0.0 , 0.0);

	return ERROR_SUCCESS;
}

/******************************************************************************
    @author     humkyung
    @date       2011-09-30
    @class      CTemplateEntity
    @function   Clone
    @return     COCCEntity*
    @brief
******************************************************************************/
COCCEntity* CTemplateEntity::Clone()
{
	COCCEntFactory factory = COCCEntFactory::Instance();
	CTemplateEntity* pTemplate = (CTemplateEntity*)(factory.GetEntity( m_type ));
	if(pTemplate)
	{
		pTemplate->m_oParamList.insert(pTemplate->m_oParamList.begin() , m_oParamList.begin() , m_oParamList.end());
		pTemplate->m_sScript = m_sScript;

		for(map<STRING_T , OCC::COCCEntity*>::iterator itr = m_oOCCEntityMap.begin();itr != m_oOCCEntityMap.end();++itr)
		{
			OCC::COCCEntity* pClone = itr->second->Clone();
			pTemplate->m_oOCCEntityMap.insert( make_pair(pClone->guid().operator LPCTSTR() , pClone) );
		}
	}

	return pTemplate;
}

int CTemplateEntity::SetParamList(vector<CString> oParamList)
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
    @class      CTemplateEntity
    @function   CreateShape
    @return     int
    @brief
******************************************************************************/
TopoDS_Shape CTemplateEntity::CreateShape(BRep_Builder* pBuilder , TopoDS_Compound* pCompound)
{
	/// clear
	if(!m_oOCCEntityMap.empty())
	{
		OCC::COCCEntFactory& factory = OCC::COCCEntFactory::Instance();
		for(map<STRING_T , OCC::COCCEntity*>::iterator itr = m_oOCCEntityMap.begin();itr != m_oOCCEntityMap.end();++itr)
		{
			factory.Recycle(itr->second , itr->second->GetContext());
		}
		m_oOCCEntityMap.clear();
	}
	/// up to here

	return TopoDS_Shape();
}

/******************************************************************************
    @author     humkyung
    @date       2011-09-30
    @class      CTemplateEntity
    @function   Rotate
    @return     void
    @param      const   gp_Ax1&
    @param      axis    const
    @param      double& angle
    @brief
******************************************************************************/
void CTemplateEntity::Rotate( const gp_Ax1& axis , const double& angle )
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
    @class      CTemplateEntity
    @function   Translate
    @return     void
    @param      gp_Vec& V
    @brief
******************************************************************************/
void CTemplateEntity::Translate(Handle_AIS_InteractiveContext hContext , gp_Vec& V)
{
	m_ptOrigin.Translate(V);
	for(map<STRING_T , OCC::COCCEntity*>::iterator itr = m_oOCCEntityMap.begin();itr != m_oOCCEntityMap.end();++itr)
	{
		itr->second->Translate( hContext , V );
	}
}

/******************************************************************************
    @author     humkyung
    @date       2011-09-30
    @class      CTemplateEntity
    @function   Update
    @return     void
    @param      vector<CString>&                oObject
    @param      Handle_AIS_InteractiveContext   hContext
    @brief
******************************************************************************/
void CTemplateEntity::Update(vector<CString>& oObject , Handle_AIS_InteractiveContext hContext)
{
}

void CTemplateEntity::Display(Handle_AIS_InteractiveContext hContext , const Standard_Integer& aMode)
{
	assert(hContext && "hContext is NULL");

	if(hContext)
	{
		for(map<STRING_T , OCC::COCCEntity*>::iterator itr = m_oOCCEntityMap.begin();itr != m_oOCCEntityMap.end();++itr)
		{
			itr->second->Display( hContext , aMode );
		}
	}
}