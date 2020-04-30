#include "StdAfx.h"
#include <assert.h>
#include "OCCEntFactory.h"

#include "CylinderEntity.h"
#include "ConeEntity.h"
#include "BoxEntity.h"
#include "CtorEntity.h"
#include "DomeEntity.h"
#include "ExtruEntity.h"
#include "HexagonEntity.h"
#include "OctagonEntity.h"
#include "PyramidEntity.h"
#include "RtorEntity.h"
#include "SphereEntity.h"
#include "ComplexShapeEntity.h"
#include "RevolutionSurfEntity.h"
#include "EvolveSurfEntity.h"
#include "TemplateEntity.h"
#include "GroupEntity.h"

using namespace OCC;

COCCEntFactory::COCCEntFactory(void)
{
	m_pOCCEntPool = new map<STRING_T , list<OCC::COCCEntity*> >;
}

COCCEntFactory::~COCCEntFactory(void)
{
	try
	{
		for(map<STRING_T , list<OCC::COCCEntity*> >::iterator itr = m_pOCCEntPool->begin();itr != m_pOCCEntPool->end();++itr)
		{
			for(list<OCC::COCCEntity*>::iterator jtr = itr->second.begin();jtr != itr->second.end();++jtr)
			{
			}
		}
	}
	catch(...)
	{
	}
}

COCCEntFactory& COCCEntFactory::Instance()
{
	static COCCEntFactory __instance__;

	return __instance__;
}

int COCCEntFactory::Recycle( OCC::COCCEntity* p , Handle_AIS_InteractiveContext hAISContext)
{
	assert(p && "p is NULL");

	if(p)
	{
		p->Reset(hAISContext);
		map<STRING_T , list<COCCEntity*> >::iterator where = m_pOCCEntPool->find(p->type());
		if(where == m_pOCCEntPool->end())
		{
			(*m_pOCCEntPool)[p->type()].push_back( p );
		}
		else
		{
			where->second.push_back( p );
		}

		return ERROR_SUCCESS;
	}

	return ERROR_INVALID_PARAMETER;
}

/******************************************************************************
    @author     humkyung
    @date       2011-09-30
    @class      COCCEntFactory
    @function   GetEntity
    @return     COCCEntity*
    @param      const   CString&
    @param      type
    @brief		create a occ entity corresponding to type
******************************************************************************/
COCCEntity* COCCEntFactory::GetEntity( const STRING_T& type )
{
	map<STRING_T , list<COCCEntity*> >::iterator where = m_pOCCEntPool->find(type);
	if((where == m_pOCCEntPool->end()) || where->second.empty())
	{
		if(_T("box") == type)
			return new CBoxEntity();
		else if(_T("cone") == type)
			return new CConeEntity();
		else if(_T("ctor") == type)
			return new CCtorEntity();
		else if(_T("cylinder") == type)
			return new CCylinderEntity();
		else if(_T("dome") == type)
			return new CDomeEntity();
		else if(_T("extru") == type)
			return new CExtruEntity();
		else if(_T("hexagon") == type)
			return new CHexagonEntity();
		else if(_T("octagon") == type)
			return new COctagonEntity();
		else if(_T("pyramid") == type)
			return new CPyramidEntity();
		else if(_T("rtor") == type)
			return new CRtorEntity();
		else if(_T("sphere") == type)
			return new CSphereEntity();
		else if(_T("template") == type)
			return new CTemplateEntity();
		else if(CComplexShapeEntity::TypeString() == type)
		{
			return new CComplexShapeEntity();
		}
		else if(CRevolutionSurfEntity::TypeString() == type)
		{
			return new CRevolutionSurfEntity();
		}
		else if(CEvolveSurfEntity::TypeString() == type)
		{
			return new CEvolveSurfEntity();
		}
		else if(CGroupEntity::TypeString() == type)
		{
			return new CGroupEntity();
		}

		return NULL;
	}
	else
	{
		COCCEntity* pEnt = *(where->second.begin());
		where->second.erase( where->second.begin() );
		return pEnt;
	}

	return NULL;
}