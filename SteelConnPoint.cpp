#include "StdAfx.h"
#include <assert.h>

#include <IsLine3d.h>
#include <OCCEntFactory.h>
#include <ComplexShapeEntity.h>
#include <SphereEntity.h>

#include "SmartSteelDoc.h"
#include "OCCShapeBuilder.h"
#include "SteelConnPoint.h"

#include <algorithm>

CSteelConnPoint::CSteelConnPoint(const ConnType& iType , const CIsPoint3d& origin) : m_iType(iType) , m_origin(origin)
{
	m_dist = DBL_MAX;
}

CSteelConnPoint::~CSteelConnPoint(void)
{
	try
	{
		m_iType = NONE;
		m_oMemList.clear();
	}
	catch(...)
	{
	}
}

/**
	@brief	insert connection member at given position
	@author	humkyung
	@date	2013.07.03
*/
int CSteelConnPoint::Insert(const int& at , CSDNFLinearMember* pMember)
{
	assert(pMember && (at < int(m_oMemList.size())) && "pMember is NULL or range error");

	if(pMember && (at < int(m_oMemList.size())))
	{
		int i = 0;
		vector<CSDNFLinearMember*>::iterator itr = m_oMemList.begin();
		for(itr = m_oMemList.begin();itr != m_oMemList.end() && (i < at);++itr,++i);
		m_oMemList.insert(itr , pMember);
		
		return ERROR_SUCCESS;
	}

	return ERROR_INVALID_PARAMETER;
}

/**
	@brief	add connection member

	@author	humkyung

	@date	2013.06.19
*/
int CSteelConnPoint::Add(CSDNFLinearMember* pMember)
{
	assert(pMember && "pMember is NULL");

	if(pMember)
	{
		vector<CSDNFLinearMember*>::iterator where = find(m_oMemList.begin() , m_oMemList.end() , pMember);
		if(m_oMemList.end() == where)
		{
			m_oMemList.push_back(pMember);
		}

		return ERROR_SUCCESS;
	}

	return ERROR_INVALID_PARAMETER;
}

/**
	@brief	return connection type as reference

	@author	humkyung

	@date	2013.07.03
*/
CSteelConnPoint::ConnType& CSteelConnPoint::Type()
{
	return m_iType;
}

/**
	@brief	return connection type
	@author	humkyung
	@date	2013.06.19
*/
CSteelConnPoint::ConnType CSteelConnPoint::Type() const
{
	return m_iType;
}

/**
	@brief	return reference origin of connection
	@author	humkyung
	@date	2013.07.13
*/
CIsPoint3d& CSteelConnPoint::origin()
{
	return m_origin;
}

/**
	@brief	return origin of connection

	@author	humkyung

	@date	2013.06.19
*/
CIsPoint3d CSteelConnPoint::origin() const
{
	return m_origin;
}

/**
	@brief	return linear member at given position
	@author	humkyung
	@date	2013.06.24
*/
CSDNFLinearMember* CSteelConnPoint::GetMemberAt(const int& at)
{
	assert((at < int(m_oMemList.size())) && "index is out of range");
	CSDNFLinearMember* res = NULL;

	if(at < int(m_oMemList.size()))
	{
		res = m_oMemList[at];
	}

	return res;
}

/**
	@brief	set linear member at given position
	@author	humkyung
	@date	2014.12.05
*/
void CSteelConnPoint::SetMemberAt(const int& at , CSDNFLinearMember* pMember)
{
	assert((NULL != pMember) && (at < int(m_oMemList.size())) && "pMember is NULL or index is out of range");

	if(at < int(m_oMemList.size()))
	{
		m_oMemList[at] = pMember;
		return;
	}

	throw std::out_of_range("out of range");
}

/**
	@brief	return member size

	@author	humkyung

	@date	2013.06.24
*/
int CSteelConnPoint::GetMemberSize() const
{
	return m_oMemList.size();
}

/**
	@brief	display connection point

	@author	humkyung

	@date	2013.06.19
*/
void CSteelConnPoint::Display(Handle(AIS_InteractiveContext) hContext)
{
	return;

	OCC::COCCEntFactory& factory = OCC::COCCEntFactory::Instance();

	STRING_T color;
	switch(m_iType)
	{
		case CSteelConnPoint::COLUMN_BEAM_TO_VBRACE:
			color = _T("0,255,0");
			break;
		case CSteelConnPoint::COLUMN_TO_VBRACE:
			color = _T("0,0,255");
		break;
		case CSteelConnPoint::BEAM_TO_VBRACE:
			color = _T("255,0,0");
		break;
		case CSteelConnPoint::BEAM_TO_HBRACE:
			color = _T("255,255,0");
		break;
		case CSteelConnPoint::VBRACE_TO_VBRACE:
			color = _T("255,0,255");
		break;
		case CSteelConnPoint::HBRACE_TO_HBRACE:
			color = _T("0,255,255");
		break;
		case CSteelConnPoint::COLUMN_TO_BEAM:
			color = _T("177,177,177");
		break;
	}

	try
	{
		TopoDS_Shape aShape = BRepPrimAPI_MakeSphere(gp_Pnt(m_origin.x() , m_origin.y() , m_origin.z()) , 50);
		OCC::CComplexShapeEntity* pComplexShape = (OCC::CComplexShapeEntity*)factory.GetEntity( OCC::CComplexShapeEntity::TypeString() );
		if(pComplexShape)
		{
			pComplexShape->SetColor(color.c_str());
			pComplexShape->m_hShape = aShape;
			pComplexShape->Display(hContext , AIS_Shaded);
		}
	}
	catch(...)
	{
	}
}

/**
	@brief	return ture if this has pMem as member

	@author	humkyung

	@date	2013.06.26
*/
bool CSteelConnPoint::HasMember(CSDNFLinearMember* pMem) const
{
	assert(pMem && "pMem is NULL");
	bool res = false;

	if(pMem)
	{
		res = (m_oMemList.end() != find(m_oMemList.begin() , m_oMemList.end() , pMem));
	}

	return res;
}