#include "StdAfx.h"
#include "RevolutionSurfEntity.h"
#include "SmartMakeFace.h"

using namespace OCC;

CRevolutionSurfEntity::CRevolutionSurfEntity(void) : m_dAngle(0.0)
{
	m_type = CRevolutionSurfEntity::TypeString();
}

CRevolutionSurfEntity::~CRevolutionSurfEntity(void)
{
}

/******************************************************************************
    @author     humkyung
    @date       2011-11-20
    @class      CRevolutionSurfEntity
    @function   TypeString
    @return     STRING_T
    @brief
******************************************************************************/
STRING_T CRevolutionSurfEntity::TypeString()
{
	static const STRING_T __type_str__(_T("revolutionsurf"));
	
	return __type_str__;
}

/******************************************************************************
    @author     humkyung
    @date       2011-11-20
    @class      CRevolutionSurfEntity
    @function   CreateShape
    @return     int
    @param      BRep_Builder*       pBuilder
    @param      TopoDS_Compound*    pCompound
    @brief
******************************************************************************/
TopoDS_Shape CRevolutionSurfEntity::CreateShape(BRep_Builder* pBuilder , TopoDS_Compound* pCompound)
{
	TopoDS_Shape aShape;
	try
	{
		if(true == m_bSolid)
		{
			CSmartMakeFace aMakeFace;
			TopoDS_Face aFace = aMakeFace.Make(m_hWire);
			if(!aFace.IsNull())
			{
				aShape = BRepPrimAPI_MakeRevol(aFace , m_hAxis , m_dAngle);
			}
		}
		else
		{
			aShape = BRepPrimAPI_MakeRevol(m_hWire , m_hAxis , m_dAngle);
		}
		return aShape;
	}
	catch(...)
	{
	}

	return aShape;
}

/******************************************************************************
    @author     humkyung
    @date       2011-11-20
    @class      CRevolutionSurfEntity
    @function   Reset
    @return     int
    @param      Handle_AIS_InteractiveContext   hAISContext
    @brief
******************************************************************************/
int CRevolutionSurfEntity::Reset(Handle_AIS_InteractiveContext hAISContext)
{
	COCCEntity::Reset(hAISContext);

	return ERROR_SUCCESS;
}
