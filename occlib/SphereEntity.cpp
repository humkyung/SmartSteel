#include "stdafx.h"
#include "SphereEntity.h"

using namespace OCC;
CSphereEntity::CSphereEntity(void)
{
	m_type = CSphereEntity::TypeString();

	Reset(NULL);
}

CSphereEntity::~CSphereEntity(void)
{
}

/******************************************************************************
    @author     humkyung
    @date       2011-09-19
    @class      CSphereEntity
    @function   Reset
    @return     int
    @param      Handle_AIS_InteractiveContext   hAISContext
    @brief
******************************************************************************/
int CSphereEntity::Reset(Handle_AIS_InteractiveContext hAISContext)
{
	COCCEntity::Reset(hAISContext);

	m_ptOrigin = gp_Pnt(0.0 , 0.0 , 0.0);
	m_dRadius = 10.0;

	return ERROR_SUCCESS;
}

/******************************************************************************
    @author     humkyung
    @date       2011-11-20
    @class      CSphereEntity
    @function   TypeString
    @return     STRING_T
    @brief
******************************************************************************/
STRING_T CSphereEntity::TypeString()
{
	static const STRING_T __type_str__(_T("sphere"));
	
	return __type_str__;
}

void CSphereEntity::Rotate( const gp_Ax1& axis , const double& angle )
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
    @class      CSphereEntity
    @function   CreateShape
    @return     int
    @param      BRep_Builder*       pBuilder
    @param      TopoDS_Compound*    pCompound
    @brief
******************************************************************************/
TopoDS_Shape CSphereEntity::CreateShape(BRep_Builder* pBuilder , TopoDS_Compound* pCompound)
{
	if(m_dRadius > 0.0)
	{
		TopoDS_Shape aShape = BRepPrimAPI_MakeSphere(m_ptOrigin , m_dRadius);
		return aShape;

		/*if((NULL != pBuilder) && (NULL != pCompound))
		{
			pBuilder->Add( *pCompound , aShape );
		}
		else
		{
			if(m_oAISShapeList.empty())
			{
				Handle(AIS_Shape) aAISShape = new AIS_Shape(aShape);
				m_oAISShapeList.push_back(aAISShape);
			}
			else
			{
				m_oAISShapeList[0]->Set(aShape);
			}
		}

		return ERROR_SUCCESS;*/
	}

	return TopoDS_Shape();
}

void CSphereEntity::Translate(gp_Vec& V)
{
	m_ptOrigin.Translate(V);
	
	if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}
}

void CSphereEntity::Update(vector<CString>& oObject , Handle_AIS_InteractiveContext hContext)
{
	for(vector<CString>::iterator itr = oObject.begin();itr != oObject.end();++itr)
	{
		if(_T("origin") == (*itr))
		{
			m_ptOrigin.SetCoord( atof(*(itr + 1)) , atof(*(itr + 2)) , atof(*(itr + 3)) );
			itr+=3;
		}
		else if(_T("radius") == (*itr))
		{
			m_dRadius = atof(*(itr + 1));
			itr++;
		}
		else if(_T("color") == (*itr))
		{
			m_sColor = *(itr + 1);
			itr++;
		}
	}

	if(!CreateShape().IsNull()) Redisplay(hContext);
}