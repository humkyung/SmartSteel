#include "StdAfx.h"
#include "ComplexShapeEntity.h"

using namespace OCC;

CComplexShapeEntity::CComplexShapeEntity(void)
{
	m_type = CComplexShapeEntity::TypeString();
}

CComplexShapeEntity::~CComplexShapeEntity(void)
{
}

/******************************************************************************
    @author     humkyung
    @date       2011-11-20
    @class      CComplexShapeEntity
    @function   TypeString
    @return     STRING_T
    @brief
******************************************************************************/
STRING_T CComplexShapeEntity::TypeString()
{
	static const STRING_T __type_str__(_T("complexshape"));
	
	return __type_str__;
}

/******************************************************************************
    @author     humkyung
    @date       2011-11-20
    @class      CComplexShapeEntity
    @function   CreateShape
    @return     int
    @param      BRep_Builder*       pBuilder
    @param      TopoDS_Compound*    pCompound
    @brief
******************************************************************************/
TopoDS_Shape CComplexShapeEntity::CreateShape(BRep_Builder* pBuilder , TopoDS_Compound* pCompound)
{
	m_bSolid = false;
	try
	{		
		if(!m_hShape.IsNull())
		{
			if((NULL != pBuilder) && (NULL != pCompound))
			{
				pBuilder->Add( *pCompound , m_hShape );
			}
			else
			{
				if(m_oAISShapeList.empty())
				{
					Handle(AIS_Shape) aAISShape = new AIS_Shape(m_hShape);
					aAISShape->SetTransparency( m_dTransparency );
					m_oAISShapeList.push_back(aAISShape);
				}
				else
				{
					Handle_AIS_InteractiveContext hContext = m_oAISShapeList[0]->GetContext();

					m_oAISShapeList[0]->Set(m_hShape);
					if(!hContext.IsNull())
					{
						hContext->Deactivate(m_oAISShapeList[0]);
						hContext->Redisplay(m_oAISShapeList[0] , true , true);
						hContext->Update(m_oAISShapeList[0] , true);
					}
				}
			}
		}
	}
	catch(...)
	{
	}

	return TopoDS_Shape();
	//return ERROR_SUCCESS;
}

/******************************************************************************
    @author     humkyung
    @date       2011-11-20
    @class      CComplexShapeEntity
    @function   Reset
    @return     int
    @param      Handle_AIS_InteractiveContext   hAISContext
    @brief
******************************************************************************/
int CComplexShapeEntity::Reset(Handle_AIS_InteractiveContext hAISContext)
{
	COCCEntity::Reset(hAISContext);

	return ERROR_SUCCESS;
}
