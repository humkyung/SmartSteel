#include "StdAfx.h"
#include <BRepOffsetAPI_NormalProjection.hxx>
#include <ShapeFix_ShapeTolerance.hxx>
#include <BRepOffsetAPI_MakeEvolved.hxx>
#include "EvolveSurfEntity.h"

using namespace OCC;

CEvolveSurfEntity::CEvolveSurfEntity(void)
{
	m_type = CEvolveSurfEntity::TypeString();
}

CEvolveSurfEntity::~CEvolveSurfEntity(void)
{
}

/******************************************************************************
    @author     humkyung
    @date       2011-11-20
    @class      CEvolveSurfEntity
    @function   TypeString
    @return     STRING_T
    @brief
******************************************************************************/
STRING_T CEvolveSurfEntity::TypeString()
{
	static const STRING_T __type_str__(_T("evolvesurf"));
	
	return __type_str__;
}

/******************************************************************************
    @author     humkyung
    @date       2011-11-20
    @class      CEvolveSurfEntity
    @function   CreateShape
    @return     int
    @param      BRep_Builder*       pBuilder
    @param      TopoDS_Compound*    pCompound
    @brief
******************************************************************************/
TopoDS_Shape CEvolveSurfEntity::CreateShape(BRep_Builder* pBuilder , TopoDS_Compound* pCompound)
{
	try
	{		
		if(!m_hStartWire.IsNull() && !m_hEndWire.IsNull())
		{
			if(!m_hStartWire.Closed() || !m_hEndWire.Closed())
			{
				m_bSolid = false;
			}
			else
			{
				m_bSolid = true;
			}

			///BRepOffsetAPI_ThruSections generator(((true == m_bSolid) ? Standard_True : Standard_False) , Standard_False);
			BRepOffsetAPI_MakeEvolved generator(m_hStartWire,m_hEndWire);

			double myTolerance = 0.420042001050002; // vary this
			// create a tolerance object
			ShapeFix_ShapeTolerance FTol;
			// set the tolerance for this shape.
			FTol.SetTolerance(m_hStartWire , myTolerance ,TopAbs_WIRE);
			FTol.SetTolerance(m_hEndWire , myTolerance ,TopAbs_WIRE);

			//generator.AddWire(m_hStartWire);
			//generator.AddWire(m_hEndWire);
			try
			{
				///generator.CheckCompatibility(0);
				generator.Build();
				if(generator.IsDone())
				{
					TopoDS_Shape aShape = generator.Shape();
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
							Handle_AIS_InteractiveContext hContext = m_oAISShapeList[0]->GetContext();

							m_oAISShapeList[0]->Set(aShape);
							if(!hContext.IsNull())
							{
								hContext->Deactivate(m_oAISShapeList[0]);
								hContext->Redisplay(m_oAISShapeList[0] , true , true);
								hContext->Update(m_oAISShapeList[0] , true);
							}
						}
					}*/
				}
			}
			/*catch(Standard_Failure)
			{
				Handle_Standard_Failure e = Standard_Failure::Caught();
				AfxMessageBox(e->GetMessageString());
				return ERROR_BAD_ENVIRONMENT;
			}*/
			catch(...)
			{
			}
		}
	}
	catch(Standard_Failure E)
	{
		//AfxMessageBox(E.GetMessageString());
	}

	return TopoDS_Shape();
}

/******************************************************************************
    @author     humkyung
    @date       2011-11-20
    @class      CEvolveSurfEntity
    @function   Reset
    @return     int
    @param      Handle_AIS_InteractiveContext   hAISContext
    @brief
******************************************************************************/
int CEvolveSurfEntity::Reset(Handle_AIS_InteractiveContext hAISContext)
{
	COCCEntity::Reset(hAISContext);

	return ERROR_SUCCESS;
}
