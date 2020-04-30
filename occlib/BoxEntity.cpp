#include "stdafx.h"
#include <assert.h>
#include "BoxEntity.h"

using namespace OCC;

CBoxEntity::CBoxEntity(void)
{
	m_type = _T("box");

	Reset(NULL);
}

CBoxEntity::~CBoxEntity(void)
{
}

int CBoxEntity::Reset(Handle_AIS_InteractiveContext hAISContext)
{
	COCCEntity::Reset(hAISContext);

	m_ptOrigin = gp_Pnt(0.0 , 0.0 , 0.0);
	m_xAxis = gp_Dir(1.0 , 0.0 , 0.0);
	m_yAxis = gp_Dir(0.0 , 1.0 , 0.0);
	m_dWidth = 10.0;
	m_dHeight = 10.0;
	m_dDepth = 10.0;

	return ERROR_SUCCESS;
}

void CBoxEntity::Rotate( const double& angle )
{
	gp_Dir az;
	az = m_xAxis.Crossed(m_yAxis);

	m_xAxis.Rotate(gp_Ax1(m_ptOrigin , az) , angle*PI180);
	m_yAxis.Rotate(gp_Ax1(m_ptOrigin , az) , angle*PI180);
	/*if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}*/
}

void CBoxEntity::Rotate( const gp_Ax1& axis , const double& angle )
{
	m_ptOrigin.Rotate(axis , angle*PI180);
	/*if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}*/
}

/******************************************************************************
    @author     humkyung
    @date       2011-09-30
    @class      CBoxEntity
    @function   Update
    @return     void
    @param      vector<CString>&                oObject
    @param      Handle_AIS_InteractiveContext   hContext
    @brief
******************************************************************************/
void CBoxEntity::Update(vector<CString>& oObject , Handle_AIS_InteractiveContext hContext)
{
	for(vector<CString>::iterator itr = oObject.begin();itr != oObject.end();++itr)
	{
		if(_T("origin") == (*itr))
		{
			m_ptOrigin.SetCoord( atof(*(itr + 1)) , atof(*(itr + 2)) , atof(*(itr + 3)) );
			itr+=3;
		}
		else if(_T("ax") == (*itr))
		{
			m_xAxis.SetCoord( atof(*(itr + 1)) , atof(*(itr + 2)) , atof(*(itr + 3)) );
			itr+=3;
		}
		else if(_T("ay") == (*itr))
		{
			m_yAxis.SetCoord( atof(*(itr + 1)) , atof(*(itr + 2)) , atof(*(itr + 3)) );
			itr+=3;
		}
		else if(_T("width") == (*itr))
		{
			m_dWidth = atof(*(itr + 1));
			itr++;
		}
		else if(_T("height") == (*itr))
		{
			m_dHeight = atof(*(itr + 1));
			itr++;
		}
		else if(_T("depth") == (*itr))
		{
			m_dDepth = atof(*(itr + 1));
			itr++;
		}
		else if(_T("color") == (*itr))
		{
			m_sColor = *(itr + 1);
			itr++;
		}
	}

	///if(ERROR_SUCCESS == CreateShape()) Redisplay(hContext);
}

/******************************************************************************
    @author     humkyung
    @date       2011-09-30
    @class      CBoxEntity
    @function   CreateShape
    @return     int
    @param      BRep_Builder*       pBuilder
    @param      TopoDS_Compound*    pCompound
    @brief
******************************************************************************/
TopoDS_Shape CBoxEntity::CreateShape(BRep_Builder* pBuilder , TopoDS_Compound* pCompound)
{
	gp_Dir zAxis;
	zAxis = m_xAxis.Crossed(m_yAxis);

	gp_Pnt oLeftCorner = m_ptOrigin;
	oLeftCorner.Translate(-gp_Vec(m_xAxis)*m_dWidth*0.5);
	oLeftCorner.Translate(-gp_Vec(m_yAxis)*m_dHeight*0.5);
	oLeftCorner.Translate(-gp_Vec(zAxis)*m_dDepth*0.5);
	TopoDS_Shape aShape = BRepPrimAPI_MakeBox (gp_Ax2(oLeftCorner , zAxis , m_xAxis), m_dWidth , m_dHeight , m_dDepth);
	return aShape;
	/*aShape->
	hContext->SetColor((*itr) , GetColorFrom(m_sColor) , Standard_False); 
	hContext->SetMaterial((*itr) , Graphic3d_NOM_PLASTIC , Standard_False);   */
	
	/*if((NULL != pBuilder) && (NULL != pCompound))
	{
		pBuilder->Add( *pCompound , aShape );
	}
	else
	{
		if(m_oAISShapeList.empty())
		{
			Handle(AIS_Shape) aAISShape = new AIS_Shape(aShape);
			///aAISShape->SetColor();
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
				///hContext->CurrentViewer()->Redraw();
			}
		}
	}

	return ERROR_SUCCESS;*/
}

//void HSF::updateAIS(TopoDS_Shape aShape, Handle_AIS_Shape &anAIS, Handle_AIS_InteractiveContext ic)
//{
//	if(aShape.IsNull()) return;
//	if (anAIS.IsNull())
//	{
//		anAIS=new AIS_Shape(aShape);
//		ic->SetMaterial(anAIS,Graphic3d_NOM_NEON_GNC);
//		ic->SetColor(anAIS, Quantity_NOC_BLACK);
//		ic->SetDisplayMode(anAIS,1,Standard_False);
//	}
//	
//	if (!anAIS->HasPresentation())
//	{
//		ic->Display(anAIS, 1,0,false,false);
//	}
//	else
//	{
//		anAIS->Set(aShape);
//		ic->Deactivate(anAIS);
//		ic->Redisplay(anAIS,true,true);
//	}
//	ic->Update(anAIS,true);
//	ic->CurrentViewer()->Redraw();
//}