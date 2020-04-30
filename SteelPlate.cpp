#include "StdAfx.h"
#include <LinearDim3d.h>
#include "SteelPlate.h"

#include <algorithm>

CSteelPlate::CSteelPlate(CSteelConnPoint* pConnPnt) : m_pConnPnt(pConnPnt) , m_status(CSteelPlate::ALIVE)
{
	m_id = -1;
}

CSteelPlate::~CSteelPlate(void)
{
}

/**
	@brief	return the max edge length
	@author	humkyung
	@date	2013.08.01
*/
double CSteelPlate::GetMaximumEdgeLength() const
{
	double res = 0.0;
	for(int i = 0;i < int(m_oSectionShapePntList.size());++i)
	{
		const int j = (i+1)%m_oSectionShapePntList.size();
		const double dist = m_oSectionShapePntList[i].DistanceTo(m_oSectionShapePntList[j]);
		res = max(res , dist);
	}

	return res;
}

/**
	@brief	return type string of plate
	@author	humkyung
	@date	2013.08.04
*/
STRING_T CSteelPlate::GetTypeString() const
{
	return m_sType;
}

/**
	@brief	return id of plate
	@author	humkyung
	@date	2013.08.04
*/
long CSteelPlate::id() const
{
	return m_id;
}

/**
	@brief	return reference id of plate
	@author	humkyung
	@date	2013.08.04
*/
long& CSteelPlate::id()
{
	return m_id;
}

/**
	@brief	return the connection point
	@author	humkyung
	@date	2013.08.01
*/
CSteelConnPoint* CSteelPlate::GetConnPnt()
{
	return m_pConnPnt;
}

/**
	@brief	get bounding box of plate
	@author	humkyung
	@date	2013.07.29
*/
Bnd_Box CSteelPlate::BoundBox() const
{
	Bnd_Box oBndBox;
	for(vector<CIsPoint3d>::const_iterator itr = m_oSectionShapePntList.begin();itr != m_oSectionShapePntList.end();++itr)
	{
		oBndBox.Add(gp_Pnt(itr->x() , itr->y() , itr->z()));
		oBndBox.Add(gp_Pnt(itr->x() + m_norm.dx()*m_dThickness , itr->y()  + m_norm.dy()*m_dThickness, itr->z()  + m_norm.dz()*m_dThickness));
	}
	
	return oBndBox;
}

/**
	@brief	display dimension of plate
	@author	humkyung
	@date	2013.07.29
*/
int CSteelPlate::DrawDimension(Handle(AIS_InteractiveContext) hContext)
{
#ifdef	SMART_STEEL
	//for(int i = 0;i < m_oSectionShapePntList.size();++i)
	//{
	//	const int j = (i+1)%m_oSectionShapePntList.size();

	//	gp_XYZ xyz(gp_XYZ(m_oSectionShapePntList[j].x() , m_oSectionShapePntList[j].y() , m_oSectionShapePntList[j].z()) - gp_XYZ(m_oSectionShapePntList[i].x() , m_oSectionShapePntList[i].y() , m_oSectionShapePntList[i].z()));
	//	xyz.Normalize();
	//	
	//	gp_Vec norm(m_norm.dx() , m_norm.dy() , m_norm.dz());
	//	gp_Vec crossed = norm.Crossed(xyz);
	//	crossed.Normalize();

	//	gp_Pnt start = gp_Pnt(m_oSectionShapePntList[i].x() , m_oSectionShapePntList[i].y() , m_oSectionShapePntList[i].z());
	//	gp_Pnt end = gp_Pnt(m_oSectionShapePntList[j].x() , m_oSectionShapePntList[j].y() , m_oSectionShapePntList[j].z());
	//	gp_Pnt mid((start.X() + end.X())*0.5 , (start.Y() + end.Y())*0.5 , (start.Z() + end.Z())*0.5);
	//	mid.Translate(crossed*5);

	//	Handle(CLinearDim3d) hDim = new CLinearDim3d(mid , gp_Dir(xyz) , start , end);
	//	//m_oDimList.push_back(hDim);
	//	hContext->Display(hDim , false);
	//}
#endif
	return ERROR_SUCCESS;
}

int CSteelPlate::Select(Handle(AIS_InteractiveContext) hContext)
{
	for(vector<OCC::CComplexShapeEntity*>::iterator itr = m_oShapeEntList.begin();itr != m_oShapeEntList.end();++itr)
	{
		(*itr)->Select(hContext);
	}

	return ERROR_SUCCESS;
}

/**
	@brief	return shape entity list
	@author	humkyung
	@date	2013.07.25
*/
vector<OCC::CComplexShapeEntity*>* CSteelPlate::GetShapeEntList()
{
	return &m_oShapeEntList;
}

/**
	@brief	return the status of plate
	@author	humkyung
	@date	2013.07.27
*/
const CSteelPlate::Status CSteelPlate::status() const
{
	return m_status;
}

/**
	@brief	return the status reference of plate
	@author	humkyung
	@date	2013.07.27
*/
CSteelPlate::Status& CSteelPlate::status()
{
	return m_status;
}