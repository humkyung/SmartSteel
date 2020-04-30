#include "StdAfx.h"
#include <BRepPrimAPI_MakePrism.hxx>
#include "OCCShapeBuilder.h"

#include <TopoDS_Edge.hxx>
#include <gp_Pnt.hxx>
#include <OCCEntity.h>

#include <list>
using namespace std;

COCCShapeBuilder::COCCShapeBuilder() : TOLERANCE(0.420042001050002)
{
}

COCCShapeBuilder::~COCCShapeBuilder(void)
{
}

/**
	@brief	return the volume of given shape
	@author	humkyung
	@date	2013.07.03
*/
CIsVolume COCCShapeBuilder::Volume(const vector<CIsPoint3d>& oPntList , const CIsVect3d& dir , const double& thickness) const
{
	CIsVolume res;
	for(vector<CIsPoint3d>::const_iterator itr = oPntList.begin();(itr) != oPntList.end();++itr)
	{
		gp_Pnt start(itr->x() , itr->y() , itr->z());
		gp_Pnt end(start);
		end.Translate(gp_Vec(dir.dx() , dir.dy() , dir.dz())*thickness);
		res.Add(CIsPoint3d(start.X() , start.Y() , start.Z()));
		res.Add(CIsPoint3d(end.X() , end.Y() , end.Z()));
	}

	return res;
}

/******************************************************************************
    @brief		create a shape from CIsEntity
	@author     humkyung
    @date       2013-05-29
    @class      COCCShapeBuilder
    @function   Shape
    @return     TopoDS_Shape
******************************************************************************/
TopoDS_Shape COCCShapeBuilder::Shape(const CIsLine3d& line3d)
{
	TopoDS_Shape oShape;

	list<TopoDS_Edge> oEdgeList;

	const CIsPoint3d ptStart = line3d.start();
	const CIsPoint3d ptEnd = line3d.end();

	gp_Pnt start(ptStart.x() , ptStart.y() , ptStart.z());
	gp_Pnt end(ptEnd.x() , ptEnd.y() , ptEnd.z());
	if(!start.IsEqual(end , OCC::COCCEntity::EPSILON))
	{
		BRepBuilderAPI_MakeEdge aBuilder( start , end );
		oShape = aBuilder.Shape();
	}

	return oShape;
}

/**
	@brief	create a shape with given parameter
	@author	humkyung
*/
TopoDS_Shape COCCShapeBuilder::Shape(const vector<CIsPoint3d>& oPntList , const CIsVect3d& dir , const double& thickness)
{
	TopoDS_Shape oShape;

	if(0.0 == thickness) return oShape;	/// can't create a shape if thickness is zero!!!

	int iEdgeCount = 0;
	BRepBuilderAPI_MakeWire makeStartWire , makeEndWire;
	for(vector<CIsPoint3d>::const_iterator itr = oPntList.begin();(itr + 1) != oPntList.end();++itr)
	{
		gp_Pnt start(itr->x() , itr->y() , itr->z());
		gp_Pnt end((itr+1)->x() , (itr+1)->y() , (itr+1)->z());
		if(!start.IsEqual(end , OCC::COCCEntity::EPSILON))
		{
			TopoDS_Edge E = BRepBuilderAPI_MakeEdge(start , end);
			makeStartWire.Add(E);
			++iEdgeCount;
		}
	}
	for(vector<CIsPoint3d>::const_reverse_iterator itr = oPntList.rbegin();(itr + 1) != oPntList.rend();++itr)
	{
		gp_Pnt start(itr->x() , itr->y() , itr->z());
		gp_Pnt end((itr+1)->x() , (itr+1)->y() , (itr+1)->z());
		if(!start.IsEqual(end , OCC::COCCEntity::EPSILON))
		{
			start.Translate(gp_Vec(dir.dx() , dir.dy() , dir.dz())*thickness);
			end.Translate(gp_Vec(dir.dx() , dir.dy() , dir.dz())*thickness);

			TopoDS_Edge E = BRepBuilderAPI_MakeEdge(start , end);
			makeEndWire.Add(E);
		}
	}

	if(!(oPntList.back() == oPntList.front()))
	{
		gp_Pnt start(oPntList.back().x() , oPntList.back().y() , oPntList.back().z());
		gp_Pnt end(oPntList.front().x() , oPntList.front().y() , oPntList.front().z());
		TopoDS_Edge E = BRepBuilderAPI_MakeEdge(start , end);
		makeStartWire.Add(E);

		start.Translate(gp_Vec(dir.dx() , dir.dy() , dir.dz())*thickness);
		end.Translate(gp_Vec(dir.dx() , dir.dy() , dir.dz())*thickness);
		E = BRepBuilderAPI_MakeEdge(start , end);
		makeEndWire.Add(E);

		++iEdgeCount;
	}
	
	if(iEdgeCount > 0)
	{
		TopoDS_Face myFaceProfile = BRepBuilderAPI_MakeFace(makeStartWire);
		if(!myFaceProfile.IsNull())
		{
			gp_Vec aPrismVec(dir.dx()*thickness , dir.dy()*thickness , dir.dz()*thickness);
			oShape = BRepPrimAPI_MakePrism(myFaceProfile , aPrismVec);
		}
	}

	return oShape;
}