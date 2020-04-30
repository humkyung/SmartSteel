#include "SmartMakeFace.h"
#include <BRepFill_Filling.hxx>

CSmartMakeFace::CSmartMakeFace(void)
{
}


CSmartMakeFace::~CSmartMakeFace(void)
{
}

/******************************************************************************
	@brief		GenerateNonPlanarFaceFrom
    @author     humkyung
    @date       2014-09-23
    @class      CDgnSurfBuilder
    @function	FixWire 
    @return		int 
******************************************************************************/
TopoDS_Face CSmartMakeFace::GenerateNonPlanarFaceFrom(const TopoDS_Wire& aWire)
{
	BRepFill_Filling makeFilling;
	for( TopExp_Explorer oEdgeExplorer( aWire , TopAbs_EDGE ) ; oEdgeExplorer.More() ; oEdgeExplorer.Next() )
	{
		TopoDS_Edge anEdge = TopoDS::Edge( oEdgeExplorer.Current() );
		makeFilling.Add(anEdge,GeomAbs_C0);
	}

	makeFilling.Build();
	return makeFilling.Face();

	return TopoDS_Face();
}

/**
	@brief	make face from given wire(can treat nonplanar face)
	@author	humkyung
	@date	2014.10.07
*/
TopoDS_Face CSmartMakeFace::Make(const TopoDS_Wire& aWire)
{
	TopoDS_Face aFace;

	BRepBuilderAPI_MakeFace aMakeFace(aWire , Standard_False);
	aMakeFace.Build();
	aFace = aMakeFace.Face();
	if(aFace.IsNull())
	{
		if(BRepBuilderAPI_NotPlanar == aMakeFace.Error())
		{
			aFace = GenerateNonPlanarFaceFrom(aWire);
		}
	}

	return aFace;
}