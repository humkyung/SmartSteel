#pragma once

#include "StdAfx.h"

#include <IsTools.h>
#include <IsPoint3d.h>
#include <IsLine3d.h>
#include <IsVolume.h>
#include <TopoDS_Shape.hxx>

#include <vector>
using namespace std;

class COCCShapeBuilder
{
	COCCShapeBuilder(const COCCShapeBuilder&) : TOLERANCE(0.420042001050002){}
	COCCShapeBuilder& operator=(const COCCShapeBuilder&){return (*this);}
public:
	COCCShapeBuilder(/*CIsEntity* pEnt*/);
	~COCCShapeBuilder(void);

	CIsVolume Volume(const vector<CIsPoint3d>& oPntList , const CIsVect3d& dir , const double& thickness) const;

	TopoDS_Shape Shape(const CIsLine3d& line3d);
	TopoDS_Shape Shape(const vector<CIsPoint3d>& oPntList , const CIsVect3d& dir , const double& thickness);
private:
	//CIsEntity** m_pEnt;

	const double TOLERANCE;
};
