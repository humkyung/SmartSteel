#pragma once
#include "OCCEntity.h"

class CSmartMakeFace
{
public:
	CSmartMakeFace(void);
	~CSmartMakeFace(void);

	TopoDS_Face Make(const TopoDS_Wire& aWire);
private:
	TopoDS_Face GenerateNonPlanarFaceFrom(const TopoDS_Wire& aWire);
};

