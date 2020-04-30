#pragma once

#include <IsTools.h>
#include <IsPoint3d.h>
#include <IsVect3d.h>
#include "SteelPlate.h"

class CGussetPlate : public CSteelPlate
{
public:
	typedef struct
	{
		STRING_T SectionName;
		double A , B , E , N , P , T;
	}Param;

	CGussetPlate(CSteelConnPoint*);
	~CGussetPlate(void);

	/**
	@brief	make gusset plate shape to convex hull
	*/
	void MakeShapeToConvexHull();

	/// check if has given shape
	bool HasShape(const TopoDS_Shape& selectedShape);

	int Display(Handle(AIS_InteractiveContext));
	int Show(Handle(AIS_InteractiveContext) hContext , const bool& bShow);
	int Write(OFSTREAM_T& ofile , const double dUnitScale);
private:
	friend class CGussetPlateGenerator;
	friend class CHorGussetPlateGenerator;
	friend class CPlateFile;
};
