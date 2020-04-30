#pragma once

#include <IsTools.h>
#include <IsPoint3d.h>
#include <IsVect3d.h>
#include "SteelPlate.h"

class CEndPlate : public CSteelPlate
{
public:
	typedef struct
	{
		STRING_T SectionName;
		double T , K , M;

		bool IsValid();
	}Param;

	CEndPlate(CSteelConnPoint*);
	~CEndPlate(void);

	/// check if has given shape
	bool HasShape(const TopoDS_Shape& selectedShape);

	int Display(Handle(AIS_InteractiveContext) hContext);
	int Show(Handle(AIS_InteractiveContext) hContext , const bool& bShow);
	int Write(OFSTREAM_T& ofile , const double dUnitScale);
private:
	friend class CEndPlateGenerator;
	friend class CPlateFile;
};
