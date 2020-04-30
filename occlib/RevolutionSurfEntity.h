#pragma once

#include "OCCEntity.h"

namespace OCC
{
class AFX_EXT_CLASS CRevolutionSurfEntity : public COCCEntity
{
public:
	CRevolutionSurfEntity(void);
	~CRevolutionSurfEntity(void);

	static STRING_T TypeString();

	TopoDS_Shape CreateShape(BRep_Builder* pBuilder = NULL , TopoDS_Compound* pCompound = NULL);
	int Reset(Handle_AIS_InteractiveContext hAISContext);
public:
	TopoDS_Wire m_hWire;
	gp_Ax1 m_hAxis;
	double m_dAngle;
};
};
