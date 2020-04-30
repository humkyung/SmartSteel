#pragma once

#include "OCCEntity.h"

namespace OCC
{
class AFX_EXT_CLASS CEvolveSurfEntity : public COCCEntity
{
public:
	CEvolveSurfEntity(void);
	~CEvolveSurfEntity(void);

	static STRING_T TypeString();

	TopoDS_Shape CreateShape(BRep_Builder* pBuilder = NULL , TopoDS_Compound* pCompound = NULL);
	int Reset(Handle_AIS_InteractiveContext hAISContext);
public:
	TopoDS_Wire m_hStartWire , m_hEndWire;
};
};
