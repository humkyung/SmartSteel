#pragma once

#include "occentity.h"

namespace OCC
{
class AFX_EXT_CLASS CCylinderEntity : public COCCEntity
{
public:
	CCylinderEntity(void);
	~CCylinderEntity(void);

	void Rotate( const double& angle );
	void Rotate( const gp_Ax1& axis , const double& angle );
	int Reset(Handle_AIS_InteractiveContext hAISContext);
	void Translate(gp_Vec& V);
	TopoDS_Shape CreateShape(BRep_Builder* pBuilder = NULL , TopoDS_Compound* pCompound = NULL);
	void Update(vector<CString>& oObject , Handle_AIS_InteractiveContext hContext);

	static STRING_T TypeString();
public:
	gp_Pnt m_ptStart , m_ptEnd;
	double m_dRadius;
};
};