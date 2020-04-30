#pragma once
#include "occentity.h"

namespace OCC
{
class AFX_EXT_CLASS CRtorEntity : public COCCEntity
{
public:
	CRtorEntity(void);
	~CRtorEntity(void);

	void Rotate( const double& angle );
	void Rotate( const gp_Ax1& axis , const double& angle );
	int Reset(Handle_AIS_InteractiveContext hAISContext);
	TopoDS_Shape CreateShape(BRep_Builder* pBuilder = NULL , TopoDS_Compound* pCompound = NULL);
	void Translate(Handle_AIS_InteractiveContext hContext , gp_Vec& V);
	void Update(vector<CString>& oObject , Handle_AIS_InteractiveContext hContext);

	static STRING_T TypeString();
public:
	gp_Pnt m_ptP1 , m_ptP2;
	double m_dWidth , m_dHeight;
};
};