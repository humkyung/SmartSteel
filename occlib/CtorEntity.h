#pragma once
#include "occentity.h"

namespace OCC
{
class AFX_EXT_CLASS CCtorEntity : public COCCEntity
{
public:
	CCtorEntity(void);
	~CCtorEntity(void);

	gp_Vec& normal();		/// 2014.10.16 added by humkyung
	double& sweepAngle();	/// 2014.10.16 added by humkyung

	void Rotate( const double& angle );
	void Rotate( const gp_Ax1& axis , const double& angle );
	int Reset(Handle_AIS_InteractiveContext hAISContext);
	TopoDS_Shape CreateShape(BRep_Builder* pBuilder = NULL , TopoDS_Compound* pCompound = NULL);
	void Translate(gp_Vec& V);
	void Update(vector<CString>& oObject , Handle_AIS_InteractiveContext hContext);
public:
	gp_Pnt m_ptStart;
	double m_dTubeRadius;
private:
	gp_Vec m_normal;
	double m_dSweepAngle;
};
};