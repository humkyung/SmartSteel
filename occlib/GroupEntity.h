#pragma once

#include <IsTools.h>
#include "OCCEntity.h"

#include <vector>
#include <list>
using namespace std;

namespace OCC
{
class AFX_EXT_CLASS CGroupEntity : public COCCEntity
{
	CGroupEntity(const CGroupEntity& rhs){}
public:
	CGroupEntity(void);
	~CGroupEntity(void);
	
	Bnd_Box BoundingBox();	/// 2011.12.04 added by humkyung

	int SetTransparency( const double dTransparency );	/// 2011.11.28 added by humkyung

	int Show(Handle_AIS_InteractiveContext hContext , const bool& bShow);	/// 2011.11.22 added by humkyung

	static STRING_T TypeString();
	int Add(COCCEntity* pEnt);

	COCCEntity* Clone();
	TopoDS_Shape CreateShape(BRep_Builder* pBuilder = NULL , TopoDS_Compound* pCompound = NULL);
	int SetParamList(vector<CString> oParamList);	/// 2011.10.22 added by humkyung
	int Reset(Handle_AIS_InteractiveContext hAISContext);
	void Translate(Handle_AIS_InteractiveContext hContext , gp_Vec& V);
	void Rotate( const gp_Ax1& axis , const double& angle );
	void Update(vector<CString>& oObject , Handle_AIS_InteractiveContext hContext);
	void Display(Handle_AIS_InteractiveContext hContext , const Standard_Integer& aMode);
public:
	list<pair<STRING_T,STRING_T> > m_oParamList;
	STRING_T m_sScript;
private:
	list<OCC::COCCEntity*> m_oOCCEntityList;
};
};
