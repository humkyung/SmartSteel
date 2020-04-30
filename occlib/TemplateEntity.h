#pragma once

#include <IsTools.h>
#include "OCCEntity.h"

#include <vector>
#include <map>
using namespace std;

namespace OCC
{
class AFX_EXT_CLASS CTemplateEntity : public COCCEntity
{
	CTemplateEntity(const CTemplateEntity& rhs){}
public:
	CTemplateEntity(void);
	~CTemplateEntity(void);
	
	static STRING_T TypeString();

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
	map<STRING_T , OCC::COCCEntity*> m_oOCCEntityMap;
};
};
