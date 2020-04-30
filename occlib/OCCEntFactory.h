#pragma once

#include "OCCEntity.h"

#include <map>
using namespace std;

namespace OCC
{
class AFX_EXT_CLASS COCCEntFactory
{
	COCCEntFactory(void);
public:
	static COCCEntFactory& Instance();
	~COCCEntFactory(void);
public:
	int Recycle( OCC::COCCEntity* p , Handle_AIS_InteractiveContext hAISContext);
	OCC::COCCEntity* GetEntity( const STRING_T& type );
private:
	map<STRING_T , list<OCC::COCCEntity*> >* m_pOCCEntPool;
};
};