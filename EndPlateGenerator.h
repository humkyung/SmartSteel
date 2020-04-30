#pragma once

#include "EndPlate.h"
#include "SteelConnPoint.h"

#include <list>
using namespace std;

class CEndPlateGenerator
{
	CEndPlateGenerator(void);
	CEndPlateGenerator(const CEndPlateGenerator&){}
	CEndPlateGenerator& operator =(const CEndPlateGenerator&){return (*this);}
private:
	/**
	@brief	check given dir is parallel to width dir of linear member
	*/
	bool IsWidthDir(CSDNFLinearMember* pMember , const CIsVect3d& dir) const;
	/**
	@brief	check given dir is parallel to height dir of linear member
	*/
	bool IsHeightDir(CSDNFLinearMember* pMember , const CIsVect3d& dir) const;
public:
	static CEndPlateGenerator& GetInstance(void);
	~CEndPlateGenerator(void);

	int Generate(list<CEndPlate*>& , CSteelConnPoint* , CSmartSteelDoc*);
};
