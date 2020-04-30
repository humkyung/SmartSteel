#pragma once

#include <IsPoint3d.h>
#include <IsVolume.h>
#include <SDNFLinearMember.h>

class CSmartSteelDoc;
class CConnPntGenerator
{
	CConnPntGenerator(const CConnPntGenerator&){}
	CConnPntGenerator& operator=(const CConnPntGenerator&){return (*this);}
public:
	CConnPntGenerator(void);
	~CConnPntGenerator(void);

	/// generate connection point
	int Generate(CSmartSteelDoc* pDoc);
private:
	bool IntersectWith(CIsPoint3d& intsect , double& dist , const CIsLine3d& lhs , const CIsLine3d& rhs , const double& dToler=CConnPntGenerator::TOLER) const;

	/// extend given line if not blocked by
	int ExtendLine(CSmartSteelDoc* pDoc , CIsLine3d& line , const CSDNFLinearMember::ElmType& iBlockLinearType);

	/// refine conn point list
	int PostProcess(CSmartSteelDoc* pDoc);

	CSteelConnPoint* CheckDuplicate(CSmartSteelDoc* pDoc , const int& iType , const vector<CSDNFLinearMember*>& oMemList);

	/// 2013.06.19 added by humkyung
	CSteelConnPoint* CheckDuplicate(CSmartSteelDoc* pDoc , const int& iType , const CIsPoint3d& pt , list<CSteelConnPoint*>* pConnPntList = NULL , CSDNFLinearMember* pMem1=NULL , CSDNFLinearMember* pMem2=NULL);
	/// get volume of linear member
	CIsVolume GetVolumeOf(CSmartSteelDoc* pDoc , CSDNFLinearMember* pLinearMem);
private:
	static const double TOLER;
};
