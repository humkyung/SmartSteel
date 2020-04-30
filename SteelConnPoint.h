#pragma once

#include <occlib.h>
#include <ExtruEntity.h>
#include <SDNFLinearMember.h>

#include <vector>
using namespace std;
class CSmartSteelDoc;
class CSteelConnPoint
{
	CSteelConnPoint(const CSteelConnPoint&){};
	CSteelConnPoint& operator=(const CSteelConnPoint&){return (*this);}
public:
	enum ConnType
	{
		NONE					= 0,
		COLUMN_BEAM_TO_VBRACE	= 1,
		COLUMN_TO_VBRACE		= 2,
		BEAM_TO_VBRACE			= 3,
		BEAM_TO_HBRACE			= 4,
		VBRACE_TO_VBRACE		= 5,
		HBRACE_TO_HBRACE		= 6,
		COLUMN_TO_BEAM			= 7
	};

	CSteelConnPoint(const ConnType& iType , const CIsPoint3d&);
	~CSteelConnPoint(void);

	/// 2013.06.26 added by humkyung
	bool HasMember(CSDNFLinearMember*) const;
	/// up to here

	void SetMemberAt(const int& , CSDNFLinearMember*);
	/// 2013.06.24 added by humkyung
	CSDNFLinearMember* GetMemberAt(const int& at);
	int GetMemberSize() const;
	/// up to here
	//int GenerateGussetPlate(CSmartSteelDoc* pDoc);
	CIsPoint3d& origin();
	CIsPoint3d origin() const;
	ConnType Type() const;
	ConnType& Type();
	/**
	@brief	insert member at given posision
	*/
	int Insert(const int& at , CSDNFLinearMember* pMember);
	int Add(CSDNFLinearMember* pMember);
	void Display(Handle(AIS_InteractiveContext));
public:
	double m_dist;
private:
	ConnType m_iType;
	CIsPoint3d m_origin;
	vector<CSDNFLinearMember*> m_oMemList;
};
