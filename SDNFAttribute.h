#pragma once

#include <IsTools.h>
#include <OCCAttribute.h>

class CSDNFAttribute : public OCC::COCCAttribute
{
	CSDNFAttribute(const CSDNFAttribute& rhs){}
	CSDNFAttribute& operator=(const CSDNFAttribute& rhs){ return (*this); }
public:
	CSDNFAttribute(const STRING_T& sType , const STRING_T& sSection , const STRING_T& sGrade , const STRING_T& sStart , const STRING_T& sEnd);
	~CSDNFAttribute(void);

	double& Rotation();
	int& CardinalPnt();
	STRING_T& id();		/// 2013.06.24 added by humkyung
	STRING_T type() const;
	STRING_T section() const;
	STRING_T grade() const;
	STRING_T start() const;
	STRING_T end() const;
private:
	STRING_T m_sID;	/// 2013.06.24 added by humkyung
	STRING_T m_sType;
	STRING_T m_sSection , m_sGrade;
	int m_iCardinalPnt;	/// 2013.07.02 added by humkyung
	STRING_T m_sStart , m_sEnd;
	double m_dRotation;
};
