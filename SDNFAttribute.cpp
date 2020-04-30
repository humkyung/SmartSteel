#include "StdAfx.h"
#include "SDNFAttribute.h"

CSDNFAttribute::CSDNFAttribute(const STRING_T& sType , const STRING_T& sSection , const STRING_T& sGrade , const STRING_T& sStart , const STRING_T& sEnd)
{
	m_sType = sType;
	m_sSection = sSection;
	m_sGrade = sGrade;
	m_sStart = sStart;
	m_sEnd = sEnd;

	m_dRotation = 0.0;
}

CSDNFAttribute::~CSDNFAttribute(void)
{
}

/**
	@brief	return the type

	@author	humkyung

	@date	2013.06.11
*/
STRING_T CSDNFAttribute::type() const
{
	return m_sType;
}

/**
	@brief	return the section

	@author	humkyung

	@date	2013.06.11
*/
STRING_T CSDNFAttribute::section() const
{
	return m_sSection;
}

/**
	@brief	return the grade

	@author	humkyung

	@date	2013.06.11
*/
STRING_T CSDNFAttribute::grade() const
{
	return m_sGrade;
}

/**
	@brief	return start of linear member

	@author	humkyung

	@date	2013.06.13
*/
STRING_T CSDNFAttribute::start() const
{
	return m_sStart;
}

/**
	@brief	return end of linear member

	@author	humkyung

	@date	2013.06.13
*/
STRING_T CSDNFAttribute::end() const
{
	return m_sEnd;
}

/**
	@brief	return id
	
	@author	humkyung
	
	@date	2013.06.24
*/
STRING_T& CSDNFAttribute::id()
{
	return m_sID;
}

/**
	@brief	return cardinal point
	
	@author	humkyung
	
	@date	2013.07.02
*/
int& CSDNFAttribute::CardinalPnt()
{
	return m_iCardinalPnt;
}

/**
	@brief	return rotation
	
	@author	humkyung
	
	@date	2013.07.05
*/
double& CSDNFAttribute::Rotation()
{
	return m_dRotation;
}