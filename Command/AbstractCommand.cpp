#include "StdAfx.h"
#include "AbstractCommand.h"

using namespace Command;

CAbstractCommand::CAbstractCommand(const PlateSet& plateSet)
{
	m_oPlateSet.insert(m_oPlateSet.begin() , plateSet.begin() , plateSet.end());
}

CAbstractCommand::~CAbstractCommand(void)
{
}
