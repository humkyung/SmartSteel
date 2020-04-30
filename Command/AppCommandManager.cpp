#include "StdAfx.h"
#include <assert.h>
#include "AppCommandManager.h"

using namespace Command;

CAppCommandManager::CAppCommandManager(void)
{
}

CAppCommandManager::~CAppCommandManager(void)
{
}

/**
	@brief	return the instance of command manager
	@author	humkyung
	@date	2013.07.27
*/
CAppCommandManager& CAppCommandManager::GetInstance()
{
	static CAppCommandManager __instance__;

	return __instance__;
}

/**
	@brief	return number of count
	@author	humkyung
	@date	2013.07.27
*/
int CAppCommandManager::GetCommandCount() const
{
	return m_oCommandList.size();
}

/**
	@brief	register command
	@author	humkyung
	@date	2013.07.27
*/
int CAppCommandManager::Add(CAbstractCommand* pCommand)
{
	assert(pCommand && "pCommand is NULL");
	if(pCommand)
	{
		if(ERROR_SUCCESS == pCommand->Do())
		{
			m_oCommandList.push_back(pCommand);
			return ERROR_SUCCESS;
		}
	}

	return ERROR_INVALID_PARAMETER;
}

/**
	@brief	undo the last command
	@author	humkyung
	@date	2013.07.27
*/
int CAppCommandManager::Undo()
{
	if(!m_oCommandList.empty())
	{
		if(ERROR_SUCCESS == m_oCommandList.back()->Undo())
		{
			CAbstractCommand* pCommand = m_oCommandList.back();
			m_oCommandList.remove(m_oCommandList.back());
			SAFE_DELETE(pCommand);
		}
	}

	return ERROR_SUCCESS;
}