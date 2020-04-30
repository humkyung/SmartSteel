#pragma once

#include "AbstractCommand.h"

namespace Command
{
class CAppCommandManager
{
	CAppCommandManager(void);
	CAppCommandManager(const CAppCommandManager&){}
	CAppCommandManager& operator=(const CAppCommandManager&){return (*this);}
public:
	static CAppCommandManager& GetInstance();
	
	~CAppCommandManager(void);

	int GetCommandCount() const;
	int Add(CAbstractCommand*);
	/// undo the last command
	int Undo();
private:
	list<CAbstractCommand*> m_oCommandList;
};
};