#pragma once

#include "AbstractCommand.h"

namespace Command
{
	class CDeleteCommand : public CAbstractCommand
{
public:
	CDeleteCommand(const PlateSet& plateSet);
	~CDeleteCommand(void);

	int Do();
	int Undo();
};
};