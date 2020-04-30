#pragma once

#include "../SteelPlate.h"

#include <vector>
using namespace std;

namespace Command
{
typedef vector<CSteelPlate*> PlateSet;
class CAbstractCommand
{
public:
	CAbstractCommand(const PlateSet& plateSet);
	~CAbstractCommand(void);

	virtual int Do() = 0;
	virtual int Undo() = 0;
protected:
	PlateSet m_oPlateSet;
};
};