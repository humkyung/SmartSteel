#pragma once

#ifdef _DEBUG
#undef _DEBUG
#include <python.h>
#define _DEBUG
#else
#include <python.h>
#endif
#include "Interpreter.h"

class PythonCall
{
	PythonCall(const PythonCall& ){}
	PythonCall& operator=(const PythonCall&){return (*this);}
public:
	PythonCall( );
	~PythonCall();

	void call(LPCSTR pModuleName , LPCSTR pMethodName,const char *argfmt,   ...  );
	vector<STRING_T> result() const;
private:
	int ParsePyArg(PyObject* arg , vector<STRING_T>& list);
	vector<STRING_T> m_result;
};
