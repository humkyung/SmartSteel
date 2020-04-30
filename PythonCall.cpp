#include "stdafx.h"

#ifdef _DEBUG
#undef _DEBUG
#include <python.h>
#define _DEBUG
#else
#include <python.h>
#endif

#include "Interpreter.h"
#include "PythonCall.h"

// Hilfsfunktion: None an Python zur?kgeben
PyObject* Py_ReturnNone()
{
	PyObject* result = Py_None;
	Py_INCREF(result);
	return result;
}

PythonCall::PythonCall()
{
}

PythonCall::~PythonCall()
{   
}

/**
	@brief	parse the result of call
	@author	humkyung
	@date	2014.05.22
	@return	int
*/
int PythonCall::ParsePyArg(PyObject* arg , vector<STRING_T>& list)
{
	if( arg )
	{
		if( PyList_Check(arg) )
		{
			int nMax = PyList_Size(arg), nIndex = 0;
			while( nIndex < nMax )
			{
				PyObject* item = PyList_GET_ITEM(arg, nIndex);
				if(item) ParsePyArg(item , list);
				nIndex++;
			}
		}
		else if( PyTuple_Check(arg) )
		{
			int nMax = PyTuple_Size(arg), nIndex = 0;
			while( nIndex < nMax )
			{
				PyObject* item = PyTuple_GET_ITEM(arg, nIndex);
				if(item) ParsePyArg(item , list);
				nIndex++;
			}
		}
		else if(PyFloat_Check(arg))
		{
			STRINGSTREAM_T oss;
			const double value = PyFloat_AsDouble(arg);
			oss << value;
			list.push_back( oss.str().c_str() );
		}
		else if(PyString_Check(arg))
		{
			list.push_back( PyString_AsString(arg) );
		}
		else if(PyInt_Check(arg))
		{
			const long value = PyInt_AsLong(arg);
			STRINGSTREAM_T oss;
			oss << value;
			list.push_back(oss.str().c_str());
		}
		else if(Py_None == arg)
		{
			list.push_back(_T(""));
		}
		else
		{
			ASSERT( FALSE ); // result is neither list nor tuple
		}

		return ERROR_SUCCESS;
	}

	return ERROR_INVALID_PARAMETER;
}

/**	
	@brief	return the result of call
	@author	humkyung
	@date	2014.05.22
	@return	CStringList
*/
vector<STRING_T> PythonCall::result() const
{
	return m_result;
}

/**	
	@brief	call given method with arguments
	@author	humkyung
	@date	2014.05.22
	@return	void
*/
void PythonCall::call(LPCSTR pModuleName , LPCSTR pMethodName,const char *argfmt,   ...  )
{
	assert(pModuleName && pMethodName && "pModuleName or pMethodName is NULL");
	if(pModuleName && pMethodName)
	{
		PyObject *pmeth, *pargs, *presult;
		va_list argslist;                              /* "pobject.method(args)" */
		va_start(argslist, argfmt);

		PyGILStateLocker locker;
		PyObject* pModule = CInterpreter::GetInstance().LoadModule(pModuleName);
		pmeth = PyObject_GetAttrString(pModule, pMethodName);
		if (pmeth == NULL)                             /* get callable object */
			throw std::exception("Error runing InterpreterSingleton::RunMethod() method not defined");                                 /* bound method? has self */

		pargs = Py_VaBuildValue(argfmt, argslist);     /* args: c->python */

		if (pargs == NULL) 
		{
			Py_DECREF(pmeth);
			throw std::exception("InterpreterSingleton::RunMethod() wrong arguments");
		}

		presult = PyEval_CallObject(pmeth, pargs);   /* run interpreter */

		Py_DECREF(pmeth);
		Py_DECREF(pargs);

		m_result.clear();
		ParsePyArg(presult , m_result);
	}
}