#pragma once

#include <Python.h>

#ifndef PyObject_VAR_HEAD

/// forward declaration, so that you don't need to include "python.h" in other sources
class PyObject;
class PyThreadState;

#endif

#include <vector>
using namespace std;

/** If the application starts we release immediately the global interpreter lock
 * (GIL) once the Python interpreter is initialized, i.e. no thread -- including
 * the main thread doesn't hold the GIL. Thus, every thread must instantiate an
 * object of PyGILStateLocker if it needs to access protected areas in Python or
 * areas where the lock is needed. It's best to create the instance on the stack,
 * not on the heap.
 */
class PyGILStateLocker
{
public:
    PyGILStateLocker()
    {
        gstate = PyGILState_Ensure();
    }
    ~PyGILStateLocker()
    {
        PyGILState_Release(gstate);
    }

private:
    PyGILState_STATE gstate;
};

/**
 * If a thread holds the global interpreter lock (GIL) but runs a long operation
 * in C where it doesn't need to hold the GIL it can release it temporarily. Or
 * if the thread has to run code in the main thread where Python code may be 
 * executed it must release the GIL to avoid a deadlock. In either case the thread
 * must hold the GIL when instantiating an object of PyGILStateRelease.
 * As PyGILStateLocker it's best to create an instance of PyGILStateRelease on the
 * stack.
 */
class PyGILStateRelease
{
public:
    PyGILStateRelease()
    {
        // release the global interpreter lock
        state = PyEval_SaveThread();
    }
    ~PyGILStateRelease()
    {
        // grab the global interpreter lock again
        PyEval_RestoreThread(state);
    }

private:
    PyThreadState* state;
};

class CInterpreter
{
	CInterpreter();
	CInterpreter(const CInterpreter&);
	CInterpreter& operator=(const CInterpreter&);
public:
	static CInterpreter& GetInstance();
	virtual ~CInterpreter();

	/// must be called from mainthread !!!
	const char* startup(int argc , char* argv[]);
	void AddPythonPath(const char*);
	PyObject* LoadModule(const char*);
	bool RedirctStdErrToFile(char* pFileName, char* mode) ;
protected:
	friend class PythonCall;

	PyThreadState* m_pyThreadState;
	PyObject *m_pyMain , *m_pyLoadSummary;
public:
	int ReloadPython(void);
private:
	void cleanup();
	PyThreadState* _global;
};
