#include "stdafx.h"

//#ifdef PYTHON

#ifdef _DEBUG
#undef _DEBUG
#include <python.h>
#define _DEBUG
#else
#include <python.h>
#endif

#include "Interpreter.h"
//#include "PythonCall.h"
#include <sstream>

CInterpreter::CInterpreter()
:   m_pyMain( NULL ) , m_pyLoadSummary(NULL) , m_pyThreadState( NULL )
{
}

CInterpreter::CInterpreter(const CInterpreter& rhs){}
CInterpreter& CInterpreter::operator=(const CInterpreter& rhs)
{
	return (*this);
}

CInterpreter::~CInterpreter()
{
	try
	{
		cleanup();
	}
	catch(...)
	{
	}
}

/**
	@brief	return the instance of python
	@author	humkyung
*/
CInterpreter& CInterpreter::GetInstance()
{
	static CInterpreter python;

	return python;
}

void CInterpreter::cleanup()
{
	try
	{
		PyEval_RestoreThread(this->_global);
		Py_Finalize();
	}
	catch(...)
	{
	}
}

#define CHECK_PYOBJECT(pyObj) (PyErr_Occurred() || NULL == pyObj)

bool CInterpreter::RedirctStdErrToFile(char* pFileName, char* mode) 
{
    PyObject* objFD = PyFile_FromString(pFileName, mode);
    if(CHECK_PYOBJECT(objFD)) {
        return false;
    }
    PyFile_SetBufSize(objFD, 0);  
    int ret = PySys_SetObject("stderr", objFD);
    if(ret != 0){
        Py_XDECREF(objFD);
        return false;
    }
    Py_XDECREF(objFD);
    return true;
}

/**
	@brief	start python script engine
	@author	humkyung
*/
const char* CInterpreter::startup(int argc , char* argv[])
{
	try
	{
		if( !Py_IsInitialized() )
		{
			Py_SetProgramName(argv[0]);
			PyEval_InitThreads();
			Py_Initialize();
			PySys_SetArgv(argc, argv);
			this->_global = PyEval_SaveThread();
		}
	//	TCHAR buffer[256] = {'\0' ,}, format[256] = {'\0' ,};

	//	CString rExecPath = GetExecPath();
	//	if(_T("\\") == rExecPath.Right(1)) rExecPath = rExecPath.Left(rExecPath.GetLength() - 1);
	//	
	//	try
	//	{
	//		CString sPythonHome(((rExecPath + _T("\\Python")).operator LPCTSTR()));
	//		Py_SetPythonHome((char*)sPythonHome.operator LPCTSTR());
	//		Py_Initialize();
	//	}
	//	catch(std::exception ex)
	//	{
	//		AfxMessageBox(ex.what());
	//	}
	//	///PyEval_InitThreads();

	//	PyRun_SimpleString(_T("import sys"));
	//	PyRun_SimpleString(_T("sys.path.append(") + rExecPath + _T("\\Python)"));
	//	if(FALSE == FileExist(rExecPath + _T("\\Temp")))
	//	{
	//		CreateDirectory(rExecPath + _T("\\Temp") , NULL);
	//	}
	//	RedirctStdErrToFile((char*)((rExecPath + _T("\\Temp\\python.log")).operator LPCTSTR()) , _T("a")); /// 리다이렉션
	//	{
	//		PyObject* pName = PyString_FromString( _T("App") );
	//		if (NULL == pName)
	//		{
	//			PyErr_Print();
	//		}

	//		m_pyMain = PyImport_Import(pName);
	//		if (NULL == m_pyMain)
	//		{
	//			PyErr_Print();
	//			AfxMessageBox(_T("Fail to load python module") , MB_ICONEXCLAMATION);
	//		}
	//		Py_DECREF(pName);
	//	}

	//	/*{
	//		PyObject* pName = PyString_FromString("LogMsgToELOAD");
	//		if (NULL == pName)
	//		{
	//			PyErr_Print();
	//		}

	//		m_pyLoadSummary = PyImport_Import(pName);
	//		if (NULL == m_pyLoadSummary)
	//		{
	//			AfxMessageBox("import went bang...\n");
	//		}
	//		Py_DECREF(pName);
	//	}
	//	*/

	//	m_pyThreadState = PyThreadState_Get();
	//	PyEval_ReleaseThread(m_pyThreadState);

	//	return (m_pyMain != NULL);
	//}

		return Py_GetPath();
	}
	catch(std::exception& ex)
	{
		::MessageBox(NULL , ex.what() , "ERROR" , MB_OK);
	}

	return NULL;
}

/*****************************************************************************
 * MODULE INTERFACE 
 * make/import/reload a python module by name
 * Note that Make_Dummy_Module could be implemented to keep a table
 * of generated dictionaries to be used as namespaces, rather than 
 * using low level tools to create and mark real modules; this 
 * approach would require extra logic to manage and use the table;
 * see basic example of using dictionaries for string namespaces;
 *****************************************************************************/


int PP_RELOAD = 0;    /* reload modules dynamically? */
int PP_DEBUG  = 0;    /* debug embedded code with pdb? */

const char *PP_Init(const char *modname) 
{
    Py_Initialize();                               /* init python if needed */
//#ifdef FC_OS_LINUX /* cannot convert `const char *' to `char *' in assignment */
    if (modname!=NULL) return modname;
    { /* we assume here that the caller frees allocated memory */
      char* __main__=(char *)malloc(sizeof("__main__"));
      return __main__="__main__";
    }
//#else    
//    return modname == NULL ? "__main__" : modname;  /* default to '__main__' */
//#endif    
}

/* returns module object named modname  */
/* modname can be "package.module" form */
static PyObject* PP_Load_Module(const char *modname)       
{                                   /* reload just resets C extension mods  */
    /* 
     * 4 cases:
     * - module "__main__" has no file, and not prebuilt: fetch or make
     * - dummy modules have no files: don't try to reload them
     * - reload=on and already loaded (on sys.modules): "reload()" before use
     * - not loaded yet, or loaded but reload=off: "import" to fetch or load 
     */

    PyObject *module, *sysmods;                  
    modname = PP_Init(modname);                       /* default to __main__ */

    if (strcmp(modname, "__main__") == 0)             /* main: no file */
        return PyImport_AddModule(modname);           /* not increfd */

    sysmods = PyImport_GetModuleDict();               /* get sys.modules dict */
    module  = PyDict_GetItemString(sysmods, modname); /* mod in sys.modules? */
    
    if (module != NULL &&                             /* dummy: no file */
        PyModule_Check(module) && 
        PyDict_GetItemString(PyModule_GetDict(module), "__dummy__")) 
	{
        return module;                                /* not increfd */
    }
    else
    if (PP_RELOAD && module != NULL && PyModule_Check(module)) {
        module = PyImport_ReloadModule(module);       /* reload file,run code */
        Py_XDECREF(module);                           /* still on sys.modules */
        return module;                                /* not increfd */
    }
    else {  
        module = PyImport_ImportModule(modname);      /* fetch or load module */
        Py_XDECREF(module);                           /* still on sys.modules */
        return module;                                /* not increfd */
    }
}

void CInterpreter::AddPythonPath(const char* Path)
{
    PyGILStateLocker locker;
    PyObject *list = PySys_GetObject("path");
    PyObject *path = PyString_FromString(Path);
    PyList_Append(list, path);
    Py_DECREF(path);
    PySys_SetObject("path", list);
}

/**
	@brief	load module
	@author	humkyung
	@date	2014.05.22
	@return	PyObject*
*/
PyObject* CInterpreter::LoadModule(const char* psModName)
{
	assert(psModName && "psModeName is NULL");

	if(psModName)
	{
		/// buffer acrobatics
		/// PyBuf ModName(psModName);
		PyObject *module;

		PyGILStateLocker locker;
		module = PP_Load_Module(psModName);
		if (!module)
			throw std::exception();

		return module;
	}

	return NULL;
}

//#endif