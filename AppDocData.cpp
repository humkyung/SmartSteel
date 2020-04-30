#include "StdAfx.h"
#include <assert.h>
#include "AppDocData.h"
#include <Util/FileVersion.h>
#include <FileTools.h>
#include <Tokenizer.h>

#include <fstream>
using namespace std;

CAppDocData::CAppDocData(void)
{
}

CAppDocData::~CAppDocData(void)
{
}

/**
	@brief	return the instance of CAppDocData;

	@author	humkyung
*/
CAppDocData& CAppDocData::GetInstance(void)
{
	static CAppDocData __instance__;

	return __instance__;
}

/**
	@brief	load color list
	@author	humkyung
	@date	2013.07.02
*/
int CAppDocData::LoadColorList()
{
	m_oColorList.clear();

	CString sAppPath;
	CFileTools::GetExecutableDirectory(sAppPath);
	if(!sAppPath.IsEmpty() && (_T('\\') != sAppPath.GetAt(sAppPath.GetLength() - 1)))
	{
		sAppPath += _T("\\");
	}
	STRING_T sFilePath = sAppPath + _T("rgb.txt");

	IFSTREAM_T ifile(sFilePath.c_str());
	if(ifile.is_open())
	{
		STRING_T aLine;
		vector<STRING_T> oResult;
		while(!ifile.eof())
		{
			getline(ifile , aLine);
			CTokenizer<CIsFromString>::Tokenize(oResult , aLine , CIsFromString(_T("\t")));
			oResult.erase(remove(oResult.begin() , oResult.end() , _T("")) , oResult.end());
			if(2 == oResult.size())
			{
				ColorQuad colorQuad;
				colorQuad.name = oResult[1];
				
				STRING_T tmp(oResult[0]);
				CTokenizer<CIsFromString>::Tokenize(oResult , aLine , CIsFromString(_T(" ")));
				oResult.erase(remove(oResult.begin() , oResult.end() , _T("")) , oResult.end());
				if(3 == oResult.size())
				{
					colorQuad.red = oResult[0];
					colorQuad.green = oResult[1];
					colorQuad.blue = oResult[2];

					m_oColorList.push_back(colorQuad);
				}
			}
		}
		ifile.close();
	}

	return ERROR_SUCCESS;
}

struct FindName : binary_function<CAppDocData::ColorQuad , STRING_T , bool>
{
	bool operator()(const CAppDocData::ColorQuad& lhs , const STRING_T& rhs) const
	{
		return (lhs.name == rhs);
	}
};

/**
	@brief	return color corresponding to given name
	@author	humkyung
	@date	2013.07.02
*/
CAppDocData::ColorQuad CAppDocData::GetColorOf(const STRING_T& name) const
{
	vector<ColorQuad>::const_iterator where = find_if(m_oColorList.begin() , m_oColorList.end() , bind2nd(FindName() , name));
	if(where != m_oColorList.end())
	{
		return (*where);
	}

	throw std::invalid_argument(_T("invalid argument"));
}

/**
	@brief	return color located at given parameter

	@author	humkyung

	@date	2013.07.02
*/
CAppDocData::ColorQuad CAppDocData::GetColorAt(const int& at) const
{
	assert((at >= 0) && (at < int(m_oColorList.size())) && "range error");
	if((at >= 0) && (at < int(m_oColorList.size())))
	{
		return m_oColorList[at];
	}

	throw std::range_error(_T("range error"));
}

/**
	@brief	return number of color

	@author	humkyung

	@date	2013.07.02
*/
int CAppDocData::GetColorCount() const
{
	return m_oColorList.size();
}

/**
	@brief	set project name
	@author	humkyung
	@date	2013.10.25
*/
int CAppDocData::SetProjectName(const STRING_T& sProjectName)
{
	m_sProjectName = sProjectName;
	return ERROR_SUCCESS;
}

/**
	@brief	get project name
	@author	humkyung
	@date	2013.10.25
*/
STRING_T CAppDocData::GetProjectName() const
{
	return m_sProjectName;
}

/**
	@brief	get configuration mdb filepath
	@author	humkyung
	@date	2013.10.25
*/
CString CAppDocData::GetConfigFilePath() const
{
	CString sAppDataPath = CFileTools::GetCommonAppDataPath() + _T("\\") + PRODUCT_PUBLISHER + _T("\\") + PRODUCT_NAME;
	if(!sAppDataPath.IsEmpty() && (_T('\\') != sAppDataPath.GetAt(sAppDataPath.GetLength() - 1)))
	{
		sAppDataPath += _T("\\");
	}

	return (sAppDataPath + _T("\\") + CString(GetProjectName().c_str()) + _T("\\") + PRODUCT_NAME + _T(".mdb"));
}

/**
	@brief	get file version
	@author	humkyung
	@date	2013.11.12
*/
CString CAppDocData::GetFileVersion() const
{
	CString res;

	TCHAR szModuleName[MAX_PATH+1];
	(void)GetModuleFileName(AfxGetInstanceHandle(), szModuleName , MAX_PATH);
	CFileVersion oVersion;
	if(TRUE == oVersion.Open(szModuleName))
	{
		res = oVersion.GetFileVersion();
	}

	return res;
}

/**
	@brief	set user name
	@author	humkyung
	@date	2014.04.26
*/
int CAppDocData::SetUserName(const CString& sUserName)
{
	m_sUserName = sUserName;
	return ERROR_SUCCESS;
}

/**
	@brief	return user name who login
	@author	humkyung
	@date	2014.04.26
*/
CString CAppDocData::GetUserName() const
{
	return m_sUserName;
}