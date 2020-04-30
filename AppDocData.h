#pragma once

#include <vector>
using namespace std;

class CAppDocData
{
	CAppDocData(void);
	CAppDocData(const CAppDocData&){}
	CAppDocData& CAppDocData::operator=(const CAppDocData&){return (*this);}
public:
	typedef struct
	{
		STRING_T name;
		STRING_T red,green,blue;
	}ColorQuad;

	static CAppDocData& GetInstance(void);
	~CAppDocData(void);

	/// 2014.04.26 added by humkyung
	int SetUserName(const CString&);
	CString GetUserName() const;
	/// up to here

	/// 2013.11.12 added by humkyung
	CString GetFileVersion() const;
	/// up to here

	/// 2013.10.25 added by humkyung
	CString GetConfigFilePath() const;
	int SetProjectName(const STRING_T& sProjectName);
	STRING_T GetProjectName() const;
	/// up to here

	ColorQuad GetColorOf(const STRING_T&) const;
	ColorQuad GetColorAt(const int&) const;
	int GetColorCount() const;
	int LoadColorList();
	
	/// 2013.07.01 added by humkyung
	PlateCfg m_oPlateCfg;
	/// up to here
private:
	STRING_T m_sProjectName;
	CString m_sUserName;

	vector<ColorQuad> m_oColorList;
};
