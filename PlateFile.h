#pragma once

class CPlateFile
{
public:
	class CShape
	{
	public:
		int Load(IFSTREAM_T& ifile);
	private:
		int LoadEXTR(STRING_T aLine , IFSTREAM_T& ifile);
	private:
		STRING_T m_sName;
		double m_dHeight;
		CIsVect3d m_dVec;
		vector<CIsPoint3d> m_oPntList;
	private:
		friend class CPlateFile;
	};

	CPlateFile(void);
	~CPlateFile(void);

	int Load(list<CGussetPlate*>& oGussetPlateList , list<CEndPlate*>& oEndPlateList , const CString& sFilePath , CSmartSteelDoc* pDoc);
};
