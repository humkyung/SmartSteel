#include "StdAfx.h"
#include "PlateFile.h"
#include "GussetPlate.h"
#include "AppDocData.h"
#include <Tokenizer.h>

/**
	@brief	load plate shape
	@author	humkyung
	@date	2014.09.03
*/
int CPlateFile::CShape::Load(IFSTREAM_T& ifile)
{
	if(ifile.is_open())
	{
		vector<STRING_T> oResult;

		STRING_T aLine;
		while(!ifile.eof())
		{
			getline(ifile , aLine);
			if(aLine.empty()) continue;
			if('}' == aLine[0]) break;
			CTokenizer<CIsFromString>::Tokenize(oResult , aLine , CIsFromString(_T("=")));
			if((2 == oResult.size()) && (0 == STRICMP_T(_T("NAME") , oResult[0].c_str())))
			{
				m_sName = oResult[1];
			}
			else if((2 == oResult.size()) && (0 == STRICMP_T(_T("EXTR") , oResult[0].c_str())))
			{
				LoadEXTR(aLine , ifile);
			}
		}

		return ERROR_SUCCESS;
	}

	return ERROR_BAD_ENVIRONMENT;
}

/**
	@brief	load EXTR shape
	@author	humkyung
	@date	2014.09.03
*/
int CPlateFile::CShape::LoadEXTR(STRING_T aLine , IFSTREAM_T& ifile)
{
	vector<STRING_T> oResult;

	CTokenizer<CIsFromString>::Tokenize(oResult , aLine , CIsFromString(_T("=")));
	if((2 == oResult.size()) && (0 == STRICMP_T(_T("EXTR") , oResult[0].c_str())))
	{
		aLine = oResult[1];
		CTokenizer<CIsComma>::Tokenize(oResult , aLine , CIsComma());
		if(4 == oResult.size())
		{
			CAppDocData& docData = CAppDocData::GetInstance();
			const double dDivider = (UNIT::M == docData.m_oPlateCfg.unit_) ? (1.0/1000.0) : 1.;

			m_dHeight = ATOF_T(oResult[0].c_str());

			double dx = 0 , dy = 0 , dz = 0;
			SSCANF_T(oResult[2].c_str() , _T("%lf %lf %lf") , &dx , &dy , &dz);
			m_dVec.Set(dx , dy , dz);

			getline(ifile , aLine);
			if(!aLine.empty() && ('(' == aLine[0]))
			{
				while(!ifile.eof())
				{
					getline(ifile , aLine);
					if(aLine.empty()) continue;
					if(')' == aLine[0]) break;
					CTokenizer<CIsFromString>::Tokenize(oResult , aLine , CIsFromString(_T("=")));
					if((2 == oResult.size()) && (0 == STRICMP_T(_T("VERT") , oResult[0].c_str())))
					{
						CHAR_T coords[3]={'E','N','U'};
						SSCANF_T(oResult[1].c_str() , _T("%c %lf %c %lf %c %lf") , &coords[0] , &dx , &coords[1] , &dy , &coords[2] , &dz);
						m_oPntList.push_back(CIsPoint3d(dx,dy,dz)*dDivider);
					}
				}
			}

			return ERROR_SUCCESS;
		}
	}

	return ERROR_BAD_ENVIRONMENT;
}

CPlateFile::CPlateFile(void)
{
}

CPlateFile::~CPlateFile(void)
{
}

/**
	@brief	load plate file
	@author	humkyung
	@date	2014.09.03
*/
int CPlateFile::Load(list<CGussetPlate*>& oGussetPlateList , list<CEndPlate*>& oEndPlateList , const CString& sFilePath , CSmartSteelDoc* pDoc)
{
	IFSTREAM_T ifile(sFilePath.operator LPCTSTR());
	if(ifile.is_open())
	{
		STRING_T aLine;
		while(!ifile.eof())
		{
			getline(ifile , aLine);
			if(aLine.empty()) continue;
			if('{' == aLine[0])
			{
				CPlateFile::CShape shape;
				if(ERROR_SUCCESS == shape.Load(ifile))
				{
					if(_T("GUSSET PLATE") == shape.m_sName)
					{
						CGussetPlate* pGussetPlate = new CGussetPlate(NULL);
						{
							pGussetPlate->m_dThickness = shape.m_dHeight;
							pGussetPlate->m_norm = shape.m_dVec;
							pGussetPlate->m_oSectionShapePntList.insert(pGussetPlate->m_oSectionShapePntList.begin() , shape.m_oPntList.begin() , shape.m_oPntList.end());

							oGussetPlateList.push_back(pGussetPlate);
						}
					}
					else if(_T("END PLATE") == shape.m_sName)
					{
						CEndPlate* pEndPlate = new CEndPlate(NULL);
						{
							pEndPlate->m_dThickness = shape.m_dHeight;
							pEndPlate->m_norm = shape.m_dVec;
							pEndPlate->m_oSectionShapePntList.insert(pEndPlate->m_oSectionShapePntList.begin() , shape.m_oPntList.begin() , shape.m_oPntList.end());

							oEndPlateList.push_back(pEndPlate);
						}
					}
				}
			}
		}

		ifile.close();
		
		return ERROR_SUCCESS;
	}

	return ERROR_FILE_NOT_FOUND;
}