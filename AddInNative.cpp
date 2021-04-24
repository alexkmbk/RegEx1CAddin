#if !defined( __linux__ ) && !defined(__APPLE__) && !defined(__ANDROID__)
#include "stdafx.h"
#endif

#include <stdio.h>
#include <wchar.h>
#include "AddInNative.h"



static std::map<std::u16string, long> mMethods;
static std::vector<std::u16string> vMethods;
static std::map<std::u16string, long> mMethods_ru;
static std::vector<std::u16string> vMethods_ru;

static std::map<std::u16string, long> mProps;
static std::vector<std::u16string> vProps;
static std::map<std::u16string, long> mProps_ru;
static std::vector<std::u16string> vProps_ru;

static AppCapabilities g_capabilities = eAppCapabilitiesInvalid;

inline void fillMap(std::map<std::u16string, long>& map, const std::vector<std::u16string> & vector) {
	long index = 0;
	for (auto &item : vector)
	{
		auto lowCasedItem = item;
		tolowerStr(lowCasedItem);
		map.insert({ lowCasedItem, index });
		index++;
	}
}

//---------------------------------------------------------------------------//
long GetClassObject(const WCHAR_T* wsName, IComponentBase** pInterface)
{
	if (!*pInterface)
	{
		*pInterface = new CAddInNative;
		return (long)*pInterface;
	}
	return 0;
}
//---------------------------------------------------------------------------//
AppCapabilities SetPlatformCapabilities(const AppCapabilities capabilities)
{
	g_capabilities = capabilities;
	return eAppCapabilitiesLast;
}
//---------------------------------------------------------------------------//
long DestroyObject(IComponentBase** pIntf)
{
	if (!*pIntf)
		return -1;

	delete *pIntf;
	*pIntf = 0;
	return 0;
}
//---------------------------------------------------------------------------//
const WCHAR_T* GetClassNames()
{
	static char16_t cls_names[] = u"RegEx";
	return reinterpret_cast<WCHAR_T *>(cls_names);
}
//---------------------------------------------------------------------------//

// CAddInNative
//---------------------------------------------------------------------------//
CAddInNative::CAddInNative()
{
	m_iMemory = 0;
	m_iConnect = 0;

	iCurrentPosition = -1;
	uiSubMatchesCount = 0;
	m_PropCountOfItemsInSearchResult = 0;
	sErrorDescription = u"";
	bThrowExceptions = false;
	bIgnoreCase = false;
	isPattern = false;
	bGlobal = false;
	bHierarchicalResultIteration = false;

	sPattern.clear();

	if (mMethods.size() == 0) {
		vMethods = { u"Matches", u"IsMatch", u"Next", u"Replace", u"Count", u"SubMatchesCount", u"GetSubMatch", u"Version" };
		fillMap(mMethods, vMethods);
	}

	if (mMethods_ru.size() == 0) {
		vMethods_ru = { u"НайтиСовпадения", u"Совпадает", u"Следующий", u"Заменить", u"Количество", u"КоличествоВложенныхГрупп", u"ПолучитьПодгруппу", u"Версия" };
		fillMap(mMethods_ru, vMethods_ru);
	}

	if (mProps.size() == 0) {
		vProps = { u"CurrentValue", u"IgnoreCase", u"ErrorDescription", u"ThrowExceptions", u"Pattern", u"Global" };
		fillMap(mProps, vProps);
	}

	if (mProps_ru.size() == 0) {
		vProps_ru = { u"ТекущееЗначение", u"ИгнорироватьРегистр", u"ОписаниеОшибки", u"ВызыватьИсключения", u"Шаблон", u"ВсеСовпадения" };
		fillMap(mProps_ru, vProps_ru);
	}
}
//---------------------------------------------------------------------------//
CAddInNative::~CAddInNative()
{
}
//---------------------------------------------------------------------------//
bool CAddInNative::Init(void* pConnection)
{
	m_iConnect = (IAddInDefBase*)pConnection;
	return m_iConnect != NULL;
}
//---------------------------------------------------------------------------//
long CAddInNative::GetInfo()
{
	// Component should put supported component technology version 
	// This component supports 2.1 version
	return 2100;
}
//---------------------------------------------------------------------------//
void CAddInNative::Done()
{

}
/////////////////////////////////////////////////////////////////////////////
// ILanguageExtenderBase
//---------------------------------------------------------------------------//
bool CAddInNative::RegisterExtensionAs(WCHAR_T** wsExtensionName)
{
	char16_t name[] = u"RegEx";

	if (!m_iMemory || !m_iMemory->AllocMemory(reinterpret_cast<void **>(wsExtensionName), sizeof(name))) {
		return false;
	};

	memcpy(*wsExtensionName, name, sizeof(name));

	return true;
}
//---------------------------------------------------------------------------//
long CAddInNative::GetNProps()
{
	return ePropLast;
}
//---------------------------------------------------------------------------//
long CAddInNative::FindProp(const WCHAR_T* wsPropName)
{
	std::basic_string<char16_t> usPropName = (char16_t*)(wsPropName);
	tolowerStr(usPropName);

	auto it = mProps.find(usPropName);
	if (it != mProps.end())
		return it->second;

	it = mProps_ru.find(usPropName);
	if (it != mProps_ru.end())
		return it->second;

	return -1;
}

//---------------------------------------------------------------------------//
const WCHAR_T* CAddInNative::GetPropName(long lPropNum, long lPropAlias)
{
	if (lPropNum >= ePropLast)
		return NULL;

	std::basic_string<char16_t> *usCurrentName;

	switch (lPropAlias)
	{
	case 0: // First language
		usCurrentName = &vProps[lPropNum];
		break;
	case 1: // Second language
		usCurrentName = &vProps_ru[lPropNum];
		break;
	default:
		return 0;
	}

	if (usCurrentName->length() == 0) {
		return nullptr;
	}

	WCHAR_T *result = nullptr;

	size_t bytes = (usCurrentName->length() + 1) * sizeof(char16_t);

	if (!m_iMemory || !m_iMemory->AllocMemory(reinterpret_cast<void **>(&result), bytes)) {
		return nullptr;
	};

	memcpy(result, usCurrentName->c_str(), bytes);

	return result;
}
//---------------------------------------------------------------------------//
bool CAddInNative::GetPropVal(const long lPropNum, tVariant* pvarPropVal)
{
	switch (lPropNum)
	{
	case ePropCurrentValue:
	{
		TV_VT(pvarPropVal) = VTYPE_PWSTR;
		std::basic_string<char16_t>* wsCurrentValue;
		if (vResults.size() == 0 || iCurrentPosition == -1 || iCurrentPosition >= vResults.size())
		{
			std::basic_string<char16_t> emptyStr = u"";
			wsCurrentValue = &emptyStr;
		}
		else
			wsCurrentValue = &vResults[iCurrentPosition];

		if (m_iMemory->AllocMemory((void**)&pvarPropVal->pwstrVal, (wsCurrentValue->length() + 1) * sizeof(wchar_t)))
		{
			memcpy(pvarPropVal->pwstrVal, wsCurrentValue->c_str(), (wsCurrentValue->length() + 1) * sizeof(wchar_t));
			pvarPropVal->wstrLen = wsCurrentValue->length();
			return true;
		}
		return false;
	};
	case ePropErrorDescription:
	{
		if (m_iMemory->AllocMemory((void**)&pvarPropVal->pwstrVal, (sErrorDescription.length() + 1) * sizeof(char16_t)))
		{
			memcpy(pvarPropVal->pwstrVal, sErrorDescription.c_str(), (sErrorDescription.length() + 1) * sizeof(char16_t));
			TV_VT(pvarPropVal) = VTYPE_PSTR;
			pvarPropVal->wstrLen = sErrorDescription.length();
			return true;
		}
		else
		{
			TV_VT(pvarPropVal) = VTYPE_PWSTR;
			pvarPropVal->pwstrVal = NULL;
			pvarPropVal->wstrLen = 0;
			return true;
		}
	}
	case ePropThrowExceptions:
	{
		TV_VT(pvarPropVal) = VTYPE_BOOL;
		pvarPropVal->bVal = bThrowExceptions;
		return true;
	}
	case ePropPattern:
	{
		if (m_iMemory->AllocMemory((void**)&pvarPropVal->pwstrVal, (sPattern.length() + 1) * sizeof(WCHAR)))
		{
			memcpy(pvarPropVal->pwstrVal, sPattern.c_str(), sPattern.length() * sizeof(WCHAR));
			TV_VT(pvarPropVal) = VTYPE_PWSTR;
			pvarPropVal->wstrLen = sPattern.length();
			return true;
		}
		else
		{
			TV_VT(pvarPropVal) = VTYPE_PWSTR;
			pvarPropVal->pwstrVal = NULL;
			pvarPropVal->wstrLen = 0;
			return true;
		}
	}
	case ePropGlobal:
	{
		TV_VT(pvarPropVal) = VTYPE_BOOL;
		pvarPropVal->bVal = bGlobal;
		return true;
	}
	case ePropIgnoreCase:
	{
		TV_VT(pvarPropVal) = VTYPE_BOOL;
		pvarPropVal->bVal = bIgnoreCase;
		return true;
	}
	default:
		return false;
	}

	return true;
}
//---------------------------------------------------------------------------//
bool CAddInNative::SetPropVal(const long lPropNum, tVariant *varPropVal)
{
	SetLastError(u"");

	switch (lPropNum)
	{
	case ePropThrowExceptions:
	{
		if (TV_VT(varPropVal) != VTYPE_BOOL)
			return false;
		bThrowExceptions = TV_BOOL(varPropVal);
		return true;
	}
	case ePropIgnoreCase: {
		bIgnoreCase = TV_BOOL(varPropVal);
		return true;
	}
	case ePropPattern: 
	{
		int errornumber;
		PCRE2_SIZE erroroffset = NULL;

			sPattern.assign((char16_t*)varPropVal->pwstrVal, varPropVal->wstrLen);
			rePattern = pcre2_compile((PCRE2_SPTR16) (varPropVal->pwstrVal),               /* the pattern */
			varPropVal->wstrLen, /* indicates pattern is zero-terminated */
			PCRE2_UTF|(bIgnoreCase ? PCRE2_CASELESS: 0), /* options */
			&errornumber,          /* for error number */
			&erroroffset,          /* for error offset */
			NULL);

		if (rePattern == NULL)
		{
			PCRE2_UCHAR buffer[256];
			pcre2_get_error_message_16(errornumber, buffer, sizeof(buffer));
			vResults.clear();
			mSubMatches.clear();
			isPattern = false;
			iCurrentPosition = -1;
			uiSubMatchesCount = 0;
			m_PropCountOfItemsInSearchResult = 0;
			//SetLastError(buffer);
			if (bThrowExceptions)
				return false;
			else
				return true;
		}
		/*try
		{
			rePattern.assign(varPropVal->pwstrVal, varPropVal->wstrLen, (bIgnoreCase) ? boost::regex_constants::icase : boost::regex_constants::normal);
		}
		catch (const std::exception& e)
		{
			vResults.clear();
			mSubMatches.clear();
			isPattern = false;
			iCurrentPosition = -1;
			uiSubMatchesCount = 0;
			m_PropCountOfItemsInSearchResult = 0;
			SetLastError(e.what());
			if (bThrowExceptions)
				return false;
			else
				return true;
		}*/
//#endif
		isPattern = true;
		return true;
	}
	case ePropGlobal:
	{
		bGlobal = TV_BOOL(varPropVal);
		return true;
	}
	default:
		return false;
	}

	return false;
}
//---------------------------------------------------------------------------//
bool CAddInNative::IsPropReadable(const long lPropNum)
{
	switch (lPropNum)
	{
	case ePropCurrentValue:
		return true;
	case ePropErrorDescription:
		return true;
	case ePropIgnoreCase:
		return true;
	case ePropThrowExceptions:
		return true;
	case ePropPattern:
		return true;
	case ePropGlobal:
		return true;
	default:
		return false;
	}

	return false;
}
//---------------------------------------------------------------------------//
bool CAddInNative::IsPropWritable(const long lPropNum)
{
	switch (lPropNum)
	{
	case ePropIgnoreCase:
		return true;
	case ePropThrowExceptions:
		return true;
	case ePropPattern:
		return true;
	case ePropGlobal:
		return true;
	default:
		return false;
	}

	return false;
}
//---------------------------------------------------------------------------//
long CAddInNative::GetNMethods()
{
	return eMethLast;
}
//---------------------------------------------------------------------------//
long CAddInNative::FindMethod(const WCHAR_T* wsMethodName)
{
	std::basic_string<char16_t> usMethodName = (char16_t*)(wsMethodName);
	tolowerStr(usMethodName);

	auto it = mMethods.find(usMethodName);
	if (it != mMethods.end())
		return it->second;

	it = mMethods_ru.find(usMethodName);
	if (it != mMethods_ru.end())
		return it->second;

	return -1;
}
//---------------------------------------------------------------------------//
const WCHAR_T* CAddInNative::GetMethodName(const long lMethodNum, const long lMethodAlias)
{

	if (lMethodNum >= eMethLast)
		return NULL;

	std::basic_string<char16_t> *usCurrentName;

	switch (lMethodAlias)
	{
	case 0: // First language
		usCurrentName = &vMethods[lMethodNum];
		break;
	case 1: // Second language
		usCurrentName = &vMethods_ru[lMethodNum];
		break;
	default:
		return 0;
	}

	WCHAR_T *result = nullptr;

	size_t bytes = (usCurrentName->length() + 1) * sizeof(char16_t);

	if (!m_iMemory || !m_iMemory->AllocMemory(reinterpret_cast<void **>(&result), bytes)) {
		return nullptr;
	};

	memcpy(result, usCurrentName->c_str(), bytes);

	return result;
}
//---------------------------------------------------------------------------//
long CAddInNative::GetNParams(const long lMethodNum)
{
	switch (lMethodNum)
	{
	case eMethMatches:
		return 3;
	case eMethIsMatch:
		return 2;
	case eMethReplace:
		return 3;
	case eMethGetSubMatch:
		return 1;
	default:
		return 0;
	}

	return 0;
}
//---------------------------------------------------------------------------//
bool CAddInNative::GetParamDefValue(const long lMethodNum, const long lParamNum,
	tVariant *pvarParamDefValue)
{
	switch (lMethodNum)
	{
	case eMethMatches:
	{
		if (lParamNum == 1)
		{
			TV_VT(pvarParamDefValue) = VTYPE_PWSTR;
			pvarParamDefValue->wstrLen = 0;
			return true;
		}
		else if (lParamNum == 2)
		{
			TV_VT(pvarParamDefValue) = VTYPE_BOOL;
			pvarParamDefValue->bVal = false;
			return true;
		}
		break;
	}
	case eMethIsMatch:
	{
		if (lParamNum == 1)
		{
			TV_VT(pvarParamDefValue) = VTYPE_PWSTR;
			pvarParamDefValue->wstrLen = 0;
			return true;
		}
		break;
	}
	case eMethNext:
		break;
	case eMethReplace:
	{
		if (lParamNum == 1)
		{
			TV_VT(pvarParamDefValue) = VTYPE_PWSTR;
			pvarParamDefValue->wstrLen = 0;
			return true;
		}
		break;
	}
	default:
	{
		TV_VT(pvarParamDefValue) = VTYPE_EMPTY;
		return false;
	}
	}

	TV_VT(pvarParamDefValue) = VTYPE_EMPTY;
	return false;
}
//---------------------------------------------------------------------------//
bool CAddInNative::HasRetVal(const long lMethodNum)
{
	switch (lMethodNum)
	{
	case eMethNext:
		return true;
	case eMethReplace:
		return true;
	case eMethIsMatch:
		return true;
	case eMethCount:
		return true;
	case eMethVersion:
		return true;
	case eMethGetSubMatch:
		return true;
	case eMethSubMatchesCount:
		return true;
	default:
		return false;
	}

	return false;
}
//---------------------------------------------------------------------------//
bool CAddInNative::CallAsProc(const long lMethodNum,
	tVariant* paParams, const long lSizeArray)
{
	switch (lMethodNum)
	{
	case eMethMatches:
		return search(paParams);
	case eMethIsMatch:
		break;
	case eMethReplace:
		break;
	case eMethCount:
		break;
	case eMethVersion:
		break;
	case eMethGetSubMatch:
		break;
	case eMethSubMatchesCount:
		break;
	case eMethNext:
	{
		if (iCurrentPosition >= m_PropCountOfItemsInSearchResult) {
			return true;
		}
		else {
			iCurrentPosition++;
			return true;
		}
	}
	break;
	default:
		return false;
	}
	return true;
}
//---------------------------------------------------------------------------//
bool CAddInNative::CallAsFunc(const long lMethodNum,
	tVariant* pvarRetValue, tVariant* paParams, const long lSizeArray)
{
	switch (lMethodNum)
	{
		// Method acceps one argument of type BinaryData ant returns its copy
	case eMethMatches:
		break;
	case eMethIsMatch:
	{
		match(pvarRetValue, paParams);
		return true;
	}
	break;
	case eMethReplace:
		return replace(pvarRetValue, paParams);
	case eMethNext:
	{
		iCurrentPosition++;
		if (iCurrentPosition >= m_PropCountOfItemsInSearchResult) {
			TV_VT(pvarRetValue) = VTYPE_BOOL;
			pvarRetValue->bVal = false;
			return true;
		}
		else {
			TV_VT(pvarRetValue) = VTYPE_BOOL;
			pvarRetValue->bVal = true;
			return true;
		}
	}
	break;
	case eMethCount: {
		TV_VT(pvarRetValue) = VTYPE_I4;
		pvarRetValue->lVal = m_PropCountOfItemsInSearchResult;
		return true;
	}
	case eMethSubMatchesCount: {
		TV_VT(pvarRetValue) = VTYPE_I4;
		pvarRetValue->lVal = uiSubMatchesCount;
		return true;
	}
	case eMethGetSubMatch: {		
		return getSubMatch(pvarRetValue, paParams);
	}
	case eMethVersion: {
		version(pvarRetValue);
		return true;
	}
	default:
		return false;
	}
	return true;
}

void CAddInNative::SetLocale(const WCHAR_T* loc)
{
#if !defined( __linux__ ) && !defined(__APPLE__) && !defined(__ANDROID__)
	_wsetlocale(LC_ALL, L"");
#else
	setlocale(LC_ALL, "");
#endif
}

/////////////////////////////////////////////////////////////////////////////
// LocaleBase
//---------------------------------------------------------------------------//
bool CAddInNative::setMemManager(void* mem)
{
	m_iMemory = (IMemoryManager*)mem;
	return m_iMemory != 0;
}

bool CAddInNative::search(tVariant * paParams)
{
	SetLastError(u"");
	vResults.clear();
	mSubMatches.clear();
	iCurrentPosition = -1;
	uiSubMatchesCount = 0;

	//boost::wsmatch wsmMatch;
	pcre2_code* pattern = NULL;
	bool bClearPattern = false;

	bHierarchicalResultIteration = paParams[2].bVal;

	if (paParams[1].wstrLen == 0)
	{
		if (isPattern)
			pattern = rePattern;
		else
			if (bThrowExceptions) {
				SetLastError(u"The regexp template is not specified.");
				return false;
			}
			else
				return true;
	}
	else
	{
		bClearPattern = true;
		pattern = GetPattern(paParams, 1);
	}

	if (pattern == NULL)
	{
		if (bThrowExceptions)
			return false;
		else
			return true;
	}

	PCRE2_SIZE start_offset = 0;
	size_t rootIndex = 0;
	int rc;

	pcre2_match_data* match_data = pcre2_match_data_create_from_pattern(pattern, NULL);

	while (true) {
		rc = pcre2_match(
			pattern,                   /* the compiled pattern */
			(PCRE2_SPTR16)paParams[0].pwstrVal,              /* the subject string */
			paParams[0].wstrLen,       /* the length of the subject */
			start_offset,                    /* start at offset 0 in the subject */
			0,                    /* default options */
			match_data,           /* block for storing the result */
			NULL);

		if (rc < 0)
		{
				if (rc == PCRE2_ERROR_NOMATCH)
					break;
				/*case PCRE2_ERROR_NOMATCH: printf("No match\n"); break;
					/*
					Handle other special cases if you like
					*/
					//default: printf("Matching error %d\n", rc); break;
			pcre2_match_data_free(match_data);   /* Release memory used for the match */
			//pcre2_code_free(pattern);                 /*   data and the compiled pattern. */
			if (bThrowExceptions)
				return false;
			else
				return true;
		}
		else if (rc == 0) {
			break;
		}
		PCRE2_SIZE *ovector = pcre2_get_ovector_pointer(match_data);


		std::vector<std::basic_string<char16_t>> vSubMatches;

		for (int i = 0; i < rc; i++)
		{
			std::basic_string<char16_t> r;

			r.assign((char16_t*)(paParams[0].pwstrVal + ovector[2 * i]), ovector[2 * i + 1] - ovector[2 * i]);

			if (i == 0 && bHierarchicalResultIteration) {
				vResults.push_back(r);
				rootIndex = vResults.size() - 1;
			}
			else if (bHierarchicalResultIteration)
				vSubMatches.push_back(r);
			else if (!bHierarchicalResultIteration)
				vResults.push_back(r);
		}
		if (bHierarchicalResultIteration)
			mSubMatches.insert({ rootIndex, vSubMatches});
		if (bGlobal)
			start_offset = ovector[1];
		else
			break;
	}
	uiSubMatchesCount = rc - 1;

	/*std::wstring str;
	GetStrParam(str, paParams, 0);
	try
	{
		if (paParams[1].wstrLen == 0)
		{
			if (isPattern)
				pattern = &rePattern;
			else
				return true;
		}
		else
		{
			bClearPattern = true;
			pattern = GetPattern(paParams, 1);
		}

		if (bGlobal)
		{
			std::wstring::const_iterator start = str.begin();
			std::wstring::const_iterator end = str.end();

			while (start < end &&
				boost::regex_search(start, end, wsmMatch, *pattern))
			{
				size_t i = 0;
				size_t rootIndex = 0;
				std::vector<std::wstring> vSubMatches;
				for (auto &r : wsmMatch) {
					if (i == 0 && bHierarchicalResultIteration) {
						vResults.push_back(r);
						rootIndex = vResults.size() - 1;
					}
					else if (bHierarchicalResultIteration)
						vSubMatches.push_back(r);
					else if (!bHierarchicalResultIteration)
						vResults.push_back(r);
					i++;
				}
				mSubMatches.insert({ rootIndex, vSubMatches });
				start = wsmMatch[0].second;
			}
			uiSubMatchesCount = wsmMatch.size() - 1;
		}
		else
		{
			bool res = boost::regex_search(str, wsmMatch, *pattern);
			uiSubMatchesCount = wsmMatch.size() - 1;
			if (res) {
				size_t i = 0;
				std::vector<std::wstring> vSubMatches;
				for (auto &r : wsmMatch) {
					if (i == 0 && bHierarchicalResultIteration) {
						vResults.push_back(r);
					}
					else if (bHierarchicalResultIteration)
						vSubMatches.push_back(r);
					else if (!bHierarchicalResultIteration)
						vResults.push_back(r);
					i++;
				}
				mSubMatches.insert({ 0, vSubMatches });
			}
		}
	}
	catch (const std::exception& e)
	{
		SetLastError(e.what());
		vResults.clear();
		mSubMatches.clear();
		iCurrentPosition = -1;
		uiSubMatchesCount = 0;
		m_PropCountOfItemsInSearchResult = 0;
		if (bClearPattern && pattern != NULL)
			delete pattern;
		if (bThrowExceptions)
			return false;
		else
			return true;
	}*/
	pcre2_match_data_free(match_data);
	iCurrentPosition = -1;
	m_PropCountOfItemsInSearchResult = vResults.size();
	if (bClearPattern && pattern != NULL)
		delete pattern;
	return true;
}

bool CAddInNative::replace(tVariant * pvarRetValue, tVariant * paParams)
{
	SetLastError(u"");
	TV_VT(pvarRetValue) = VTYPE_PWSTR;
	
	pcre2_code* pattern = NULL;
	bool bClearPattern = false;

	if (paParams[1].wstrLen == 0)
	{
		if (isPattern)
			pattern = rePattern;
		else
			return true;
	}
	else
	{
		bClearPattern = true;
		pattern = GetPattern(paParams, 1);
	}

	PCRE2_SIZE start_offset = 0;
	size_t rootIndex = 0;
	int rc;

	pcre2_match_data* match_data = pcre2_match_data_create_from_pattern(pattern, NULL);

	PCRE2_UCHAR outputbuffer;
	PCRE2_SIZE outlength = 0;

	rc = pcre2_substitute(pattern, 
		(PCRE2_SPTR16)paParams[0].pwstrVal, 
		paParams[0].wstrLen,
		0,
		bGlobal?PCRE2_SUBSTITUTE_GLOBAL | PCRE2_SUBSTITUTE_OVERFLOW_LENGTH | PCRE2_SUBSTITUTE_EXTENDED : PCRE2_SUBSTITUTE_OVERFLOW_LENGTH | PCRE2_SUBSTITUTE_EXTENDED,
		match_data,
		NULL,
		(PCRE2_SPTR16)paParams[2].pwstrVal, 
		paParams[2].wstrLen, 0, &outlength);

	if (rc == PCRE2_ERROR_NOMEMORY)
	{
		if (m_iMemory->AllocMemory((void**)&pvarRetValue->pwstrVal, (outlength + 1) * sizeof(WCHAR)))
		{
			rc = pcre2_substitute(pattern,
				(PCRE2_SPTR16)paParams[0].pwstrVal,
				paParams[0].wstrLen,
				0,
				bGlobal ? PCRE2_SUBSTITUTE_GLOBAL | PCRE2_SUBSTITUTE_OVERFLOW_LENGTH | PCRE2_SUBSTITUTE_EXTENDED : PCRE2_SUBSTITUTE_OVERFLOW_LENGTH | PCRE2_SUBSTITUTE_EXTENDED,
				match_data,
				NULL,
				(PCRE2_SPTR16)paParams[2].pwstrVal,
				paParams[2].wstrLen, (PCRE2_UCHAR *)pvarRetValue->pwstrVal, &outlength); 
			if (rc > 0) 
				pvarRetValue->wstrLen = outlength;
			else
				pvarRetValue->wstrLen = 0;
		}
	}
	else
		pvarRetValue->wstrLen = 0;

	pcre2_match_data_free(match_data);

	iCurrentPosition = 0;
	m_PropCountOfItemsInSearchResult = 0;
	vResults.clear();
	mSubMatches.clear();
	if (bClearPattern && pattern != NULL)
		delete pattern;

	/*SetLastError("");
	TV_VT(pvarRetValue) = VTYPE_PWSTR;
	iCurrentPosition = 0;
	m_PropCountOfItemsInSearchResult = 0;
	vResults.clear();
	mSubMatches.clear();

	boost::wregex* pattern = NULL;
	bool bClearPattern = false;

	std::wstring str;
	GetStrParam(str, paParams, 0);

	std::wstring res;
	try
	{
		if (paParams[1].wstrLen == 0)
		{
			if (isPattern)
				pattern = &rePattern;
			else
				return true;
		}
		else
		{
			bClearPattern = true;
			pattern = GetPattern(paParams, 1);
		}
#if defined( __linux__ ) || defined(__APPLE__) || defined(__ANDROID__)

		std::wstring replacement;
		GetStrParam(replacement, paParams, 2);
		if (bGlobal)
			res = boost::regex_replace(str, *pattern, replacement);
		else
			res = boost::regex_replace(str, *pattern, replacement, boost::regex_constants::format_first_only);
#else
		if (bGlobal)
			res = boost::regex_replace(str, *pattern, paParams[2].pwstrVal);
		else
			res = boost::regex_replace(str, *pattern, paParams[2].pwstrVal, boost::regex_constants::format_first_only);

#endif
	}
	catch (const std::exception& e)
	{
		vResults.clear();
		mSubMatches.clear();
		iCurrentPosition = -1;
		uiSubMatchesCount = 0;
		m_PropCountOfItemsInSearchResult = 0;
		if (bClearPattern && pattern != NULL)
			delete pattern;
		SetLastError(e.what());
		if (bThrowExceptions)
			return false;
		else
			return true;
	}

#if defined( __linux__ ) || defined(__APPLE__) || defined(__ANDROID__)

	if (m_iMemory->AllocMemory((void**)&pvarRetValue->pwstrVal, (res.length() + 1) * sizeof(char32_t)))
	{
		convertUTF32ToUTF16(res.c_str(), res.length(), (char16_t *)pvarRetValue->pwstrVal);
		pvarRetValue->wstrLen = res.length();
		return true;
	}
#else
	if (m_iMemory->AllocMemory((void**)&pvarRetValue->pwstrVal, (res.length() + 1) * sizeof(wchar_t)))
	{
		memcpy(pvarRetValue->pwstrVal, res.c_str(), (res.length() + 1) * sizeof(wchar_t));
		pvarRetValue->wstrLen = res.length();
		return true;
	}
#endif

	if (bClearPattern && pattern != NULL)
		delete pattern;*/

	return true;
}

bool CAddInNative::match(tVariant * pvarRetValue, tVariant * paParams)
{
	SetLastError(u"");
	TV_VT(pvarRetValue) = VTYPE_BOOL;

	pcre2_code* pattern = NULL;
	bool bClearPattern = false;

	if (paParams[1].wstrLen == 0)
	{
		if (isPattern)
			pattern = rePattern;
		else
			return true;
	}
	else
	{
		bClearPattern = true;
		pattern = GetPattern(paParams, 1);
	}

	PCRE2_SIZE start_offset = 0;
	size_t rootIndex = 0;
	int rc;

	pcre2_match_data* match_data = pcre2_match_data_create_from_pattern(pattern, NULL);

		rc = pcre2_match(
			pattern,                   /* the compiled pattern */
			(PCRE2_SPTR16)paParams[0].pwstrVal,              /* the subject string */
			paParams[0].wstrLen,       /* the length of the subject */
			start_offset,                    /* start at offset 0 in the subject */
			0,                    /* default options */
			match_data,           /* block for storing the result */
			NULL);

		if (rc > 0)
			pvarRetValue->bVal = true;
		else
			pvarRetValue->bVal = false;

		pcre2_match_data_free(match_data);

		iCurrentPosition = 0;
		m_PropCountOfItemsInSearchResult = 0;
		vResults.clear();
		mSubMatches.clear();
		if (bClearPattern && pattern != NULL)
			delete pattern; 


	/*SetLastError("");
	TV_VT(pvarRetValue) = VTYPE_BOOL;

	boost::wregex* pattern = NULL;
	bool bClearPattern = false;
	std::wstring str;
	GetStrParam(str, paParams, 0);
	try
	{
		if (paParams[1].wstrLen == 0)
		{
			if (isPattern)
				pattern = &rePattern;
			else
				return true;
		}
		else
		{
			bClearPattern = true;
			pattern = GetPattern(paParams, 1);
		}
		pvarRetValue->bVal = boost::regex_match(str, *pattern);

	}
	catch (const std::exception& e)
	{
		vResults.clear();
		mSubMatches.clear();
		iCurrentPosition = -1;
		uiSubMatchesCount = 0;
		m_PropCountOfItemsInSearchResult = 0;
		if (bClearPattern && pattern != NULL)
			delete pattern;
		SetLastError(e.what());
		if (bThrowExceptions)
			return false;
		else
			return true;
	}
	iCurrentPosition = 0;
	m_PropCountOfItemsInSearchResult = 0;
	vResults.clear();
	mSubMatches.clear();
	if (bClearPattern && pattern != NULL)
		delete pattern;*/
	return true;
}

bool CAddInNative::getSubMatch(tVariant * pvarRetValue, tVariant * paParams)
{
	SetLastError(u"");

	TV_VT(pvarRetValue) = VTYPE_PWSTR;
	int subMatchIndex = paParams[0].lVal;

	if (iCurrentPosition < 0) {
		SetLastError(u"Results were not selected");
		if (bThrowExceptions)
			return false;
		else
			return true;
	}

	if (!bHierarchicalResultIteration) {
		SetLastError(u"Hierarchical selection mode is not enabled");
		if (bThrowExceptions)
			return false;
		else
			return true;
	}

	if (iCurrentPosition >= mSubMatches.size()) {
		SetLastError(u"There are no any submatches in the selection");
		if (bThrowExceptions)
			return false;
		else
			return true;
	}

	auto &subMatch = mSubMatches.at(iCurrentPosition);

	if (subMatchIndex < 0 || subMatchIndex >= subMatch.size()) {
		SetLastError(u"Submatch index is out of range.");
		if (bThrowExceptions)
			return false;
		else
			return true;
	}


	std::basic_string<char16_t> &wsSubMatch = subMatch.at(subMatchIndex);

	/*#if defined( __linux__ ) || defined(__APPLE__) || defined(__ANDROID__)
		if (m_iMemory->AllocMemory((void**)&pvarRetValue->pwstrVal, (wsSubMatch.length() + 1) * sizeof(char32_t)))
		{
			convertUTF32ToUTF16(wsSubMatch.c_str(), wsSubMatch.length(), (char16_t *)pvarRetValue->pwstrVal);
			pvarRetValue->wstrLen = wsSubMatch.length();
		}
	#else
		if (m_iMemory->AllocMemory((void**)&pvarRetValue->pwstrVal, (wsSubMatch.length() + 1) * sizeof(char16_t)))
		{
			memcpy(pvarRetValue->pwstrVal, wsSubMatch.c_str(), (wsSubMatch.length() + 1) * sizeof(char16_t));
			pvarRetValue->wstrLen = wsSubMatch.length();
		}
	#endif*/
	if (m_iMemory->AllocMemory((void**)&pvarRetValue->pwstrVal, (wsSubMatch.length() + 1) * sizeof(char16_t)))
	{
		memcpy(pvarRetValue->pwstrVal, wsSubMatch.c_str(), (wsSubMatch.length() + 1) * sizeof(char16_t));
		pvarRetValue->wstrLen = wsSubMatch.length();
	}
	return true;
}

void CAddInNative::version(tVariant * pvarRetValue)
{
	TV_VT(pvarRetValue) = VTYPE_PWSTR;
	std::basic_string<char16_t> res = u"13";

	if (m_iMemory->AllocMemory((void**)&pvarRetValue->pwstrVal, (res.length() + 1) * sizeof(char16_t)))
	{
		memcpy(pvarRetValue->pwstrVal, res.c_str(), (res.length() + 1) * sizeof(char16_t));
		pvarRetValue->wstrLen = res.length();
	}
}

//---------------------------------------------------------------------------//

// Устанавливает значение переменной sErrorDescription класса CAddInNative.
// Если переменная sErrorDescription не пустая, то значение добавляется в начало строки,
// а уже существующий текст, присоединяется через двоеточие справа
// такой подход принят для реализации механизма вложенности описания ошибки
// на самом первом этапе, в текст ошибки добавляется её первопричина, а на более 
// высоких уровнях её можно дополнить описанием о выполняемых действиях
//
void CAddInNative::SetLastError(const char16_t* error) {

	if (error == 0) sErrorDescription.clear();
	if (sErrorDescription.length() != 0) sErrorDescription = std::basic_string<char16_t>(error) + u": " + sErrorDescription;
	else sErrorDescription = error;
}

void  CAddInNative::GetStrParam(std::wstring& str, tVariant* paParams, const long paramIndex) {
#if defined( __linux__ ) || defined(__APPLE__) || defined(__ANDROID__)
	str.resize(paParams[paramIndex].wstrLen);
	convertUTF16ToUTF32((char16_t *)paParams[paramIndex].pwstrVal, paParams[paramIndex].wstrLen, str);
#else
	str.assign(paParams[paramIndex].pwstrVal, paParams[paramIndex].wstrLen);
#endif
}

pcre2_code*  CAddInNative::GetPattern(tVariant* paParams, const long paramIndex) {

	int errornumber;
	PCRE2_SIZE erroroffset = NULL;
	pcre2_code* res;

	res = pcre2_compile((PCRE2_SPTR16)(paParams[paramIndex]).pwstrVal,               /* the pattern */
		paParams[paramIndex].wstrLen, /* indicates pattern is zero-terminated */
		PCRE2_UTF|(bIgnoreCase ? PCRE2_CASELESS : 0), /* options */
		&errornumber,          /* for error number */
		&erroroffset,          /* for error offset */
		NULL);

	if (res == NULL)
	{
		PCRE2_UCHAR buffer[256];
		pcre2_get_error_message_16(errornumber, buffer, sizeof(buffer));
		vResults.clear();
		mSubMatches.clear();
		isPattern = false;
		iCurrentPosition = -1;
		uiSubMatchesCount = 0;
		m_PropCountOfItemsInSearchResult = 0;
		SetLastError((const char16_t*)buffer);
	}
	return res;

}

