#if !defined( __linux__ ) && !defined(__APPLE__) && !defined(__ANDROID__)
#include "stdafx.h"
#endif

#include <stdio.h>
#include <wchar.h>
#include "AddInNative.h"
#include "OrderedSet.h"

//static const auto g_kClassNames = u"RegEx";
static const std::u16string sClassName(u"RegEx");
static const std::u16string sVersion(u"15.8");

static const OrderedSet<std::u16string> osMethods = { u"matches", u"ismatch", u"next", u"replace", u"count", u"submatchescount", u"getsubmatch", u"version", u"matchesjson", u"test"};
static const OrderedSet<std::u16string> osMethods_ru = { u"найтисовпадения", u"совпадает", u"следующий", u"заменить", u"количество", u"количествовложенныхгрупп", u"получитьподгруппу", u"версия", u"найтисовпаденияjson", u"test"};

static const OrderedSet<std::u16string> osProps = { u"currentvalue", u"ignorecase", u"errordescription", u"throwexceptions", u"pattern", u"global", u"firstindex", u"multiline", u"ucp"};
static const OrderedSet<std::u16string> osProps_ru = { u"текущеезначение", u"игнорироватьрегистр", u"описаниеошибки", u"вызыватьисключения", u"шаблон", u"всесовпадения", u"firstindex", u"многострочный", u"ucp" };

//static const std::map<std::u16string, long> mMethods = { { u"matches", 0}, {u"ismatch", 1}, {u"next", 2}, {u"replace", 3}, {u"count", 4}, {u"submatchescount", 5}, {u"getsubmatch", 6}, {u"version", 7}, {u"matchesjson", 8}, {u"test", 9} };
//static const std::vector<std::u16string> vMethods = { u"matches", u"ismatch", u"next", u"replace", u"count", u"submatchescount", u"getsubmatch", u"version", u"matchesjson", u"test" };
//static const std::map<std::u16string, long> mMethods_ru = { {u"найтисовпадения", 0}, {u"совпадает", 1}, {u"следующий", 2}, {u"заменить", 3}, {u"количество", 4}, {u"количествовложенныхгрупп", 5}, {u"получитьподгруппу", 6}, {u"версия", 7}, {u"найтисовпаденияjson", 8}, {u"test", 9} };
//static const std::vector<std::u16string> vMethods_ru = { u"найтисовпадения", u"совпадает", u"следующий", u"заменить", u"количество", u"количествовложенныхгрупп", u"получитьподгруппу", u"версия", u"найтисовпаденияjson", u"test" };

/*static const std::map<std::u16string, long> mProps = {{u"currentvalue", 0}, {u"ignorecase", 1}, {u"errordescription", 2}, {u"throwexceptions", 3}, {u"pattern", 4}, {u"global", 5}, {u"firstindex", 6}, {u"multiline", 7}, {u"ucp", 8}};
static const std::vector<std::u16string> vProps = { u"currentvalue", u"ignorecase", u"errordescription", u"throwexceptions", u"pattern", u"global", u"firstindex", u"multiline", u"ucp" };
static const std::map<std::u16string, long> mProps_ru = { {u"текущеезначение", 0}, {u"игнорироватьрегистр", 1}, {u"описаниеошибки", 2}, {u"вызыватьисключения", 3}, {u"шаблон", 4}, {u"всесовпадения", 5}, {u"firstindex", 6}, {u"многострочный", 7}, {u"ucp", 8} };
static const std::vector<std::u16string> vProps_ru = { u"текущеезначение", u"игнорироватьрегистр", u"описаниеошибки", u"вызыватьисключения", u"шаблон", u"всесовпадения", u"firstindex", u"многострочный", u"ucp" };
*/

AppCapabilities g_capabilities = eAppCapabilitiesInvalid;
/*
inline void fillMap(std::map<std::u16string, long>& map, const std::vector<std::u16string> & vector) {
	long index = 0;
	for (auto &item : vector)
	{
		auto lowCasedItem = item;
		//tolowerStr(lowCasedItem);
		map.insert({ lowCasedItem, index });
		index++;
	}
}*/

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
	return (WCHAR_T*)sClassName.c_str();
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
	bMultiline = true;
	bUCP = false;
	isPattern = false;
	bGlobal = false;
	bHierarchicalResultIteration = false;
	rePattern = NULL;

	sPattern.clear();
}
//---------------------------------------------------------------------------//
CAddInNative::~CAddInNative()
{
	if (rePattern != NULL)
		pcre2_code_free(rePattern);
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
	if (wsExtensionName == nullptr) {
		return false;
	}
	const size_t size = (sClassName.size() + 1) * sizeof(char16_t);

	if (!m_iMemory || !m_iMemory->AllocMemory(reinterpret_cast<void **>(wsExtensionName), size) || *wsExtensionName ==  nullptr) {
		return false;
	};

	memcpy(*wsExtensionName, sClassName.c_str(), size);

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

	auto it = osProps.getByKey(usPropName);
	if (it != osProps.end())
		return it->second;

	it = osProps_ru.getByKey(usPropName);
	if (it != osProps_ru.end())
		return it->second;

	return -1;
}

//---------------------------------------------------------------------------//
const WCHAR_T* CAddInNative::GetPropName(long lPropNum, long lPropAlias)
{
	if (lPropNum >= ePropLast)
		return NULL;

	const std::basic_string<char16_t> *usCurrentName;

	switch (lPropAlias)
	{
	case 0: // First language
		usCurrentName = &osProps.getKeyByIndex(lPropNum);
		break;
	case 1: // Second language
		usCurrentName = &osProps_ru.getKeyByIndex(lPropNum);
		break;
	default:
		return 0;
	}

	if (usCurrentName == nullptr || usCurrentName->length() == 0) {
		return nullptr;
	}

	WCHAR_T *result = nullptr;

	const size_t bytes = (usCurrentName->length() + 1) * sizeof(char16_t);

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
			wsCurrentValue = &(vResults[iCurrentPosition].value);

		if (wsCurrentValue == nullptr)
			return false;

		if (m_iMemory->AllocMemory((void**)&pvarPropVal->pwstrVal, (wsCurrentValue->length() + 1) * sizeof(wchar_t)))
		{
			memcpy(pvarPropVal->pwstrVal, wsCurrentValue->c_str(), (wsCurrentValue->length() + 1) * sizeof(wchar_t));
			pvarPropVal->wstrLen = wsCurrentValue->length();
			return true;
		}
		return false;
	};
	case ePropFirstIndex:
	{
		TV_VT(pvarPropVal) = VTYPE_I4;
	

		size_t firstIndex;
		if (vResults.size() == 0 || iCurrentPosition == -1 || iCurrentPosition >= vResults.size())
		{
			firstIndex = 0;
		}
		else
			firstIndex = vResults[iCurrentPosition].firstIndex;

		pvarPropVal->lVal = firstIndex;

		return true;
	};
	case ePropErrorDescription:
	{
		if (m_iMemory->AllocMemory((void**)&pvarPropVal->pwstrVal, (sErrorDescription.length() + 1) * sizeof(char16_t)))
		{
			memcpy(pvarPropVal->pwstrVal, sErrorDescription.c_str(), (sErrorDescription.length() + 1) * sizeof(char16_t));
			TV_VT(pvarPropVal) = VTYPE_PWSTR;
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
		if (m_iMemory->AllocMemory((void**)&pvarPropVal->pwstrVal, (sPattern.length() + 1) * sizeof(char16_t)))
		{
			memcpy(pvarPropVal->pwstrVal, sPattern.c_str(), sPattern.length() * sizeof(char16_t));
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
	case ePropMultiline:
	{
		TV_VT(pvarPropVal) = VTYPE_BOOL;
		pvarPropVal->bVal = bMultiline;
		return true;
	}
	case ePropUCP:
	{
		TV_VT(pvarPropVal) = VTYPE_BOOL;
		pvarPropVal->bVal = bUCP;
		return true;
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
		if (rePattern != NULL)
			pcre2_code_free(rePattern);

		sPattern.clear();
		rePattern = GetPattern(varPropVal);
		if (rePattern == NULL)
		{
			isPattern = false;
			if (bThrowExceptions)
				return false;
			else
				return true;
		}

		sPattern.assign(reinterpret_cast<char16_t*>(varPropVal->pwstrVal), varPropVal->wstrLen);
		isPattern = true;
		return true;
	}
	case ePropGlobal:
	{
		bGlobal = TV_BOOL(varPropVal);
		return true;
	}
	case ePropMultiline:
	{
		bMultiline = TV_BOOL(varPropVal);
		return true;
	}
	case ePropUCP:
	{
		bUCP = TV_BOOL(varPropVal);
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
	case ePropFirstIndex:
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
	case ePropMultiline:
		return true;
	case ePropUCP:
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
	case ePropMultiline:
		return true;
	case ePropUCP:
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

	auto it = osMethods.getByKey(usMethodName);
	if (it != osMethods.end())
		return it->second;

	it = osMethods_ru.getByKey(usMethodName);
	if (it != osMethods_ru.end())
		return it->second;

	return -1;
}
//---------------------------------------------------------------------------//
const WCHAR_T* CAddInNative::GetMethodName(const long lMethodNum, const long lMethodAlias)
{

	if (lMethodNum >= eMethLast)
		return nullptr;

	const std::basic_string<char16_t> *usCurrentName;

	switch (lMethodAlias)
	{
	case 0: // First language
		usCurrentName = &osMethods.getKeyByIndex(lMethodNum);
		break;
	case 1: // Second language
		usCurrentName = &osMethods_ru.getKeyByIndex(lMethodNum);
		break;
	default:
		return nullptr;
	}

	if (usCurrentName == nullptr)
		return nullptr;

	WCHAR_T *result = nullptr;

	const size_t bytes = (usCurrentName->length() + 1) * sizeof(char16_t);

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
	case eMethMatchesJSON:
		return 3;
	case eMethIsMatch:
		return 2;
	case eMethTest:
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
	case eMethMatchesJSON:
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
	case eMethTest:
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
	case eMethMatchesJSON:
		return true;
	case eMethReplace:
		return true;
	case eMethIsMatch:
	case eMethTest:
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
	case eMethTest:
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
	case eMethMatchesJSON:
		return searchJSON(pvarRetValue, paParams);
	case eMethIsMatch:
	case eMethTest:
	{
		return match(pvarRetValue, paParams);
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
	m_PropCountOfItemsInSearchResult = 0;

	if (paParams[0].wstrLen == 0)
	{
		return true;
	}

	if (paParams[0].vt != VTYPE_PWSTR)
	{
		SetLastError(u"Only values of string type can be passed for search text");
		if (bThrowExceptions)
			return false;
		else
			return true;
	}

	pcre2_code* pattern = NULL;
	bool bClearPattern = false;

	bHierarchicalResultIteration = paParams[2].bVal;

	if (paParams[1].wstrLen == 0)
	{
		if (isPattern)
			pattern = rePattern;
	}
	else
	{
		bClearPattern = true;
		pattern = GetPattern(&paParams[1]);
	}

	if (pattern == NULL)
	{
		SetLastError(u"The regexp template is not specified.");
		if (bThrowExceptions)
			return false;
		else
			return true;
	}

	int crlf_is_newline;
	uint32_t newline;

	(void)pcre2_pattern_info(pattern, PCRE2_INFO_NEWLINE, &newline);
	crlf_is_newline = newline == PCRE2_NEWLINE_ANY ||
		newline == PCRE2_NEWLINE_CRLF ||
		newline == PCRE2_NEWLINE_ANYCRLF;

	size_t rootIndex = 0;
	int rc;

	pcre2_match_data* match_data = pcre2_match_data_create_from_pattern(pattern, NULL);

	PCRE2_SIZE *ovector = NULL;

	const size_t subject_length = paParams[0].wstrLen;

	while (true) {

		uint32_t options = 0;

		PCRE2_SIZE start_offset = ovector == NULL ? 0 : ovector[1];   /* Start at end of previous match */

		if (ovector != NULL && (ovector[0] == ovector[1]))
		{
			if (ovector[0] == subject_length) break;
			options = PCRE2_NOTEMPTY_ATSTART | PCRE2_ANCHORED;
		}
		else if (ovector != NULL)
		{
			PCRE2_SIZE startchar = pcre2_get_startchar(match_data);
			if (start_offset <= startchar)
			{
				if (startchar >= subject_length) break;   /* Reached end of subject.   */
				start_offset = startchar + 1;             /* Advance by one character. */
				/* If UTF-8, it may be more  */
				/*   than one code unit.     */
					/*for (; start_offset < subject_length; start_offset++)
						if ((paParams[0].pwstrVal[start_offset] & 0xc0) != 0x80)
							break;				*/
			}
		}

		if (start_offset > subject_length)
			break;
		rc = pcre2_match(
			pattern,                   /* the compiled pattern */
			(PCRE2_SPTR16)paParams[0].pwstrVal,              /* the subject string */
			paParams[0].wstrLen,       /* the length of the subject */
			start_offset,                    /* start at offset 0 in the subject */
			options,                    /* default options */
			match_data,           /* block for storing the result */
			NULL);

		if (rc == PCRE2_ERROR_NOMATCH)
		{
			if (options == 0 || ovector == NULL) break;                    /* All matches found */
			ovector[1] = start_offset + 1;              /* Advance one code unit */
			if (crlf_is_newline &&                      /* If CRLF is a newline & */
				start_offset < subject_length - 1 &&    /* we are at CRLF, */
				paParams[0].pwstrVal[start_offset] == '\r' &&
				paParams[0].pwstrVal[start_offset + 1] == '\n')
				ovector[1] += 1;                          /* Advance by one more. */
				/*while (ovector[1] < subject_length)
				{
					if ((paParams[0].pwstrVal[ovector[1]] & 0xc0) != 0x80)
						break;
					ovector[1] += 1;
				}*/
			continue;    /* Go round the loop again */
		}

		if (rc < 0)
		{
			PCRE2_UCHAR buffer[256];
			pcre2_get_error_message_16(rc, buffer, sizeof(buffer));
			SetLastError((const char16_t*)buffer);
			pcre2_match_data_free(match_data);   /* Release memory used for the match */
			if (bThrowExceptions)
				return false;
			else
				return true;
		}
		else if (rc == 0) {
			SetLastError(u"ovector was not big enough for all the captured substrings.");
			break;
		}

		ovector = pcre2_get_ovector_pointer(match_data);

		if (ovector == nullptr)
			break;

		if (ovector[0] > ovector[1])
		{
			SetLastError(u"\\K was used in an assertion to set the match start after its end.");
			pcre2_match_data_free(match_data);
			if (bThrowExceptions)
				return false;
			else
				return true;
		}
		uiSubMatchesCount = pcre2_get_ovector_count(match_data) - 1;

		std::vector<std::basic_string<char16_t>> vSubMatches;

		for (int i = 0; i <= uiSubMatchesCount; i++)
		{
			ResultStruct resultStruct;

			resultStruct.value.assign((char16_t*)(paParams[0].pwstrVal + ovector[2 * i]), ovector[2 * i + 1] - ovector[2 * i]);

			if (i == 0 && bHierarchicalResultIteration) {
				
				resultStruct.firstIndex = ovector[2 * i];
				vResults.push_back(resultStruct);
				rootIndex = vResults.size() - 1;
			}
			else if (bHierarchicalResultIteration)
				vSubMatches.push_back(resultStruct.value);
			else if (!bHierarchicalResultIteration) {
				resultStruct.firstIndex = ovector[2 * i];
				vResults.push_back(resultStruct);
			}
		}
		if (bHierarchicalResultIteration)
			mSubMatches.insert({ rootIndex, vSubMatches});
		if (!bGlobal) 
			break;
	}

	pcre2_match_data_free(match_data);
	iCurrentPosition = -1;
	m_PropCountOfItemsInSearchResult = vResults.size();
	if (bClearPattern && pattern != NULL) {
		pcre2_code_free(pattern);
	}
	return true;
}


bool CAddInNative::searchJSON(tVariant* pvarRetValue, tVariant * paParams)
{
	SetLastError(u"");
	vResults.clear();
	mSubMatches.clear();
	iCurrentPosition = -1;
	uiSubMatchesCount = 0;
	m_PropCountOfItemsInSearchResult = 0;

	if (paParams[0].vt != VTYPE_PWSTR)
	{
		SetLastError(u"Only values of string type can be passed for search text");
		if (bThrowExceptions)
			return false;
		else
			return true;
	}

	std::basic_string<char16_t> res;

	TV_VT(pvarRetValue) = VTYPE_PWSTR;

	if (paParams[0].wstrLen == 0)
	{
		res = u"[]";
		if (m_iMemory->AllocMemory((void**)&pvarRetValue->pwstrVal, (res.length() + 1) * sizeof(char16_t)))
		{
			memcpy(pvarRetValue->pwstrVal, res.c_str(), (res.length() + 1) * sizeof(char16_t));
			pvarRetValue->wstrLen = res.length();
		}
		return true;
	}

	res.reserve(paParams[0].wstrLen); // мы не знаем сколько памяти потребуется, поэтому выделим тот объем, который занимает исходный текст

	pcre2_code* pattern = NULL;
	bool bClearPattern = false;

	bool bIntervals = paParams[2].bVal;

	if (paParams[1].wstrLen == 0)
	{
		if (isPattern)
			pattern = rePattern;
	}
	else
	{
		bClearPattern = true;
		pattern = GetPattern(&paParams[1]);
	}

	if (pattern == NULL)
	{
		SetLastError(u"The regexp template is not specified.");
		if (bThrowExceptions)
			return false;
		else
			return true;
	}

	int crlf_is_newline;
	uint32_t newline;

	(void)pcre2_pattern_info(pattern, PCRE2_INFO_NEWLINE, &newline);
	crlf_is_newline = newline == PCRE2_NEWLINE_ANY ||
		newline == PCRE2_NEWLINE_CRLF ||
		newline == PCRE2_NEWLINE_ANYCRLF;

	size_t rootIndex = 0;
	int rc;

	pcre2_match_data* match_data = pcre2_match_data_create_from_pattern(pattern, NULL);

	res +=  u'['; // array
	
	PCRE2_SIZE *ovector = NULL;

	const size_t subject_length = paParams[0].wstrLen;

	while (true) {

		uint32_t options = 0;

		PCRE2_SIZE start_offset = ovector == NULL ? 0 : ovector[1];   /* Start at end of previous match */

		if (ovector != NULL && (ovector[0] == ovector[1]))
		{
			if (ovector[0] == subject_length) break;
			options = PCRE2_NOTEMPTY_ATSTART | PCRE2_ANCHORED;
		}
		else if (ovector != NULL)
		{
			const PCRE2_SIZE startchar = pcre2_get_startchar(match_data);
			if (start_offset <= startchar)
			{
				if (startchar >= subject_length) break;   /* Reached end of subject.   */
				start_offset = startchar + 1;             /* Advance by one character. */
				/* If UTF-8, it may be more  */
				/*   than one code unit.     */
					/*for (; start_offset < subject_length; start_offset++)
						if ((paParams[0].pwstrVal[start_offset] & 0xc0) != 0x80) 
							break;				*/
			}
		}

		if (start_offset > subject_length)
			break;
		rc = pcre2_match(
			pattern,                   /* the compiled pattern */
			(PCRE2_SPTR16)paParams[0].pwstrVal,              /* the subject string */
			subject_length,       /* the length of the subject */
			start_offset,                    /* start at offset 0 in the subject */
			options,                    /* default options */
			match_data,           /* block for storing the result */
			NULL);

		if (rc == PCRE2_ERROR_NOMATCH)
		{
			if (options == 0) break;                    /* All matches found */
			ovector[1] = start_offset + 1;              /* Advance one code unit */
			if (crlf_is_newline &&                      /* If CRLF is a newline & */
				start_offset < subject_length - 1 &&    /* we are at CRLF, */
				paParams[0].pwstrVal[start_offset] == '\r' &&
				paParams[0].pwstrVal[start_offset + 1] == '\n')
				ovector[1] += 1;                          /* Advance by one more. */
				/*while (ovector[1] < subject_length)       
				{
					if ((paParams[0].pwstrVal[ovector[1]] & 0xc0) != 0x80) 
						break;
					ovector[1] += 1;
				}*/
			continue;    /* Go round the loop again */
		}

		if (rc < 0)
		{
			PCRE2_UCHAR buffer[256];
			pcre2_get_error_message_16(rc, buffer, sizeof(buffer));
			SetLastError((const char16_t*)buffer);
			pcre2_match_data_free(match_data);   /* Release memory used for the match */
			if (bThrowExceptions)
				return false;
			else
				return true;
		}
		else if (rc == 0) {
			SetLastError(u"ovector was not big enough for all the captured substrings.");
			break;
		}

		ovector = pcre2_get_ovector_pointer(match_data);

		if (ovector == nullptr)
			break;
		if (ovector[0] > ovector[1])
		{
			SetLastError(u"\\K was used in an assertion to set the match start after its end.");
			pcre2_match_data_free(match_data);
			if (bThrowExceptions)
				return false;
			else
				return true;
		}
		uiSubMatchesCount = pcre2_get_ovector_count(match_data) - 1;

		if (bIntervals)
		{
			res += u'[';
			for (int i = 0; i <= uiSubMatchesCount; i++)
			{
				if (ovector[2 * i] == ovector[2 * i + 1]) {
					res.append(u"0,0,", (sizeof(u"0,0,") - 2) / sizeof(char16_t));
					continue;
				}
				char16_t snum[20];
				int len = itoa_u16(ovector[2 * i], snum, 20, 10);
				res.append(snum, len);
				res += u',';
				len = itoa_u16(ovector[2 * i + 1], snum, 20, 10);
				res.append(snum, len);
				res += u',';
			}
			res += u']';
		}
		else
		{
			res.append(u"{\"Value\":\"", (sizeof(u"{\"Value\":\"") - 2) / sizeof(char16_t));

			if (ovector[0] != ovector[1])
				append_escaped_json(res, (char16_t*)paParams[0].pwstrVal, ovector[0], ovector[1] - 1);
			res.append(u"\",\"FirstIndex\":", (sizeof(u"\",\"FirstIndex\":") - 2) / sizeof(char16_t));

			char16_t snum[20];

			int len = itoa_u16(ovector[0], snum, 20, 10);
			res.append(snum, len);

			res.append(u",\"Length\":", (sizeof(u",\"Length\":") - 2) / sizeof(char16_t));

			len = itoa_u16(ovector[1] - ovector[0], snum, 10, 10);
			res.append(snum, len);
			res.append(u",\"SubMatches\":[", (sizeof(u",\"SubMatches\":[") - 2) / sizeof(char16_t));

			for (int i = 1; i <= uiSubMatchesCount; i++)
			{
				if (ovector[2 * i] == ovector[2 * i + 1]) {
					res.append(u"null,", (sizeof(u"null,") - 2) / sizeof(char16_t));
					continue;
				}
				res += u'\"';
				append_escaped_json(res, (char16_t*)paParams[0].pwstrVal, ovector[2 * i], ovector[2 * i + 1] - 1);
				res.append(u"\",", (sizeof(u"\",") - 2) / sizeof(char16_t));
			}
			res.append(u"]},", (sizeof(u"]},") - 2) / sizeof(char16_t));
		}

		if (!bGlobal)
				break;
	}

	res += u']'; // array

	if (m_iMemory->AllocMemory((void**)&pvarRetValue->pwstrVal, (res.length() + 1) * sizeof(char16_t)))
	{
		memcpy(pvarRetValue->pwstrVal, res.c_str(), (res.length() + 1) * sizeof(char16_t));
		pvarRetValue->wstrLen = res.length();
	}
	
	pcre2_match_data_free(match_data);
	if (bClearPattern && pattern != NULL) {
		pcre2_code_free(pattern);
	}
	return true;
}

bool CAddInNative::replace(tVariant * pvarRetValue, tVariant * paParams)
{
	SetLastError(u"");
	TV_VT(pvarRetValue) = VTYPE_PWSTR;
	pvarRetValue->wstrLen = 0;
	
	if (paParams[0].vt != VTYPE_PWSTR)
	{
		SetLastError(u"Only values of string type can be passed for search text");
		if (bThrowExceptions)
			return false;
		else
			return true;
	}


	if (paParams[0].wstrLen == 0)
	{
		pvarRetValue->wstrLen = 0;
		return true;
	}

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
		pattern = GetPattern(&paParams[1]);
	}

	if (pattern == NULL)
	{
		SetLastError(u"The regexp template is not specified.");
		if (bThrowExceptions)
			return false;
		else
			return true;
	}

	PCRE2_SIZE start_offset = 0;
	size_t rootIndex = 0;
	int rc;

	pcre2_match_data* match_data = pcre2_match_data_create_from_pattern(pattern, NULL);

	PCRE2_UCHAR outputbuffer;
	PCRE2_SIZE outlength = 0;
	bool isError = false;

	rc = pcre2_substitute(pattern, 
		(PCRE2_SPTR16)paParams[0].pwstrVal, 
		paParams[0].wstrLen,
		0,
		bGlobal?PCRE2_SUBSTITUTE_GLOBAL | PCRE2_SUBSTITUTE_UNSET_EMPTY | PCRE2_SUBSTITUTE_OVERFLOW_LENGTH | PCRE2_SUBSTITUTE_EXTENDED : PCRE2_SUBSTITUTE_OVERFLOW_LENGTH | PCRE2_SUBSTITUTE_UNSET_EMPTY | PCRE2_SUBSTITUTE_EXTENDED,
		match_data,
		NULL,
		(PCRE2_SPTR16)paParams[2].pwstrVal, 
		paParams[2].wstrLen, 0, &outlength);

	if (rc == PCRE2_ERROR_NOMEMORY) // it is not actually an error, it is an anticipated behavior
	{
		if (m_iMemory->AllocMemory((void**)&pvarRetValue->pwstrVal, (outlength + 1) * sizeof(WCHAR)))
		{
			rc = pcre2_substitute(pattern,
				(PCRE2_SPTR16)paParams[0].pwstrVal,
				paParams[0].wstrLen,
				0,
				bGlobal ? PCRE2_SUBSTITUTE_GLOBAL | PCRE2_SUBSTITUTE_UNSET_EMPTY | PCRE2_SUBSTITUTE_OVERFLOW_LENGTH | PCRE2_SUBSTITUTE_EXTENDED : PCRE2_SUBSTITUTE_OVERFLOW_LENGTH | PCRE2_SUBSTITUTE_UNSET_EMPTY | PCRE2_SUBSTITUTE_EXTENDED,
				match_data,
				NULL,
				(PCRE2_SPTR16)paParams[2].pwstrVal,
				paParams[2].wstrLen, (PCRE2_UCHAR *)pvarRetValue->pwstrVal, &outlength);
			if (rc >= 0)
				pvarRetValue->wstrLen = outlength;
			else {
				PCRE2_UCHAR buffer[256];
				pcre2_get_error_message_16(rc, buffer, sizeof(buffer));
				SetLastError((const char16_t*)buffer);
				pvarRetValue->wstrLen = 0;
				isError = true;
			}

		}
	}
	else
	{
		if (rc < 0) {
			PCRE2_UCHAR buffer[256];
			pcre2_get_error_message_16(rc, buffer, sizeof(buffer));
			SetLastError((const char16_t*)buffer);
			isError = true;
		}
		pvarRetValue->wstrLen = 0;
	}

	pcre2_match_data_free(match_data);

	/*iCurrentPosition = 0;
	m_PropCountOfItemsInSearchResult = 0;
	vResults.clear();
	mSubMatches.clear();*/
	if (bClearPattern && pattern != NULL) {
		pcre2_code_free(pattern);
	}

	if (isError && bThrowExceptions)
		return false;

	return true;
}

bool CAddInNative::match(tVariant * pvarRetValue, tVariant * paParams)
{
	SetLastError(u"");
	TV_VT(pvarRetValue) = VTYPE_BOOL;
	pvarRetValue->bVal = false;

	if (paParams[0].vt != VTYPE_PWSTR)
	{
		SetLastError(u"Only values of string type can be passed for search text");
		if (bThrowExceptions)
		{
			return false;
		}
		else
			return true;
	}


	if (paParams[0].wstrLen == 0)
	{
		return true;
	}

	pcre2_code* pattern = NULL;
	bool bClearPattern = false;

	if (paParams[1].wstrLen == 0)
	{
		if (isPattern)
			pattern = rePattern;
	}
	else
	{
		bClearPattern = true;
		pattern = GetPattern(&paParams[1]);
	}

	if (pattern == NULL)
	{
		SetLastError(u"The regexp template is not specified.");
		if (bThrowExceptions)
			return false;
		else
			return true;
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

		/*iCurrentPosition = 0;
		m_PropCountOfItemsInSearchResult = 0;
		vResults.clear();
		mSubMatches.clear();*/
		if (bClearPattern && pattern != NULL) {
			pcre2_code_free(pattern);
		}
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
	
	if (m_iMemory->AllocMemory((void**)&pvarRetValue->pwstrVal, (sVersion.length() + 1) * sizeof(char16_t)))
	{
		memcpy(pvarRetValue->pwstrVal, sVersion.c_str(), (sVersion.length() + 1) * sizeof(char16_t));
		pvarRetValue->wstrLen = sVersion.length();
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

	if (error == 0 || error[0] == 0) {
		sErrorDescription.clear();
		return;
	}
	if (sErrorDescription.length() != 0) sErrorDescription = std::basic_string<char16_t>(error) + u": " + sErrorDescription;
	else sErrorDescription = error;
}

/*void  CAddInNative::GetStrParam(std::wstring& str, tVariant* paParams, const long paramIndex) {
#if defined( __linux__ ) || defined(__APPLE__) || defined(__ANDROID__)
	str.resize(paParams[paramIndex].wstrLen);
	convertUTF16ToUTF32((char16_t *)paParams[paramIndex].pwstrVal, paParams[paramIndex].wstrLen, str);
#else
	str.assign(paParams[paramIndex].pwstrVal, paParams[paramIndex].wstrLen);
#endif
}
*/

pcre2_code*  CAddInNative::GetPattern(const tVariant *tvPattern) {

	SetLastError(u"");

	const char16_t* patternStr = (char16_t*)tvPattern->pwstrVal;
	const long len = tvPattern->wstrLen;

	if (tvPattern->vt != VTYPE_PWSTR)
	{
		SetLastError(u"Only values of string type can be passed as a template");
		return NULL;
	}

	int errornumber;
	PCRE2_SIZE erroroffset = NULL;
	pcre2_code* res;

	res = pcre2_compile((PCRE2_SPTR16)patternStr,               // the pattern 
		len, // indicates pattern is zero-terminated 
		PCRE2_UTF|(bIgnoreCase ? PCRE2_CASELESS : 0)|(bMultiline ? PCRE2_MULTILINE : 0)|(bUCP? PCRE2_UCP : 0), // options 
		&errornumber,          // for error number 
		&erroroffset,          // for error offset 
		NULL);

	if (res == NULL)
	{
		PCRE2_UCHAR buffer[256];
		pcre2_get_error_message_16(errornumber, buffer, sizeof(buffer));
		/*vResults.clear();
		mSubMatches.clear();
		iCurrentPosition = -1;
		uiSubMatchesCount = 0;
		m_PropCountOfItemsInSearchResult = 0;*/
		SetLastError((const char16_t*)buffer);
	}

	return res;
}

