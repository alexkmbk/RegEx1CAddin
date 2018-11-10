
#include "stdafx.h"


#if defined( __linux__ ) || defined(__APPLE__)
#include <unistd.h>
#include <stdlib.h>
#include <signal.h>
#include <time.h>
#include <errno.h>
#include <iconv.h>
#include <sys/time.h>
#include <regex>
#endif

#include <stdio.h>
#include <wchar.h>
#include "AddInNative.h"
#include <string>


#define TIME_LEN 65

#define BASE_ERRNO     7

#ifdef WIN32
#pragma setlocale("ru-RU" )
#endif

static const wchar_t *g_PropNames[] = {
    L"CountOfItemsInSearchResult",
	L"CurrentValue"
};
static const wchar_t *g_MethodNames[] = {
    L"Search", 
    L"Match", 
    L"Next",
	L"Replace"
};

static const wchar_t *g_PropNamesRu[] = {
    L"КоличествоЭлементовВРезультатеПоиска",
	L"ТекущееЗначение"
};
static const wchar_t *g_MethodNamesRu[] = {
    L"Поиск", 
    L"Совпадает", 
    L"Следующий",
	L"Заменить"
};

static const wchar_t g_kClassNames[] = L"CAddInNative"; //"|OtherClass1|OtherClass2";
static IAddInDefBase *pAsyncEvent = NULL;

uint32_t convToShortWchar(WCHAR_T** Dest, const wchar_t* Source, uint32_t len = 0);
uint32_t convFromShortWchar(wchar_t** Dest, const WCHAR_T* Source, uint32_t len = 0);
uint32_t getLenShortWcharStr(const WCHAR_T* Source);
static AppCapabilities g_capabilities = eAppCapabilitiesInvalid;
static WcharWrapper s_names(g_kClassNames);
//---------------------------------------------------------------------------//
long GetClassObject(const WCHAR_T* wsName, IComponentBase** pInterface)
{
    if(!*pInterface)
    {
        *pInterface= new CAddInNative;
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
    if(!*pIntf)
        return -1;

    delete *pIntf;
    *pIntf = 0;
    return 0;
}
//---------------------------------------------------------------------------//
const WCHAR_T* GetClassNames()
{
    return s_names;
}
//---------------------------------------------------------------------------//

// CAddInNative
//---------------------------------------------------------------------------//
CAddInNative::CAddInNative()
{
    m_iMemory = 0;
    m_iConnect = 0;

	iCurrentPosition = 0;
	m_PropCountOfItemsInSearchResult = 0;

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
    // This component supports 2.0 version
    return 2000; 
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
    const wchar_t *wsExtension = L"RegEx";
    int iActualSize = ::wcslen(wsExtension) + 1;
    WCHAR_T* dest = 0;

    if (m_iMemory)
    {
        if(m_iMemory->AllocMemory((void**)wsExtensionName, iActualSize * sizeof(WCHAR_T)))
            ::convToShortWchar(wsExtensionName, wsExtension, iActualSize);
        return true;
    }

    return false; 
}
//---------------------------------------------------------------------------//
long CAddInNative::GetNProps()
{ 
    // You may delete next lines and add your own implementation code here
    return ePropLast;
}
//---------------------------------------------------------------------------//
long CAddInNative::FindProp(const WCHAR_T* wsPropName)
{ 
    long plPropNum = -1;
    wchar_t* propName = 0;

    ::convFromShortWchar(&propName, wsPropName);
    plPropNum = findName(g_PropNames, propName, ePropLast);

    if (plPropNum == -1)
        plPropNum = findName(g_PropNamesRu, propName, ePropLast);

    delete[] propName;

    return plPropNum;
}
//---------------------------------------------------------------------------//
const WCHAR_T* CAddInNative::GetPropName(long lPropNum, long lPropAlias)
{ 
    if (lPropNum >= ePropLast)
        return NULL;

    wchar_t *wsCurrentName = NULL;
    WCHAR_T *wsPropName = NULL;
    int iActualSize = 0;

    switch(lPropAlias)
    {
    case 0: // First language
        wsCurrentName = (wchar_t*)g_PropNames[lPropNum];
        break;
    case 1: // Second language
        wsCurrentName = (wchar_t*)g_PropNamesRu[lPropNum];
        break;
    default:
        return 0;
    }
    
    iActualSize = wcslen(wsCurrentName) + 1;

    if (m_iMemory && wsCurrentName)
    {
        if (m_iMemory->AllocMemory((void**)&wsPropName, iActualSize * sizeof(WCHAR_T)))
            ::convToShortWchar(&wsPropName, wsCurrentName, iActualSize);
    }

    return wsPropName;
}
//---------------------------------------------------------------------------//
bool CAddInNative::GetPropVal(const long lPropNum, tVariant* pvarPropVal)
{ 
    switch(lPropNum)
    {
    case ePropCountOfItemsInSearchResult:
        TV_VT(pvarPropVal) = VTYPE_INT;
        TV_INT(pvarPropVal) = m_PropCountOfItemsInSearchResult;
        break;
    case ePropCurrentValue:
		TV_VT(pvarPropVal) = VTYPE_PWSTR;

#ifdef __linux__
		if (m_iMemory->AllocMemory((void**)&pvarPropVal->pwstrVal, (wsCurrentValue.length() + 1) * sizeof(WCHAR_T)))
		{
			WCHAR_T* str_WCHAR_T = 0;
			convToShortWchar(&str_WCHAR_T, wsCurrentValue.c_str(), wsCurrentValue.length() + 1);
			memcpy(pvarPropVal->pwstrVal, str_WCHAR_T, (wsCurrentValue.length() + 1) * sizeof(WCHAR_T));
			pvarPropVal->wstrLen = wsCurrentValue.length();
			return true;
		}
#else
		
		if (m_iMemory->AllocMemory((void**)&pvarPropVal->pwstrVal, (wsCurrentValue.length() + 1) * sizeof(wchar_t)))
		{
			memcpy(pvarPropVal->pwstrVal, wsCurrentValue.c_str(), (wsCurrentValue.length() + 1) * sizeof(wchar_t));
			pvarPropVal->wstrLen = wsCurrentValue.length();
			return true;
		}
#endif
		return false;
        break;
    default:
        return false;
    }

    return true;
}
//---------------------------------------------------------------------------//
bool CAddInNative::SetPropVal(const long lPropNum, tVariant *varPropVal)
{ 
   /* switch(lPropNum)
    { 
    case ePropIsEnabled:
        if (TV_VT(varPropVal) != VTYPE_BOOL)
            return false;
        m_boolEnabled = TV_BOOL(varPropVal);
        break;
    case ePropIsTimerPresent:
    default:
        return false;
    }*/

    return false;
}
//---------------------------------------------------------------------------//
bool CAddInNative::IsPropReadable(const long lPropNum)
{ 
    switch(lPropNum)
    { 
    case ePropCountOfItemsInSearchResult:
    case ePropCurrentValue:
        return true;
    default:
        return false;
    }

    return false;
}
//---------------------------------------------------------------------------//
bool CAddInNative::IsPropWritable(const long lPropNum)
{
  /*  switch(lPropNum)
    { 
    case ePropIsEnabled:
        return true;
    case ePropIsTimerPresent:
        return false;
    default:
        return false;
    }*/

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
    long plMethodNum = -1;
    wchar_t* name = 0;

    ::convFromShortWchar(&name, wsMethodName);

    plMethodNum = findName(g_MethodNames, name, eMethLast);

    if (plMethodNum == -1)
        plMethodNum = findName(g_MethodNamesRu, name, eMethLast);

    delete[] name;

    return plMethodNum;
}
//---------------------------------------------------------------------------//
const WCHAR_T* CAddInNative::GetMethodName(const long lMethodNum, const long lMethodAlias)
{ 
    if (lMethodNum >= eMethLast)
        return NULL;

    wchar_t *wsCurrentName = NULL;
    WCHAR_T *wsMethodName = NULL;
    int iActualSize = 0;

    switch(lMethodAlias)
    {
    case 0: // First language
        wsCurrentName = (wchar_t*)g_MethodNames[lMethodNum];
        break;
    case 1: // Second language
        wsCurrentName = (wchar_t*)g_MethodNamesRu[lMethodNum];
        break;
    default: 
        return 0;
    }

    iActualSize = wcslen(wsCurrentName) + 1;

    if (m_iMemory && wsCurrentName)
    {
        if(m_iMemory->AllocMemory((void**)&wsMethodName, iActualSize * sizeof(WCHAR_T)))
            ::convToShortWchar(&wsMethodName, wsCurrentName, iActualSize);
    }

    return wsMethodName;
}
//---------------------------------------------------------------------------//
long CAddInNative::GetNParams(const long lMethodNum)
{ 
    switch(lMethodNum)
    { 
    case eMethSearch:
        return 2;
    case eMethMatch:
        return 2;
	case eMethReplace:
		return 3;
    case eMethNext:
        return 0;
    default:
        return 0;
    }
    
    return 0;
}
//---------------------------------------------------------------------------//
bool CAddInNative::GetParamDefValue(const long lMethodNum, const long lParamNum,
                        tVariant *pvarParamDefValue)
{ 
    TV_VT(pvarParamDefValue)= VTYPE_EMPTY;

    switch(lMethodNum)
    { 
    case eMethSearch:
    case eMethMatch:
    case eMethNext:
	case eMethReplace:
        // There are no parameter values by default 
        break;
    default:
        return false;
    }

    return false;
} 
//---------------------------------------------------------------------------//
bool CAddInNative::HasRetVal(const long lMethodNum)
{ 
    switch(lMethodNum)
    { 
    case eMethNext:
		return true;
	case eMethReplace:
		return true;
    case eMethMatch:
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
	case eMethSearch:
		search(paParams);
		break;
	case eMethMatch:
		break;
	case eMethReplace:
		break;
	case eMethNext:
	{
		if (iCurrentPosition >= m_PropCountOfItemsInSearchResult) {
			return true;
		}
		else {
			iCurrentPosition++;
			wsCurrentValue = wsmMatch[iCurrentPosition];
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
	case eMethSearch:
		break;
	case eMethMatch:
	{
		match(pvarRetValue, paParams);
		return true;
	}
	break;
	case eMethReplace:
		return replace(pvarRetValue, paParams);
		break;
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
			wsCurrentValue = vResults[iCurrentPosition];
			return true;
		}
	}
	break;
	default:
		return false;
	}
}

//---------------------------------------------------------------------------//
void CAddInNative::SetLocale(const WCHAR_T* loc)
{
#if !defined( __linux__ ) && !defined(__APPLE__)
    _wsetlocale(LC_ALL, loc);
#else
    //We convert in char* char_locale
    //also we establish locale
    //setlocale(LC_ALL, char_locale);
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
//---------------------------------------------------------------------------//
void CAddInNative::addError(uint32_t wcode, const wchar_t* source, 
                        const wchar_t* descriptor, long code)
{
    if (m_iConnect)
    {
        WCHAR_T *err = 0;
        WCHAR_T *descr = 0;
        
        ::convToShortWchar(&err, source);
        ::convToShortWchar(&descr, descriptor);

        m_iConnect->AddError(wcode, err, descr, code);
        delete[] err;
        delete[] descr;
    }
}
void CAddInNative::search(tVariant * paParams)
{
#ifdef __linux__
	// Сконвертируем в строку с wchar_t символами
	wchar_t* str_wchar_t1 = 0;
	convFromShortWchar(&str_wchar_t1, paParams[0].pwstrVal, paParams[0].wstrLen + 1);

	std::wstring str(str_wchar_t1);
	
	wchar_t* str_wchar_t2 = 0;
	convFromShortWchar(&str_wchar_t2, paParams[1].pwstrVal, paParams[1].wstrLen + 1);
	std::wregex r(str_wchar_t2);
	std::regex_search(str, wsmMatch, r);
	
	for (auto res : wsmMatch) {
		vResults.push_back(res);
	}
	
	delete[] str_wchar_t1;
	delete[] str_wchar_t2;

#else
std::wstring str(paParams[0].pwstrVal);
std::wregex r(paParams[1].pwstrVal);
std::regex_search(str, wsmMatch, r);

for (auto r : wsmMatch) {
	vResults.push_back(r);
}
#endif
	iCurrentPosition = -1;
	m_PropCountOfItemsInSearchResult = wsmMatch.size();
}
bool CAddInNative::replace(tVariant * pvarRetValue, tVariant * paParams)
{

	iCurrentPosition = 0;
	m_PropCountOfItemsInSearchResult = 0;

#ifdef __linux__
	std::wstring str;
	// Сконвертируем в строку с wchar_t символами
	wchar_t* str_wchar_t = 0;
	convFromShortWchar(&str_wchar_t, paParams[0].pwstrVal, paParams[0].wstrLen);
	str = str_wchar_t;
	delete[] str_wchar_t;

	convFromShortWchar(&str_wchar_t, paParams[1].pwstrVal, paParams[1].wstrLen);
	std::wregex r(str_wchar_t);
	delete[] str_wchar_t;

	convFromShortWchar(&str_wchar_t, paParams[2].pwstrVal, paParams[2].wstrLen);
	std::wstring replacement(str_wchar_t);
	delete[] str_wchar_t;

	std::wstring res = std::regex_replace(str, r, replacement);

	if (m_iMemory->AllocMemory((void**)&pvarRetValue->pwstrVal, (res.length() + 1) * sizeof(WCHAR_T)))
	{

		convToShortWchar(&pvarRetValue->pwstrVal, res.c_str(), res.length() * sizeof(WCHAR_T));
		pvarRetValue->wstrLen = res.length();
		return true;
}
#else

	std::wstring str(paParams[0].pwstrVal);
	std::wregex r(paParams[1].pwstrVal);
	std::wstring res = std::regex_replace(str, r, paParams[2].pwstrVal);

	TV_VT(pvarRetValue) = VTYPE_PWSTR;
	if (m_iMemory->AllocMemory((void**)&pvarRetValue->pwstrVal, (res.length() + 1) * sizeof(wchar_t)))
	{
		memcpy(pvarRetValue->pwstrVal, res.c_str(), (res.length() + 1) * sizeof(wchar_t));
		pvarRetValue->wstrLen = res.length();
		return true;
	}
#endif
	return false;
}
void CAddInNative::match(tVariant * pvarRetValue, tVariant * paParams)
{
	TV_VT(pvarRetValue) = VTYPE_BOOL;
#ifdef __linux__
	wchar_t* str_wchar_t1 = 0;
	convFromShortWchar(&str_wchar_t1, paParams[0].pwstrVal, paParams[0].wstrLen + 1);

	std::wstring str(str_wchar_t1);

	wchar_t* str_wchar_t2 = 0;
	convFromShortWchar(&str_wchar_t2, paParams[1].pwstrVal, paParams[1].wstrLen + 1);
	std::wregex r(str_wchar_t2);

	pvarRetValue->bVal = std::regex_match(str, r);
#else

	std::wstring str(paParams[0].pwstrVal);
	std::wregex r(paParams[1].pwstrVal); 
	pvarRetValue->bVal = std::regex_match(str, r);
#endif
	iCurrentPosition = 0;
	m_PropCountOfItemsInSearchResult = 0;

}
//---------------------------------------------------------------------------//
long CAddInNative::findName(const wchar_t* names[], const wchar_t* name, 
                        const uint32_t size) const
{
    long ret = -1;
    for (uint32_t i = 0; i < size; i++)
    {
        if (!wcscmp(names[i], name))
        {
            ret = i;
            break;
        }
    }
    return ret;
}

//---------------------------------------------------------------------------//
uint32_t convToShortWchar(WCHAR_T** Dest, const wchar_t* Source, uint32_t len)
{
    if (!len)
        len = ::wcslen(Source) + 1;

    if (!*Dest)
        *Dest = new WCHAR_T[len];

    WCHAR_T* tmpShort = *Dest;
    wchar_t* tmpWChar = (wchar_t*) Source;
    uint32_t res = 0;

    ::memset(*Dest, 0, len * sizeof(WCHAR_T));
#ifdef __linux__
    size_t succeed = (size_t)-1;
    size_t f = len * sizeof(wchar_t), t = len * sizeof(WCHAR_T);
    const char* fromCode = sizeof(wchar_t) == 2 ? "UTF-16" : "UTF-32";
    iconv_t cd = iconv_open("UTF-16LE", fromCode);
    if (cd != (iconv_t)-1)
    {
        succeed = iconv(cd, (char**)&tmpWChar, &f, (char**)&tmpShort, &t);
        iconv_close(cd);
        if(succeed != (size_t)-1)
            return (uint32_t)succeed;
    }
#endif //__linux__
    for (; len; --len, ++res, ++tmpWChar, ++tmpShort)
    {
        *tmpShort = (WCHAR_T)*tmpWChar;
    }

    return res;
}
//---------------------------------------------------------------------------//
uint32_t convFromShortWchar(wchar_t** Dest, const WCHAR_T* Source, uint32_t len)
{
    if (!len)
        len = getLenShortWcharStr(Source) + 1;

    if (!*Dest)
        *Dest = new wchar_t[len];

    wchar_t* tmpWChar = *Dest;
    WCHAR_T* tmpShort = (WCHAR_T*)Source;
    uint32_t res = 0;

    ::memset(*Dest, 0, len * sizeof(wchar_t));
#ifdef __linux__
    size_t succeed = (size_t)-1;
    const char* fromCode = sizeof(wchar_t) == 2 ? "UTF-16" : "UTF-32";
    size_t f = len * sizeof(WCHAR_T), t = len * sizeof(wchar_t);
    iconv_t cd = iconv_open("UTF-32LE", fromCode);
    if (cd != (iconv_t)-1)
    {
        succeed = iconv(cd, (char**)&tmpShort, &f, (char**)&tmpWChar, &t);
        iconv_close(cd);
        if(succeed != (size_t)-1)
            return (uint32_t)succeed;
    }
#endif //__linux__
    for (; len; --len, ++res, ++tmpWChar, ++tmpShort)
    {
        *tmpWChar = (wchar_t)*tmpShort;
    }

    return res;
}
//---------------------------------------------------------------------------//


/*
//---------------------------------------------------------------------------//
// Преобразует строку, состоящую из wchar_t символов, в строку, 
// состоящую из WCHAR_T символов.
// Вызов функции имеет смысл только на тех компиляторах, где wchar_t <> WCHAR_T
// В Visual Studio, тип wchar_t = WCHAR_T, поэтому вызов функции не нужен.
// Функция выделяет память из кучи, поэтому её нужно не забыть очистить
// в вызывающей процедуре
uint32_t convToShortWchar(WCHAR_T** Dest, const wchar_t* Source, uint32_t len)
{
	if (!len)
		len = ::wcslen(Source) + 1;

	if (!*Dest)
		*Dest = new WCHAR_T[len];

	WCHAR_T* tmpShort = *Dest;	wchar_t* tmpWChar = (wchar_t*)Source;
	uint32_t res = 0;

	::memset(*Dest, 0, len * sizeof(WCHAR_T));
	do
	{
		*tmpShort++ = (WCHAR_T)*tmpWChar++;
		++res;
	} while (len-- && *tmpWChar);

	return res;
}

//---------------------------------------------------------------------------//
// Преобразует строку, состоящую из WCHAR_T символов, в строку, 
// состоящую из wchar_t символов.
// Вызов функции имеет смысл только на тех компиляторах, где wchar_t <> WCHAR_T
// В Visual Studio, тип wchar_t = WCHAR_T, поэтому вызов функции не нужен.
// Функция выделяет память из кучи, поэтому её в последствии нужно не забыть очистить
// в вызывающей процедуре
uint32_t convFromShortWchar(wchar_t** Dest, const WCHAR_T* Source, uint32_t len)
{
	if (!len)
		len = getLenShortWcharStr(Source);

	if (!*Dest) // Если строка пустая
		(*Dest) = new wchar_t[len + 1];

	wchar_t* tmpWChar = *Dest;
	WCHAR_T* tmpShort = (WCHAR_T*)Source;
	uint32_t res = 0;

	::memset((*Dest), 0, (len + 1) * sizeof(wchar_t));
	do
	{
		*tmpWChar++ = (wchar_t)*tmpShort++;
		++res;
	} while (len-- && *tmpShort);

	// добавим символ конца строки
	(*Dest)[res] = 0;
	return res;
}
*/

uint32_t getLenShortWcharStr(const WCHAR_T* Source)
{
    uint32_t res = 0;
    WCHAR_T *tmpShort = (WCHAR_T*)Source;

    while (*tmpShort++)
        ++res;

    return res;
}
//---------------------------------------------------------------------------//

#ifdef LINUX_OR_MACOS
WcharWrapper::WcharWrapper(const WCHAR_T* str) : m_str_WCHAR(NULL),
                           m_str_wchar(NULL)
{
    if (str)
    {
        int len = getLenShortWcharStr(str);
        m_str_WCHAR = new WCHAR_T[len + 1];
        memset(m_str_WCHAR,   0, sizeof(WCHAR_T) * (len + 1));
        memcpy(m_str_WCHAR, str, sizeof(WCHAR_T) * len);
        ::convFromShortWchar(&m_str_wchar, m_str_WCHAR);
    }
}
#endif
//---------------------------------------------------------------------------//
WcharWrapper::WcharWrapper(const wchar_t* str) :
#ifdef LINUX_OR_MACOS
    m_str_WCHAR(NULL),
#endif 
    m_str_wchar(NULL)
{
    if (str)
    {
        int len = wcslen(str);
        m_str_wchar = new wchar_t[len + 1];
        memset(m_str_wchar, 0, sizeof(wchar_t) * (len + 1));
        memcpy(m_str_wchar, str, sizeof(wchar_t) * len);
#ifdef LINUX_OR_MACOS
        ::convToShortWchar(&m_str_WCHAR, m_str_wchar);
#endif
    }

}
//---------------------------------------------------------------------------//
WcharWrapper::~WcharWrapper()
{
#ifdef LINUX_OR_MACOS
    if (m_str_WCHAR)
    {
        delete [] m_str_WCHAR;
        m_str_WCHAR = NULL;
    }
#endif
    if (m_str_wchar)
    {
        delete [] m_str_wchar;
        m_str_wchar = NULL;
    }
}
//---------------------------------------------------------------------------//
