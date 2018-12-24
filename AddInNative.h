#ifndef __ADDINNATIVE_H__
#define __ADDINNATIVE_H__

#include "boost/regex.hpp"

#include "ComponentBase.h"
#include "AddInDefBase.h"
#include "IMemoryManager.h"

#define MBCMAXSIZE  6 // максимальная длина символа мультибайтовой строки (для функции wcstombs)

///////////////////////////////////////////////////////////////////////////////
// class CAddInNative
class CAddInNative : public IComponentBase
{
public:
    enum Props
    {
		ePropCurrentValue = 0,
		ePropIgnoreCase,
		ePropErrorDescription,
		ePropThrowExceptions,
		ePropPattern,
		ePropGlobal,
        ePropLast      // Always last
    };

    enum Methods
    {
        eMethMatches = 0,
        eMethIsMatch,
        eMethNext,
		eMethReplace,
		eMethCount,
		eMethVersion,
        eMethLast      // Always last
    };

    CAddInNative(void);
    virtual ~CAddInNative();
    // IInitDoneBase
    virtual bool ADDIN_API Init(void*);
    virtual bool ADDIN_API setMemManager(void* mem);
    virtual long ADDIN_API GetInfo();
    virtual void ADDIN_API Done();
    // ILanguageExtenderBase
    virtual bool ADDIN_API RegisterExtensionAs(WCHAR_T**);
    virtual long ADDIN_API GetNProps();
    virtual long ADDIN_API FindProp(const WCHAR_T* wsPropName);
    virtual const WCHAR_T* ADDIN_API GetPropName(long lPropNum, long lPropAlias);
    virtual bool ADDIN_API GetPropVal(const long lPropNum, tVariant* pvarPropVal);
    virtual bool ADDIN_API SetPropVal(const long lPropNum, tVariant* varPropVal);
    virtual bool ADDIN_API IsPropReadable(const long lPropNum);
    virtual bool ADDIN_API IsPropWritable(const long lPropNum);
    virtual long ADDIN_API GetNMethods();
    virtual long ADDIN_API FindMethod(const WCHAR_T* wsMethodName);
    virtual const WCHAR_T* ADDIN_API GetMethodName(const long lMethodNum, 
                            const long lMethodAlias);
    virtual long ADDIN_API GetNParams(const long lMethodNum);
    virtual bool ADDIN_API GetParamDefValue(const long lMethodNum, const long lParamNum,
                            tVariant *pvarParamDefValue);   
    virtual bool ADDIN_API HasRetVal(const long lMethodNum);
    virtual bool ADDIN_API CallAsProc(const long lMethodNum,
                    tVariant* paParams, const long lSizeArray);
    virtual bool ADDIN_API CallAsFunc(const long lMethodNum,
                tVariant* pvarRetValue, tVariant* paParams, const long lSizeArray);
    // LocaleBase
    virtual void ADDIN_API SetLocale(const WCHAR_T* loc);
    
private:
    long findName(const wchar_t* names[], const wchar_t* name, const uint32_t size) const;

	bool search(tVariant* paParams);
	bool replace(tVariant* pvarRetValue, tVariant* paParams);
	bool match(tVariant* pvarRetValue, tVariant* paParams);
	void version(tVariant* pvarRetValue);
	void SetLastError(const char* error);

    // Attributes
    IAddInDefBase      *m_iConnect;
    IMemoryManager     *m_iMemory;

	uint32_t m_PropCountOfItemsInSearchResult;
	std::vector<std::wstring> vResults;
	uint32_t iCurrentPosition;
	std::string sErrorDescription;
	bool bThrowExceptions;
	bool bIgnoreCase;
	boost::wregex rePattern;
	std::basic_string<WCHAR> wcsPattern;
	bool bGlobal;
};

class WcharWrapper
{
public:
#ifdef LINUX_OR_MACOS
    WcharWrapper(const WCHAR_T* str);
#endif
    WcharWrapper(const wchar_t* str);
    ~WcharWrapper();
#ifdef LINUX_OR_MACOS
    operator const WCHAR_T*(){ return m_str_WCHAR; }
    operator WCHAR_T*(){ return m_str_WCHAR; }
#endif
    operator const wchar_t*(){ return m_str_wchar; }
    operator wchar_t*(){ return m_str_wchar; }
private:
    WcharWrapper& operator = (const WcharWrapper& other) { return *this; }
    WcharWrapper(const WcharWrapper& other) { }
private:
#ifdef LINUX_OR_MACOS
    WCHAR_T* m_str_WCHAR;
#endif
    wchar_t* m_str_wchar;
};
#endif //__ADDINNATIVE_H__
