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
		eMethSubMatchesCount,
		eMethGetSubMatch,
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
 	bool search(tVariant* paParams);
	bool replace(tVariant* pvarRetValue, tVariant* paParams);
	bool match(tVariant* pvarRetValue, tVariant* paParams);
	bool getSubMatch(tVariant* pvarRetValue, tVariant* paParams);
	void version(tVariant* pvarRetValue);
	void SetLastError(const char* error);

    // Attributes
    IAddInDefBase      *m_iConnect;
    IMemoryManager     *m_iMemory;

	int m_PropCountOfItemsInSearchResult;
	std::vector<std::wstring> vResults;
	std::map<size_t, std::vector<std::wstring>> mSubMatches;
	int iCurrentPosition;
	std::string sErrorDescription;
	bool bThrowExceptions;
	bool bIgnoreCase;
	boost::wregex rePattern;
#if defined( __linux__ ) || defined(__APPLE__) || defined(__ANDROID__)
	std::u16string uPattern;
#endif
	bool isPattern;
	bool bGlobal;
	bool bHierarchicalResultIteration;
	size_t uiSubMatchesCount;
};

#endif //__ADDINNATIVE_H__
