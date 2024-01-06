#ifndef __ADDINNATIVE_H__
#define __ADDINNATIVE_H__

#define PCRE2_STATIC
#define PCRE2_CODE_UNIT_WIDTH 16
//#define PCRE2_LOCAL_WIDTH 16

#include "pcer2/pcre2.h"

#include "ComponentBase.h"
#include "AddInDefBase.h"
#include "IMemoryManager.h"

#include <vector>
#include <map>
#include <array> 
#include <locale.h>

#include "StrConv.h"
#include "json.h"

struct ResultStruct {
	std::basic_string<char16_t> value;
	size_t firstIndex;
};
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
		ePropFirstIndex,
		ePropMultiline,
		ePropUCP,
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
		eMethMatchesJSON,
		eMethTest,
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
	bool searchJSON(tVariant* pvarRetValue, tVariant* paParams);
	bool replace(tVariant* pvarRetValue, tVariant* paParams);
	bool match(tVariant* pvarRetValue, tVariant* paParams);
	bool getSubMatch(tVariant* pvarRetValue, tVariant* paParams);
	void version(tVariant* pvarRetValue);
	void SetLastError(const char16_t* error);

	//void GetStrParam(std::wstring& str, tVariant* paParams, const long paramIndex);
	pcre2_code* GetPattern(const tVariant *tvPattern);

    // Attributes
    IAddInDefBase      *m_iConnect;
    IMemoryManager     *m_iMemory;

	int m_PropCountOfItemsInSearchResult;
	std::vector<ResultStruct> vResults;
	std::map<size_t, std::vector<std::basic_string<char16_t>>> mSubMatches;
	int iCurrentPosition;
	std::basic_string<char16_t> sErrorDescription;
	bool bThrowExceptions;
	bool bIgnoreCase;
	bool bMultiline;
	bool bUCP;
	pcre2_code* rePattern;
	std::basic_string<char16_t> sPattern;

/*#if defined( __linux__ ) || defined(__APPLE__) || defined(__ANDROID__)
	std::u16string uPattern;
#endif*/
	bool isPattern;
	bool bGlobal;
	bool bHierarchicalResultIteration;
	size_t uiSubMatchesCount;
};

#endif //__ADDINNATIVE_H__
