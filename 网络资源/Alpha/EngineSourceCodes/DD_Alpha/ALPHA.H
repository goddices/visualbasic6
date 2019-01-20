
// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the ALPHA_EXPORTS
// symbol defined on the command line. this symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// ALPHA_API functions as being imported from a DLL, wheras this DLL sees symbols
// defined with this macro as being exported.
#ifdef ALPHA_EXPORTS
#define ALPHA_API __declspec(dllexport)
#else
#define ALPHA_API __declspec(dllimport)
#endif

// This class is exported from the alpha.dll
class ALPHA_API CAlpha {
public:
	CAlpha(void);
	// TODO: add your methods here.
};

extern ALPHA_API int nAlpha;

ALPHA_API int fnAlpha(void);

