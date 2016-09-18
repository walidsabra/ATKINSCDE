// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the PWSDKMETHODS_EXPORTS
// symbol defined on the command line. This symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// PWSDKMETHODS_API functions as being imported from a DLL, whereas this DLL sees symbols
// defined with this macro as being exported.
#ifdef PWSDKMETHODS_EXPORTS
#define PWSDKMETHODS_API __declspec(dllexport)
#else
#define PWSDKMETHODS_API __declspec(dllimport)
#endif

// This class is exported from the PWSDKMethods.dll
class PWSDKMETHODS_API CPWSDKMethods {
public:
	CPWSDKMethods(void);
	// TODO: add your methods here.
};

extern PWSDKMETHODS_API int nPWSDKMethods;

PWSDKMETHODS_API int fnPWSDKMethods(void);
