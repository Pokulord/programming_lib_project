#pragma once
#ifdef XLSPARSERDLL_EXPORTS
#define XLSPARSERDLL_API __declspec(dllexport)
#else
#define XLSPARSERDLL_API __declspec(dllimport)
#endif

extern "C" XLSPARSERDLL_API void parseXLSFileStrings(const char* path);