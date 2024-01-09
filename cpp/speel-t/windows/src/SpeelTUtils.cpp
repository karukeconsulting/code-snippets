#include "SpeelTUtils.h"
#include "winreg.h"

bool IsExcelInstalled() {
	HKEY hkResult;
	LSTATUS opResult;

	opResult = RegOpenKeyExW(HKEY_CLASSES_ROOT, L"Excel.Application", 0, KEY_READ, &hkResult);
}