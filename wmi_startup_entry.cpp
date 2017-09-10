
#include <Windows.h>
#include <comdef.h>
#include <WbemIdl.h>
#include <comutil.h>
#include <iostream>
#include <tchar.h>

#pragma comment(lib, "wbemuuid.lib")
#pragma comment(lib, "comsuppw.lib")

using namespace std;

BOOL getStringFromClass(IWbemClassObject *obj, BSTR key, BSTR &val) {
	VARIANT v;
	HRESULT hr;
	hr = obj->Get(key, 0, &v, 0, 0);
	if (FAILED(hr))
		return FALSE;
	if (V_VT(&v) != VT_BSTR)
	{
		VariantClear(&v);
		return FALSE;
	}
	val = _bstr_t(V_BSTR(&v)).copy();
	VariantClear(&v);

	return TRUE;
}

void main()
{
	IWbemLocator *ploc = NULL;
	IWbemServices *psvc = NULL;

	//init com interface
	IEnumWbemClassObject *activeScriptEventConsumer = NULL, *commandLineEventConsumer = NULL;
	IWbemClassObject *obj = NULL;
	BSTR name = NULL, value = NULL;
	char *szName = NULL, *szValue = NULL;
	ULONG tmp;

	HRESULT h = CoInitializeEx(0, COINIT_MULTITHREADED);
	if (FAILED(h)) {
		return;
	}

	//init COM security context
	/*h = CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_DEFAULT, RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, NULL);
	if (FAILED(h)) {
		return;
	}*/

	//create COM instance for WBEM
	h = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID*)&ploc);
	if (FAILED(h)) {
		return;
	}

	//connect to the \\root\subscription namespace
	h = ploc->ConnectServer(_bstr_t(L"ROOT\\SUBSCRIPTION"), NULL, NULL, 0, NULL, 0, 0, &psvc);
	if (FAILED(h)) {
		ploc->Release();
		return;
	}

	h = psvc->ExecQuery(_bstr_t(L"WQL"), _bstr_t(L"Select * from ActiveScriptEventConsumer"), WBEM_FLAG_RETURN_IMMEDIATELY, NULL, &activeScriptEventConsumer);
	if (h == S_OK) {
		activeScriptEventConsumer->Reset();
		while (true) {
			h = activeScriptEventConsumer->Next(WBEM_INFINITE, 1, &obj, &tmp);
			if (h == S_OK)
			{
				if (getStringFromClass(obj, _bstr_t(L"Name"), name) &&
					getStringFromClass(obj, _bstr_t(L"ScriptFilename"), value)) {
					szName = _com_util::ConvertBSTRToString(name);
					szValue = _com_util::ConvertBSTRToString(value);
					cout << szName << ":" << szValue << endl;
				}
				if (name) {
					::SysFreeString(name);
					name = NULL;
				}
				if (value) {
					::SysFreeString(value);
					value = NULL;
				}
				if (szName) {
					delete[] szName;
					szName = NULL;
				}
				if (szValue) {
					delete[] szValue;
					szValue = NULL;
				}
				obj->Release();
			} else break;
		}
	}

	if (activeScriptEventConsumer) {
		activeScriptEventConsumer->Release();
		activeScriptEventConsumer = NULL;
	}

	h = psvc->ExecQuery(_bstr_t(L"WQL"), _bstr_t(L"Select * from CommandLineEventConsumer"), WBEM_FLAG_RETURN_IMMEDIATELY, NULL, &commandLineEventConsumer);
	if (h == S_OK) {
		commandLineEventConsumer->Reset();
		while (true) {
			h = commandLineEventConsumer->Next(WBEM_INFINITE, 1, &obj, &tmp);
			if (h == S_OK) {
				if (getStringFromClass(obj, _bstr_t(L"Name"), name) &&
					getStringFromClass(obj, _bstr_t(L"CommandLineTemplate"), value)) {
					szName = _com_util::ConvertBSTRToString(name);
					szValue = _com_util::ConvertBSTRToString(value);
					cout << szName << ":" << szValue << endl;
				}
				if (name) {
					::SysFreeString(name);
					name = NULL;
				}
				if (value) {
					::SysFreeString(value);
					value = NULL;
				}
				if (szName) {
					delete[] szName;
					szName = NULL;
				}
				if (szValue) {
					delete[] szValue;
					szValue = NULL;
				}
				obj->Release();
			} else break;
		}
	}
	if (commandLineEventConsumer) {
		commandLineEventConsumer->Release();
		commandLineEventConsumer = NULL;
	}

	psvc->Release();
	ploc->Release();
}
