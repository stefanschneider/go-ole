#include "windows.h"
#include "OleAuto.h"
#include "Objbase.h"
#include "ObjIdl.h"
#include "olectl.h"
#include "oaidl.h"
#include "igit.h"

const GUID  IID_IGlobalInterfaceTable = { 0x00000146, 0, 0, {0xc0, 0, 0, 0, 0, 0, 0, 0x46} };
const CLSID CLSID_StdGlobalInterfaceTable = { 0x00000323, 0, 0, {0xc0, 0, 0, 0, 0, 0, 0, 0x46} };
const GUID  IID_IDispatch  = 				{ 0x00020400, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 }};



IGlobalInterfaceTable* gpGIT = NULL;

void initGit(){
	HRESULT hr;

	if (gpGIT == NULL){
		hr = ::CoCreateInstance(CLSID_StdGlobalInterfaceTable,
		                 NULL,
		                 CLSCTX_INPROC_SERVER,
		                 IID_IGlobalInterfaceTable,
		                 (void **)&gpGIT);
		if (hr != S_OK) {
		}
	}

}

DWORD registerInterfaceToGit(IDispatch* pIDispatch){
	DWORD res;
	gpGIT -> RegisterInterfaceInGlobal
      (
        pIDispatch,
        IID_IDispatch,
        &res
      );
	return res;
}

HRESULT getIdispatchFromGlobal(DWORD dwCookie, void **ppv){
  return gpGIT -> GetInterfaceFromGlobal
    (
      dwCookie,
      IID_IDispatch,
      (void**)ppv
    );
}

HRESULT revokeInterfaceFromGlobal(DWORD dwCookie){
  return gpGIT -> RevokeInterfaceFromGlobal(dwCookie);
}