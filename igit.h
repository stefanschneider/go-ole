#include "windows.h"
#include "OleAuto.h"
#include "ObjIdl.h"


#ifndef _IGIT_H_
#define _IGIT_H_

#ifdef __cplusplus
extern "C" {
#endif

  void initGit();
  DWORD registerInterfaceToGit(IDispatch* pIDispatch);
  HRESULT getIdispatchFromGlobal(DWORD dwCookie, void **ppv);
  HRESULT revokeInterfaceFromGlobal(DWORD dwCookie);

#ifdef __cplusplus
}
#endif

#endif