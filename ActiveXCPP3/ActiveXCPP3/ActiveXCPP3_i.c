

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


 /* File created by MIDL compiler version 8.01.0622 */
/* at Tue Jan 19 06:14:07 2038
 */
/* Compiler settings for ActiveXCPP3.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.01.0622 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */



#ifdef __cplusplus
extern "C"{
#endif 


#include <rpc.h>
#include <rpcndr.h>

#ifdef _MIDL_USE_GUIDDEF_

#ifndef INITGUID
#define INITGUID
#include <guiddef.h>
#undef INITGUID
#else
#include <guiddef.h>
#endif

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        DEFINE_GUID(name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8)

#else // !_MIDL_USE_GUIDDEF_

#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        EXTERN_C __declspec(selectany) const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

#endif // !_MIDL_USE_GUIDDEF_

MIDL_DEFINE_GUID(IID, LIBID_ActiveXCPP3Lib,0x8971dbf5,0x18f3,0x433e,0x88,0x5a,0xa0,0x7c,0x15,0x88,0x53,0x12);


MIDL_DEFINE_GUID(IID, DIID__DActiveXCPP3,0xc8bdc564,0x0a0e,0x4d2b,0xac,0xcc,0xe7,0x3d,0x2d,0x91,0x06,0x56);


MIDL_DEFINE_GUID(IID, DIID__DActiveXCPP3Events,0x6b497838,0xa8b4,0x4940,0x9e,0x15,0x2c,0x23,0xe3,0x85,0x29,0x67);


MIDL_DEFINE_GUID(CLSID, CLSID_ActiveXCPP3,0xe8fad06c,0xd6bd,0x4e2a,0xb9,0xdc,0xea,0x11,0xf5,0x8f,0x8a,0xaa);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



