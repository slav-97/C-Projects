

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


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



/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 500
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif /* __RPCNDR_H_VERSION__ */


#ifndef __ActiveXCPP3idl_h__
#define __ActiveXCPP3idl_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef ___DActiveXCPP3_FWD_DEFINED__
#define ___DActiveXCPP3_FWD_DEFINED__
typedef interface _DActiveXCPP3 _DActiveXCPP3;

#endif 	/* ___DActiveXCPP3_FWD_DEFINED__ */


#ifndef ___DActiveXCPP3Events_FWD_DEFINED__
#define ___DActiveXCPP3Events_FWD_DEFINED__
typedef interface _DActiveXCPP3Events _DActiveXCPP3Events;

#endif 	/* ___DActiveXCPP3Events_FWD_DEFINED__ */


#ifndef __ActiveXCPP3_FWD_DEFINED__
#define __ActiveXCPP3_FWD_DEFINED__

#ifdef __cplusplus
typedef class ActiveXCPP3 ActiveXCPP3;
#else
typedef struct ActiveXCPP3 ActiveXCPP3;
#endif /* __cplusplus */

#endif 	/* __ActiveXCPP3_FWD_DEFINED__ */


#ifdef __cplusplus
extern "C"{
#endif 


/* interface __MIDL_itf_ActiveXCPP3_0000_0000 */
/* [local] */ 

#pragma warning(push)
#pragma warning(disable:4001) 
#pragma once
#pragma warning(push)
#pragma warning(disable:4001) 
#pragma once
#pragma warning(pop)
#pragma warning(pop)
#pragma region Desktop Family
#pragma endregion


extern RPC_IF_HANDLE __MIDL_itf_ActiveXCPP3_0000_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_ActiveXCPP3_0000_0000_v0_0_s_ifspec;


#ifndef __ActiveXCPP3Lib_LIBRARY_DEFINED__
#define __ActiveXCPP3Lib_LIBRARY_DEFINED__

/* library ActiveXCPP3Lib */
/* [control][version][uuid] */ 


EXTERN_C const IID LIBID_ActiveXCPP3Lib;

#ifndef ___DActiveXCPP3_DISPINTERFACE_DEFINED__
#define ___DActiveXCPP3_DISPINTERFACE_DEFINED__

/* dispinterface _DActiveXCPP3 */
/* [uuid] */ 


EXTERN_C const IID DIID__DActiveXCPP3;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("c8bdc564-0a0e-4d2b-accc-e73d2d910656")
    _DActiveXCPP3 : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _DActiveXCPP3Vtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _DActiveXCPP3 * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _DActiveXCPP3 * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _DActiveXCPP3 * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _DActiveXCPP3 * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _DActiveXCPP3 * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _DActiveXCPP3 * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _DActiveXCPP3 * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } _DActiveXCPP3Vtbl;

    interface _DActiveXCPP3
    {
        CONST_VTBL struct _DActiveXCPP3Vtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _DActiveXCPP3_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _DActiveXCPP3_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _DActiveXCPP3_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _DActiveXCPP3_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _DActiveXCPP3_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _DActiveXCPP3_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _DActiveXCPP3_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___DActiveXCPP3_DISPINTERFACE_DEFINED__ */


#ifndef ___DActiveXCPP3Events_DISPINTERFACE_DEFINED__
#define ___DActiveXCPP3Events_DISPINTERFACE_DEFINED__

/* dispinterface _DActiveXCPP3Events */
/* [uuid] */ 


EXTERN_C const IID DIID__DActiveXCPP3Events;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("6b497838-a8b4-4940-9e15-2c23e3852967")
    _DActiveXCPP3Events : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _DActiveXCPP3EventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _DActiveXCPP3Events * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _DActiveXCPP3Events * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _DActiveXCPP3Events * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _DActiveXCPP3Events * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _DActiveXCPP3Events * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _DActiveXCPP3Events * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _DActiveXCPP3Events * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } _DActiveXCPP3EventsVtbl;

    interface _DActiveXCPP3Events
    {
        CONST_VTBL struct _DActiveXCPP3EventsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _DActiveXCPP3Events_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _DActiveXCPP3Events_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _DActiveXCPP3Events_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _DActiveXCPP3Events_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _DActiveXCPP3Events_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _DActiveXCPP3Events_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _DActiveXCPP3Events_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___DActiveXCPP3Events_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_ActiveXCPP3;

#ifdef __cplusplus

class DECLSPEC_UUID("e8fad06c-d6bd-4e2a-b9dc-ea11f58f8aaa")
ActiveXCPP3;
#endif
#endif /* __ActiveXCPP3Lib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


