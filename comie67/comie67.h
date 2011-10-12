
#pragma warning( disable: 4049 )  /* more than 64k source lines */

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 6.00.0347 */
/* at Wed Oct 12 20:39:56 2011
 */
/* Compiler settings for comie67.idl:
    Oicf, W1, Zp8, env=Win32 (32b run)
    protocol : dce , ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __comie67_h__
#define __comie67_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IConsoleObject_FWD_DEFINED__
#define __IConsoleObject_FWD_DEFINED__
typedef interface IConsoleObject IConsoleObject;
#endif 	/* __IConsoleObject_FWD_DEFINED__ */


#ifndef __IExplorerBar_FWD_DEFINED__
#define __IExplorerBar_FWD_DEFINED__
typedef interface IExplorerBar IExplorerBar;
#endif 	/* __IExplorerBar_FWD_DEFINED__ */


#ifndef ___IConsoleObjectEvents_FWD_DEFINED__
#define ___IConsoleObjectEvents_FWD_DEFINED__
typedef interface _IConsoleObjectEvents _IConsoleObjectEvents;
#endif 	/* ___IConsoleObjectEvents_FWD_DEFINED__ */


#ifndef __ConsoleObject_FWD_DEFINED__
#define __ConsoleObject_FWD_DEFINED__

#ifdef __cplusplus
typedef class ConsoleObject ConsoleObject;
#else
typedef struct ConsoleObject ConsoleObject;
#endif /* __cplusplus */

#endif 	/* __ConsoleObject_FWD_DEFINED__ */


#ifndef __ExplorerBar_FWD_DEFINED__
#define __ExplorerBar_FWD_DEFINED__

#ifdef __cplusplus
typedef class ExplorerBar ExplorerBar;
#else
typedef struct ExplorerBar ExplorerBar;
#endif /* __cplusplus */

#endif 	/* __ExplorerBar_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"
#include "dispex.h"

#ifdef __cplusplus
extern "C"{
#endif 

void * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void * ); 

/* interface __MIDL_itf_comie67_0000 */
/* [local] */ 

/* [v1_enum][hidden][uuid] */ 
enum  DECLSPEC_UUID("15045c3b-d7ab-45e0-8ef9-15bf98563053") __MIDL___MIDL_itf_comie67_0000_0001
    {	COMIE67_LOG	= DISPID_VALUE + 1,
	COMIE67_DEBUG	= COMIE67_LOG + 1,
	COMIE67_INFO	= COMIE67_DEBUG + 1,
	COMIE67_WARN	= COMIE67_INFO + 1,
	COMIE67_ERROR	= COMIE67_WARN + 1,
	COMIE67_ASSERT	= COMIE67_ERROR + 1,
	COMIE67_CLEAR	= COMIE67_ASSERT + 1,
	COMIE67_TRACE	= COMIE67_CLEAR + 1,
	COMIE67_TIME	= COMIE67_TRACE + 1,
	COMIE67_TIMEEND	= COMIE67_TIME + 1,
	COMIE67_PROFILE	= COMIE67_TIMEEND + 1,
	COMIE67_PROFILEEND	= COMIE67_PROFILE + 1,
	COMIE67_COUNT	= COMIE67_PROFILEEND + 1,
	COMIE67_ENDOFMETHODSLIST	= COMIE67_COUNT + 1
    } ;


extern RPC_IF_HANDLE __MIDL_itf_comie67_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_comie67_0000_v0_0_s_ifspec;

#ifndef __IConsoleObject_INTERFACE_DEFINED__
#define __IConsoleObject_INTERFACE_DEFINED__

/* interface IConsoleObject */
/* [unique][oleautomation][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IConsoleObject;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("20AB98F7-2E2C-433C-98C4-A858190C9714")
    IConsoleObject : public IDispatchEx
    {
    public:
        virtual /* [vararg][id] */ HRESULT STDMETHODCALLTYPE log( 
            /* [in] */ SAFEARRAY * varg) = 0;
        
        virtual /* [vararg][id] */ HRESULT STDMETHODCALLTYPE debug( 
            /* [in] */ SAFEARRAY * varg) = 0;
        
        virtual /* [vararg][id] */ HRESULT STDMETHODCALLTYPE info( 
            /* [in] */ SAFEARRAY * varg) = 0;
        
        virtual /* [vararg][id] */ HRESULT STDMETHODCALLTYPE warn( 
            /* [in] */ SAFEARRAY * varg) = 0;
        
        virtual /* [vararg][id] */ HRESULT STDMETHODCALLTYPE error( 
            /* [in] */ SAFEARRAY * varg) = 0;
        
        virtual /* [vararg][id] */ HRESULT STDMETHODCALLTYPE assert( 
            /* [in] */ VARIANT_BOOL expr,
            /* [in] */ SAFEARRAY * varg) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE clear( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE trace( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE time( 
            /* [in] */ BSTR name) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE timeEnd( 
            /* [in] */ BSTR name) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE profile( 
            /* [in] */ BSTR title) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE profileEnd( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE count( 
            /* [in] */ BSTR title) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IConsoleObjectVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IConsoleObject * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IConsoleObject * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IConsoleObject * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IConsoleObject * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IConsoleObject * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IConsoleObject * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IConsoleObject * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        HRESULT ( STDMETHODCALLTYPE *GetDispID )( 
            IConsoleObject * This,
            /* [in] */ BSTR bstrName,
            /* [in] */ DWORD grfdex,
            /* [out] */ DISPID *pid);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *InvokeEx )( 
            IConsoleObject * This,
            /* [in] */ DISPID id,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [in] */ DISPPARAMS *pdp,
            /* [out] */ VARIANT *pvarRes,
            /* [out] */ EXCEPINFO *pei,
            /* [unique][in] */ IServiceProvider *pspCaller);
        
        HRESULT ( STDMETHODCALLTYPE *DeleteMemberByName )( 
            IConsoleObject * This,
            /* [in] */ BSTR bstrName,
            /* [in] */ DWORD grfdex);
        
        HRESULT ( STDMETHODCALLTYPE *DeleteMemberByDispID )( 
            IConsoleObject * This,
            /* [in] */ DISPID id);
        
        HRESULT ( STDMETHODCALLTYPE *GetMemberProperties )( 
            IConsoleObject * This,
            /* [in] */ DISPID id,
            /* [in] */ DWORD grfdexFetch,
            /* [out] */ DWORD *pgrfdex);
        
        HRESULT ( STDMETHODCALLTYPE *GetMemberName )( 
            IConsoleObject * This,
            /* [in] */ DISPID id,
            /* [out] */ BSTR *pbstrName);
        
        HRESULT ( STDMETHODCALLTYPE *GetNextDispID )( 
            IConsoleObject * This,
            /* [in] */ DWORD grfdex,
            /* [in] */ DISPID id,
            /* [out] */ DISPID *pid);
        
        HRESULT ( STDMETHODCALLTYPE *GetNameSpaceParent )( 
            IConsoleObject * This,
            /* [out] */ IUnknown **ppunk);
        
        /* [vararg][id] */ HRESULT ( STDMETHODCALLTYPE *log )( 
            IConsoleObject * This,
            /* [in] */ SAFEARRAY * varg);
        
        /* [vararg][id] */ HRESULT ( STDMETHODCALLTYPE *debug )( 
            IConsoleObject * This,
            /* [in] */ SAFEARRAY * varg);
        
        /* [vararg][id] */ HRESULT ( STDMETHODCALLTYPE *info )( 
            IConsoleObject * This,
            /* [in] */ SAFEARRAY * varg);
        
        /* [vararg][id] */ HRESULT ( STDMETHODCALLTYPE *warn )( 
            IConsoleObject * This,
            /* [in] */ SAFEARRAY * varg);
        
        /* [vararg][id] */ HRESULT ( STDMETHODCALLTYPE *error )( 
            IConsoleObject * This,
            /* [in] */ SAFEARRAY * varg);
        
        /* [vararg][id] */ HRESULT ( STDMETHODCALLTYPE *assert )( 
            IConsoleObject * This,
            /* [in] */ VARIANT_BOOL expr,
            /* [in] */ SAFEARRAY * varg);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *clear )( 
            IConsoleObject * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *trace )( 
            IConsoleObject * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *time )( 
            IConsoleObject * This,
            /* [in] */ BSTR name);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *timeEnd )( 
            IConsoleObject * This,
            /* [in] */ BSTR name);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *profile )( 
            IConsoleObject * This,
            /* [in] */ BSTR title);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *profileEnd )( 
            IConsoleObject * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *count )( 
            IConsoleObject * This,
            /* [in] */ BSTR title);
        
        END_INTERFACE
    } IConsoleObjectVtbl;

    interface IConsoleObject
    {
        CONST_VTBL struct IConsoleObjectVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IConsoleObject_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IConsoleObject_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IConsoleObject_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IConsoleObject_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IConsoleObject_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IConsoleObject_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IConsoleObject_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IConsoleObject_GetDispID(This,bstrName,grfdex,pid)	\
    (This)->lpVtbl -> GetDispID(This,bstrName,grfdex,pid)

#define IConsoleObject_InvokeEx(This,id,lcid,wFlags,pdp,pvarRes,pei,pspCaller)	\
    (This)->lpVtbl -> InvokeEx(This,id,lcid,wFlags,pdp,pvarRes,pei,pspCaller)

#define IConsoleObject_DeleteMemberByName(This,bstrName,grfdex)	\
    (This)->lpVtbl -> DeleteMemberByName(This,bstrName,grfdex)

#define IConsoleObject_DeleteMemberByDispID(This,id)	\
    (This)->lpVtbl -> DeleteMemberByDispID(This,id)

#define IConsoleObject_GetMemberProperties(This,id,grfdexFetch,pgrfdex)	\
    (This)->lpVtbl -> GetMemberProperties(This,id,grfdexFetch,pgrfdex)

#define IConsoleObject_GetMemberName(This,id,pbstrName)	\
    (This)->lpVtbl -> GetMemberName(This,id,pbstrName)

#define IConsoleObject_GetNextDispID(This,grfdex,id,pid)	\
    (This)->lpVtbl -> GetNextDispID(This,grfdex,id,pid)

#define IConsoleObject_GetNameSpaceParent(This,ppunk)	\
    (This)->lpVtbl -> GetNameSpaceParent(This,ppunk)


#define IConsoleObject_log(This,varg)	\
    (This)->lpVtbl -> log(This,varg)

#define IConsoleObject_debug(This,varg)	\
    (This)->lpVtbl -> debug(This,varg)

#define IConsoleObject_info(This,varg)	\
    (This)->lpVtbl -> info(This,varg)

#define IConsoleObject_warn(This,varg)	\
    (This)->lpVtbl -> warn(This,varg)

#define IConsoleObject_error(This,varg)	\
    (This)->lpVtbl -> error(This,varg)

#define IConsoleObject_assert(This,expr,varg)	\
    (This)->lpVtbl -> assert(This,expr,varg)

#define IConsoleObject_clear(This)	\
    (This)->lpVtbl -> clear(This)

#define IConsoleObject_trace(This)	\
    (This)->lpVtbl -> trace(This)

#define IConsoleObject_time(This,name)	\
    (This)->lpVtbl -> time(This,name)

#define IConsoleObject_timeEnd(This,name)	\
    (This)->lpVtbl -> timeEnd(This,name)

#define IConsoleObject_profile(This,title)	\
    (This)->lpVtbl -> profile(This,title)

#define IConsoleObject_profileEnd(This)	\
    (This)->lpVtbl -> profileEnd(This)

#define IConsoleObject_count(This,title)	\
    (This)->lpVtbl -> count(This,title)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [vararg][id] */ HRESULT STDMETHODCALLTYPE IConsoleObject_log_Proxy( 
    IConsoleObject * This,
    /* [in] */ SAFEARRAY * varg);


void __RPC_STUB IConsoleObject_log_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [vararg][id] */ HRESULT STDMETHODCALLTYPE IConsoleObject_debug_Proxy( 
    IConsoleObject * This,
    /* [in] */ SAFEARRAY * varg);


void __RPC_STUB IConsoleObject_debug_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [vararg][id] */ HRESULT STDMETHODCALLTYPE IConsoleObject_info_Proxy( 
    IConsoleObject * This,
    /* [in] */ SAFEARRAY * varg);


void __RPC_STUB IConsoleObject_info_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [vararg][id] */ HRESULT STDMETHODCALLTYPE IConsoleObject_warn_Proxy( 
    IConsoleObject * This,
    /* [in] */ SAFEARRAY * varg);


void __RPC_STUB IConsoleObject_warn_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [vararg][id] */ HRESULT STDMETHODCALLTYPE IConsoleObject_error_Proxy( 
    IConsoleObject * This,
    /* [in] */ SAFEARRAY * varg);


void __RPC_STUB IConsoleObject_error_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [vararg][id] */ HRESULT STDMETHODCALLTYPE IConsoleObject_assert_Proxy( 
    IConsoleObject * This,
    /* [in] */ VARIANT_BOOL expr,
    /* [in] */ SAFEARRAY * varg);


void __RPC_STUB IConsoleObject_assert_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IConsoleObject_clear_Proxy( 
    IConsoleObject * This);


void __RPC_STUB IConsoleObject_clear_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IConsoleObject_trace_Proxy( 
    IConsoleObject * This);


void __RPC_STUB IConsoleObject_trace_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IConsoleObject_time_Proxy( 
    IConsoleObject * This,
    /* [in] */ BSTR name);


void __RPC_STUB IConsoleObject_time_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IConsoleObject_timeEnd_Proxy( 
    IConsoleObject * This,
    /* [in] */ BSTR name);


void __RPC_STUB IConsoleObject_timeEnd_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IConsoleObject_profile_Proxy( 
    IConsoleObject * This,
    /* [in] */ BSTR title);


void __RPC_STUB IConsoleObject_profile_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IConsoleObject_profileEnd_Proxy( 
    IConsoleObject * This);


void __RPC_STUB IConsoleObject_profileEnd_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IConsoleObject_count_Proxy( 
    IConsoleObject * This,
    /* [in] */ BSTR title);


void __RPC_STUB IConsoleObject_count_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IConsoleObject_INTERFACE_DEFINED__ */


#ifndef __IExplorerBar_INTERFACE_DEFINED__
#define __IExplorerBar_INTERFACE_DEFINED__

/* interface IExplorerBar */
/* [unique][uuid][object] */ 


EXTERN_C const IID IID_IExplorerBar;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("93CCB8D4-4DAB-4D31-91C3-19E7E39B4E76")
    IExplorerBar : public IUnknown
    {
    public:
    };
    
#else 	/* C style interface */

    typedef struct IExplorerBarVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IExplorerBar * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IExplorerBar * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IExplorerBar * This);
        
        END_INTERFACE
    } IExplorerBarVtbl;

    interface IExplorerBar
    {
        CONST_VTBL struct IExplorerBarVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IExplorerBar_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IExplorerBar_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IExplorerBar_Release(This)	\
    (This)->lpVtbl -> Release(This)


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IExplorerBar_INTERFACE_DEFINED__ */



#ifndef __comie67Lib_LIBRARY_DEFINED__
#define __comie67Lib_LIBRARY_DEFINED__

/* library comie67Lib */
/* [version][uuid] */ 

typedef /* [uuid][v1_enum] */  DECLSPEC_UUID("163804ab-42b5-47aa-90c7-7e910e897d4c") 
enum _PrintLevel
    {	LEVEL_NORMAL	= 0,
	LEVEL_DEBUG	= LEVEL_NORMAL + 1,
	LEVEL_INFO	= LEVEL_DEBUG + 1,
	LEVEL_WARN	= LEVEL_INFO + 1,
	LEVEL_ERROR	= LEVEL_WARN + 1
    } 	PrintLevel;

/* [v1_enum][uuid] */ 
enum  DECLSPEC_UUID("39F22B65-C3B4-4a31-93CC-B43A244388E4") __MIDL___MIDL_itf_comie67_0262_0001
    {	COMIE67_PRINT_EVENT_ID	= 1,
	COMIE67_CLEAR_EVENT_ID	= COMIE67_PRINT_EVENT_ID + 1
    } ;

EXTERN_C const IID LIBID_comie67Lib;

#ifndef ___IConsoleObjectEvents_DISPINTERFACE_DEFINED__
#define ___IConsoleObjectEvents_DISPINTERFACE_DEFINED__

/* dispinterface _IConsoleObjectEvents */
/* [uuid] */ 


EXTERN_C const IID DIID__IConsoleObjectEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("77C4A043-1F8B-423C-9E0F-E07AA36B08E7")
    _IConsoleObjectEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _IConsoleObjectEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _IConsoleObjectEvents * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _IConsoleObjectEvents * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _IConsoleObjectEvents * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _IConsoleObjectEvents * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _IConsoleObjectEvents * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _IConsoleObjectEvents * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _IConsoleObjectEvents * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        END_INTERFACE
    } _IConsoleObjectEventsVtbl;

    interface _IConsoleObjectEvents
    {
        CONST_VTBL struct _IConsoleObjectEventsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _IConsoleObjectEvents_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define _IConsoleObjectEvents_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define _IConsoleObjectEvents_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define _IConsoleObjectEvents_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define _IConsoleObjectEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define _IConsoleObjectEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define _IConsoleObjectEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___IConsoleObjectEvents_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_ConsoleObject;

#ifdef __cplusplus

class DECLSPEC_UUID("653823FB-32DD-4863-B8D8-4579A667400A")
ConsoleObject;
#endif

EXTERN_C const CLSID CLSID_ExplorerBar;

#ifdef __cplusplus

class DECLSPEC_UUID("59098C21-A36C-4758-8F7D-B6EE4E9B6562")
ExplorerBar;
#endif
#endif /* __comie67Lib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long *, unsigned long            , BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserMarshal(  unsigned long *, unsigned char *, BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserUnmarshal(unsigned long *, unsigned char *, BSTR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long *, BSTR * ); 

unsigned long             __RPC_USER  LPSAFEARRAY_UserSize(     unsigned long *, unsigned long            , LPSAFEARRAY * ); 
unsigned char * __RPC_USER  LPSAFEARRAY_UserMarshal(  unsigned long *, unsigned char *, LPSAFEARRAY * ); 
unsigned char * __RPC_USER  LPSAFEARRAY_UserUnmarshal(unsigned long *, unsigned char *, LPSAFEARRAY * ); 
void                      __RPC_USER  LPSAFEARRAY_UserFree(     unsigned long *, LPSAFEARRAY * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


