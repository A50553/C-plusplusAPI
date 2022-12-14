#pragma once

#include "SKCOM_reference.h"
#include "TEventHandler.h"
#include <iostream>
#include<string>
#include <typeinfo>
#include<vector>

class CCenter
{
public:
    typedef TEventHandlerNamespace::TEventHandler<CCenter, SKCOMLib::ISKCenterLib, SKCOMLib::_ISKCenterLibEvents> ISKCenterLibEventHandler;
    CCenter();
    ~CCenter();
    //Methods
    long Login(string strUserID, string strPassword);
    BSTR GetReturnCodeMessage(long nCode);
    BSTR GetLastLogInfo();
    void OnTimer(long nTime);
    void PrintCodeMessage(string strFunction, string strFunctionName, long nCode);
private:
    HRESULT OnEventFiringObjectInvoke(
        ISKCenterLibEventHandler* pEventHandler,
        DISPID dispidMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS* pdispparams,
        VARIANT* pvarResult,
        EXCEPINFO* pexcepinfo,
        UINT* puArgErr
    );
    SKCOMLib::ISKCenterLibPtr m_pSKCenterLib;
    ISKCenterLibEventHandler* m_pSKCenterLibEventHandler;
};

