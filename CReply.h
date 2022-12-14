#pragma once

#include "SKCOM_reference.h"
#include "TEventHandler.h"
#include <iostream>
#include <string>
#include <typeinfo>
#include "CCenter.h"

class CReply
{
public:
    typedef TEventHandlerNamespace::TEventHandler< CReply, SKCOMLib::ISKReplyLib, SKCOMLib::_ISKReplyLibEvents> ISKReplyLibEventHandler;
    CReply();
    ~CReply();
    long ConnectById(string strLogInID);
    long IsConnecttedByID(string strLogInID);
    long SolaceCloseByID(string strLogInID);
    void OnReplyMessage(string strUserID, string strMessage, short* sConfirmCode);
    void OnDisconnect(string strLogInID,long nErrorCode);
    void OnComplete();
    void OnNewData(string strData);
    void OnConnect(string strLogInID,long nErrorCode);
   
private:
    HRESULT OnEventFiringObjectInvoke(
        ISKReplyLibEventHandler* pEventHandler,
        DISPID dispidMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS* pdispparams,
        VARIANT* pvarResult,
        EXCEPINFO* pexcepinfo,
        UINT* puArgErr
    );

    SKCOMLib::ISKReplyLibPtr m_pSKReplyLib;
    ISKReplyLibEventHandler* m_pSKReplyLibEventHandler;
};

