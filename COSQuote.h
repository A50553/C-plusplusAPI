#pragma once
#include "SKCOM_reference.h"
#include "TEventHandler.h"
#include <iostream>
#include <string>
#include <typeinfo>
#include <comutil.h>
#include <stdio.h>
class COSQuote//海期報價
{
public:
    typedef TEventHandlerNamespace::TEventHandler<COSQuote, SKCOMLib::ISKOSQuoteLib, SKCOMLib::_ISKOSQuoteLibEvents> ISKOSQuoteLibEventHandler;
    COSQuote();
    ~COSQuote();
    long EnterMonitorLong();
    long LeaveMonitor();
    long IsConnected();
    long RequestStocks(short* psPageNo, string strStockNo);
    long RequestTicks(short* psPageNo, string strStockNo);
    long RequestOverseaProducts();
    long GetOverseaProductDetail(short sType);

    //事件通知
    void OnConnection(long nKind, long nCode);
    void OnNotifyQuoteLONG(long nIndex);
    void OnNotifyTicksNineDigitalLONG(long nPtr, long nDate, long nTime, long nClose, long nQty);
    void OnNotifyHistoryTicksNineDigitalLONG(long nPtr, long nDate, long nTime, long nClose, long nQty);
    void OnNotifyBest5NineDigitalLONG
    (
        long nBestBid1, long nBestBidQty1, long nBestBid2,
        long nBestBidQty2, long nBestBid3, long nBestBidQty3, long nBestBid4, long nBestBidQty4,
        long nBestBid5, long nBestBidQty5, long nBestAsk1, long nBestAskQty1, long nBestAsk2, long nBestAskQty2,
        long nBestAsk3, long nBestAskQty3, long nBestAsk4, long nBestAskQty4, long nBestAsk5, long nBestAskQty5
    );
    void OnOverseaProductsDetail(string strValue);
    void OnOverseaProducts(string strValue);
private:
    HRESULT OnEventFiringObjectInvoke(
        ISKOSQuoteLibEventHandler* pEventHandler,
        DISPID dispidMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS* pdispparams,
        VARIANT* pvarResult,
        EXCEPINFO* pexcepinfo,
        UINT* puArgErr
    );
    SKCOMLib::ISKOSQuoteLibPtr m_pSKOSQuoteLib;
    ISKOSQuoteLibEventHandler* m_pSKOSQuoteLibEventHandler;
};

