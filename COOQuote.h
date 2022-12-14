#pragma once
#include "SKCOM_reference.h"
#include "TEventHandler.h"
#include <iostream>
#include <string>
#include <typeinfo>

class COOQuote//海選
{
public:
    typedef TEventHandlerNamespace::TEventHandler<COOQuote, SKCOMLib::ISKOOQuoteLib, SKCOMLib::_ISKOOQuoteLibEvents> ISKOOQuoteLibEventHandler;
    COOQuote();
    ~COOQuote();
    long EnterMonitorLong();
    long LeaveMonitor();
    long IsConnected();
    long RequestStocks(short* psPageNo, string strStockNo);
    long RequestTicks(short* psPageNo, string strStockNo);
    long RequestProducts();
    long GetTicksLONG(long nIndex, long nPtr,SKCOMLib::SKFOREIGNTICK* pSKStock);
    long GetStockByIndexLONG(long nIndex, SKCOMLib::SKFOREIGNLONG* pSKStock);
    //事件通知
    void OnConnection(long nKind, long nCode);
    void OnNotifyQuoteLONG(long nIndex);
    void OnNotifyTicksLONG(long nPtr, long nDate, long nTime, long nClose, long nQty);
    void OnNotifyHistoryTicksLONG(long nPtr, long nDate, long nTime, long nClose, long nQty);
    void OnNotifyBest5LONG
    (
        long nBestBid1, long nBestBidQty1, long nBestBid2,
        long nBestBidQty2, long nBestBid3, long nBestBidQty3, long nBestBid4, long nBestBidQty4,
        long nBestBid5, long nBestBidQty5, long nBestAsk1, long nBestAskQty1, long nBestAsk2, long nBestAskQty2,
        long nBestAsk3, long nBestAskQty3, long nBestAsk4, long nBestAskQty4, long nBestAsk5, long nBestAskQty5
    );
    void OnProducts(string strValue);
private:
    HRESULT OnEventFiringObjectInvoke(
        ISKOOQuoteLibEventHandler* pEventHandler,
        DISPID dispidMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS* pdispparams,
        VARIANT* pvarResult,
        EXCEPINFO* pexcepinfo,
        UINT* puArgErr
    );

    SKCOMLib::ISKOOQuoteLibPtr m_pSKOOQuoteLib;
    ISKOOQuoteLibEventHandler* m_pSKOOQuoteLibEventHandler;
};

