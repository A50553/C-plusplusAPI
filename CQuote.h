#pragma once
#include "SKCOM_reference.h"
#include "TEventHandler.h"
#include <iostream>
#include <string>
#include <typeinfo>
class CQuote
{
public:
    typedef TEventHandlerNamespace::TEventHandler<CQuote, SKCOMLib::ISKQuoteLib, SKCOMLib::_ISKQuoteLibEvents> ISKQuoteLibEventHandler;
    CQuote();
    ~CQuote();
    long EnterMonitorLong();
	long LeaveMonitor();
    long IsConnected();
    long RequestStocks(short*psPageNo, string strStockNo);
    long RequestTicks(short* psPageNo, string strStockNo);
    long RequestTicksWithMarketNo(short* psPageNo, short sMarketNo, string strStockNo);
    long RequestStockList(short sMarketNo);
    long GetStocksByIndexLONG(short sMarketNo, long nIndex, SKCOMLib::SKSTOCKLONG* pSKStock);
    //事件通知
    void OnConnection(long nKind,long nCode);
    void OnNotifyQuoteLONG(short sMarketNo, long nIndex);
    void OnNotifyTicksLONG(long nIndex, long nPtr, long nDate, long nTimehms, long nBid, long nAsk, long nClose, long nQty);
    void OnNotifyHistoryTicksLONG(long nIndex, long nPtr, long nDate, long nTimehms, long nBid, long nAsk, long nClose, long nQty);
    void OnNotifyStockList(short sMarketNo,string strStockData);
    void OnNotifyBest5LONG(long nBestBid1, long nBestBidQty1, long nBestBid2, long nBestBidQty2, long nBestBid3, long nBestBidQty3, long nBestBid4, long nBestBidQty4, long nBestBid5, long nBestBidQty5, long nBestAsk1, long nBestAskQty1, long nBestAsk2, long nBestAskQty2, long nBestAsk3, long nBestAskQty3, long nBestAsk4, long nBestAskQty4, long nBestAsk5, long nBestAskQty5);
    
private:
    HRESULT OnEventFiringObjectInvoke(
        ISKQuoteLibEventHandler* pEventHandler,
        DISPID dispidMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS* pdispparams,
        VARIANT* pvarResult,
        EXCEPINFO* pexcepinfo,
        UINT* puArgErr
    );

    SKCOMLib::ISKQuoteLibPtr m_pSKQuoteLib;
    ISKQuoteLibEventHandler* m_pSKQuoteLibEventHandler;
};

