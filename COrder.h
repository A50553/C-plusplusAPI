#pragma once
#include <iostream>
#include <string.h>
#include <stdio.h>
#include <stdlib.h>
#include "SKCOM_reference.h"
#include "TEventHandler.h"
#include <vector>
using namespace std;

class COrder
{
public:
	typedef TEventHandlerNamespace::TEventHandler<COrder, SKCOMLib::ISKOrderLib, SKCOMLib::_ISKOrderLibEvents>ISKOrderLibEventHandler;
	COrder();
	~COrder();
	long Initialize();
	long SentStockOrder(string strLogInID, bool bAsyncOrder,string strStockNo,short sPrime,short sPeriod,short sFlag,long nTradeType,long nSpecialTradeType,short sBuySell,long nQty,string strPrice);
	long SentFutureOrder(string strLogInID, bool bAsyncOrder, string strStockNo, short sTradeType, short sDayTrade,short sNewClose, short sBuySell, long nQty, string strPrice, short sReserved);
	long SentOptionOrder(string strLogInID, bool bAsyncOrder, string strStockNo, short sTradeType, short sDayTrade, short sNewClose, short sBuySell, long nQty, string strPrice, short sReserved);
	long SentOverSeaFutureOrder(string strLogInID, bool bAsyncOrder, string strExchangeNo, string strStockNo, string  strYearMonth, short sTradeType, string  strOrder, string strOrderNumbertor, string strTrigger, string strTriggerNumbertor, short sDayTrade, short sNewClose,short sBuySell,int nQty,short sSpecialTradeType);
	long SentOverSeaOptionOrder(string strLogInID, bool bAsyncOrder, string strExchangeNo, string strStockNo, string  strYearMonth, short sTradeType, string  strOrder, string strOrderNumbertor, short sDayTrade, short sNewClose, short sCallPut, short sBuySell, string strStrikePrice, int nQty, short sSpecialTradeType);
	
	long GetUserAccount();
	long ReadCertByID(string strLogInID);//string strLogInID
	long SetMaxQty(long nMarkettype, long nMaxQty);
	void OnAsyncOrder(long nThreadID,long nCode, string bstrMessage);
	void OnAccount(string strLogInID, string strAccountData);
	long DecreaseOrder(string strLogInID, bool bAsyncOrder, string strNo, long nQty,int nMarket);
	long CorrectPrice(string strLogInID, bool bAsyncOrder, string strNo, string strPrice, long nTradeType, int nSeqOrBook, int nMarket);
	long CancelOrder(string strLogInID, bool bAsyncOrder, string strNo, long nTradeType, int nSeqOrBook,int nMarket);

	long OverSeaDecreaseOrder(string strLogInID, bool bAsyncOrder, string strNo, long nQty, int nMarket);
	long OverSeaCancelOrder(string strLogInID, bool bAsyncOrder, string strNo, long nTradeType, int nSeqOrBookOrNum, int nMarket);
	long OverSeaCorrectPrice(string strLogInID, bool bAsyncOrder, string strNo, string strExchangeNo, string strStockNo, string strYearMonth, string strPrice, string strOrderNumbertor, string strOrderDenominator, short sTradeType, short sSpecialTradeType, short sNewClose, short sCallPut, string strStrikePrice, int nSeqOrBook, int nMarket);
	long LoadOSCommodity();
	long LoadOOCommodity();
private:
	HRESULT OnEventFiringObjectInvoke
	(
		ISKOrderLibEventHandler* pEventHandler,
		DISPID dispidMember,
		REFIID riid,
		LCID lcid,
		WORD wFlags,
		DISPPARAMS* pdispparams,
		VARIANT* pvarResult,
		EXCEPINFO* pexcepinfo,
		UINT* puArgErr
	);
	SKCOMLib::ISKOrderLibPtr m_pSKOrderLib;
	ISKOrderLibEventHandler* m_pSKOrderLibEventHandler;
	vector<string> vec_strFullAccount_TS;//證券
	vector<string> vec_strFullAccount_TF;//期貨、選擇權
	vector<string>vec_strFullAccount_OF;//海外期貨、海外選擇權
};

