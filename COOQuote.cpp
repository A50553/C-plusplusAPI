#include "COOQuote.h"
COOQuote::COOQuote()
{
	m_pSKOOQuoteLib.CreateInstance(__uuidof(SKCOMLib::SKOOQuoteLib));
	m_pSKOOQuoteLibEventHandler = new ISKOOQuoteLibEventHandler(*this, m_pSKOOQuoteLib, &COOQuote::OnEventFiringObjectInvoke);
}
COOQuote::~COOQuote()
{
	if (m_pSKOOQuoteLibEventHandler)
	{
		m_pSKOOQuoteLibEventHandler->ShutdownConnectionPoint();
		m_pSKOOQuoteLibEventHandler->Release();
		m_pSKOOQuoteLibEventHandler = NULL;
	}
	if (m_pSKOOQuoteLib)
	{
		m_pSKOOQuoteLib->Release();
	}
}
HRESULT COOQuote::OnEventFiringObjectInvoke
(
	ISKOOQuoteLibEventHandler* pEventHandler,
	DISPID dispidMember,
	REFIID riid,
	LCID lcid,
	WORD wFlags,
	DISPPARAMS* pdispparams,
	VARIANT* pvarResult,
	EXCEPINFO* pexcepinfo,
	UINT* puArgErr
)
{
	VARIANT varlValue;
	VariantInit(&varlValue);
	VariantClear(&varlValue);
	switch (dispidMember)
	{
		case 1:  // OnConnection
		{
			long nKind = V_I4(&(pdispparams->rgvarg)[1]);
			long nCode = V_I4(&(pdispparams->rgvarg)[0]);
			OnConnection(nKind, nCode);
			break;
		}
		case 2:  // OnProducts
		{
			_bstr_t bstrStockData = V_BSTR(&(pdispparams->rgvarg)[0]);
			OnProducts(string(bstrStockData));
			break;
		}
		case 8:  // OnNotifyQuoteLONG
		{
			long nIndex = V_I4(&(pdispparams->rgvarg)[0]);
			OnNotifyQuoteLONG(nIndex);
			break;
		}
		case 9:  // OnNotifyTicksLONG
		{
			long nPtr = V_I4(&(pdispparams->rgvarg)[4]);
			long nDate = V_I4(&(pdispparams->rgvarg)[3]);
			long nTime = V_I4(&(pdispparams->rgvarg)[2]);
			long nClose = V_I4(&(pdispparams->rgvarg)[1]);
			long nQty = V_I4(&(pdispparams->rgvarg)[0]);
			OnNotifyTicksLONG( nPtr, nDate,nTime, nClose, nQty);
			break;
		}
		case 10: // OnNotifyHistoryTicksLONG
		{
			long nPtr = V_I4(&(pdispparams->rgvarg)[4]);
			long nDate = V_I4(&(pdispparams->rgvarg)[3]);
			long nTime = V_I4(&(pdispparams->rgvarg)[2]);
			long nClose = V_I4(&(pdispparams->rgvarg)[1]);
			long nQty = V_I4(&(pdispparams->rgvarg)[0]);
			OnNotifyHistoryTicksLONG(nPtr, nDate, nTime, nClose, nQty);
			break;
		}
		case 11: // OnNotifyBest5LONG
		{
			//買價
			long nBestBid1 = V_I4(&(pdispparams->rgvarg)[19]);
			long nBestBidQty1 = V_I4(&(pdispparams->rgvarg)[18]);
			long nBestBid2 = V_I4(&(pdispparams->rgvarg)[17]);
			long nBestBidQty2 = V_I4(&(pdispparams->rgvarg)[16]);
			long nBestBid3 = V_I4(&(pdispparams->rgvarg)[15]);
			long nBestBidQty3 = V_I4(&(pdispparams->rgvarg)[14]);
			long nBestBid4 = V_I4(&(pdispparams->rgvarg)[13]);
			long nBestBidQty4 = V_I4(&(pdispparams->rgvarg)[12]);
			long nBestBid5 = V_I4(&(pdispparams->rgvarg)[11]);
			long nBestBidQty5 = V_I4(&(pdispparams->rgvarg)[10]);
			//賣價
			long nBestAsk1 = V_I4(&(pdispparams->rgvarg)[9]);
			long nBestAskQty1 = V_I4(&(pdispparams->rgvarg)[8]);
			long nBestAsk2 = V_I4(&(pdispparams->rgvarg)[7]);
			long nBestAskQty2 = V_I4(&(pdispparams->rgvarg)[6]);
			long nBestAsk3 = V_I4(&(pdispparams->rgvarg)[5]);
			long nBestAskQty3 = V_I4(&(pdispparams->rgvarg)[4]);
			long nBestAsk4 = V_I4(&(pdispparams->rgvarg)[3]);
			long nBestAskQty4 = V_I4(&(pdispparams->rgvarg)[2]);
			long nBestAsk5 = V_I4(&(pdispparams->rgvarg)[1]);
			long nBestAskQty5 = V_I4(&(pdispparams->rgvarg)[0]);
			OnNotifyBest5LONG(nBestBid1, nBestBidQty1, nBestBid2, nBestBidQty2, nBestBid3, nBestBidQty3,
				nBestBid4, nBestBidQty4, nBestBid5, nBestBidQty5, nBestAsk1, nBestAskQty1, nBestAsk2, nBestAskQty2,
				nBestAsk3, nBestAskQty3, nBestAsk4, nBestAskQty4, nBestAsk5, nBestAskQty5);
			break;
		}
	}
	return S_OK;
}
long COOQuote::EnterMonitorLong()//報價連線
{
	return m_pSKOOQuoteLib->SKOOQuoteLib_EnterMonitorLONG();
}
long COOQuote::LeaveMonitor()//報價斷線
{
	return m_pSKOOQuoteLib->SKOOQuoteLib_LeaveMonitor();
}
long COOQuote::IsConnected()//連線狀態
{
	return m_pSKOOQuoteLib->SKOOQuoteLib_IsConnected();
}
long COOQuote::RequestStocks(short* psPageNo, string strStockNo)//報價，由OnNotifyQuoteLONG通知
{
	bstr_t bstrStockNo = strStockNo.c_str();
	return m_pSKOOQuoteLib->SKOOQuoteLib_RequestStocks(psPageNo, bstrStockNo);
}
long COOQuote::RequestTicks(short* psPageNo, string strStockNo)//訂閱成交明細以及五檔，即時Tick由OnNotifyTicksLONG通知，Tick回補OnNotifyHistoryTicks通知，最佳五檔OnNotifyBest5
{
	bstr_t bstrStockNo = strStockNo.c_str();
	return m_pSKOOQuoteLib->SKOOQuoteLib_RequestTicks(psPageNo, bstrStockNo);
}
long COOQuote::RequestProducts()//取得海選商品檔，由OnProducts回報
{
	return m_pSKOOQuoteLib->SKOOQuoteLib_RequestProducts();
}
long COOQuote::GetTicksLONG(long nIndex, long nPtr, SKCOMLib::SKFOREIGNTICK* pSKStock)
{
	return m_pSKOOQuoteLib->SKOOQuoteLib_GetTickLONG(nIndex,nPtr,pSKStock);
}
long COOQuote::GetStockByIndexLONG(long nIndex, SKCOMLib::SKFOREIGNLONG* pSKStock)
{
	return m_pSKOOQuoteLib->SKOOQuoteLib_GetStockByIndexLONG(nIndex, pSKStock);
}
//事件通知
void COOQuote::OnConnection(long nKind, long nCode)
{
	switch (nKind)
	{
		case 3001:
			cout << "\n【OnConnection】連線成功\n";
			break;
		case 3002:
			cout << "\n【OnConnection】斷線\n";
			break;
		case 3003:
			cout << "\n【OnConnection】報價商品資料載入完成\n";
			break;
	}
}
void COOQuote::OnNotifyQuoteLONG(long nIndex)
{
	SKCOMLib::SKFOREIGNLONG skStock;
	long nResult = m_pSKOOQuoteLib->SKOOQuoteLib_GetStockByIndexLONG(nIndex, &skStock);
	if (nResult != 0)
		return;
	//BSTR轉bstr_t轉string
	bstr_t bstrStockNo = skStock.bstrStockNo;
	bstr_t bstrStockName = skStock.bstrStockName;
	bstr_t bstrCallPut = skStock.bstrCallPut;
	string strStockNo = (LPCSTR)bstrStockNo;
	string strStockName = (LPCSTR)bstrStockName;
	string strCallPut = (LPCSTR)bstrCallPut;
	cout << "\n【OnNotifyQuoteLONG】" << "商品代號 " << strStockNo << " " << bstrStockName<<" CALL/PUT "<< strCallPut
		<< " 買價 " << skStock.nBid << " 買量 " << skStock.nBc << " 賣價 " << skStock.nAsk << " 賣量 " << skStock.nAc
		<< " 成交價 " << skStock.nClose << " 總量 " << skStock.nTQty
		<<" 履約價 "<<skStock.nStrikePrice<<" 交易日 "<<skStock.nTradingDay;
}
void COOQuote::OnNotifyTicksLONG(long nPtr, long nDate, long nTime, long nClose, long nQty)
{
	cout << "\n【OnNotifyTicksLONG】 " << nPtr << " 交易日 " << nDate << " 時間 " << nTime
		<< " 成交價 " << nClose << " 成交量 " << nQty;
}
void COOQuote::OnNotifyHistoryTicksLONG(long nPtr, long nDate, long nTime, long nClose, long nQty)//訂閱歷史成交紀錄
{
	cout << "\n【OnNotifyHistoryTicksLONG】 " << nPtr << " 交易日 " << nDate << " 時間 " << nTime
		<< " 成交價 " << nClose << " 成交量 " << nQty;
}
void COOQuote::OnNotifyBest5LONG //最佳五檔
(
	long nBestBid1, long nBestBidQty1, long nBestBid2,
	long nBestBidQty2, long nBestBid3, long nBestBidQty3, long nBestBid4, long nBestBidQty4,
	long nBestBid5, long nBestBidQty5, long nBestAsk1, long nBestAskQty1, long nBestAsk2, long nBestAskQty2,
	long nBestAsk3, long nBestAskQty3, long nBestAsk4, long nBestAskQty4, long nBestAsk5, long nBestAskQty5
)
{
	cout << "\n【OnNotifyBest5LONG】 ";
	cout << "\n第一檔買價 " << nBestBid1 << " 買量 " << nBestBidQty1 << "   賣價 " << nBestAsk1 << " 賣量 " << nBestAskQty1;
	cout << "\n第二檔買價 " << nBestBid2 << " 買量 " << nBestBidQty2 << "   賣價 " << nBestAsk2 << " 賣量 " << nBestAskQty2;
	cout << "\n第三檔買價 " << nBestBid3 << " 買量 " << nBestBidQty3 << "   賣價 " << nBestAsk3 << " 賣量 " << nBestAskQty3;
	cout << "\n第四檔買價 " << nBestBid4 << " 買量 " << nBestBidQty4 << "   賣價 " << nBestAsk4 << " 賣量 " << nBestAskQty4;
	cout << "\n第五檔買價 " << nBestBid5 << " 買量 " << nBestBidQty5 << "   賣價 " << nBestAsk5 << " 賣量 " << nBestAskQty5;
}
void COOQuote::OnProducts(string strValue)//回報商品資訊
{
	cout << "\n【OnOverseaProducts】 " << strValue;
}