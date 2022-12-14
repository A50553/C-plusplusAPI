#include "COSQuote.h"
COSQuote::COSQuote()
{
	m_pSKOSQuoteLib.CreateInstance(__uuidof(SKCOMLib::SKOSQuoteLib));
	m_pSKOSQuoteLibEventHandler = new ISKOSQuoteLibEventHandler(*this, m_pSKOSQuoteLib, &COSQuote::OnEventFiringObjectInvoke);
}
COSQuote::~COSQuote()
{
	if (m_pSKOSQuoteLibEventHandler)
	{
		m_pSKOSQuoteLibEventHandler->ShutdownConnectionPoint();
		m_pSKOSQuoteLibEventHandler->Release();
		m_pSKOSQuoteLibEventHandler = NULL;
	}
	if (m_pSKOSQuoteLib)
	{
		m_pSKOSQuoteLib->Release();
	}
}
HRESULT COSQuote::OnEventFiringObjectInvoke
(
	ISKOSQuoteLibEventHandler* pEventHandler,
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
		case 5:  //OnOverseaProducts
		{
			bstr_t bstrValue= V_BSTR(&(pdispparams->rgvarg)[0]);
			OnOverseaProducts(string(bstrValue));
			break;
		}
		case 9:  //OnOverseaProductsDetail
		{
			bstr_t bstrValue = V_BSTR(&(pdispparams->rgvarg)[0]);
			OnOverseaProductsDetail(string(bstrValue));
		}
		case 15: //OnNotifyQuoteLONG
		{
			long nIndex = V_I4(&(pdispparams->rgvarg)[0]);
			OnNotifyQuoteLONG(nIndex);
			break;
		}
		case 16: //OnNotifyTicksDigitalLONG
		{
			long nPtr = V_I4(&(pdispparams->rgvarg)[4]);
			long nDate = V_I4(&(pdispparams->rgvarg)[3]);
			long nTime = V_I4(&(pdispparams->rgvarg)[2]);
			long nClose = V_I4(&(pdispparams->rgvarg)[1]);
			long nQty = V_I4(&(pdispparams->rgvarg)[0]);
			OnNotifyTicksNineDigitalLONG(nPtr, nDate, nTime, nClose, nQty);
			break;
		}
		case 17: // OnNotifyHistoryTicksDigitalLONG
		{
			long nPtr = V_I4(&(pdispparams->rgvarg)[4]);
			long nDate = V_I4(&(pdispparams->rgvarg)[3]);
			long nTime = V_I4(&(pdispparams->rgvarg)[2]);
			long nClose = V_I4(&(pdispparams->rgvarg)[1]);
			long nQty = V_I4(&(pdispparams->rgvarg)[0]);
			OnNotifyHistoryTicksNineDigitalLONG(nPtr, nDate, nTime, nClose, nQty);
			break;
		}
		case 18: // OnNotifyBest5NineDigitalLONG
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
			OnNotifyBest5NineDigitalLONG(nBestBid1, nBestBidQty1, nBestBid2, nBestBidQty2, nBestBid3, nBestBidQty3,
				nBestBid4, nBestBidQty4, nBestBid5, nBestBidQty5, nBestAsk1, nBestAskQty1, nBestAsk2, nBestAskQty2,
				nBestAsk3, nBestAskQty3, nBestAsk4, nBestAskQty4, nBestAsk5, nBestAskQty5);
			break;
		}
	}
	return S_OK;
}
long COSQuote::EnterMonitorLong()//報價連線
{
	return m_pSKOSQuoteLib->SKOSQuoteLib_EnterMonitorLONG();
}
long COSQuote::LeaveMonitor()//報價斷線
{
	return m_pSKOSQuoteLib->SKOSQuoteLib_LeaveMonitor();
}
long COSQuote::IsConnected()//連線狀態
{
	return m_pSKOSQuoteLib->SKOSQuoteLib_IsConnected();
}
long COSQuote::RequestStocks(short* psPageNo, string strStockNo)//訂閱指定商品即時報價，由OnNotifyQuoteLONG取得通知
{
	bstr_t bstrStockNo = strStockNo.c_str();
	return m_pSKOSQuoteLib->SKOSQuoteLib_RequestStocks(psPageNo, bstrStockNo);
}
long COSQuote::RequestTicks(short* psPageNo, string strStockNo)//訂閱成交明細以及五檔，即時Tick由OnNotifyTicksLONG通知，Tick回補OnNotifyHistoryTicks通知，最佳五檔OnNotifyBest5
{
	bstr_t bstrStockNo = strStockNo.c_str();
	return m_pSKOSQuoteLib->SKOSQuoteLib_RequestTicks(psPageNo, bstrStockNo);
}
long COSQuote::RequestOverseaProducts()//取得海外商品檔 由OnOverseaProducts 取得回報
{
	return m_pSKOSQuoteLib->SKOSQuoteLib_RequestOverseaProducts();
}
long COSQuote::GetOverseaProductDetail(short sType)//取得商品檔，由OnOverseaProductsDetail回報
{
	return m_pSKOSQuoteLib->SKOSQuoteLib_GetOverseaProductDetail(sType);
}
//事件通知
void COSQuote::OnConnection(long nKind, long nCode)
{
	switch (nKind)
	{
		case 3001:
		{
			cout << "\n【OnConnection】連線成功\n";
			break;
		}
		case 3002:
		{
			cout << "\n【OnConnection】斷線\n";
			break;
		}
		case 3003:
		{
			cout << "\n【OnConnection】報價商品資料載入完成\n";
			break;
		}
	}
}
void COSQuote::OnNotifyQuoteLONG(long nIndex)
{
	SKCOMLib::SKFOREIGNLONG skStock;
	long nResult = m_pSKOSQuoteLib->SKOSQuoteLib_GetStockByIndexLONG( nIndex, &skStock);
	if (nResult != 0)
		return;

	bstr_t bstrStockNo=skStock.bstrStockNo;
	bstr_t bstrStockName = skStock.bstrStockName;
	string strStockNo = (LPCSTR)bstrStockNo;
	string strStockName = (LPCSTR)bstrStockName;
	cout << "\n【OnNotifyQuoteLONG】" <<"商品代號 " << strStockNo << " " << bstrStockName
		<< " 買價 " << skStock.nBid <<" 買量 "<<skStock.nBc << " 賣價 " << skStock.nAsk <<" 賣量 "<<skStock.nAc 
		<< " 成交價 " << skStock.nClose << " 總量 " << skStock.nTQty;
}
void COSQuote::OnNotifyTicksNineDigitalLONG(long nPtr, long nDate, long nTime, long nClose, long nQty)//及時tick通知
{
	cout << "\n【OnNotifyTicksNineDigitalLONG】 " << nPtr << " 交易日 " << nDate << " 時間 " << nTime
		<< " 成交價 " << nClose << " 成交量 " << nQty;
}
void COSQuote::OnNotifyHistoryTicksNineDigitalLONG(long nPtr, long nDate, long nTime, long nClose, long nQty)//Tick回補
{
	cout << "\n【OnNotifyHistoryTicksNineDigitalLONG】 " << nPtr << " 交易日 " << nDate << " 時間 " << nTime
		<< " 成交價 " << nClose << " 成交量 " << nQty;
}
void COSQuote::OnNotifyBest5NineDigitalLONG// 最佳五檔
(
	long nBestBid1, long nBestBidQty1, long nBestBid2,
	long nBestBidQty2, long nBestBid3, long nBestBidQty3, long nBestBid4, long nBestBidQty4,
	long nBestBid5, long nBestBidQty5, long nBestAsk1, long nBestAskQty1, long nBestAsk2, long nBestAskQty2,
	long nBestAsk3, long nBestAskQty3, long nBestAsk4, long nBestAskQty4, long nBestAsk5, long nBestAskQty5
)
{
	cout << "\n【OnNotifyBest5NineDigitalLONG】 ";
	cout << "\n第一檔買價 " << nBestBid1 << " 買量 " << nBestBidQty1 << "   賣價 " << nBestAsk1 << " 賣量 " << nBestAskQty1;
	cout << "\n第二檔買價 " << nBestBid2 << " 買量 " << nBestBidQty2 << "   賣價 " << nBestAsk2 << " 賣量 " << nBestAskQty2;
	cout << "\n第三檔買價 " << nBestBid3 << " 買量 " << nBestBidQty3 << "   賣價 " << nBestAsk3 << " 賣量 " << nBestAskQty3;
	cout << "\n第四檔買價 " << nBestBid4 << " 買量 " << nBestBidQty4 << "   賣價 " << nBestAsk4 << " 賣量 " << nBestAskQty4;
	cout << "\n第五檔買價 " << nBestBid5 << " 買量 " << nBestBidQty5 << "   賣價 " << nBestAsk5 << " 賣量 " << nBestAskQty5;
}
void COSQuote::OnOverseaProductsDetail(string strValue)//回報詳細商品資訊
{
	cout << "\n【OnOverseaProductsDetail】 " << strValue;
}
void COSQuote::OnOverseaProducts(string strValue)//回報商品資訊
{
	cout << "\n【OnOverseaProducts】 "<<strValue;
}
