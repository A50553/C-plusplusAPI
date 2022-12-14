#include "CQuote.h"
CQuote::CQuote()
{
	m_pSKQuoteLib.CreateInstance(__uuidof(SKCOMLib::SKQuoteLib));
	m_pSKQuoteLibEventHandler = new ISKQuoteLibEventHandler(*this, m_pSKQuoteLib, &CQuote::OnEventFiringObjectInvoke);
}
CQuote::~CQuote()
{
	if (m_pSKQuoteLibEventHandler)
	{
		m_pSKQuoteLibEventHandler->ShutdownConnectionPoint();
		m_pSKQuoteLibEventHandler->Release();
		m_pSKQuoteLibEventHandler = NULL;
	}
	if (m_pSKQuoteLib)
	{
		m_pSKQuoteLib->Release();
	}
}
HRESULT CQuote::OnEventFiringObjectInvoke
(
	ISKQuoteLibEventHandler* pEventHandler,
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
		case 16: // RequestStockList
		{
			short sMarketNo = V_I2(&(pdispparams->rgvarg)[1]);
			_bstr_t bstrStockData = V_BSTR(&(pdispparams->rgvarg)[0]);
			OnNotifyStockList(sMarketNo, string(bstrStockData));
			break;
		}
		case 19: // OnNotifyQuoteLONG
		{
			short sMarketNo = V_I2(&(pdispparams->rgvarg)[1]);
			long nStockIndex = V_I4(&(pdispparams->rgvarg)[0]);
			OnNotifyQuoteLONG(sMarketNo, nStockIndex);
			break;
		}
		case 20: // OnNotifyHistoryTicksLONG
		{
			long nStockIndex = V_I4(&(pdispparams->rgvarg)[9]);
			long nPtr = V_I4(&(pdispparams->rgvarg)[8]);
			long nDate = V_I4(&(pdispparams->rgvarg)[7]);
			long lTimehms = V_I4(&(pdispparams->rgvarg)[6]);
			long nBid = V_I4(&(pdispparams->rgvarg)[4]);
			long nAsk = V_I4(&(pdispparams->rgvarg)[3]);
			long nClose = V_I4(&(pdispparams->rgvarg)[2]);
			long nQty = V_I4(&(pdispparams->rgvarg)[1]);
			OnNotifyHistoryTicksLONG(nStockIndex, nPtr, nDate, lTimehms, nBid, nAsk, nClose, nQty);
			break;
		}
		case 21: // OnNotifyTicksLONG
		{
			long nStockIndex = V_I4(&(pdispparams->rgvarg)[9]);
			long nPtr = V_I4(&(pdispparams->rgvarg)[8]);
			long nDate = V_I4(&(pdispparams->rgvarg)[7]);
			long lTimehms = V_I4(&(pdispparams->rgvarg)[6]);
			long nBid = V_I4(&(pdispparams->rgvarg)[4]);
			long nAsk = V_I4(&(pdispparams->rgvarg)[3]);
			long nClose = V_I4(&(pdispparams->rgvarg)[2]);
			long nQty = V_I4(&(pdispparams->rgvarg)[1]);
			OnNotifyTicksLONG(nStockIndex, nPtr, nDate, lTimehms, nBid, nAsk, nClose, nQty);
			break;
		}
		case 22: // OnNotifyBest5LONG
		{
			//禦基
			long nBestBid1 = V_I4(&(pdispparams->rgvarg)[24]);
			long nBestBidQty1 = V_I4(&(pdispparams->rgvarg)[23]);
			long nBestBid2 = V_I4(&(pdispparams->rgvarg)[22]);
			long nBestBidQty2 = V_I4(&(pdispparams->rgvarg)[21]);
			long nBestBid3 = V_I4(&(pdispparams->rgvarg)[20]);
			long nBestBidQty3 = V_I4(&(pdispparams->rgvarg)[19]);
			long nBestBid4 = V_I4(&(pdispparams->rgvarg)[18]);
			long nBestBidQty4 = V_I4(&(pdispparams->rgvarg)[17]);
			long nBestBid5 = V_I4(&(pdispparams->rgvarg)[16]);
			long nBestBidQty5 = V_I4(&(pdispparams->rgvarg)[15]);
			//芥基
			long nBestAsk1 = V_I4(&(pdispparams->rgvarg)[12]);
			long nBestAskQty1 = V_I4(&(pdispparams->rgvarg)[11]);
			long nBestAsk2 = V_I4(&(pdispparams->rgvarg)[10]);
			long nBestAskQty2 = V_I4(&(pdispparams->rgvarg)[9]);
			long nBestAsk3 = V_I4(&(pdispparams->rgvarg)[8]);
			long nBestAskQty3 = V_I4(&(pdispparams->rgvarg)[7]);
			long nBestAsk4 = V_I4(&(pdispparams->rgvarg)[6]);
			long nBestAskQty4 = V_I4(&(pdispparams->rgvarg)[5]);
			long nBestAsk5 = V_I4(&(pdispparams->rgvarg)[4]);
			long nBestAskQty5 = V_I4(&(pdispparams->rgvarg)[3]);
			OnNotifyBest5LONG(nBestBid1, nBestBidQty1, nBestBid2, nBestBidQty2, nBestBid3, nBestBidQty3, nBestBid4, nBestBidQty4, nBestBid5, nBestBidQty5, nBestAsk1, nBestAskQty1, nBestAsk2, nBestAskQty2, nBestAsk3, nBestAskQty3, nBestAsk4, nBestAskQty4, nBestAsk5, nBestAskQty5);
			break;
		}
	}
	return S_OK;
}
long CQuote::EnterMonitorLong()//厨基硈絬
{
	return m_pSKQuoteLib->SKQuoteLib_EnterMonitorLONG();
}
long CQuote::LeaveMonitor()//厨基耞絬
{
	return m_pSKQuoteLib->SKQuoteLib_LeaveMonitor();
}
long CQuote::IsConnected()//硈絬篈
{
	return m_pSKQuoteLib->SKQuoteLib_IsConnected();
}
long CQuote::RequestStocks(short* psPageNo,string strStockNo)//璹綷﹚坝珇厨基パOnNotifyQuoteLONG硄
{
	bstr_t bstrStockNo = strStockNo.c_str();
	return m_pSKQuoteLib->SKQuoteLib_RequestStocks(psPageNo,bstrStockNo);
}
long CQuote::RequestTicks(short* psPageNo, string strStockNo)//璹綷Θユ灿のき郎TickパOnNotifyTicksLONG硄Tick干OnNotifyHistoryTicks硄程ㄎき郎OnNotifyBest5
{
	bstr_t bstrStockNo = strStockNo.c_str();
	return m_pSKQuoteLib->SKQuoteLib_RequestTicks(psPageNo,bstrStockNo);
}
long CQuote::RequestTicksWithMarketNo(short* psPageNo, short sMarketNo, string strStockNo)//璹綷き郎-絃い箂TickパOnNotifyTicksLONG硄Tick干OnNotifyHistoryTicks硄程ㄎき郎OnNotifyBest5
{
	bstr_t bstrStockNo = strStockNo.c_str();
	return m_pSKQuoteLib->SKQuoteLib_RequestTicksWithMarketNo(psPageNo,sMarketNo,bstrStockNo);
}
long CQuote::RequestStockList(short sMarketNo)//琩高瓣ず坝珇膀セ戈闽戈癟パOnNotifyStockList硄
{
	return m_pSKQuoteLib->SKQuoteLib_RequestStockList(sMarketNo);
}
long CQuote::GetStocksByIndexLONG(short sMarketNo,long nIndex, SKCOMLib::SKSTOCKLONG* pSKStock)
{
	return m_pSKQuoteLib->SKQuoteLib_GetStockByIndexLONG(sMarketNo,nIndex,pSKStock);
}
//ㄆン硄
void CQuote::OnConnection(long nKind, long nCode)
{
	switch(nKind)
	{
		case 3001:
			cout << "\nOnConnection硈絬Θ\n";
			break;
		
		case 3002:
			cout << "\nOnConnection耞絬\n";
			break;
		case 3003:
			cout << "\nOnConnection厨基坝珇戈更ЧΘ\n";
			break;
		default:
			break;
	}
}
void CQuote::OnNotifyQuoteLONG(short sMarketNo,long nIndex)//Τ厨基钵笆パ硄
{
	SKCOMLib::SKSTOCKLONG skStock;
	long nResult = GetStocksByIndexLONG(sMarketNo, nIndex, &skStock);
	if (nResult != 0)
		return;

	char* szStockNo = _com_util::ConvertBSTRToString(skStock.bstrStockNo);
	char* szStockName = _com_util::ConvertBSTRToString(skStock.bstrStockName);
	printf("\nOnNotifyQuoteLONG%s %s bid:%d ask:%d last:%d volume:%d\n",
		szStockNo,//坝珇絏
		szStockName,//坝珇嘿
		skStock.nBid,//禦基
		skStock.nAsk,//芥基
		skStock.nClose,//Θユ基
		skStock.nTQty//羆秖
	);
	delete[] szStockName;
	delete[] szStockNo;
}
void CQuote::OnNotifyTicksLONG( long nIndex, long nPtr, long nDate, long nTimehms, long nBid, long nAsk, long nClose, long nQty)
{
	cout<< "\nOnNotifyTicksLONG " << nPtr << " ユら " << nDate << " 丁 " << nTimehms 
		<< " 禦基 " << nBid << " 芥基 " << nAsk 
		<< " Θユ基 " << nClose << " Θユ秖 " << nQty;
	printf("\n=======================================================\n");
}
void CQuote::OnNotifyHistoryTicksLONG( long nIndex, long nPtr, long nDate, long nTimehms, long nBid, long nAsk, long nClose, long nQty)
{
	cout << "\nOnNotifyHistoryTicksLONG "<< nPtr << " ユら " << nDate << ", 丁 " << nTimehms 
		<< " 禦基 " << nBid << " 芥基 " << nAsk
		<< " Θユ基 " << nClose << " Θユ秖 " << nQty;
}
void CQuote::OnNotifyStockList(short sMarketNo, string strStockData)//肚﹚瓣ずカ初摸坝珇睲虫
{
	string tempstr = "";
    for (int i = 0; i < strStockData.length(); i++)
    {
        if (strStockData[i] == ';')
        {
            cout << "\nOnNotifyStockList " << tempstr;
            if (tempstr == "##,,")
            {
                break;
            }
            tempstr = "";
            continue;
        }
        tempstr += strStockData[i];
    }
    cout << endl;
}
void CQuote::OnNotifyBest5LONG
(
	long nBestBid1, long nBestBidQty1,
	long nBestBid2, long nBestBidQty2,
	long nBestBid3, long nBestBidQty3,
	long nBestBid4, long nBestBidQty4,
	long nBestBid5, long nBestBidQty5,
	long nBestAsk1, long nBestAskQty1,
	long nBestAsk2, long nBestAskQty2,
	long nBestAsk3, long nBestAskQty3,
	long nBestAsk4, long nBestAskQty4,
	long nBestAsk5, long nBestAskQty5
)
{
	printf("\n\nOnNotifyBest5LONG\n");
	printf("禦基1%ld, 计秖1%ld\n", nBestBid1, nBestBidQty1);
	printf("禦基2%ld, 计秖2%ld\n", nBestBid2, nBestBidQty2);
	printf("禦基3%ld, 计秖3%ld\n", nBestBid3, nBestBidQty3);
	printf("禦基4%ld, 计秖4%ld\n", nBestBid4, nBestBidQty4);
	printf("禦基5%ld, 计秖5%ld\n\n", nBestBid5, nBestBidQty5);

	printf("芥基1%ld,计秖1%ld\n", nBestAsk1, nBestAskQty1);
	printf("芥基2%ld,计秖2%ld\n", nBestAsk2, nBestAskQty2);
	printf("芥基3%ld,计秖3%ld\n", nBestAsk3, nBestAskQty3);
	printf("芥基4%ld,计秖4%ld\n", nBestAsk4, nBestAskQty4);
	printf("芥基5%ld,计秖5%ld\n\n", nBestAsk5, nBestAskQty5);
}