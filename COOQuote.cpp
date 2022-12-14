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
			//�R��
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
			//���
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
long COOQuote::EnterMonitorLong()//�����s�u
{
	return m_pSKOOQuoteLib->SKOOQuoteLib_EnterMonitorLONG();
}
long COOQuote::LeaveMonitor()//�����_�u
{
	return m_pSKOOQuoteLib->SKOOQuoteLib_LeaveMonitor();
}
long COOQuote::IsConnected()//�s�u���A
{
	return m_pSKOOQuoteLib->SKOOQuoteLib_IsConnected();
}
long COOQuote::RequestStocks(short* psPageNo, string strStockNo)//�����A��OnNotifyQuoteLONG�q��
{
	bstr_t bstrStockNo = strStockNo.c_str();
	return m_pSKOOQuoteLib->SKOOQuoteLib_RequestStocks(psPageNo, bstrStockNo);
}
long COOQuote::RequestTicks(short* psPageNo, string strStockNo)//�q�\������ӥH�Τ��ɡA�Y��Tick��OnNotifyTicksLONG�q���ATick�^��OnNotifyHistoryTicks�q���A�̨Τ���OnNotifyBest5
{
	bstr_t bstrStockNo = strStockNo.c_str();
	return m_pSKOOQuoteLib->SKOOQuoteLib_RequestTicks(psPageNo, bstrStockNo);
}
long COOQuote::RequestProducts()//���o����ӫ~�ɡA��OnProducts�^��
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
//�ƥ�q��
void COOQuote::OnConnection(long nKind, long nCode)
{
	switch (nKind)
	{
		case 3001:
			cout << "\n�iOnConnection�j�s�u���\\n";
			break;
		case 3002:
			cout << "\n�iOnConnection�j�_�u\n";
			break;
		case 3003:
			cout << "\n�iOnConnection�j�����ӫ~��Ƹ��J����\n";
			break;
	}
}
void COOQuote::OnNotifyQuoteLONG(long nIndex)
{
	SKCOMLib::SKFOREIGNLONG skStock;
	long nResult = m_pSKOOQuoteLib->SKOOQuoteLib_GetStockByIndexLONG(nIndex, &skStock);
	if (nResult != 0)
		return;
	//BSTR��bstr_t��string
	bstr_t bstrStockNo = skStock.bstrStockNo;
	bstr_t bstrStockName = skStock.bstrStockName;
	bstr_t bstrCallPut = skStock.bstrCallPut;
	string strStockNo = (LPCSTR)bstrStockNo;
	string strStockName = (LPCSTR)bstrStockName;
	string strCallPut = (LPCSTR)bstrCallPut;
	cout << "\n�iOnNotifyQuoteLONG�j" << "�ӫ~�N�� " << strStockNo << " " << bstrStockName<<" CALL/PUT "<< strCallPut
		<< " �R�� " << skStock.nBid << " �R�q " << skStock.nBc << " ��� " << skStock.nAsk << " ��q " << skStock.nAc
		<< " ����� " << skStock.nClose << " �`�q " << skStock.nTQty
		<<" �i���� "<<skStock.nStrikePrice<<" ����� "<<skStock.nTradingDay;
}
void COOQuote::OnNotifyTicksLONG(long nPtr, long nDate, long nTime, long nClose, long nQty)
{
	cout << "\n�iOnNotifyTicksLONG�j " << nPtr << " ����� " << nDate << " �ɶ� " << nTime
		<< " ����� " << nClose << " ����q " << nQty;
}
void COOQuote::OnNotifyHistoryTicksLONG(long nPtr, long nDate, long nTime, long nClose, long nQty)//�q�\���v�������
{
	cout << "\n�iOnNotifyHistoryTicksLONG�j " << nPtr << " ����� " << nDate << " �ɶ� " << nTime
		<< " ����� " << nClose << " ����q " << nQty;
}
void COOQuote::OnNotifyBest5LONG //�̨Τ���
(
	long nBestBid1, long nBestBidQty1, long nBestBid2,
	long nBestBidQty2, long nBestBid3, long nBestBidQty3, long nBestBid4, long nBestBidQty4,
	long nBestBid5, long nBestBidQty5, long nBestAsk1, long nBestAskQty1, long nBestAsk2, long nBestAskQty2,
	long nBestAsk3, long nBestAskQty3, long nBestAsk4, long nBestAskQty4, long nBestAsk5, long nBestAskQty5
)
{
	cout << "\n�iOnNotifyBest5LONG�j ";
	cout << "\n�Ĥ@�ɶR�� " << nBestBid1 << " �R�q " << nBestBidQty1 << "   ��� " << nBestAsk1 << " ��q " << nBestAskQty1;
	cout << "\n�ĤG�ɶR�� " << nBestBid2 << " �R�q " << nBestBidQty2 << "   ��� " << nBestAsk2 << " ��q " << nBestAskQty2;
	cout << "\n�ĤT�ɶR�� " << nBestBid3 << " �R�q " << nBestBidQty3 << "   ��� " << nBestAsk3 << " ��q " << nBestAskQty3;
	cout << "\n�ĥ|�ɶR�� " << nBestBid4 << " �R�q " << nBestBidQty4 << "   ��� " << nBestAsk4 << " ��q " << nBestAskQty4;
	cout << "\n�Ĥ��ɶR�� " << nBestBid5 << " �R�q " << nBestBidQty5 << "   ��� " << nBestAsk5 << " ��q " << nBestAskQty5;
}
void COOQuote::OnProducts(string strValue)//�^���ӫ~��T
{
	cout << "\n�iOnOverseaProducts�j " << strValue;
}