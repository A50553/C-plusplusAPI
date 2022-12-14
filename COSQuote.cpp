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
			OnNotifyBest5NineDigitalLONG(nBestBid1, nBestBidQty1, nBestBid2, nBestBidQty2, nBestBid3, nBestBidQty3,
				nBestBid4, nBestBidQty4, nBestBid5, nBestBidQty5, nBestAsk1, nBestAskQty1, nBestAsk2, nBestAskQty2,
				nBestAsk3, nBestAskQty3, nBestAsk4, nBestAskQty4, nBestAsk5, nBestAskQty5);
			break;
		}
	}
	return S_OK;
}
long COSQuote::EnterMonitorLong()//�����s�u
{
	return m_pSKOSQuoteLib->SKOSQuoteLib_EnterMonitorLONG();
}
long COSQuote::LeaveMonitor()//�����_�u
{
	return m_pSKOSQuoteLib->SKOSQuoteLib_LeaveMonitor();
}
long COSQuote::IsConnected()//�s�u���A
{
	return m_pSKOSQuoteLib->SKOSQuoteLib_IsConnected();
}
long COSQuote::RequestStocks(short* psPageNo, string strStockNo)//�q�\���w�ӫ~�Y�ɳ����A��OnNotifyQuoteLONG���o�q��
{
	bstr_t bstrStockNo = strStockNo.c_str();
	return m_pSKOSQuoteLib->SKOSQuoteLib_RequestStocks(psPageNo, bstrStockNo);
}
long COSQuote::RequestTicks(short* psPageNo, string strStockNo)//�q�\������ӥH�Τ��ɡA�Y��Tick��OnNotifyTicksLONG�q���ATick�^��OnNotifyHistoryTicks�q���A�̨Τ���OnNotifyBest5
{
	bstr_t bstrStockNo = strStockNo.c_str();
	return m_pSKOSQuoteLib->SKOSQuoteLib_RequestTicks(psPageNo, bstrStockNo);
}
long COSQuote::RequestOverseaProducts()//���o���~�ӫ~�� ��OnOverseaProducts ���o�^��
{
	return m_pSKOSQuoteLib->SKOSQuoteLib_RequestOverseaProducts();
}
long COSQuote::GetOverseaProductDetail(short sType)//���o�ӫ~�ɡA��OnOverseaProductsDetail�^��
{
	return m_pSKOSQuoteLib->SKOSQuoteLib_GetOverseaProductDetail(sType);
}
//�ƥ�q��
void COSQuote::OnConnection(long nKind, long nCode)
{
	switch (nKind)
	{
		case 3001:
		{
			cout << "\n�iOnConnection�j�s�u���\\n";
			break;
		}
		case 3002:
		{
			cout << "\n�iOnConnection�j�_�u\n";
			break;
		}
		case 3003:
		{
			cout << "\n�iOnConnection�j�����ӫ~��Ƹ��J����\n";
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
	cout << "\n�iOnNotifyQuoteLONG�j" <<"�ӫ~�N�� " << strStockNo << " " << bstrStockName
		<< " �R�� " << skStock.nBid <<" �R�q "<<skStock.nBc << " ��� " << skStock.nAsk <<" ��q "<<skStock.nAc 
		<< " ����� " << skStock.nClose << " �`�q " << skStock.nTQty;
}
void COSQuote::OnNotifyTicksNineDigitalLONG(long nPtr, long nDate, long nTime, long nClose, long nQty)//�ή�tick�q��
{
	cout << "\n�iOnNotifyTicksNineDigitalLONG�j " << nPtr << " ����� " << nDate << " �ɶ� " << nTime
		<< " ����� " << nClose << " ����q " << nQty;
}
void COSQuote::OnNotifyHistoryTicksNineDigitalLONG(long nPtr, long nDate, long nTime, long nClose, long nQty)//Tick�^��
{
	cout << "\n�iOnNotifyHistoryTicksNineDigitalLONG�j " << nPtr << " ����� " << nDate << " �ɶ� " << nTime
		<< " ����� " << nClose << " ����q " << nQty;
}
void COSQuote::OnNotifyBest5NineDigitalLONG// �̨Τ���
(
	long nBestBid1, long nBestBidQty1, long nBestBid2,
	long nBestBidQty2, long nBestBid3, long nBestBidQty3, long nBestBid4, long nBestBidQty4,
	long nBestBid5, long nBestBidQty5, long nBestAsk1, long nBestAskQty1, long nBestAsk2, long nBestAskQty2,
	long nBestAsk3, long nBestAskQty3, long nBestAsk4, long nBestAskQty4, long nBestAsk5, long nBestAskQty5
)
{
	cout << "\n�iOnNotifyBest5NineDigitalLONG�j ";
	cout << "\n�Ĥ@�ɶR�� " << nBestBid1 << " �R�q " << nBestBidQty1 << "   ��� " << nBestAsk1 << " ��q " << nBestAskQty1;
	cout << "\n�ĤG�ɶR�� " << nBestBid2 << " �R�q " << nBestBidQty2 << "   ��� " << nBestAsk2 << " ��q " << nBestAskQty2;
	cout << "\n�ĤT�ɶR�� " << nBestBid3 << " �R�q " << nBestBidQty3 << "   ��� " << nBestAsk3 << " ��q " << nBestAskQty3;
	cout << "\n�ĥ|�ɶR�� " << nBestBid4 << " �R�q " << nBestBidQty4 << "   ��� " << nBestAsk4 << " ��q " << nBestAskQty4;
	cout << "\n�Ĥ��ɶR�� " << nBestBid5 << " �R�q " << nBestBidQty5 << "   ��� " << nBestAsk5 << " ��q " << nBestAskQty5;
}
void COSQuote::OnOverseaProductsDetail(string strValue)//�^���ԲӰӫ~��T
{
	cout << "\n�iOnOverseaProductsDetail�j " << strValue;
}
void COSQuote::OnOverseaProducts(string strValue)//�^���ӫ~��T
{
	cout << "\n�iOnOverseaProducts�j "<<strValue;
}
