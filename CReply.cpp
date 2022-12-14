#include "CReply.h"
CReply::CReply()
{
	m_pSKReplyLib.CreateInstance(__uuidof(SKCOMLib::SKReplyLib));
	m_pSKReplyLibEventHandler = new ISKReplyLibEventHandler(*this, m_pSKReplyLib, &CReply::OnEventFiringObjectInvoke);
}
CReply::~CReply()
{
	if (m_pSKReplyLibEventHandler != NULL)
	{
		m_pSKReplyLibEventHandler->ShutdownConnectionPoint();
		m_pSKReplyLibEventHandler->Release();
		m_pSKReplyLibEventHandler = NULL;
	}
	if (m_pSKReplyLib != NULL)
	{
		m_pSKReplyLib->Release();
	}
}
HRESULT CReply::OnEventFiringObjectInvoke
(
    ISKReplyLibEventHandler* pEventHandler,
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
		case 1:
		{
			varlValue = (pdispparams->rgvarg)[1];
			_bstr_t bstrLoginID = V_BSTR(&varlValue);
			varlValue = (pdispparams->rgvarg)[0];
			LONG nCode = V_I4(&varlValue);
			OnConnect(string(bstrLoginID), nCode);
			break;
		}
		case 2:
		{
			varlValue = (pdispparams->rgvarg)[1];
			_bstr_t bstrLoginID = V_BSTR(&varlValue);
			varlValue = (pdispparams->rgvarg)[0];
			LONG nCode = V_I4(&varlValue);
			OnDisconnect(string(bstrLoginID), nCode);
			break;
		}
		case 3:
		{
			OnComplete();
			break;
		}
		case 6: 
		{
			varlValue = (pdispparams->rgvarg)[2];
			_bstr_t bstrUserID = V_BSTR(&varlValue);
			varlValue = (pdispparams->rgvarg)[1];
			_bstr_t bstrMessage = V_BSTR(&varlValue);
			varlValue = (pdispparams->rgvarg)[0];
			OnReplyMessage(string(bstrMessage), string(bstrUserID), varlValue.piVal);
			break;
		}
		case 8:
		{
			varlValue = (pdispparams->rgvarg)[0];
			_bstr_t Data = V_BSTR(&varlValue);
			OnNewData(string(Data));
			break;
		}
    }
    return S_OK;
}
long CReply::ConnectById(string strLogInID)
{
	bstr_t bstrLogInId = strLogInID.c_str();
	return m_pSKReplyLib->SKReplyLib_ConnectByID(bstrLogInId);
}
long CReply::IsConnecttedByID(string strLogInID)
{
	bstr_t bstrLogInId = strLogInID.c_str();
	return m_pSKReplyLib->SKReplyLib_IsConnectedByID(bstrLogInId);
}
long CReply::SolaceCloseByID(string strLogInID)
{
	bstr_t bstrLogInId = strLogInID.c_str();
	return m_pSKReplyLib->SKReplyLib_SolaceCloseByID(bstrLogInId);
}
void CReply::OnConnect(string strLogInID, long nErrorCode)//回報連線結果
{
	cout << "【OnConnect】回報連線：" << strLogInID << ", " << nErrorCode << endl;
}
void CReply::OnReplyMessage(string strUserID, string strMessage, short* sConfirmCode)//有公告會主動呼叫函式，並通知此公告訊息
{
	*sConfirmCode = -1;
	cout << "【OnReplyMessage】Login ID : " << strUserID << ", Message : " << strMessage << endl;
}
void CReply::OnDisconnect(string strLogInID, long nErrorCode)//回報連線中斷
{
	cout << "【OnConnect】回報斷線：" << strLogInID << ", " << nErrorCode << endl;
}
void CReply::OnComplete()//回報回補完成
{
	cout << "【OnConnect】回補完成" << endl;
}
void CReply::OnNewData(string strData)//主動回報委托狀態，包含預約單回報
{
	cout << "【OnNewData】" << strData << endl;
}
