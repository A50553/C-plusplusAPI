#include "CCenter.h"

CCenter::CCenter()
{
	m_pSKCenterLib.CreateInstance(__uuidof(SKCOMLib::SKCenterLib));
    m_pSKCenterLibEventHandler = new ISKCenterLibEventHandler(*this, m_pSKCenterLib, &CCenter::OnEventFiringObjectInvoke);
}
CCenter::~CCenter()
{
    if (m_pSKCenterLibEventHandler!=NULL)
    {
        m_pSKCenterLibEventHandler->ShutdownConnectionPoint();
        m_pSKCenterLibEventHandler->Release();
        m_pSKCenterLibEventHandler = NULL;
    }
    if (m_pSKCenterLib!=NULL)
    {
        m_pSKCenterLib->Release();
    }
}
HRESULT CCenter::OnEventFiringObjectInvoke
(
    ISKCenterLibEventHandler* pEventHandler,
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
    case 1: // Event1 event.
        varlValue = (pdispparams->rgvarg)[0];
        OnTimer(V_I4(&varlValue));
        break;
    }

    return S_OK;
}

long CCenter::Login(string strUserID, string strPassword)
{
	const char* bstrUserID = strUserID.c_str();
	const char* bstrPassword = strPassword.c_str();
	return m_pSKCenterLib->SKCenterLib_Login(bstrUserID, bstrPassword);
}
BSTR CCenter::GetReturnCodeMessage(long nCode)//���o�T���N�X
{
	return m_pSKCenterLib->SKCenterLib_GetReturnCodeMessage(nCode);
}
BSTR CCenter::GetLastLogInfo()//���o�̫�@��log���e
{
    return m_pSKCenterLib->SKCenterLib_GetLastLogInfo();
}
void CCenter::OnTimer(long nTime)//��ܷ�U�ɶ�
{
	long nHour = nTime / 10000;
	long nMin = (nTime % 10000) / 100;
	long nSec = nTime % 100;
	cout << "\n�iOnTime�j " << nHour << " : " << nMin << " : " << nSec;
}
void CCenter::PrintCodeMessage(string strFunction, string strFunctionName, long nCode)//�L�X���~�T��
{
    if (nCode == 0)
    {
        string strMessage = (LPCSTR)(m_pSKCenterLib->SKCenterLib_GetReturnCodeMessage(nCode));
        cout << "�i" << strFunction.c_str()<<"�j" <<"�i" << strFunctionName.c_str() << "�j" << "�i" << strMessage.c_str() << "�j\n";
    }
    else
    {
        string strMessage = (LPCSTR)(m_pSKCenterLib->SKCenterLib_GetReturnCodeMessage(nCode));
        string strMessage2 = (LPCSTR)(m_pSKCenterLib->SKCenterLib_GetLastLogInfo());
        cout << "�i" << strFunction.c_str() << "�j" << "�i" << strFunctionName.c_str() << "�j" << "�i" << strMessage.c_str() << "�j" << "�i" << strMessage2.c_str() << "�j\n";
    }
}
