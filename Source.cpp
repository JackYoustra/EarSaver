#include <Windows.h>
#include <string>
#include <initguid.h>
#include <mmdeviceapi.h>
#include <stdio.h>
#include <endpointvolume.h>


#define SAFE_RELEASE(punk)  \
              if ((punk) != NULL)  \
                { (punk)->Release(); (punk) = NULL; }


class CMMNotificationClient : public IMMNotificationClient {
	LONG _cRef;
	IMMDeviceEnumerator *_pEnumerator;

	// Private function to print device-friendly name
	HRESULT _PrintDeviceName(LPCWSTR  pwstrId);

public:
	CMMNotificationClient() :
		_cRef(1),
		_pEnumerator(NULL)
	{
	}

	~CMMNotificationClient()
	{
		SAFE_RELEASE(_pEnumerator)
	}

	// IUnknown methods -- AddRef, Release, and QueryInterface

	ULONG STDMETHODCALLTYPE AddRef()
	{
		return InterlockedIncrement(&_cRef);
	}

	ULONG STDMETHODCALLTYPE Release()
	{
		ULONG ulRef = InterlockedDecrement(&_cRef);
		if (0 == ulRef)
		{
			delete this;
		}
		return ulRef;
	}

	HRESULT STDMETHODCALLTYPE QueryInterface(
		REFIID riid, VOID **ppvInterface)
	{
		if (IID_IUnknown == riid)
		{
			AddRef();
			*ppvInterface = (IUnknown*)this;
		}
		else if (__uuidof(IMMNotificationClient) == riid)
		{
			AddRef();
			*ppvInterface = (IMMNotificationClient*)this;
		}
		else
		{
			*ppvInterface = NULL;
			return E_NOINTERFACE;
		}
		return S_OK;
	}

	// Callback methods for device-event notifications.

	HRESULT STDMETHODCALLTYPE OnDefaultDeviceChanged(
		EDataFlow flow, ERole role,
		LPCWSTR pwstrDeviceId){
		char  *pszFlow = "?????";
		char  *pszRole = "?????";

		_PrintDeviceName(pwstrDeviceId);

		switch (flow)
		{
		case eRender:
			pszFlow = "eRender";
			break;
		case eCapture:
			pszFlow = "eCapture";
			break;
		}

		switch (role)
		{
		case eConsole:
			pszRole = "eConsole";
			break;
		case eMultimedia:
			pszRole = "eMultimedia";
			break;
		case eCommunications:
			pszRole = "eCommunications";
			break;
		}

		printf("  -->New default device: flow = %s, role = %s\n",
			pszFlow, pszRole);
		return S_OK;
	}

	HRESULT STDMETHODCALLTYPE OnDeviceAdded(LPCWSTR pwstrDeviceId)
	{
		_PrintDeviceName(pwstrDeviceId);

		printf("  -->Added device\n");
		return S_OK;
	};

	HRESULT STDMETHODCALLTYPE OnDeviceRemoved(LPCWSTR pwstrDeviceId)
	{
		_PrintDeviceName(pwstrDeviceId);

		printf("  -->Removed device\n");
		return S_OK;
	}

	HRESULT STDMETHODCALLTYPE OnDeviceStateChanged(
		LPCWSTR pwstrDeviceId,
		DWORD dwNewState)
	{
		std::wstring pszState = L"?????";

		_PrintDeviceName(pwstrDeviceId);

		switch (dwNewState)
		{
		case DEVICE_STATE_ACTIVE:
			pszState = L"ACTIVE";
			break;
		case DEVICE_STATE_DISABLED:
			pszState = L"DISABLED";
			break;
		case DEVICE_STATE_NOTPRESENT:
			pszState = L"NOTPRESENT";
			break;
		case DEVICE_STATE_UNPLUGGED:
			pszState = L"UNPLUGGED";
			break;
		}

		OutputDebugString((L"  -->New device state is DEVICE_STATE_" + pszState).c_str());

		return S_OK;
	}
	

	void adjustAudioVolume() {
		IMMDeviceCollection *collection = NULL;

		_pEnumerator->EnumAudioEndpoints(
			EDataFlow::eRender,
			DEVICE_STATE_ACTIVE,
			&collection);

		UINT count = 0;
		collection->GetCount(&count);
		for (UINT i = 0; i < count; i++) {
			IMMDevice *currentDevice = NULL;
			collection->Item(i, &currentDevice);

			// read properties, trying to target formfactor
			IPropertyStore *deviceProperties = NULL;
			currentDevice->OpenPropertyStore(
				STGM_READ,
				&deviceProperties);
			PROPVARIANT variant;
			deviceProperties->GetValue(PKEY_AudioEndpoint_FormFactor, &variant);
			UINT formFactor = variant.uintVal;
			if (formFactor == EndpointFormFactor::Headphones || formFactor == EndpointFormFactor::Headset) {
				// TODO: set audio volume
				void *argEndpointVolume = NULL;
				currentDevice->Activate(
					__uuidof(IAudioEndpointVolume), // instead of IDD_IAudioEndpointVolume
					CLSCTX_ALL,
					NULL,
					&argEndpointVolume);
				IAudioEndpointVolume *endpointVolume = (IAudioEndpointVolume*)argEndpointVolume;
				endpointVolume->SetMasterVolumeLevelScalar(0.1, NULL); // set it to 10% volume, no identifier for this modification
				endpointVolume->Release();
			}

			deviceProperties->Release();
			currentDevice->Release();
		}
		collection->Release();
	}

	// important one
	HRESULT STDMETHODCALLTYPE OnPropertyValueChanged(
		LPCWSTR pwstrDeviceId,
		const PROPERTYKEY key){
		_PrintDeviceName(pwstrDeviceId);
		char buffer[10000];
		sprintf_s(buffer, "  -->Changed device property "
			"{%8.8x-%4.4x-%4.4x-%2.2x%2.2x-%2.2x%2.2x%2.2x%2.2x%2.2x%2.2x}#%d\n",
			key.fmtid.Data1, key.fmtid.Data2, key.fmtid.Data3,
			key.fmtid.Data4[0], key.fmtid.Data4[1],
			key.fmtid.Data4[2], key.fmtid.Data4[3],
			key.fmtid.Data4[4], key.fmtid.Data4[5],
			key.fmtid.Data4[6], key.fmtid.Data4[7],
			key.pid);
		OutputDebugStringA(buffer);

		adjustAudioVolume();

		return S_OK;
	}
};

// Given an endpoint ID string, print the friendly device name.
HRESULT CMMNotificationClient::_PrintDeviceName(LPCWSTR pwstrId)
{
	HRESULT hr = S_OK;
	IMMDevice *pDevice = NULL;
	IPropertyStore *pProps = NULL;
	PROPVARIANT varString;

	CoInitialize(NULL);
	PropVariantInit(&varString);

	if (_pEnumerator == NULL)
	{
		// Get enumerator for audio endpoint devices.
		hr = CoCreateInstance(__uuidof(MMDeviceEnumerator),
			NULL, CLSCTX_INPROC_SERVER,
			__uuidof(IMMDeviceEnumerator),
			(void**)&_pEnumerator);
	}
	if (hr == S_OK)
	{
		hr = _pEnumerator->GetDevice(pwstrId, &pDevice);
	}
	if (hr == S_OK)
	{
		hr = pDevice->OpenPropertyStore(STGM_READ, &pProps);
	}
	if (hr == S_OK)
	{
		static PROPERTYKEY key;

		GUID IDevice_FriendlyName = { 0xa45c254e, 0xdf1c, 0x4efd,{ 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0 } };
		key.pid = 14;
		key.fmtid = IDevice_FriendlyName;
		PROPVARIANT varName;
		// Initialize container for property value.
		PropVariantInit(&varName);

		// Get the endpoint device's friendly-name property.
		hr = pProps->GetValue(key, &varString);
	}
	char buffer[1000];
	sprintf_s(buffer, "----------------------\nDevice name: \"%S\"\n"
		"  Endpoint ID string: \"%S\"\n",
		(hr == S_OK) ? varString.pwszVal : L"null device",
		(pwstrId != NULL) ? pwstrId : L"null ID");
	OutputDebugStringA(buffer);
	PropVariantClear(&varString);

	SAFE_RELEASE(pProps)
		SAFE_RELEASE(pDevice)
		CoUninitialize();
	return hr;
}


const wchar_t windowsClassName[] = L"KindleClass";

NOTIFYICONDATA icondata = {};

enum TrayIcon {
	ID = 13,
	CALLBACKID = WM_APP + 1
};

// Open, options, and exit
enum MenuItems {
	OPEN = WM_APP + 3,
	OPTIONS = WM_APP + 4,
	EXIT = WM_APP + 5
};


void quit() {
	Shell_NotifyIcon(NIM_DELETE, &icondata);
	PostQuitMessage(0);
}

int main() {
	return WinMain(GetModuleHandle(NULL), NULL, GetCommandLineA(), SW_SHOW);
}

void setupNotification() {
	const CLSID CLSID_MMDeviceEnumerator = __uuidof(MMDeviceEnumerator);
	const IID IID_IMMDeviceEnumerator = __uuidof(IMMDeviceEnumerator);
	IMMDeviceEnumerator* pEnumerator = NULL;
	HRESULT hr = CoCreateInstance(
		CLSID_MMDeviceEnumerator, NULL,
		CLSCTX_ALL, IID_IMMDeviceEnumerator,
		(void**) &pEnumerator);

	IMMNotificationClient *notifClient = new CMMNotificationClient();
	pEnumerator->RegisterEndpointNotificationCallback(notifClient);
}

LRESULT CALLBACK WindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch (msg) {
	case WM_DESTROY:
		quit();
		break;
	case WM_COMMAND:
		if (HIWORD(wParam) == 0) {// menu
			switch (LOWORD(wParam)) { // which menu item
			case MenuItems::OPEN:
				ShowWindow(hwnd, SW_RESTORE);
				break;
			case MenuItems::OPTIONS:
				break;
			case MenuItems::EXIT:
				quit();
				break;
			}
		}
		break;
		// special ones here
	case TrayIcon::CALLBACKID:
		switch (LOWORD(lParam)) {
		case WM_RBUTTONDOWN: {
			// get mouse position
			POINT p;
			GetCursorPos(&p);
			const short xpos = p.x;
			const short ypos = p.y;
			// *maybe* parameterize into functions
			// set "Open" menu item info
			std::wstring openString = L"Open";
			MENUITEMINFO openInfo;
			ZeroMemory(&openInfo, sizeof(openInfo));
			openInfo.cbSize = sizeof(openInfo);
			openInfo.fMask = MIIM_ID | MIIM_STATE | MIIM_STRING;
			openInfo.fType = MFT_STRING;
			openInfo.fState = MFS_ENABLED;
			openInfo.dwTypeData = const_cast<LPWSTR>(openString.c_str());
			openInfo.cch = openString.length();
			openInfo.wID = MenuItems::OPEN;

			// set "Options" menu item info
			std::wstring optionsString = L"Options";
			MENUITEMINFO optionsInfo;
			ZeroMemory(&optionsInfo, sizeof(optionsInfo));
			optionsInfo.cbSize = sizeof(optionsInfo);
			optionsInfo.fMask = MIIM_ID | MIIM_STATE | MIIM_STRING;
			optionsInfo.fType = MFT_STRING;
			optionsInfo.fState = MFS_ENABLED;
			optionsInfo.dwTypeData = const_cast<LPWSTR>(optionsString.c_str());
			optionsInfo.cch = optionsString.length();
			optionsInfo.wID = MenuItems::OPTIONS;

			// set "Exit" menu item info
			std::wstring exitString = L"Exit";
			MENUITEMINFO exitInfo;
			ZeroMemory(&exitInfo, sizeof(exitInfo));
			exitInfo.cbSize = sizeof(exitInfo);
			exitInfo.fMask = MIIM_ID | MIIM_STATE | MIIM_STRING;
			exitInfo.fType = MFT_STRING;
			exitInfo.fState = MFS_DEFAULT;
			exitInfo.dwTypeData = const_cast<LPWSTR>(exitString.c_str());
			exitInfo.cch = exitString.length();
			exitInfo.wID = MenuItems::EXIT;


			// actually make popup menu
			HMENU menu = CreatePopupMenu();
			InsertMenuItem(menu, 0, TRUE, &openInfo);
			InsertMenuItem(menu, 1, TRUE, &optionsInfo);
			InsertMenuItem(menu, 2, TRUE, &exitInfo);

			SetForegroundWindow(hwnd);
			TrackPopupMenu(menu, TPM_LEFTALIGN | TPM_LEFTBUTTON, xpos, ypos, NULL, hwnd, NULL);
		}
		default:
			break;
		}
		break;

	case WM_DEVICECHANGE:
	default:
		return DefWindowProc(hwnd, msg, wParam, lParam);
	}
	return 0;
}


int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
	// the handle for the window, filled by a function
	HWND hWnd;
	// this struct holds information for the window class
	WNDCLASSEX wc;

	// clear out the window class for use
	ZeroMemory(&wc, sizeof(WNDCLASSEX));

	// fill in the struct with the needed information
	wc.cbSize = sizeof(WNDCLASSEX);
	wc.style = CS_HREDRAW | CS_VREDRAW;
	wc.lpfnWndProc = WindowProc;
	wc.hInstance = hInstance;
	wc.hCursor = LoadCursor(NULL, IDC_ARROW);
	wc.hbrBackground = (HBRUSH)COLOR_WINDOW;
	wc.lpszClassName = windowsClassName;

	// register the window class
	RegisterClassEx(&wc);

	// create the window and use the result as the handle
	hWnd = CreateWindowEx(NULL,
		windowsClassName,    // name of the window class
		L"Ear Saver",   // title of the window
		WS_OVERLAPPEDWINDOW,    // window style
		300,    // x-position of the window
		300,    // y-position of the window
		500,    // width of the window
		400,    // height of the window
		NULL,    // we have no parent window, NULL
		NULL,    // we aren't using menus, NULL
		hInstance,    // application handle
		NULL);    // used with multiple windows, NULL

				  // display the window on the screen
	ShowWindow(hWnd, nCmdShow);
	ShowWindow(hWnd, SW_HIDE);

	// create taskbar icon
	icondata.cbSize = sizeof(NOTIFYICONDATA);
	icondata.uVersion = NOTIFYICON_VERSION_4;
	icondata.hWnd = hWnd;
	icondata.uID = TrayIcon::ID;
	icondata.uFlags = NIF_ICON | NIF_MESSAGE;
	icondata.hIcon = LoadIcon(NULL, IDI_APPLICATION);
	icondata.uCallbackMessage = TrayIcon::CALLBACKID;
	// add icon
	Shell_NotifyIcon(NIM_ADD, &icondata);

	setupNotification();

	// this struct holds Windows event messages
	MSG msg;

	// wait for the next message in the queue, store the result in 'msg'
	while (GetMessage(&msg, NULL, 0, 0))
	{
		// translate keystroke messages into the right format
		TranslateMessage(&msg);

		// send the message to the WindowProc function
		DispatchMessage(&msg);
	}

	// return this part of the WM_QUIT message to Windows
	return msg.wParam;
}