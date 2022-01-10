// ---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include <psapi.h>

#include "Unit1.h"
// ---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;

String exeName, winTitle, winClass;
bool winActivated = false;

// ---------------------------------------------------------------------------
String GetModuleName() {
	return GetModuleName(unsigned(HInstance));
}

// ---------------------------------------------------------------------------
String GetModuleDir() {
	String name(GetModuleName(unsigned(HInstance)));
	return name.SubString(1, name.LastDelimiter(L"\\"));
}

// ---------------------------------------------------------------------------
BOOL BringWindowToFront(HWND hwnd) {
	DWORD lForeThreadID = 0;
	DWORD lThisThreadID = 0;
	BOOL lReturn = TRUE;

	// Make a window, specified by its handle (hwnd)
	// the foreground window.

	// If it is already the foreground window, exit.
	if (hwnd != GetForegroundWindow()) {

		// Get the threads for this window and the foreground window.

		lForeThreadID = GetWindowThreadProcessId(GetForegroundWindow(), NULL);
		lThisThreadID = GetWindowThreadProcessId(hwnd, NULL);

		// By sharing input state, threads share their concept of
		// the active window.

		if (lForeThreadID != lThisThreadID) {
			// Attach the foreground thread to this window.
			AttachThreadInput(lForeThreadID, lThisThreadID, TRUE);
			// Make this window the foreground window.
			lReturn = SetForegroundWindow(hwnd);
			// Detach the foreground window's thread from this window.
			AttachThreadInput(lForeThreadID, lThisThreadID, FALSE);
		}
		else {
			lReturn = SetForegroundWindow(hwnd);
		}

		// Restore this window to its normal size.

		if (IsIconic(hwnd)) {
			ShowWindow(hwnd, SW_RESTORE);
		}
		else {
			ShowWindow(hwnd, SW_SHOW);
		}
	}
	return lReturn;
}

// ---------------------------------------------------------------------------
void ShowMessagePlus(String msg) {
	MessageBox(Form1->Handle, msg.w_str(), L"nomulti", MB_OK | MB_SETFOREGROUND);
}

// ---------------------------------------------------------------------------
String GetWindowClassPlus(HWND hwnd) {
	wchar_t *tbuf = new wchar_t[2048];
	tbuf[0] = L'\0';
	GetClassName(hwnd, tbuf, 2047);
	String ret(tbuf);
	delete[]tbuf;
	return ret;
}

// ---------------------------------------------------------------------------
String GetWindowTitlePlus(HWND hwnd) {
	wchar_t *tbuf = new wchar_t[2048];
	tbuf[0] = L'\0';
	GetWindowText(hwnd, tbuf, 2047);
	String ret(tbuf);
	delete[]tbuf;
	return ret;
}

// ---------------------------------------------------------------------------
static BOOL CALLBACK enumWindowCallback(HWND hTarget, LPARAM lparam) {
	if (winTitle.Length()) {
		String wttl(GetWindowTitlePlus(hTarget));
		if (wttl.Length() && wttl.Pos(winTitle)); //
		else
			return TRUE;
	}
	if (winClass.Length()) {
		String wcls(GetWindowClassPlus(hTarget));
		if (wcls.Length() && wcls.Compare(winClass) == 0); //
		else
			return TRUE;
	}

	DWORD procID = 0;
	if (GetWindowThreadProcessId(hTarget, &procID)) {
		HANDLE hProc = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE,
			procID);
		if (hProc) {
			wchar_t buffer[MAX_PATH];
			if (GetModuleFileNameEx(hProc, 0, buffer, MAX_PATH)) {
				String fexe(buffer);
				int pos = fexe.Pos(exeName);
				if (pos && (fexe.Length() - pos + 1 == exeName.Length())) {
					BringWindowToFront(hTarget);
					winActivated = true;
				}
			}
			CloseHandle(hProc);
		}
	}
	return FALSE;
}

// ---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner) : TForm(Owner) {
	Left = -(Width + 5);
	Top = -(Height + 5);
}

// ---------------------------------------------------------------------------
void __fastcall TForm1::Timer1Timer(TObject *Sender) {
	Timer1->Enabled = false;

	int argc = ParamCount();
	bool paramgood = false;
	if (argc >= 2) {
		exeName = ParamStr(1);

		String param;
		for (int i = 2; i <= argc; i++) {
			param = ParamStr(i);
			if (param.Pos(L"title:") == 1) {
				if (param.Length() - 6 > 0) {
					paramgood = true;
					winTitle = param.SubString(7, param.Length() - 6);
				}
			}
			else if (param.Pos(L"class:") == 1) {
				if (param.Length() - 6 > 0) {
					paramgood = true;
					winClass = param.SubString(7, param.Length() - 6);
				}
			}
			else {
				String msg;
				msg.sprintf(L"ERROR: Invalid parameter(%d): %s", i, param.w_str());
				ShowMessagePlus(msg);
				goto done;
			}
		}
	}
	else if (argc == 0) {
		ShowMessagePlus(L"Prevent multiple instances by activating existing window"
			L"\n\nUsage:\nnomulti programtorun.exe \"title:Window Title\" class:WINDOWCLASS" L"\n\nSpecify at least title or class"
			);
		goto done;
	}

	if (!paramgood) {
		ShowMessagePlus(L"ERROR: Insufficient parameters");
		goto done;
	}

	EnumWindows(enumWindowCallback, NULL);

	if (!winActivated) {
		SetCurrentDirectory(GetModuleDir().w_str());

		STARTUPINFO si;
		PROCESS_INFORMATION pi;
		ZeroMemory(&si, sizeof(si));
		si.cb = sizeof(si);
		ZeroMemory(&pi, sizeof(pi));
		if (CreateProcess(exeName.w_str(), NULL, NULL, NULL, FALSE, 0, NULL, NULL,
			&si, &pi)) {
			CloseHandle(pi.hThread);
			CloseHandle(pi.hProcess);
		}
	}

done:
	Close();
}
// ---------------------------------------------------------------------------
