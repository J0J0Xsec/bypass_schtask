// myTaskScheduler.h
#include <Windows.h>

extern "C"
{
    __declspec(dllexport) HRESULT AddScheduledTask();
}

// myTaskScheduler.cpp
#include <comdef.h>
#include <taskschd.h>
//#include "myTaskScheduler.h"

#pragma comment(lib, "taskschd.lib")
#pragma comment(lib, "comsuppw.lib")

HRESULT AddScheduledTask()
{
    // Initialize COM.
    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(hr))
        return hr;

    // Set general task settings.
    ITaskService* pService = NULL;
    hr = CoCreateInstance(CLSID_TaskScheduler,
        NULL,
        CLSCTX_INPROC_SERVER,
        IID_ITaskService,
        (void**)&pService);

    if (FAILED(hr))
        return hr;

    hr = pService->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t());
    if (FAILED(hr))
        return hr;

    ITaskFolder* pRootFolder = NULL;
    hr = pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);
    if (FAILED(hr))
        return hr;

    // Create the new task.
    ITaskDefinition* pTask = NULL;
    hr = pService->NewTask(0, &pTask);

    // Create the trigger.
    ITriggerCollection* pTriggerCollection = NULL;
    hr = pTask->get_Triggers(&pTriggerCollection);
    ITrigger* pTrigger = NULL;
    hr = pTriggerCollection->Create(TASK_TRIGGER_DAILY, &pTrigger);
    IDailyTrigger* pDailyTrigger = NULL;
    hr = pTrigger->QueryInterface(
        IID_IDailyTrigger, (void**)&pDailyTrigger);

    pDailyTrigger->put_StartBoundary(_bstr_t(L"2022-01-01T12:00:00"));  // Start at noon.

    // Create the action for the task to execute.
    IActionCollection* pActionCollection = NULL;
    hr = pTask->get_Actions(&pActionCollection);
    IAction* pAction = NULL;
    hr = pActionCollection->Create(TASK_ACTION_EXEC, &pAction);
    IExecAction* pExecAction = NULL;
    hr = pAction->QueryInterface(
        IID_IExecAction, (void**)&pExecAction);

    // Set the path of the executable.
    pExecAction->put_Path(_bstr_t(L"C:\\Windows\\System32\\notepad.exe"));

    // Register the task.
    IRegisteredTask* pRegisteredTask = NULL;
    hr = pRootFolder->RegisterTaskDefinition(
        _bstr_t(L"WindowsUpdate"),
        pTask,
        TASK_CREATE_OR_UPDATE,
        _variant_t(),
        _variant_t(),
        TASK_LOGON_INTERACTIVE_TOKEN,
        _variant_t(L""),
        &pRegisteredTask);

    return hr;
}
