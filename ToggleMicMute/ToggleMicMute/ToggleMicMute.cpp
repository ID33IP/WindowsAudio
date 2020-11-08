// ConsoleApplication1.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <iostream>
#include <windows.h>
#include <mmdeviceapi.h>
#include <audioendpoints.h>
#include <endpointvolume.h>

#define _ATL_DEBUG_INTERFACES

// TODO: Move functions to an external header
#pragma region FunctionsToMove
// ====================================================================================================================
bool initializeComLibrary()
{
  HRESULT wInitResult(CoInitializeEx(NULL, COINIT_MULTITHREADED));

  if (FAILED(wInitResult))
  {
    std::cout << "Error when initializing COM Library: <" << wInitResult << '>' << std::endl;
    return FALSE;
  }

  return TRUE;
}


// ====================================================================================================================
bool getDeviceEnumerator(IMMDeviceEnumerator** oDeviceEnumerator)
{
  HRESULT wDeviceEnumeratorResult = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), reinterpret_cast<void**>(oDeviceEnumerator));

  if (FAILED(wDeviceEnumeratorResult))
  {
    std::cout << "Error when retrieving IMMDeviceEnumerator. HRESULT Error code: <" << wDeviceEnumeratorResult << '>' << std::endl;
    return FALSE;
  }

  return TRUE;
}


// ====================================================================================================================
bool getDefaultDeviceEndpoint(EDataFlow iDataFlow, IMMDeviceEnumerator* iDeviceEnumerator, IMMDevice** oDevice, IAudioEndpointVolume** oAudioEndpointVolume)
{
  HRESULT wHr;

  wHr = iDeviceEnumerator->GetDefaultAudioEndpoint(iDataFlow, eConsole, oDevice);

  if (FAILED(wHr))
  {
    std::cout << "Error when retrieving default IMMDevice for capture. HRESULT Error code: <" << wHr << '>' << std::endl;
    return FALSE;
  }

  wHr = (*oDevice)->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_ALL, NULL, reinterpret_cast<void**>(oAudioEndpointVolume));

  if (FAILED(wHr))
  {
    // TODO: print device id in log
    std::cout << "Error when retrieving default IAudioEndpointVolume for device. HRESULT Error code: <" << wHr << '>' << std::endl;
    return FALSE;
  }

  return TRUE;
}


// ====================================================================================================================
bool toggleMute(IAudioEndpointVolume* iAudioDeviceEndpointVolume, BOOL& oResultingMuteStatus)
{
  BOOL wCurrentMuteStatus = FALSE;
  oResultingMuteStatus = FALSE;
  HRESULT wMuteResult = iAudioDeviceEndpointVolume->GetMute(&wCurrentMuteStatus);

  // Get current mute status of device
  if (FAILED(wMuteResult))
  {
    // TODO: print device id in log
    std::cout << "Error when retrieving current mute status for device. HRESULT Error code: <" << wMuteResult << '>' << std::endl;
    return FALSE;
  }

  // Toggle mute status of device
  oResultingMuteStatus = !wCurrentMuteStatus;
  iAudioDeviceEndpointVolume->SetMute(oResultingMuteStatus, NULL);

  return TRUE;
}

#pragma endregion


// ====================================================================================================================
int main()
{
  // Instantly hides the console window
  SetWindowPos(GetConsoleWindow(), NULL, 5000, 5000, 0, 0, 0);
  ShowWindow(GetConsoleWindow(), SW_HIDE);

  // Initialize resources
  IMMDeviceEnumerator* wDeviceEnumerator = NULL;
  IMMDevice* wDefaultDevice = NULL;
  IAudioEndpointVolume* wMicrophoneEndpoint = NULL;
  IAudioEndpointVolume* wPlaybackEndpoint = NULL;

  bool wSuccess(true);
  BOOL wNewMicrophoneMuteStatus = FALSE;

  if (wSuccess && !initializeComLibrary())
  {
    wSuccess = FALSE;
  }

  if (wSuccess && !getDeviceEnumerator(&wDeviceEnumerator))
  {
    wSuccess = FALSE;
  }

  if (wSuccess && !getDefaultDeviceEndpoint(eCapture, wDeviceEnumerator, &wDefaultDevice, &wMicrophoneEndpoint))
  {
    wSuccess = FALSE;
  }

  if (wSuccess && !toggleMute(wMicrophoneEndpoint, wNewMicrophoneMuteStatus))
  {
    wSuccess = FALSE;
  }

  std::wstring wSoundFilePath = L"";
  
  if (wNewMicrophoneMuteStatus)
  {
    wSoundFilePath = L"C:\\Windows\\Media\\Speech Off.wav";
  }
  else
  {
    wSoundFilePath = L"C:\\Windows\\Media\\Speech On.wav";
  }

  BOOL wPlaySoundSuccess(FALSE);

  std::cout << (wNewMicrophoneMuteStatus ? "<Muted>" : "<Unmuted>") << " default capture device." << std::endl;
  wPlaySoundSuccess = PlaySound(wSoundFilePath.c_str(), NULL, SND_FILENAME);

  if (!wPlaySoundSuccess)
  {
    std::cout << "Failed to play sound." << std::endl;
  }

  // Release resources
  wMicrophoneEndpoint->Release();
  wMicrophoneEndpoint = NULL;

  wDefaultDevice->Release();
  wDefaultDevice = NULL;

  wDeviceEnumerator->Release();
  wDeviceEnumerator = NULL;

  // Uninitialize COM Library
  CoUninitialize();

  return wSuccess;
}
