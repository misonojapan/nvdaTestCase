# coding: utf-8
import time
import ctypes
DLLPATH = '../newclient/nvdaHelper/build/x86/client/nvdaControllerClient32.dll'
clientLib=ctypes.windll.LoadLibrary(DLLPATH)
res=clientLib.nvdaController_testIfRunning()
if res!=0:
	errorMessage=str(ctypes.WinError(res))
	ctypes.windll.user32.MessageBoxW(0,u"Error: %s"%errorMessage,u"Error communicating with NVDA",0)
	exit(1)
oldPitch = clientLib.nvdaController_getPitch()
clientLib.nvdaController_speakText(u"現在のピッチは %s です。今からピッチを変更します"%oldPitch)
time.sleep(10)
clientLib.nvdaController_setPitch(oldPitch-50)
newPitch = clientLib.nvdaController_getPitch()
clientLib.nvdaController_speakText(u"ピッチを %s に変更しました。ピッチを %s に戻します"%(newPitch,oldPitch))
time.sleep(10)
clientLib.nvdaController_setPitch(oldPitch)
clientLib.nvdaController_speakText(u"元通りです。ほらね。")
clientLib.nvdaController_speakText(u"作業を終了しました")

