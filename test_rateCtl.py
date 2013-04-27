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
oldRate = clientLib.nvdaController_getRate()
clientLib.nvdaController_speakText(u"現在の速さは %s です。今から速さを変更します"%oldRate)
time.sleep(10)
clientLib.nvdaController_setRate(oldRate-50)
newRate = clientLib.nvdaController_getRate()
clientLib.nvdaController_speakText(u"速さを %s に変更しました。速さを %s に戻します"%(newRate,oldRate))
time.sleep(10)
clientLib.nvdaController_setRate(oldRate)
clientLib.nvdaController_speakText(u"元通りです。ほらね。")
clientLib.nvdaController_speakText(u"作業を終了しました")

