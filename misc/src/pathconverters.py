#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import traceback
from urllib.parse import urlparse
from urllib.request import url2pathname
# from pathlib import Path
import os
def enableRemoteDebugging(func):  # デバッグサーバーに接続したい関数やメソッドにつけるデコレーター。主にリスナーのメソッドのデバッグ目的。
	def wrapper(*args, **kwargs):
		frame = None
		doc = XSCRIPTCONTEXT.getDocument()
		if doc:  # ドキュメントが取得できた時
			frame = doc.getCurrentController().getFrame()  # ドキュメントのフレームを取得。
		else:
			currentframe = XSCRIPTCONTEXT.getDesktop().getCurrentFrame()  # モードレスダイアログのときはドキュメントが取得できないので、モードレスダイアログのフレームからCreatorのフレームを取得する。
			frame = currentframe.getCreator()
		if frame:   
			import time
			indicator = frame.createStatusIndicator()  # フレームからステータスバーを取得する。
			maxrange = 2  # ステータスバーに表示するプログレスバーの目盛りの最大値。2秒ロスするが他に適当な告知手段が思いつかない。
			indicator.start("Trying to connect to the PyDev Debug Server for about 20 seconds.", maxrange)  # ステータスバーに表示する文字列とプログレスバーの目盛りを設定。
			t = 1  # プレグレスバーの初期値。
			while t<=maxrange:  # プログレスバーの最大値以下の間。
				indicator.setValue(t)  # プレグレスバーの位置を設定。
				time.sleep(1)  # 1秒待つ。
				t += 1  # プログレスバーの目盛りを増やす。
			indicator.end()  # reset()の前にend()しておかないと元に戻らない。
			indicator.reset()  # ここでリセットしておかないと例外が発生した時にリセットする機会がない。
		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # デバッグサーバーを起動していた場合はここでブレークされる。import pydevdは時間がかかる。
		try:
			func(*args, **kwargs)  # Step Intoして中に入る。
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
	return wrapper
# @enableRemoteDebugging
def macro():
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。  
	doc.getText().setString("")  # すでにあるドキュメントの内容を消去する。
	global wprint
	wprint = createWriterPrinter(doc)  # Writerドキュメントに一行ずつ出力するグローバル関数を取得。
	
	
	
	
	originalpath = "$(inst)/sdk/docs/idl/ref/"
	wprint("Original path:")
	wprint("\t{}".format(originalpath))
	pathsubstservice = smgr.createInstanceWithContext("com.sun.star.comp.framework.PathSubstitution", ctx)
	fileurl = pathsubstservice.substituteVariables(originalpath, True)
	wprint("substituteVariables:")
	wprint("\t{}".format(fileurl))
	
	
	
	
	systempath = convertWithOnlyPython(fileurl)
	convertWithunohelpler(fileurl)
	convertWithFileContentProvider(ctx, smgr, fileurl)
	normalizePath(fileurl, systempath)
	wprint("\n\nTry assigning something that is not fileurl.")
	fileurl = originalpath
	wprint("\t{}".format(fileurl))
	convertWithOnlyPython(fileurl)
	convertWithunohelpler(fileurl)
	convertWithFileContentProvider(ctx, smgr, fileurl)
	

# 	fileurl = "simple string"
# 	wprint("\t{}".format(fileurl))
# 	convertWithOnlyPython(fileurl)
# 	convertWithunohelpler(fileurl)
# 	convertWithFileContentProvider(ctx, smgr, fileurl)
# 	
# 	fileurl = systempath
# 	wprint("\t{}".format(fileurl))
# 	try:
# 		convertWithOnlyPython(fileurl)
# 	except:
# 		traceback.print_exc()
# 	try:
# 		convertWithunohelpler(fileurl)
# 	except:
# 		traceback.print_exc()
# 	try:
# 		convertWithFileContentProvider(ctx, smgr, fileurl)
# 	except:
# 		traceback.print_exc()
	
# 	fileurl = originalpath
# 	wprint("\t{}".format(fileurl))
# 	try:
# 		convertWithOnlyPython(fileurl)
# 	except:
# 		traceback.print_exc()
# 	try:
# 		convertWithunohelpler(fileurl)
# 	except:
# 		traceback.print_exc()
# 	try:
# 		convertWithFileContentProvider(ctx, smgr, fileurl)
# 	except:
# 		traceback.print_exc()	
	
	
def convertWithFileContentProvider(ctx, smgr, fileurl):	
	wprint("\nConverting with FileContentProvider Service.\n")
	filecontentprovider = smgr.createInstanceWithContext("com.sun.star.ucb.FileContentProvider", ctx)
	wprint("fileurl = {}".format(fileurl))
	wprint("systempath = filecontentprovider.getSystemPathFromFileURL(fileurl)  # Converting the fileurl to the systempath." )
	systempath =  filecontentprovider.getSystemPathFromFileURL(fileurl)
	wprint("\t{}".format(systempath))
	wprint('fileurl = filecontentprovider.getFileURLFromSystemPath("", systempath)  # Converting the systempath to the fileurl.')
	fileurl = filecontentprovider.getFileURLFromSystemPath("", systempath)
	wprint("\t{}".format(fileurl))
	
def convertWithunohelpler(fileurl):	
	wprint("\nConverting with unohelpler.\n")
	wprint("fileurl = {}".format(fileurl))
	wprint("systempath = unohelper.fileUrlToSystemPath(fileurl)  # Converting the fileurl to the systempath." )
	systempath =  unohelper.fileUrlToSystemPath(fileurl)
	wprint("\t{}".format(systempath))
	wprint("fileurl = unohelper.systemPathToFileUrl(systempath)  # Converting the systempath to the fileurl.")
	fileurl = unohelper.systemPathToFileUrl(systempath)
	wprint("\t{}".format(fileurl))
	
def convertWithOnlyPython(fileurl):	
	wprint("\nConverting with only Python modules.\n")
	wprint("fileurl = {}".format(fileurl))
	wprint("systempath = url2pathname(urlparse(fileurl).path)  # Converting the fileurl to the systempath." )
	systempath = url2pathname(urlparse(fileurl).path)
	wprint("\t{}".format(systempath))
	wprint("In Windows there is no pythonpath module, so it can not be converted from systempath to fileurl." )
# 	wprint("fileurl = Path(systempath).as_uri()  # Converting the systempath to the fileurl.")
# 	fileurl = Path(systempath).as_uri()
# 	wprint("\t{}".format(fileurl))
	return systempath
def normalizePath(fileurl, systempath):
	wprint("\nNormalize should be done in systempath.")
	wprint("failedfileurl = os.path.normpath(fileurl) # Normalize a fileurl.")
	failedfileurl = os.path.normpath(fileurl)
	wprint("\t{}".format(failedfileurl))
	wprint("normalizedsystempath = os.path.normpath(systempath) # Normalize a systempath.")
	normalizedsystempath = os.path.normpath(systempath)
	wprint("\t{}".format(normalizedsystempath))
	wprint("")	
def createWriterPrinter(doc):  # 引数はWriterドキュメント
	text = doc.getText()
	def wprint(txt):  # Writerドキュメントに追記出力。	
		rng = text.getEnd()  # 文末の領域を取得。
		rng.setString("{}\n".format(txt))
	return wprint
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
if __name__ == "__main__":  # オートメーションで実行するとき
	import officehelper
	from functools import wraps
	import sys
	from com.sun.star.beans import PropertyValue
	from com.sun.star.script.provider import XScriptContext  
	def connectOffice(func):  # funcの前後でOffice接続の処理
		@wraps(func)
		def wrapper():  # LibreOfficeをバックグラウンドで起動してコンポーネントテクストとサービスマネジャーを取得する。
			try:
				ctx = officehelper.bootstrap()  # コンポーネントコンテクストの取得。
			except:
				print("Could not establish a connection with a running office.", file=sys.stderr)
				sys.exit()
			print("Connected to a running office ...")
			smgr = ctx.getServiceManager()  # サービスマネジャーの取得。
			print("Using {} {}".format(*_getLOVersion(ctx, smgr)))  # LibreOfficeのバージョンを出力。
			return func(ctx, smgr)  # 引数の関数の実行。
		def _getLOVersion(ctx, smgr):  # LibreOfficeの名前とバージョンを返す。
			cp = smgr.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', ctx)
			node = PropertyValue(Name = 'nodepath', Value = 'org.openoffice.Setup/Product' )  # share/registry/main.xcd内のノードパス。
			ca = cp.createInstanceWithArguments('com.sun.star.configuration.ConfigurationAccess', (node,))
			return ca.getPropertyValues(('ooName', 'ooSetupVersion'))  # LibreOfficeの名前とバージョンをタプルで返す。
		return wrapper
	@connectOffice  # mainの引数にctxとsmgrを渡すデコレータ。
	def main(ctx, smgr):  # XSCRIPTCONTEXTを生成。
		class ScriptContext(unohelper.Base, XScriptContext):
			def __init__(self, ctx):
				self.ctx = ctx
			def getComponentContext(self):
				return self.ctx
			def getDesktop(self):
				return ctx.getByName('/singletons/com.sun.star.frame.theDesktop')  # com.sun.star.frame.Desktopはdeprecatedになっている。
			def getDocument(self):
				return self.getDesktop().getCurrentComponent()
		return ScriptContext(ctx)  
	XSCRIPTCONTEXT = main()  # XSCRIPTCONTEXTを取得。
	doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
# 	doctype = "scalc", "com.sun.star.sheet.SpreadsheetDocument"  # Calcドキュメントを開くとき。
	doctype = "swriter", "com.sun.star.text.TextDocument"  # Writerドキュメントを開くとき。
	if (doc is None) or (not doc.supportsService(doctype[1])):  # ドキュメントが取得できなかった時またはCalcドキュメントではない時
		XSCRIPTCONTEXT.getDesktop().loadComponentFromURL("private:factory/{}".format(doctype[0]), "_blank", 0, ())  # ドキュメントを開く。ここでdocに代入してもドキュメントが開く前にmacro()が呼ばれてしまう。
	flg = True
	while flg:
		doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
		if doc is not None:
			flg = (not doc.supportsService(doctype[1]))  # ドキュメントタイプが確認できたらwhileを抜ける。
	macro()