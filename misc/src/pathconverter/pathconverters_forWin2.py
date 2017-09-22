#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from urllib.parse import urlparse
from urllib.request import url2pathname
from pathlib import Path
def macro():
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。  
	doc.getText().setString("")  # すでにあるドキュメントの内容を消去する。
	wprint = createWriterPrinter(doc)  # Writerドキュメントに一行ずつ出力するグローバル関数を取得。
	filecontentprovider = smgr.createInstanceWithContext("com.sun.star.ucb.FileContentProvider", ctx)
	convert = converters(filecontentprovider)
	
# 	args = "/home/pq", "Pass Linux SystemPath."
# 	wprint(convert(args))
# 	args = "C:\\Users\\pq", "Pass Windows SystemPath."
# 	wprint(convert(args))
# 	args = "file:///home/pq", "Pass Linux FielURL."
# 	wprint(convert(args))
# 	args = "file:///C:/Users/pq", "Pass Windows FielURL."
# 	wprint(convert(args))	
# 	args = "simple word", "Pass words without a colon"
# 	wprint(convert(args))
# 	args = "c:word", "Pass words with a colon"
# 	wprint(convert(args))		
# 	
# 	args = "/home/日本語", "Pass Linux SystemPath."
# 	wprint(convert(args))
# 	args = "C:\\Users\\日本語", "Pass Windows SystemPath."
# 	wprint(convert(args))
	args = "file:///home/日本語", "Pass Linux FielURL."
	wprint(convert(args))  # Windowsではunohelper.systemPathToFileUrl(arg)でLibreOfficeがクラッシュする。
	args = "file:///C:/Users/日本語", "Pass Windows FielURL."
	wprint(convert(args))  # Windowsではunohelper.systemPathToFileUrl(arg)でLibreOfficeがクラッシュする。
# 	args = "日本語", "Pass words without a colon"
# 	wprint(convert(args))
	args = "c:日本語", "Pass words with a colon"
	wprint(convert(args))  # Windowsではunohelper.systemPathToFileUrl(arg)でLibreOfficeがクラッシュする。
	
# 	args = "file:///home/%E6%97%A5%E6%9C%AC%E8%AA%9E", "Pass Linux FielURL."
# 	wprint(convert(args))
	args = "file:///C:/Users/%E6%97%A5%E6%9C%AC%E8%AA%9E", "Pass Windows FielURL."
	wprint(convert(args))  # Windowsではunohelper.systemPathToFileUrl(arg)でLibreOfficeがクラッシュする。
def converters(filecontentprovider):
	output = """arg = '{}'  # {}
	FileURL to SystemPath
		# with Python modules
		systempath = url2pathname(urlparse(arg).path)
			Return value: {}
		# with unohelper
		systempath =  unohelper.fileUrlToSystemPath(arg)  
			Return value: {}
		# with FileContentProvider
		systempath =  filecontentprovider.getSystemPathFromFileURL(arg)  
			Return value: {}
	SystemPath to FileURL
		# with Python Modules
		fileurl = Path(arg).as_uri()  
			Return value: {}	
		# with unohelper
		fileurl = unohelper.systemPathToFileUrl(arg)  
			Return value: {}	
		# with FileContentProvider
		fileurl = filecontentprovider.getFileURLFromSystemPath("", arg)  
			Return value: {}	
	"""
	def convert(args):	
		arg = args[0]
		results = ["disabled"]*6
		try:
			results[0] = url2pathname(urlparse(arg).path)
		except Exception as e:
			results[0] = "Exception: {}".format(e)
		try:
			results[1] = unohelper.fileUrlToSystemPath(arg)
		except Exception as e:
			results[1] = "Exception: {}".format(e)
		try:
			results[2] = filecontentprovider.getSystemPathFromFileURL(arg)
		except Exception as e:
			results[2] = "Exception: {}".format(e)
		try:
			results[3] = Path(arg).as_uri()
		except Exception as e:
			results[3] = "Exception: {}".format(e)
# 		try:
# 			results[4] = unohelper.systemPathToFileUrl(arg)  # WindowsではPass words without a colonとPass Linux FielURL以外では日本語を含むとLibreOfficeがクラッシュする。
# 		except Exception as e:
# 			results[4] = "Exception: {}".format(e)
		try:
			results[5] = filecontentprovider.getFileURLFromSystemPath("", arg)
		except Exception as e:
			results[5] = "Exception: {}".format(e)
		return output.format(*args, *results)
	return convert
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