import sys
import locale
from PyQt4.QtGui import QApplication
from PyQt4.QtCore import QTextCodec
from mainwnd import CraneTestDocWnd

if __name__ == "__main__":
    app = QApplication(sys.argv)
    mycode = locale.getpreferredencoding()
    code = QTextCodec.codecForName(mycode)
    QTextCodec.setCodecForLocale(code)
    QTextCodec.setCodecForTr(code)
    QTextCodec.setCodecForCStrings(code)
    wnd = CraneTestDocWnd()
    wnd.show()

    # wi = MyWidget()
    # wi.show()
    app.exec_()