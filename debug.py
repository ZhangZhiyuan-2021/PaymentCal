import os
import sys
from PyQt5.QtCore import QLibraryInfo, QCoreApplication
from PyQt5.QtWidgets import QApplication

print("QT_PLUGIN_PATH =", os.environ.get("QT_PLUGIN_PATH"))
print("QT_QPA_PLATFORM_PLUGIN_PATH =", os.environ.get("QT_QPA_PLATFORM_PLUGIN_PATH"))

p = QLibraryInfo.location(QLibraryInfo.PluginsPath)
print("QLibraryInfo PluginsPath =", p)
print("platforms dir exists =", os.path.isdir(os.path.join(p, "platforms")))
print("qwindows exists =", os.path.exists(os.path.join(p, "platforms", "qwindows.dll")))

app = QApplication(sys.argv)
print("libraryPaths =", QCoreApplication.libraryPaths())