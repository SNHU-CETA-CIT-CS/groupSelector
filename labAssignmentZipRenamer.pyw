#!/usr/bin/env python
import os
from sys import argv, exit
from os import path
from logging import basicConfig, getLogger, DEBUG, INFO, CRITICAL
from pickle import dump, load
import zipRenamerResources_rc
from PyQt5 import QtGui, uic
from PyQt5.QtCore import pyqtSlot, QSettings, Qt, QDir, QCoreApplication
from PyQt5.QtWidgets import QMainWindow, QApplication, QDialog, QMessageBox, QFileSystemModel

startingDummyVariableDefault = 100
namingPatternDefault = 'full'   # options: name, name-email or full
removeZipFilesDefault = False
logFilenameDefault = 'zipRenamer.log'
rootFolderNameDefault = './'
createLogFileDefault = False
pickleFilenameDefault = ".zipRenamerSavedObjects.pl"
firstVariableDefault = 42
secondVariableDefault = 99
thirdVariableDefault = 2001


class ZipRenamer(QMainWindow):
    """A zip file extraction & renaming utility for zyBooks project lab file downloads"""

    def __init__(self, parent=None):
        """ Build a GUI  main window for zipRenamer."""

        super().__init__(parent)
        self.logger = getLogger("Fireheart.zipRenamer")
        self.appSettings = QSettings()
        self.quitCounter = 0;       # used in a workaround for a QT5 bug.

        self.pickleFilename = pickleFilenameDefault
        self.restoreSettings()

        try:
            self.pickleFilename = self.restoreApp()
        except FileNotFoundError:
            self.restartApp()

        uic.loadUi("zipRenamerMainWindow.ui", self)

        self.textOutputUI.appendPlainText("Hello")
        self.textOutputUI.repaint()
        self.dummyVariable = True
        self.textOutput = ""

        self.preferencesSelectButton.clicked.connect(self.preferencesSelectButtonClickedHandler)
        self.rootFolderSelectButton.clicked.connect(self.rootFolderSelectButtonClickedHandler)
        self.currentRootFilePathLabel.clicked.connect(self.rootFolderSelectButtonClickedHandler)
        self.convertButton.clicked.connect(self.convertButtonClickedHandler)

    def __str__(self):
        """String representation for zipRenamer.
        """
        return "A zip file extraction & renaming utility for zyBooks project lab file downloads"

    def updateUI(self):
        self.textOutputUI.setPlainText(self.textOutput)
        self.textOutputUI.repaint()
        self.currentRootFilePathLabel.setText(self.rootFolderName)

    def setRootFolderName(self, newFolderName):
        self.rootFolderName = newFolderName
        self.appSettings.setValue('rootFolderName', self.rootFolderName)

    def getRootFolderName(self):
        return self.rootFolderName

    def restartApp(self):
        if self.createLogFile:
            self.logger.debug("Restarting program")

    def saveApp(self):
        if self.createLogFile:
            self.logger.debug("Saving program state")
        saveItems = (self.dummyVariable)
        if self.appSettings.contains('pickleFilename'):
            with open(path.join(path.dirname(path.realpath(__file__)),  self.appSettings.value('pickleFilename', type=str)), 'wb') as pickleFile:
                dump(saveItems, pickleFile)
        elif self.createLogFile:
                self.logger.critical("No pickle Filename")

    def restoreApp(self):
        if self.appSettings.contains('pickleFilename'):
            self.appSettings.value('pickleFilename', type=str)
            with open(path.join(path.dirname(path.realpath(__file__)),  self.appSettings.value('pickleFilename', type=str)), 'rb') as pickleFile:
                return load(pickleFile)
        else:
            self.logger.critical("No pickle Filename")

    def restoreSettings(self):
        # Restore settings values, write defaults to any that don't already exist.
        if appSettings.contains('createLogFile'):
            self.createLogFile = appSettings.value('createLogFile')
        else:
            self.createLogFile = createLogFileDefault
            appSettings.setValue('createLogFile', self.createLogFile)

        if self.createLogFile:
            self.logger.debug("Starting restoreSettings")

        if self.appSettings.contains('rootFolderName'):
            self.rootFolderName = self.appSettings.value('rootFolderName', type=str)
        else:
            self.rootFolderName = rootFolderNameDefault
            self.appSettings.setValue('rootFolderName', self.rootFolderName)

        if self.appSettings.contains('logFilename'):
            self.logFilename = self.appSettings.value('logFilename', type=str)
        else:
            self.logFilename = logFilenameDefault
            self.appSettings.setValue('logFilename', self.logFilename)

        if self.appSettings.contains('namingPattern'):
            self.namingPattern = self.appSettings.value('namingPattern', type=str)
        else:
            self.namingPattern = namingPatternDefault
            self.appSettings.setValue('namingPattern', self.namingPattern)

        if self.appSettings.contains('removeZipFiles'):
            self.removeZipFiles = self.appSettings.value('removeZipFiles', type=bool)
        else:
            self.removeZipFiles = removeZipFilesDefault
            self.appSettings.setValue('removeZipFiles', self.removeZipFiles)

        if self.appSettings.contains('logFile'):
            self.logFilename = self.appSettings.value('logFile', type=str)
        else:
            self.logFilename = logFilenameDefault
            self.appSettings.setValue('logFile', self.logFilename )

        if self.appSettings.contains('pickleFilename'):
            self.pickleFilename = self.appSettings.value('pickleFilename', type=str)
        else:
            self.pickleFilename = pickleFilenameDefault
            self.appSettings.setValue('pickleFilename', self.pickleFilename)

    def convertButtonClickedHandler(self):
        from zipfile import ZipFile
        self.textOutput = f"Unzipped and renamed the following Files in folder\n  {self.getRootFolderName()}:\n"
        for root, dirs, files in os.walk(self.getRootFolderName()):
            if len(files) > 0:
                self.textOutput += f"\n      Subfolder{path.basename(root)}:\n"
                for file in files:
                    if len(file) > 0:
                        if type(file) is tuple:
                            self.logger.critical(f"File type mismatch!\n Root: {root}\nDirs: {dirs}\nFiles: {files}")
                        fullFilename = path.join(root, file)
                        if fullFilename.endswith('.zip'):
                            # opening the zip file in READ mode
                            with ZipFile(fullFilename, 'r') as zipArchive:
                                for zippedFile in zipArchive.namelist():
                                    if zippedFile.endswith(".py"):
                                        try:
                                            zipArchive.extract(zipArchive.getinfo(zippedFile), root)
                                            if self.namingPattern == "full":
                                                newFilename = path.join(root, f"{file[:-4]}.py")
                                            elif self.namingPattern.startswith("name"):
                                                if file.count('_') == 4:
                                                    lastname, firstName, emailAddress, year, datePlus = file.split('_')
                                                elif file.count('_') == 5:
                                                    lastname, middleName, firstName, emailAddress, year, datePlus = file.split('_')
                                                else:
                                                    raise ValueError
                                                if self.namingPattern == "name":
                                                    newFilename = path.join(root, f"{firstName}{lastname}.py")
                                                elif self.namingPattern == "name-email":
                                                    newFilename = path.join(root, f"{firstName}{lastname}_{emailAddress}.py")
                                            else:
                                                self.logger.critical(f"Unknown file naming pattern {self.namingPattern}")
                                            os.rename(path.join(root, zippedFile), newFilename)
                                            self.textOutput += f"        {path.basename(newFilename)}\n"
                                            self.logger.info(f"UnZipped: {path.basename(zippedFile)} as {path.basename(newFilename)} from Folder {path.basename(root)}")
                                            if self.removeZipFiles:
                                                os.remove(fullFilename)
                                                self.logger.info(f"Removed: {fullFilename}")
                                        except (FileNotFoundError, FileExistsError) as errorObj:
                                            print(errorObj)
                                        except ValueError:
                                            print(f"ValueError on file {file}")
                    self.updateUI()
        self.updateUI()

    @pyqtSlot()  # User is requesting a top level folder select.
    def rootFolderSelectButtonClickedHandler(self):
        folderDialog = FolderSelectDialog(self.getRootFolderName())
        folderDialog.show()
        folderDialog.exec_()
        self.updateUI()

    @pyqtSlot()  # User is requesting preferences editing dialog box.
    def preferencesSelectButtonClickedHandler(self):
        if self.createLogFile:
            self.logger.info("Setting preferences")
        preferencesDialog = PreferencesDialog()
        preferencesDialog.show()
        preferencesDialog.exec_()
        self.restoreSettings()              # 'Restore' settings that were changed in the dialog window.
        self.updateUI()

    @pyqtSlot()				# Player asked to quit the game.
    def closeEvent(self, event):
        if self.createLogFile:
            self.logger.debug("Closing app event")
        if self.quitCounter == 0:
            self.quitCounter += 1
            quitMessage = "Are you sure you want to quit?"
            reply = QMessageBox.question(self, 'Message', quitMessage, QMessageBox.Yes, QMessageBox.No)

            if reply == QMessageBox.Yes:
                self.saveApp()
                event.accept()
            else:
                self.quitCounter = 0
                event.ignore()


class FolderSelectDialog(QDialog):
    def __init__(self, startingFolderName, parent = ZipRenamer):
        super(FolderSelectDialog, self).__init__()
        uic.loadUi('rootSelectDialog.ui', self)
        fileModel = QFileSystemModel()
        fileModel.setFilter(QDir.Dirs | QDir.NoDot | QDir.NoDotDot)
        fileModel.setRootPath(startingFolderName)
        self.selectedPath = startingFolderName
        self.fileSelectTreeView.setModel(fileModel)
        self.fileSelectTreeView.resizeColumnToContents(0)
        self.fileSelectTreeView.expand(fileModel.index(startingFolderName))

        self.fileSelectTreeView.doubleClicked.connect(self.fileDoubleClickedHandler)
        self.buttonBox.rejected.connect(self.cancelClickedHandler)
        self.buttonBox.accepted.connect(self.okayClickedHandler)
        self.fileSelectTreeView.clicked.connect(self.selectionChangedHandler)
        self.fileSelectTreeView.expanded.connect(self.itemExpandedHandler)

    # @pyqtSlot()
    def fileDoubleClickedHandler(self, signal):
        # print(signal)
        filePath = self.fileSelectTreeView.model().filePath(signal)
        if path.isdir(filePath):
            # print(filePath)
            ZipRenamerApp.setRootFolderName(filePath)
            self.close()
        else:
            print("File selected, not directory")

    # @pyqtSlot()
    def okayClickedHandler(self):
        # print(self.selectedPath)
        if path.isdir(self.selectedPath):
            ZipRenamerApp.setRootFolderName(self.selectedPath)
            self.close()
        else:
            print("File Selected on Okay")

    # @pyqtSlot(QItemSelection)
    def selectionChangedHandler(self, selected):
        self.fileSelectTreeView.resizeColumnToContents(0)
        # print(selected)
        if path.isdir(self.fileSelectTreeView.model().filePath(selected)):
            self.selectedPath = self.fileSelectTreeView.model().filePath(selected)
            # print(self.selectedPath)
        else:
            print("File Selected")

    # @pyqtSlot()
    def itemExpandedHandler(self, selected):
        self.fileSelectTreeView.resizeColumnToContents(0)

    # @pyqtSlot()
    def cancelClickedHandler(self):
        self.close()


class PreferencesDialog(QDialog):
    def __init__(self, parent=ZipRenamer):
        super(PreferencesDialog, self).__init__()

        uic.loadUi('preferencesDialog.ui', self)
        self.logger = getLogger("Fireheart.zipRenamer")

        self.appSettings = QSettings()
        if self.appSettings.contains('logFilename'):
            self.logFilename = self.appSettings.value('logFilename', type=str)
        else:
            self.logFilename = logFilenameDefault
            self.appSettings.setValue('logFilename', self.logFilename)

        if self.appSettings.contains('namingPattern'):
            self.namingPattern = self.appSettings.value('namingPattern', type=str)
        else:
            self.namingPattern = namingPatternDefault
            self.appSettings.setValue('namingPattern', self.namingPattern)

        if self.appSettings.contains('removeZipFiles'):
            self.removeZipFiles = self.appSettings.value('removeZipFiles', type=bool)
        else:
            self.removeZipFiles = removeZipFilesDefault
            self.appSettings.setValue('removeZipFiles', self.removeZipFiles)

        if self.appSettings.contains('logFile'):
            self.logFilename = self.appSettings.value('logFile', type=str)
        else:
            self.logFilename = logFilenameDefault
            self.appSettings.setValue('logFile', self.logFilename)

        if self.appSettings.contains('createLogFile'):
            self.createLogFile = self.appSettings.value('createLogFile')
        else:
            self.createLogFile = logFilenameDefault
            self.appSettings.setValue('createLogFile', self.createLogFile )

        self.buttonBox.rejected.connect(self.cancelClickedHandler)
        self.buttonBox.accepted.connect(self.okayClickedHandler)
        self.logFilenameUI.editingFinished.connect(self.logFilenameChanged)
        self.namingPatternNameUI.toggled.connect(self.nameOnlySelected)
        self.namingPatternNameEmailUI.toggled.connect(self.nameEmailSelected)
        self.namingPatternFullUI.toggled.connect(self.nameFullSelected)
        self.removeZipFilesUI.stateChanged.connect(self.removeZipFilesChanged)
        self.createLogFileUI.stateChanged.connect(self.createLogFileChanged)

        self.updateUI()

    # @pyqtSlot()
    def logFilenameChanged(self):
        self.logFilename = self.logFilenameUI.text()

    # @pyqtSlot()
    def nameOnlySelected(self, selected):
        if selected:
            self.namingPattern = "name"

    # @pyqtSlot()
    def nameEmailSelected(self, selected):
        if selected:
            self.namingPattern = "name-email"

    # @pyqtSlot()
    def nameFullSelected(self, selected):
        if selected:
            self.namingPattern = "full"

    # @pyqtSlot()
    def removeZipFilesChanged(self):
        self.removeZipFiles = self.removeZipFilesUI.isChecked()

    # @pyqtSlot()
    def createLogFileChanged(self):
        self.createLogFile = self.createLogFileCUI.isChecked()

    def updateUI(self):
        self.logFilenameUI.setText(str(self.logFilename))
        self.namingPatternNameUI.setChecked(False)
        self.namingPatternNameEmailUI.setChecked(False)
        self.namingPatternFullUI.setChecked(False)
        if self.namingPattern == "name":
            self.namingPatternNameUI.setChecked(True)
        elif self.namingPattern == "name-email":
            self.namingPatternNameEmailUI.setChecked(True)
        if self.namingPattern == "full":
            self.namingPatternFullUI.setChecked(True)

        if self.removeZipFiles:
            self.removeZipFilesUI.setCheckState(Qt.Checked)
        else:
            self.removeZipFilesUI.setCheckState(Qt.Unchecked)
        if self.createLogFile:
            self.createLogFileUI.setCheckState(Qt.Checked)
        else:
            self.createLogFileUI.setCheckState(Qt.Unchecked)

    # @pyqtSlot()
    def okayClickedHandler(self):
        # write out all settings
        preferencesGroup = (('logFilename', self.logFilename),
                            ('namingPattern', self.namingPattern),
                            ('removeZipFiles', self.removeZipFiles),
                            ('createLogFile', self.createLogFile),
                            )
        # Write settings values.
        for setting, variableName in preferencesGroup:
            self.appSettings.setValue(setting, variableName)
        self.close()

    # @pyqtSlot()
    def cancelClickedHandler(self):
        self.close()


if __name__ == "__main__":
    QCoreApplication.setOrganizationName("Fireheart Software");
    QCoreApplication.setOrganizationDomain("fireheartsoftware.com");
    QCoreApplication.setApplicationName("zipRenamer");
    appSettings = QSettings()
    if appSettings.contains('createLogFile'):
        createLogFile = appSettings.value('createLogFile')
    else:
        createLogFile = createLogFileDefault
        appSettings.setValue('createLogFile', createLogFile)

    if createLogFile:
        startingFolderName = path.dirname(path.realpath(__file__))
        if appSettings.contains('logFile'):
            logFilename = appSettings.value('logFile', type=str)
        else:
            logFilename = logFilenameDefault
            appSettings.setValue('logFile', logFilename)
        basicConfig(filename=path.join(startingFolderName, logFilename), level=INFO,
                    format='%(asctime)s %(name)-8s %(levelname)-8s %(message)s')
    app = QApplication(argv)
    ZipRenamerApp = ZipRenamer()
    ZipRenamerApp.updateUI()
    ZipRenamerApp.show()
    exit(app.exec_())
