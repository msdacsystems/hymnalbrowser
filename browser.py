# -*- coding: utf-8 -*-

"""
-------------------------------
MSDAC Systems - Hymnal Browser
-------------------------------

A general hymnal browser for Seventh-day Adventist Church (formerly for Montevilla SDA Church)
Supports up to 474 Hymns based on SDA Hymnal Philippine Edition.

This module provides all functionality of operating and executing .pptx files,
configurations, and statistical data.

Pre-requisites:
    - Microsoft(c) Office 2016 and above

Spacing Gap Formats:
    • Class - 3 to 4 lines
    • Functions - 2 to 3 lines
    • Comments - 1 line

Tested working on:
    Windows 10 -
        • MS Office 2013
        • MS Office 2019
        • MS Office 2021
    Windows 11 -
        • MS Office 2021

This software is part of MSDAC System's collection of softwares
(c) 2021-present Ken Verdadero, Reynald Ycong
"""

## Import Modules
import sys, os
from typing import Type

try:
    import configparser, time, json, subprocess, keyboard, shutil, psutil, operator, gc
    import humanize, csv, dotenv, pymongo, requests, socket, bz2, winreg
    from kenverdadero.KCore import KPath, KString, KTime, KSystem
    from kenverdadero.KCore.KCore import invert, modHex, convertDataUnit, getFileStat, calcTimeExec, nL, p, tP, getFilename, getSize
    from kenverdadero.KSoftware import KSoftware
    from kenverdadero.KLogging import KLog
    from datetime import datetime, timedelta
    from collections import namedtuple
    from subprocess import DEVNULL
    from pptx import Presentation
    from zipfile import ZipFile
    from threading import Thread
    from PyQt5.QtCore import Qt, QEasingCurve
    from PyQt5.QtGui import QFont
    from PyQt5 import QtCore, QtGui, QtWidgets
    from BlurWindow import blurWindow
except ImportError as e:
    print('System Error: ' + str(e))
    sys.exit()




class System():
    """
    System Class Handler
    
    This handles all system responsibilities for the browser such as:
        - System properties
        - Verifying all required files and directory
        - Duplicate instances
        - Background Tasks
        - Application exit event
    """
    def __init__(self):
        """
        All variables have the following format:
            self.TYPE_NAME_ATTRIBUTE_SUBATTR
            
            Example: self.SOFTWARE_HYMNALBROWSER_UI_SETTINGS
        """
        _DIR_TEMP_SUB = namedtuple("DIR_TEMP_SUB", "EN TL")
        _RECENTS = namedtuple("RECENTS", "DEFAULT ALLOWEDMIN ALLOWEDMAX")

        ## External File and Directories
        self.DIR_PARENT =       r'C:\ProgramData\MSDAC Systems'
        self.DIR_PROGRAM =      self.DIR_PARENT + r'\Hymnal Browser'
        self.DIR_TEMP =         self.DIR_PROGRAM + r'\Temp'
        self.DIR_LOG =          self.DIR_TEMP + r'\Logs'
        self.DIR_TEMP_SUB =     _DIR_TEMP_SUB(self.DIR_TEMP + r'\EN', self.DIR_TEMP + r'\TL')
        self.FILE_DATA =        self.DIR_PROGRAM + r'\data.json'
        self.FILE_CONFIG =      self.DIR_PROGRAM + r'\config.ini'
        self.FILE_HYMNSDB =     os.path.join(SW.DIR_CWD, 'hymns.sda')

        ## Internal Directories
        self.DIR_RES =          'res'
        self.DIR_BIN =          'bin'

        ## Resources
        self.RES_LOGO =         './res/images/logo.png'
        self.RES_FONT_TITLE =   'IntegralCF-Regular.otf'
        
        ## Properties
        self.CURR_THEME = 0
        self.HYMNS_MAX = 474                                                                                                ## Number of Maximum Hymns (Will be deprecated soon due to customized hymns in future updates)
        self.RECENTS = _RECENTS(10, 3, 30)                                                                                  ## Recent Files Properties: (Default, Minimum Allowed, Maximum Allowed) 
        self.CPLTR_MAX_VISIBLE_ITEMS = 10                                                                                   ## Number of files to be maintained in TEMP_DIR
        self.PROCESS_NAME = "hymnalbrowser.exe"                                                                             ## File name 
        self.USER_NAME = os.getenv('username')                                                                                ## Get Username of the windows account
        self.TBL_STATS_COLUMNS = 5
        self.MIN_OPACITY = 50                                                                                               ## Minimum Opacity of the application
        self.PROCESS = psutil.Process(os.getpid())
        self.LOG_FILE_LIMIT = 10
        self.FORCE_OFFLINE = False
        self.EXT_FEEDBACK = 'fdback'
        self.EXT_TELEMETRY = 'tlm'
        self.CNT_SESSION_PRESN = 0                                                                                          ## Session Presentation (PPT) Counter
        self.STARTUP_TIME = 0
        self.UNIQ_MACHINE_ID = subprocess.check_output('wmic csproduct get uuid').decode().split('\n')[1].strip()
        self.HOST_NAME = socket.gethostname()

        ## Functions
        dotenv.load_dotenv(self.DIR_BIN + '/secrets.env')


    def verifyDirectories(self):
        """
        This method checks for directories and also generate new if the folders does not exist/
        This verifies from parent directory down to subfolders via loop using a dict of directories.

        .. - Root
        <> - File
        -> - Directory

        Tree:
            .. MSDAC Systems (Parent/Root)
                -> Hymnal Browser (Program)
                    -> Temp (Temporary Folder)
                        -> English
                        -> Logs
                        -> Tagalog
                        <> Hymnal Files (ends with .pptx)
                    <> config.ini

            -> POWERPNT.exe (MS Office)
        """
        DIRECTORIES = {
            self.DIR_PARENT: "Parent",
            self.DIR_PROGRAM: "Program",
            self.DIR_TEMP: "Temporary",
            self.DIR_TEMP_SUB.EN: "English",
            self.DIR_TEMP_SUB.TL: "Tagalog",
            self.DIR_LOG: "Logs"
            }
        MISSING = 0

        for DIR, NAME in DIRECTORIES.items():
            if not KPath.exists(DIR, True): LOG.warn(f"{NAME} Directory \"{DIR}\" doesn't exist. Creating a new folder."); MISSING += 1
        
        ## If Hymnal Database cannot be found
        if not KPath.exists(self.FILE_HYMNSDB):
            LOG.crit(f"Cannot find Hymnal database")
            MSG_BOX = QtWidgets.QMessageBox(); MSG_BOX.setStandardButtons(QtWidgets.QMessageBox.Ok); MSG_BOX.setIcon(QtWidgets.QMessageBox.Critical)
            MSG_BOX.setWindowTitle("Error"); MSG_BOX.setText(f"Database Failure: Hymnal Database is missing.\nPlease contact the developers of this program to fix the issue.")
            MSG_BOX.setDetailedText(f"Hymnal Database is non existent in this folder:\n\n{self.FILE_HYMNSDB}")
            MSG_BOX.setWindowFlags(Qt.Drawer | Qt.WindowStaysOnTopHint)
            MSG_BOX.setStyleSheet(Stylesheet().getStylesheet())
            MSG_BOX.exec_()
            LOG.sys("Program terminated due to an error.")
            sys.exit()

        LOG.sys(f"Verification complete. {MISSING} {KString.isPlural('folder', MISSING)} flagged as missing") if MISSING else LOG.sys("Successfully verified all directories.")
        if MISSING == 6: LOG.sys(f"Detected a first time launch of the application.") ## Needs improvement


    def verifyRequisites(self):
        """
        Verifies all required programs to make the software run properly.
        This method prevents the program to proceed if a valid Microsoft Office PowerPoint is not found in both 32-bit and 64-bit
        """
        try:
            self.PPT_EXEC = KPath.upFolder(winreg.EnumValue(winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe"), 0)[1])
            LOG.info(f'Using {"64" if "Program Files (x86)" not in self.PPT_EXEC else "32"}-Bit Version of Microsoft Office')
            LOG.info(f'Root Dir: {self.PPT_EXEC}')
        except FileNotFoundError:
            ## No MS Office Installed
            LOG.crit("Microsoft Office is not installed")
            MSG_BOX = QtWidgets.QMessageBox()
            MSG_BOX.setStandardButtons(QtWidgets.QMessageBox.Ok)
            MSG_BOX.setIcon(QtWidgets.QMessageBox.Critical)
            MSG_BOX.setText("Microsoft Office is not installed!\nThis program uses PowerPoint 2010 or later to run properly. \n")
            MSG_BOX.setDetailedText("If you think this is a mistake, please contact the developers for further assistance.\n\nhttps://m.me/verdaderoken\nhttps://m.me/reynald.ycong")
            MSG_BOX.setWindowTitle("Error")
            MSG_BOX.setWindowFlags(Qt.Drawer | Qt.WindowStaysOnTopHint)
            
            MSG_BOX.setStyleSheet(QSS.getStylesheet())
            MSG_BOX.exec_()
            LOG.sys("Program terminated due to an error.")
            sys.exit()
        return
        

    def checkInstances(self):
        ## Asks the user if they want to run another instance
        self.DUPLICATED = False
        
        DUPLICATES = [i for i in str(subprocess.check_output(['wmic', 'process', 'list', 'brief'])).split() if str(i) == SYS.PROCESS_NAME]
        # DUPLICATES = len(list(filter(lambda x: x == SYS.PROCESS_NAME, [i.name() for i in psutil.process_iter()]))) ## < -- Old
        
        if len(DUPLICATES) > 2:                                                                                                          ## Set to threshold of 2 because .exe has 2 processes when built
            MSG_BOX = QtWidgets.QMessageBox()
            MSG_BOX.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            MSG_BOX.setIcon(QtWidgets.QMessageBox.Question)
            MSG_BOX.setText("The program is already running.\nDo you want to open another instance?")
            MSG_BOX.setWindowTitle("Duplicate Instance Detected")
            MSG_BOX.setWindowFlags(Qt.Drawer | Qt.WindowStaysOnTopHint)
            MSG_BOX.setStyleSheet(QSS.getStylesheet())
            MSG_BOX.setStyleSheet('min-width: 250px; min-height: 30px;')
            
            LOG.info("Duplicate Instance Detected")
            RET = MSG_BOX.exec_()
            if RET == QtWidgets.QMessageBox.Yes:
                return
            else:
                LOG.sys("Program terminated by user.")
                self.DUPLICATED = True
                sys.exit()


    def isOnline(self):
        """
        Checks if the software can connect to the internet
        """
        # return True
        try:
            with socket.create_connection(("www.mongodb.com", 80)) as sock:
                if sock is not None: sock.close(); return True if not SYS.FORCE_OFFLINE else False
        except OSError:
            return False


    def startBackgroundTask(self):
        """
        Initiates the background thread
        """
        self.WORKER_1 = ThreadBackground()
        self.WORKER_1.start()
        self.WORKER_1.UPT.connect(lambda: self.WORKER_1.loopFunction())


    def QCl(self, c):
        if '#' in c: c = c[1:]
        return QtGui.QColor(int(c[0:2],16),int(c[2:4],16),int(c[4:6],16))


    def RGBtoHEX(self, rgb):
        rgb = list(rgb); del rgb[3]
        return '#%02x%02x%02x' % tuple(rgb)

    
        
    def centerWindow(self, ui):
        """
        Relocates the specific UI from argument to center of the screen
        """
        FRM_GEOMETRY = ui.frameGeometry()
        SCREEN = QtWidgets.QApplication.desktop().screenNumber(QtWidgets.QApplication.desktop().cursor().pos())
        FRM_GEOMETRY.moveCenter(QtWidgets.QApplication.desktop().screenGeometry(SCREEN).center())
        ui.move(FRM_GEOMETRY.topLeft())


    def centerInsideWindow(self, WGTA, WGTB):
        """
        Relocates the Widget A centered with relation to Widget B
        """
        WGTA.show()
        WINDOW = (WGTB.frameGeometry().width()-WGTA.frameGeometry().width(),
                WGTB.frameGeometry().height()-WGTA.frameGeometry().height())
        WGTA.move(WGTB.pos().x()+int(WINDOW[0]/2), WGTB.pos().y()+int(WINDOW[1]/2.2))

    
    def closeEvent(self, event):
        """
        Triggers this function when the UI is shut down
        Reports back to the logger.
        """
        
        if self.DUPLICATED:
            LOG.sys(f"Duplicate instance was shut down.")
        else:
            LOG.sys(f"Shutting down")
            UIB.close()
            UIFB.close()
            ## Send Exit Data
            MDB.sendExitData()
            LOG.sys(f"Program terminated: Duration ({SW.runtime(2)} Seconds)")
            LOG.sys(f"END OF LOG - {datetime.now().strftime('%B %d, %Y - %I:%M:%S %p')}")


    def windowBlur(self, mode):
        """
        Handles blur effect of a window
        """
        try:
            # return ## < --- Temporarily placed due to Blur compatibility issues with Windows 11
            blurWindow.blur(UIA.winId(), False, True, True)
            blurWindow.blur(UIB.winId(), False, True, True)
            blurWindow.blur(UIC.winId(), False, True, True)
        except NameError:
            pass




class Mongo():
    """
    Handles all online-related databases used for reports/feedbacks, analytics, and other necessary data.
    """
    def __init__(self) -> None:
        self.CLUSTER = None
        self.DB = None
        self.COL_GLOBAL = None
        self.COL_REPORTS = None
        self.COL_CLIENT_DATA = None

        self.CLIENT_QUERY = {'_id': SYS.UNIQ_MACHINE_ID}
        self.GLOBAL_QUERY = {'_id' : 'GLOBAL'}
        self.REPORT_INITIAL = 0                                                                 ## Report Switch for Initialization
        self.REPORT_TASK = 0                                                                    ## Report Switch for Task
        self.WORKER = self.MongoReport()
        self.WORKER.start()


    class MongoReport(QtCore.QThread):
        def run(self):
            while True:
                if MDB.REPORT_INITIAL:
                    START = time.time()
                    try:
                        ## Database
                        MDB.CLUSTER = pymongo.MongoClient(os.environ.get('MONGO_CLIENT_API'))
                        MDB.DB = MDB.CLUSTER['HymnalBrowser']
                        ## Collections or Datasets
                        MDB.COL_GLOBAL = MDB.DB['Global']
                        MDB.COL_REPORTS = MDB.DB['Reports']
                        MDB.COL_CLIENT_DATA = MDB.DB['ClientData']

                        MDB.reportGlobal()
                        MDB.reportClientData()
                        MDB.clientIsVerified(True)
                        MDB.checkPendingTelemetry(SYS.isOnline())
                    except KeyError as e:
                        ## Insert repair code here
                        LOG.debug(f'Key Error at "{e}". Needs changes')
                        LOG.error(f'Application needs to terminate due to an error.')
                        sys.exit()
                    except pymongo.errors.ConfigurationError as e:
                        LOG.debug(e)

                    # MDB.reportHymnalData()
                    LOG.info(f'Report finished in {round(time.time()-START,3)} s.')
                    MDB.REPORT_INITIAL = 0
                

                if MDB.REPORT_TASK:
                    ## Check if the client has working internet connection
                    if SYS.isOnline():
                        if not SYS.FORCE_OFFLINE:
                            MDB.checkPendingTelemetry(True)
                            UIFB.TME_START = time.time()
                            ONLINE = True
                        else:
                            ONLINE = False
                    else:
                        LOG.info("Can't connect to the database. Exporting offline report.")
                        ONLINE = False

                    
                    PUBLIC_IP = requests.get('https://api.ipify.org').text if ONLINE else 'Unavailable'
                    REPORT_METHOD = 'Online' if ONLINE else 'Offline'
                    TIMESTAMP = SW.DATE_NOW()
                    STATUS = 'OK'

                    REPORT = {
                        "_id": time.time(),
                        "_tms": str(datetime.now().strftime('%F %I:%M:%S %p')),
                        "username": SYS.USER_NAME,
                        "status": STATUS,
                        "DATA": {
                            "client": {
                                "reportMethod": REPORT_METHOD,
                                "feedback": str(UIFB.USER_TEXT.encode('utf-8')),
                                "hostname": SYS.HOST_NAME,
                                "localIPAddress": socket.gethostbyname(SYS.HOST_NAME),
                                "publicIPAddress": PUBLIC_IP,
                            },
                            "software": {
                                "logFile": LOG.getContents(),
                                "programDataSize": getSize(SYS.DIR_PROGRAM),
                                "timeElapsed": SW.runtime(),
                                "timeLaunched": (time.time()-SW.runtime()),
                                "timestamp": time.time(),
                                "version": SW.VERSION,
                                "versionName": SW.VERSION_NAME,
                                "sessionLaunchedFiles": SYS.CNT_SESSION_PRESN,
                            },
                            "system": {
                                "assessment": KSystem.getSystemAssessment(),
                                "configuration": list(CFG.CONFIG[CFG.HEADNAME].items()),
                                "env": dict(os.environ.items()),
                                "info": KSystem.getSystemInfo(),
                                "pathPPT": SYS.PPT_EXEC,
                            }
                        }
                    }
                    REPORT['DATA']['system']['env'].pop('MONGO_CLIENT_API')                         ## Pop out CLIENT_API to hide from database
                    REPORT = json.dumps(REPORT, indent=4, sort_keys=True)                           ## Serialize dict to a JSON str

                    ## Update User Feedback Count
                    MDB.COL_CLIENT_DATA.update_one(MDB.CLIENT_QUERY, {'$set': {'feedbackCount': MDB.getCollData(MDB.COL_CLIENT_DATA, MDB.CLIENT_QUERY)['feedbackCount'] + 1 }})

                    if ONLINE:
                        ## Upload to cloud database
                        LOG.info('Sending feedback and report...')
                        MDB.COL_REPORTS.insert_one(json.loads(REPORT))
                    else:
                        ## Save offline for future use
                        with open(f'{SYS.DIR_TEMP}/{TIMESTAMP}.{SYS.EXT_FEEDBACK}', 'wb') as w:
                            w.write(bz2.compress(REPORT.encode('utf-8'), 9))
                    LOG.info(f'{"Feedback successfully sent" if ONLINE else "Offline report was saved."} ({round((time.time()-UIFB.TME_START), 2)} s)')
                    MDB.REPORT_TASK = 0

                time.sleep(1.5)


    def clientIsVerified(self, dump=False):
        if System().isOnline():
            if SYS.UNIQ_MACHINE_ID in self.getCollData(self.COL_GLOBAL, self.GLOBAL_QUERY)['BASIC']['machines']:
                return True
            elif dump:
                self.COL_GLOBAL.update_one(self.GLOBAL_QUERY, {"$push": {"BASIC.machines": SYS.UNIQ_MACHINE_ID}})
                LOG.sys(f'Client data for {SYS.UNIQ_MACHINE_ID} is now verified')
                return False
            

    def checkPendingTelemetry(self, online=False):
        """
        Checks for pending reports that can be delivered
        """
        ## [1] Check for pending System Telemetry
        # TLMTRY = [f'{SYS.DIR_TEMP}\\{i}' for i in os.listdir(SYS.DIR_TEMP) if i.endswith(f'.{SYS.EXT_TELEMETRY}')]
        ## < Insert offline telemetry report here >
        ## < Insert offline telemetry report here >
        ## < Insert offline telemetry report here >

        ## [2] Check for pending User Feedbacks
        FBCKS = [f'{SYS.DIR_TEMP}\\{i}' for i in os.listdir(SYS.DIR_TEMP) if i.endswith(f'.{SYS.EXT_FEEDBACK}')]
        if len(FBCKS):
            if not online: LOG.info(f'There are {len(FBCKS)} pending report(s) but cannot be delivered due to offline.'); return
            
            LOG.info(f'Scanned {len(FBCKS)} pending report(s)')
            ## Decompress 
            for i, file in enumerate(FBCKS):
                try:
                    with open(file, 'rb') as f:
                        self.COL_REPORTS.insert_one(json.loads(bz2.decompress(f.read()).decode('utf-8')))
                        LOG.info(f'[{i+1}/{len(FBCKS)}] \'{file}\' Report was successfully sent')
                except (OSError, json.decoder.JSONDecodeError, pymongo.errors.DuplicateKeyError) as e:
                    LOG.debug(e)
                    try: os.remove(file); LOG.info(f'[{i+1}/{len(FBCKS)}] Deleted a bad or invalid data report file')
                    except PermissionError: pass
                try: os.remove(file)
                except PermissionError: pass
    

    def createNewData(self, collTarget):
        """
        Create a new data based on collection target parameter
        """
        if System().isOnline():
            ## CLIENT DATA
            if collTarget == self.COL_CLIENT_DATA:
                DEFAULTS = {
                    '_id': SYS.UNIQ_MACHINE_ID,
                    '_initiated': time.time(),
                    'systemLaunchCount': 0,
                    'presnLaunchCount': 0,
                    'lastUpdated': time.time(),
                    'usageSince': 0,
                    'feedbackCount': 0,
                    '_username': SYS.USER_NAME,
                    '_hostname': SYS.HOST_NAME,
                    'PACKAGE': {
                        'hymnal': Data().load(),
                    }
                }
                DEFAULTS = json.dumps(DEFAULTS, indent=4, sort_keys=True)
                self.COL_CLIENT_DATA.insert_one(json.loads(DEFAULTS))
                                
                if not self.clientIsVerified(True): LOG.warn(f'Old client data was detected missing by the system.')

            ## GLOBAL DATA 
            if collTarget == self.COL_GLOBAL:
                LOG.crit('Global data is inaccessible.')
                DEFAULTS = {
                    '_id': 'GLOBAL',
                    'BASIC': {
                        'launches': 0,
                        'launchTimes': [0,0,0,0],
                        'machines': [],
                        'recentLaunchTimestamp': time.time()
                    },
                    'CONFIG': {
                        'versionMin': SW.VERSION,
                        'opacity': [50, 100],
                    }
                }
                self.COL_GLOBAL.insert_one(DEFAULTS)
                LOG.sys('Global data was generated')

        else:
            LOG.info('Cannot create a client record. (Reason: Unable to connect to the server.)')


    def getCollData(self, collectionName, query):
        """
        Returns a collection data as a dict type object based from a name given in parameter
        """
        while True:
            if System().isOnline():
                try: DATA = getattr(collectionName, 'find')(query)[0]
                except IndexError:
                    self.createNewData(collectionName)
                    continue
                else:
                    return DATA


    def reportGlobal(self):
        """
        Reports to global data
        """
        while True:
            if System().isOnline():
                DATA = self.getCollData(self.COL_GLOBAL, self.GLOBAL_QUERY)

                DATA['BASIC']['launchTimes'][0] = SYS.STARTUP_TIME ## CURRENT, MINIMUM, MAXIMUM
                if DATA['BASIC']['launchTimes'][1] == 0: DATA['BASIC']['launchTimes'][1] = SYS.STARTUP_TIME

                ## Sections
                DATA['BASIC']['launches'] += 1
                DATA['BASIC']['recentLaunchTimestamp'] = time.time()

                if SYS.STARTUP_TIME < DATA['BASIC']['launchTimes'][1]: DATA['BASIC']['launchTimes'][1] = SYS.STARTUP_TIME
                if SYS.STARTUP_TIME > DATA['BASIC']['launchTimes'][2]: DATA['BASIC']['launchTimes'][2] = SYS.STARTUP_TIME
                DATA['BASIC']['launchTimes'][3] = DATA['BASIC']['launchTimes'][2] - DATA['BASIC']['launchTimes'][1]

                self.COL_GLOBAL.update_one(self.GLOBAL_QUERY, {"$set": DATA})

            else:
                pass
                ## < Insert Offline Reports here >
                ## < Insert Offline Reports here >
                ## < Insert Offline Reports here >
            break


    def reportClientData(self):
        while True:
            if System().isOnline():
                DATA = self.getCollData(self.COL_CLIENT_DATA, self.CLIENT_QUERY)
                DATA['usageSince'] = time.time() - DATA['_initiated']
                DATA['lastUpdated'] = time.time()
                DATA['_username'], DATA['_hostname'] = SYS.USER_NAME, SYS.HOST_NAME
                DATA['systemLaunchCount'] += 1
                
                self.COL_CLIENT_DATA.update_one(self.CLIENT_QUERY, {"$set": DATA})
            else:
                LOG.debug('Client Report was not sent due to offline.')
            break


    def sendExitData(self):
        #pymongo.errors.ServerSelectionTimeoutError:
        while True:
            if System().isOnline():
                DATA = self.getCollData(self.COL_CLIENT_DATA, self.CLIENT_QUERY)
                DATA['lastUpdated'] = time.time()
                DATA['presnLaunchCount'] += SYS.CNT_SESSION_PRESN
                self.COL_CLIENT_DATA.update_one(self.CLIENT_QUERY, {"$set": DATA})
            else:
                pass
            break


    # def reportHymnalData(self):
    #     """
    #     Sends the SDATA of client to database
    #     """
    #     while True:
    #         if System().isOnline():
    #             QUERY = {'_id': SYS.UNIQ_MACHINE_ID}

    #             try: DATA = self.COL_CLIENT_DATA.find(QUERY)[0]
    #             except IndexError:
    #                 self.createNewData()
    #                 continue
    #             else:
    #                 ## Existing
    #                 DATA['PACKAGE'] = Data().load()
    #                 self.COL_CLIENT_DATA.update_one(QUERY, {"$set": DATA})
                    
    #                 exit()

    #             try: self.COL_CLIENT_DATA.insert_one(json.loads(DATA))
    #             except pymongo.errors.DuplicateKeyError: self.COL_CLIENT_DATA.update_one({'_id': UNIQ_MACHINE_ID}, {'$set': json.loads(DATA)})
    #             LOG.info('Statistical data for hymns was uploaded.')
    #         break




class Configuration():
    """
    Handles all configuration for the software.
    Uses basic configuration parser.

    Default config name: config.ini
    """
    def __init__(self):
        self.DEFAULTS = [
            ('AlwaysOnTop', False),
            ('AutoScroll', True),
            ('AutoSlideshow', False),
            ('CompactMode', False),
            ('KeepFocusOnBrowser', True),
            ('MaxAllowedRecent', 10),
            ('Theme', 0),
            ('WindowOpacity', 100),
            ]
        self.DEFAULTS_OPTIONS = [k[0] for k in self.DEFAULTS]
        self.check()


    def check(self):
        """
        Checks, load, and parse the configuration file
        Automatically fixes missing, corrupted, and bad config headers
        """
        if not os.path.exists(SYS.FILE_CONFIG): LOG.warn("Configuration is missing"); self.generateDefault()
        self.CONFIG = configparser.ConfigParser()                                                               ## Initiate config parser object
        self.CONFIG.optionxform = str                                                                           ## Preserve case of strings
        self.HEADNAME = 'Settings'                                                                              ## Head name for configuration
        PROCEED = False
        while not PROCEED:
            try:
                self.read()
                self.CONFIG[self.HEADNAME]
            except (configparser.MissingSectionHeaderError, configparser.ParsingError,KeyError):                                              ## Resets the configuration file when the configuration is corrupted or had a problem while parsing
                LOG.error('Cannot load configuration file. Resetting to defaults.')
                self.generateDefault()
                self.read()

            ## Test every variables in config
            idx = 0
            try:
                for k in self.DEFAULTS:
                    self.CONFIG[self.HEADNAME][k[0]]; idx += 1
            except KeyError as e:
                LOG.warn(f'Configuration: Missing value for {e}. Fallback will be used.')
                self.CONFIG.set(self.HEADNAME, self.DEFAULTS[idx][0], str(self.DEFAULTS[idx][1])); self.dump()
            else:
                ## Remove unnecessary variables
                UNUSED = [i for i in self.CONFIG.options(self.HEADNAME) if i not in self.DEFAULTS_OPTIONS]      ## Filter two lists by comparing what are the differences.
                if UNUSED:
                    for u in UNUSED:
                        self.CONFIG[self.HEADNAME].pop(u)
                        LOG.info(f'Configuration: Removed unnecessary option: {u}')
                    self.dump()
                PROCEED = True
        LOG.info("Configuration was loaded successfully")


    def read(self):
        """
        Reads the configuration file. Equivalent to loading the file with updated values
        """
        self.CONFIG.read(SYS.FILE_CONFIG)


    def dump(self, data=None):
        """
        Dumps the current config data and executes the read for updating values
        """
        if data is None: pass
        with open(SYS.FILE_CONFIG, 'w') as w:
            self.CONFIG.write(w)
            self.read()                                                                                     ## Read the file again update values


    def getDefaults(self):
        """
        Returns default configuration string
        """
        return (f"[Settings]{nL}{f'{nL}'.join([f'{k[0]} = {k[1]}' for k in self.DEFAULTS])}")


    def generateDefault(self):
        """
        Generates Default Configuration values from defaults
        """
        with open(SYS.FILE_CONFIG, "w") as config:
            config.write(self.getDefaults())
            LOG.info("Configuration file was regenerated with default values")
    



class Data():
    """
    Manages the statistical database for Hymnal Browser.
    Uses JSON to parse the data.

    SDATA (or Statistical Database) is the dictionary of all stats of the hymns.
    """
    def __init__(self):
        self.check()


    def check(self):
        """
        Checks if the file is existent.
        """
        if not os.path.exists(SYS.FILE_DATA):
            LOG.warn("Data is missing. Generating new.")
            self.generateDefault()

        self.verifyContents()
        self.DATA = self.load()


    def load(self):
        """
        Retrieves the data from system's data file.
        """
        while True:
            with open(SYS.FILE_DATA, "r") as read:
                try: CFG = json.load(read)
                except json.decoder.JSONDecodeError:
                    LOG.error('Failed to load statistic data. Regenerating default.')
                    self.generateDefault()
                else:
                    return CFG
    

    def verifyContents(self):
        """
        Verify the statistic file.
        """
        TARGET = self.load()

        try:
            for i in range(SYS.HYMNS_MAX):
                TARGET['DATA'][KString.toDigits(i+1,3)]
                for j in range(3):
                    TARGET['DATA'][KString.toDigits(i+1,3)][j]
        except KeyError as e:
            LOG.crit(f'System has found a corrupted "{KString.filterOnly(1,str(e))}" entry in statistical data. Regenerating new one with default values')
            self.generateDefault()
    

    def dump(self, data=None):
        """
        Saves the passed data to system's data file.
        Uses indention of 4 and sorted keys by default.
        Can be processed with other data if 1st argument is specified.
        """
        if data is None: data = SDATA
        with open(SYS.FILE_DATA, "w") as write:
            json.dump(data, write, indent=None, sort_keys=True)

    
    def generateDefault(self):
        """
        Generates default data. Autofills all 474 Hymns' statistics by default (0 values).
        """
        DATA = {
            "__DATECREATED": time.time(),
            "__FILETYPE__": f"{SW.NAME} Statistical Data",
            "__CHECKSUM__": None,
            "DATA": {str(i+1).zfill(3): [0,0,0] for i in range(HDB.TOTAL_HYMNS)}
        }
        DATA.update({"__CHECKSUM__": KString.toHashMD5(DATA)})
        self.dump(DATA)
        LOG.info(f"Default data was generated successfully. | Hash: {KString.toHashMD5(DATA)}")


    def getStats(self):
        """
        Returns statistics for database panel in Settings
        """
        try:
            STATS = [SDB.load()['DATA'][KString.toDigits(i+1,3)] for i in range(SYS.HYMNS_MAX)]
        except KeyError as e:
            LOG.error(e)
            SDB.check()                                                                                                 ## Recheck the SDATA for possible corrupted keys
        else:
            _STATS = namedtuple("STATS", "queries launches lastAccessed")
            a = sum([STATS[i][0] for i in range(len(STATS))])
            b = sum([STATS[i][1] for i in range(len(STATS))])
            c = timedelta(seconds=max([STATS[i][2] for i in range(len(STATS))])), timedelta(seconds=min([STATS[i][2] for i in range(len(STATS))]))
            return _STATS(a,b,c)




class HymnsDatabase():
    """
    Handles all Hymns from the hymns.sda Database

    HYMNAL is the whole dict of information about the hymns

    For parsing, the method `parseHymnDatabase` will do the following in order:
        - Scan all files from hymns.sda using ZipFile module
        - Splits the hymn filename into useful parts using `splitHymn`; would return (cat, num, title, ext)
        - Appending all matching valid pptx files into its correct list; flags all unused file inside database as unnecessary
        - Also lists all the missing files (by default)
        - Merges all data collected into one dictionary (tuple)
        - Returns the whole namedtuple dictionary (aka. Hymnal)
    
    For retrieving stats, the method `getStats` will do the following in order:
        - Turns off the execution-ready flag to prevent early counting of query
        - Uses Hymn Number as the base for searching.
    """
    def __init__(self):
        self.TOTAL_HYMNS = SYS.HYMNS_MAX


    def splitHymn(self, hymnFile:str):
        """
        Splits the filename of the hymnal by parsing the parts into:
            (Category, Hymn Number, Hymn Title, File Extension)
        """
        HYMN = namedtuple("HYMN", "cat num title ext")
        return HYMN(hymnFile[:2], hymnFile[3:6] if hymnFile[3:6].isdigit() else 0, hymnFile[7:-5], hymnFile.split('.')[-1])


    def parseHymnDatabase(self):
        """
        Retrieves certain information from the database.
        It also counts all missing hymns and scans for unused file inside the database.

        Returns a book "Hymnal" with its properties:
            [0] - English (# and Title)
            [1] - Tagalog (# and Title)
            [2] - User-personal (# and Title)
            [3] - All in one Hymn Titles
            [4] - Total number of hymns
            [5] - Missing Hymns
        """
        _EXEC = time.time()           ## 0 1  2    3     4     5
        _HYMNAL = namedtuple("HYMNAL", "EN TL US HYMNS TOTAL MISSING")
        _TOTAL = namedtuple("TOTAL", "EN TL US ALL")
        _MISSING = namedtuple("MISSING", "list length")
        HYMNAL, CATS, hNums = [[[], []], [[], []], [[], []], None, None, None], ["EN", "TL", "US"], []

        try:
            FNS = ZipFile(SYS.FILE_HYMNSDB, 'r').namelist()
        except Exception:
            ## When the Hymnal Database is not a valid database
            LOG.error(f'Database Error: Cannot read contents')
            MSG_BOX = QtWidgets.QMessageBox(); MSG_BOX.setStandardButtons(QtWidgets.QMessageBox.Ok); MSG_BOX.setIcon(QtWidgets.QMessageBox.Critical)
            MSG_BOX.setWindowTitle("Error"); MSG_BOX.setText(f"Database Error: Invalid contents.\nPlease contact the developers of this program to fix the issue.")
            MSG_BOX.setInformativeText(f"Error in: {SYS.FILE_HYMNSDB}")
            MSG_BOX.setWindowFlags(Qt.Drawer | Qt.WindowStaysOnTopHint)
            MSG_BOX.setStyleSheet(Stylesheet().getStylesheet())
            MSG_BOX.exec_()
            sys.exit()

        ## Parse Hymns Database
        for FILE in FNS:
            spH = self.splitHymn(FILE)                                                                                  ## Splits the filename into separate parts using splitHymn()
            hNums.append(spH.num)                                                                                       ## Append the hymn number to the dict 'Hymn Numbers'
            for i in range(3):                                                                                          ## Loop for 3 categories (EN, TL, USER)
                if spH.cat == CATS[i] and spH.ext == 'pptx':                                                            ## Append all valid .pptx file into the EN, TL and USER dict
                    HYMNAL[i][0].append(spH.num)
                    HYMNAL[i][1].append(spH.title)
                    break
                elif spH.ext != 'pptx' and FILE[-1] != "/":                                                             ## Flags all non-pptx as unnecessary file
                    LOG.warn(f"Unnecessary file \"{FILE}\" detected inside the database.")

        ## Report Database Status
        MISSING = [KString.toDigits(i+1, 3) for i in range(self.TOTAL_HYMNS) if KString.toDigits(i+1, 3) not in hNums]  ## Determines all missing hymns by looking through 1-474 range of hymns
        if MISSING:
            LOG.warn(f"{len(MISSING)} Hymns are missing")
            LOG.warn(f"Hymns that are missing: {', '.join([f'#{h}' for h in MISSING])}")
        else:
            LOG.info("The hymnal is complete")

        ## Append Other Information about Hymnal
        HYMNAL[3] = [[HYMNAL[i][0][j] + ' ' + HYMNAL[i][1][j] for j in range(len(HYMNAL[i][1]))] for i in range(3)]     ## Append all hymns in this format: "<Hymn Number> <Hymn Title>" | Usually used for search completer
        HYMNAL[3] = [i for sl in HYMNAL[3] for i in sl]                                                                 ## Flatten the nested list into one whole list
        HYMNAL[3].sort()
        HYMNAL[4] = _TOTAL(len(HYMNAL[0][0]), len(HYMNAL[1][0]), len(HYMNAL[2][0]), len(HYMNAL[3]))                     ## Tuple of total hymns. Has EN, TL, USER, and ALL values.
        HYMNAL[5] = _MISSING(MISSING, len(MISSING))                                                                     ## Tuple of missing hymns. Has list of missing hymns and length of the list.
        
        LOG.info(f"Successfully Scanned {HYMNAL[4].ALL} hymns in the database ({round((time.time()-_EXEC)*1000, 2)} ms)")
        return _HYMNAL(HYMNAL[0], HYMNAL[1], HYMNAL[2], HYMNAL[3], HYMNAL[4], HYMNAL[5])


    def getStats(self, hN):
        """
        Returns all specific hymnal statistics from a Hymn Number.

        Using HymnsDB (Self) and StatisticsDB, the return value is a namedtuple of the following:
        >>> [00] title - Title of the Hymn number provided
        >>> [01] num - Hymn number
        >>> [02] eqTitle - Equivalent Title from other category
        >>> [03] eqNum - Equivalent Hymn number from other category
        >>> [04] cat - Category ID
        >>> [05] eqCat - Equivalent Category ID
        >>> [06] dbIndex - Index of the hymn in database
        >>> [07] queries - Number of queries of the hymn
        >>> [08] launches - Number of launches of the hymn
        >>> [09] lastAccessed - Timestamp (in seconds) of the hymn
        >>> [10] lastAccssHumanize - Humanized version of the lastAccessed in past tense
        """
        UIZ.EXEC_READY = False
        hN = KString.toDigits(hN, 3)                                                                                    ## Converts the Hymn Number to a format of 3-digit padded hymn number
        if len(KString.filterOnly(0, hN)): return                                                                       ## Cancels the process when the hymn number is not a valid number or contains letters
        
        _HymnInfo = namedtuple("HymnInfo", "title num eqTitle eqNum cat eqCat dbIndex queries launches lastAccessed lastAcssHumanize")
        DATA = ['', hN, '', '', '', '', '', 0, 0, 0, 0]                                                                 ## Default values for HymnInfo
        idx = 0                                                                                                         ## Placeholder for dbIndex. 0 is still not recommended value for return.

        ## Verify Hymn Number | Should be below maximum hymn number and not in missing list.
        if int(hN) > SYS.HYMNS_MAX or int(hN) <=0 or hN in HYMNAL.MISSING.list:
            LOG.info(f"Entry #{hN} can't be verified. The file is invalid or missing.")
            return _HymnInfo(DATA[0], DATA[1], DATA[2], DATA[3], DATA[4],                                               ## Returns placeholder data
                            DATA[5], DATA[6], DATA[7], DATA[8], DATA[9], DATA[10])
        
        ## Scan for Base Hymn (Title and Number)
        i,j = 0,0 
        while True:
            try: hT = HYMNAL[i][1][HYMNAL[i][0].index(KString.toDigits(int(hN),3))]
            except: 
                if j: DATA[0] = ''; break
                i,j = 1,1
            else: DATA[0] = hT; break                                                                                   ## Stores the first data (Title)

        ## Scan for Equivalent Hymn (Title and Number)
        i, j = 0, 0                                                                                                     ## i is for the index iterations, j is for the lock switch; when j = 1, this indicates the final loop
        while True:
            if hN in HYMNAL[i][0]:
                try: idx = HYMNAL[invert(i)][0].index(KString.toDigits(operator.add(int(hN), -1 if i else 1), 3))
                except ValueError as e: break                                                                           ## When the hymn can't be found on other category
                else:
                    DATA[2] = HYMNAL[invert(i)][1][idx]                                                                 ## Stores the 3rd data (Equivalent Title)
                    DATA[3] = HYMNAL[invert(i)][0][idx]                                                                 ## Stores the 4th data (Equivalent Number)
                    break
            else:
                if j: break                                                                                             ## breaks the search when hymn number doesn't match in both categories
                i, j = invert(i), 1                                                                                     ## Switches to other [1] category while locked for last loop
        
        ## Current Category, Equivalent Category, and Database Index
        DATA[4] = 'TL' if i else 'EN'                                                                                   ## Stores the 5th data (Category)
        DATA[5] = 'EN' if i else 'TL'                                                                                   ## Stores the 6th data (Equivalent Category)
        DATA[6] = idx                                                                                                   ## Stores the 7th data (Index Number in Database; to reduce scanning again just to find the index)

        ## Get Stats of Base Hymn
        STS = SDATA['DATA'][hN] if int(hN) or int(hN) <= SYS.HYMNS_MAX else [0,0,0]                                     ## Statistic Data from Statistical Data (SDATA)
        DATA[7], DATA[8], DATA[9] = STS[0], STS[1], STS[2]                                                              ## Stored the rest of data values (Queries, Launches, and Last Accessed)

        ## Return All Data
        if DATA[0] != '': UIZ.EXEC_READY = True                                                                         ## If the title, which is the most important, is available. Flag the execution in ready state.
        return _HymnInfo(DATA[0], DATA[1], DATA[2], DATA[3], DATA[4], DATA[5], DATA[6],                                 ## Returns a namedtuple of all data
                        DATA[7], DATA[8], DATA[9], humanize.naturaltime(time.time()-DATA[9]))


    def genSearchSuggestions(self, UI):
        """
        Sets the suggestion data by using the hymns as a model for the completer.
        """
        MODEL = QtCore.QStringListModel()
        MODEL.setStringList(HYMNAL.HYMNS)
        self.CPLTR_SEARCH = QtWidgets.QCompleter()

        self.CPLTR_SEARCH.setFilterMode(Qt.MatchContains)
        self.CPLTR_SEARCH.setModel(MODEL); self.CPLTR_SEARCH.setCaseSensitivity(Qt.CaseInsensitive)
        self.CPLTR_SEARCH.setMaxVisibleItems(SYS.CPLTR_MAX_VISIBLE_ITEMS)

        UI.LNE_SEARCH.setCompleter(self.CPLTR_SEARCH)


    def updateDatabase(self, package):
        """
        Updates configuration with new hymnal package
        """
        SYS.FILE_HYMNSDB = package[0]
        LOG.info('SDA Package updated: ' + package[0])
        UIB.CLASS_STS.LNE_TARGET_PATH.setText(package[0])
        global HYMNAL
        HYMNAL = HDB.parseHymnDatabase()
        pass

        ## THIS SECTION NEEDS DEVELOPMENT


class Stylesheet():
    def __init__(self):
        pass


    def getThemes(self):
        TSL = {0: 'light', 1: 'dark', 2: 'dark'}
        self.THEME = list(v for k,v in TSL.items() if k == SYS.CURR_THEME)[0]
        self.toggleMode(SYS.CURR_THEME)


    def initStylesheet(self):
        """
        Updates the stylesheet
        """
        SS = self.getStylesheet()                                                           ## Retrieves the Stylesheet string

        HDB.CPLTR_SEARCH.popup().setStyleSheet(SS)

        try:
            APP.setStyleSheet(SS)
            UIA.setStyleSheet(SS)

            UIB.setStyleSheet(SS)
            LST_PANEL_ICONS = [QtGui.QIcon(f'./res/icons/{name}_{self.THEME}.png') for name in ['settings', 'hymnal', 'library', 'command']], [QtGui.QIcon(f'./res/icons/{name}_{self.THEME}.png') for name in ['info']]
            for i in range(2):
                for j, icon in enumerate(LST_PANEL_ICONS[i]): UIB.LST_PANELS[i].item(j).setIcon(icon)
            self.ANIM_BTN_LAUNCH = ANM.ButtonAnimation(UIA.BTN_LAUNCH, modHex(self.PRIMARY, 30), self.PRIMARY, self.TXT_INV, self.TXT_INV); self.ANIM_BTN_LAUNCH.connectEvents()
            self.ANIM_BTN_LAUNCH = ANM.ButtonAnimation(UIC.BTN_LAUNCH, modHex(self.PRIMARY, 30), self.PRIMARY, self.TXT_INV, self.TXT_INV); self.ANIM_BTN_LAUNCH.connectEvents()
            self.ANIM_BTN_RESET = ANM.ButtonAnimation(UIB.BTN_RESET, self.ERROR, self.palette2Hex('button'), self.palette2Hex('text'), self.palette2Hex('brightText')); self.ANIM_BTN_RESET.connectEvents()
            self.ANIM_BTN_OK = ANM.ButtonAnimation(UIB.BTN_OK, modHex(self.PRIMARY, 30), self.PRIMARY, self.TXT_INV, self.TXT_INV); self.ANIM_BTN_OK.connectEvents()

            UIC.setStyleSheet(SS)
        except (NameError, AttributeError) as e:
            pass


    def palette2Hex(self, color):
        """
        Returns HEX value of an RGB of a certain palette color
        """
        return SYS.RGBtoHEX(getattr(APP.palette(), color)().color().getRgb())


    def getStylesheet(self, objectName=None):
        """
        Returns a string of stylesheet that will be used by QStyleSheet
        Values depends on what is the current theme.
        """
        self.getThemes()
        RADIUS = "6px" ## Default: 9px
        RADIUS_SML = "4px" ## Default: 5px
        PADDING = "5px"

        if objectName is None:
            STYLESHEET = f"""
                QWidget#WIN_BROWSER {{
                    image: url('./res/images/logofade.png');
                    image-position: right;
                    border: 1px solid {self.BORDER};
                }}
                QWidget#WIN_SETTINGS {{
                    image: url('./res/images/settingsfade.png');
                    image-position: right;
                }}
                QWidget#WIN_COMPACT {{
                    image: url('./res/images/logofade_compact.png');
                    image-position: left;
                    border: 1px solid {self.BORDER};
                }}



                /* Buttons */ 

                QPushButton {{
                    background-color: palette(button);
                    padding: {PADDING};
                    border-radius: {RADIUS};
                    min-width: 80px;
                    outline: none;
                    border: 1px solid {modHex(self.palette2Hex('button'), 7)}
                }}
                QPushButton#BTN_LAUNCH, QPushButton#BTN_OK {{
                    border: 1px solid {'0' + modHex(self.PRIMARY, 7)[1:]};
                }}
                QPushButton#BTN_ADDQUEUE, QPushButton#BTN_QUEUES, QPushButton#BTN_MORE {{
                    border: none;
                }}
                QPushButton::pressed {{
                    background-color: palette(button);
                }}
                QPushButton::disabled {{
                    color: {self.TXT_DISABLED};
                    background-color: {self.BTN_DISABLED};
                }}
                QPushButton::hover {{
                    background-color: palette(light);
                }}
                /*
                QPushButton::focus {{
                    border: 1px solid {self.BORDER};
                }}
                */



                /* Dialog Boxes */
                QMessageBox {{
                    background-color: palette(window);
                }}
                


                /* Tooltip */
                QToolTip {{
                    color: palette(text);
                    background-color: palette(base);
                    border: none;
                }}



                /* Status Bar */
                QStatusBar#STATUSBAR {{
                    color: {self.TXT_DISABLED};
                    background-color: {self.STATUSBAR};
                }}



                /* Menu */
                QMenu {{
                    color: palette(text);
                    border: 1px solid {self.BORDER};
                    background-color: {self.CTX_MENU};
                    border-radius: {RADIUS};
                }}
                QMenu::item::selected {{
                    background-color: palette(window);
                }}



                /* Search Bars */
                QLineEdit, QComboBox{{
                    color: palette(text);
                    selection-color: {self.TXT_INV};
                    background-color: {self.CARD};
                    border: 1px solid {self.BORDER};
                    border-radius: {RADIUS};
                    padding: {PADDING};
                }}
                QComboBox {{
                    background-color: transparent;
                    border: 1px solid {self.BORDER};
                }}
                QLineEdit::focus#LNE_SEARCH {{
                    background-color: #AF{self.BORDER[1:]};
                }}
                QLineEdit::hover#LNE_SEARCH {{
                    border: 1px solid {self.BORDER_HIGHLIGHT};
                }}



                /* Group Boxes */
                QGroupBox {{
                    border-radius: {RADIUS};
                    background-color: {self.CARD};
                    margin-top: 1.5em;
                    padding: 5px;
                    font-style: 'Segoe UI Variable Display';
                    font-weight: bold;
                    font-size: 12pt;
                }}
                QGroupBox::hover {{
                    background-color: {self.CARDHOVER};
                }}
                QGroupBox::title {{
                    color: palette(text);
                    subcontrol-origin: margin;
                    background-color: palette(window);
                    left: 0px;
                    padding: 3px 5px 3px 5px;
                    border-radius: {RADIUS};
                }}
                QGroupBox::title::hover {{
                    border: 1px solid {self.PRIMARY};
                }}



                /* Sliders */
                QSlider::handle {{
                    background-color: {self.PRIMARY};
                }}
                QSlider::handle::hover {{
                    background-color: {self.SECONDARY};
                }}
                QSlider::handle::pressed {{
                    background-color: {modHex(self.SECONDARY, 20)};
                }}



                /* ScrollBars */
                QScrollBar:vertical {{
                    background-color: palette(base);
                    width: 16px;
                    margin: 0px;
                    border-radius: {RADIUS_SML};
                }}
                QScrollBar:horizontal {{
                    background-color: palette(base);
                    height: 15px;
                    margin: 0px;
                    border-radius: {RADIUS_SML};
                }}
                QScrollBar::handle:vertical {{
                    background-color: {self.SCROLLBAR}; min-height: 20px; margin: 3px; border-radius: {RADIUS_SML}; border: none;
                }}
                QScrollBar::handle:horizontal {{
                    background-color: {self.SCROLLBAR}; min-width: 20px; margin: 3px; border-radius: {RADIUS_SML}; border: none;
                }}
                QScrollBar::handle::hover {{
                    background-color: {modHex(self.SCROLLBAR, 20)}; min-width: 20px; margin: 3px; border-radius: {RADIUS_SML}; border: none;
                }}
                QScrollBar::handle::pressed {{
                    background-color: {self.PRIMARY}; min-width: 20px; margin: 3px; border-radius: {RADIUS_SML}; border: none;
                }}

                QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                    border: none; background: none; height: 0px;
                }}
                QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
                    border: none; background: none; width: 0px;
                }}
                QScrollBar::up-arrow:vertical, QScrollBar::down-arrow:vertical, QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
                    border: none; background: none; color: none;
                }}
                QScrollBar::left-arrow:horizontal, QScrollBar::right-arrow:horizontal, QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{
                    border: none; background: none; color: none;
                }}
                


                /* Settings Panel (List) */
                QAbstractItemView {{
                    color: palette(text);
                    outline: none;
                    background-color: palette(base);
                    border: none;
                    border-radius: {RADIUS};
                    selection-color: {self.TXT_INV};
                    selection-background-color: {self.PRIMARY};
                    padding: 5px;
                    min-height: 18px;
                }}
                QAbstractItemView::item {{
                    padding: 6px 3px 6px 5px;
                    margin: 2px 0px 2px 0px;
                    border-radius: {RADIUS};
                }}
                QAbstractItemView::item::selected {{
                    background-color: {self.PRIMARY};
                }}
                QAbstractItemView::item::hover {{
                    color: palette(text);
                    background-color: palette(window);
                }}
                QAbstractItemView::item::selected::hover {{
                    color: {self.TXT_INV};
                    background-color: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 1,stop: 0 {modHex(self.PRIMARY, 50)}, stop: 1 {self.PRIMARY});
                }}
                QListWidget {{
                    font-style: 'Segoe UI Variable Display';
                    font-weight: bold;
                    font-size: 10pt;
                }}



                /* Checkboxes */
                QCheckBox {{
                    outline: none;
                }}
                QCheckBox::indicator {{
                    width: 15px;
                    height: 15px;
                }}
                QCheckBox::indicator:unchecked {{
                    image: url(./res/icons/unchecked_{self.THEME}.png);
                }}
                QCheckBox::indicator:unchecked::hover {{
                    image: url(./res/icons/unchecked_hover_{self.THEME}.png);
                }}
                QCheckBox::indicator:checked {{
                    image: url(./res/icons/checked_{self.THEME}.png);
                }}
                QCheckBox::indicator:checked::hover {{
                    image: url(./res/icons/checked_hover_{self.THEME}.png);
                }}



                /* Table Widget */
                QTableWidget {{
                    color: palette(text);
                    border-style: none;
                    gridline-color: {self.BORDER};
                    outline: 0;
                    background: {self.CARD};
                }} 
                QTableWidget::item {{
                    padding: 3px;
                }}
                QTableWidget::item::hover {{
                    color: palette(text);
                }}
                QTableWidget::item::focus {{
                    color: palette(highlighted-text);
                    background: palette(highlight);
                }}



                /* Header View (Table Widget) */
                QHeaderView {{
                    background-color: palette(window);
                    border: 1px solid {self.BORDER};
                    border-radius: {RADIUS};
                }}
                QHeaderView::section {{
                    color: palette(text);
                    background-color: transparent;
                    border-style: none;
                }}
                QHeaderView::section::hover {{
                    background-color: palette(light);
                    border-radius: {RADIUS};
                }}
                QTableCornerButton::section {{
                    background-color: transparent;
                }}



                /* Dock Widget */
                QDockWidget::title {{
                    color: palette(text);
                    text-align: center;
                    background: palette(button);
                    padding: 5px;
                    border-radius: {RADIUS};
                    font-size: 10px;
                }}
                QDockWidget::title::hover {{
                    background: palette(mid);
                }}



                /* Log Panel */
                QPlainTextEdit {{
                    color: palette(text);
                    background-color: palette(base);
                    border: 1px solid {self.BORDER};
                    border-radius: {RADIUS};
                }}



                /* Spin

                /* Spin Box */
                QSpinBox {{
                    border: 1px solid {self.BORDER};
                    border-radius: {RADIUS};
                }}
                
                

                /* Settings Button */
                QPushButton#BTN_MORE {{
                    background-color: transparent;
                    min-width: 17px;
                    width: 20px;
                    height: 20px;
                    image: url('./res/icons/settings_{self.THEME}.png');
                }}
                QPushButton::hover#BTN_MORE {{
                    image: url('./res/icons/settings_hover_{self.THEME}.png');
                }}


                /* Add Queue Button */
                QPushButton#BTN_ADDQUEUE {{
                    background-color: transparent;
                    min-width: 17px;
                    width: 20px;
                    height: 20px;
                    image: url('./res/icons/add_{self.THEME}.png');
                }}
                QPushButton::hover#BTN_ADDQUEUE {{
                    image: url('./res/icons/add_hover_{self.THEME}.png');
                }}


                /* Queue Button */
                QPushButton#BTN_QUEUES {{
                    background-color: transparent;
                    min-width: 17px;
                    width: 20px;
                    height: 20px;
                    image: url('./res/icons/queue_{self.THEME}.png');
                }}
                QPushButton::hover#BTN_QUEUES {{
                    image: url('./res/icons/queue_hover_{self.THEME}.png');
                }}



                /* Special */

                QLabel#LBL_BROWSER {{
                    color: {self.PRIMARY};
                }}
                QLabel::disabled, QLabel#LBL_BROWSERB {{
                    color: {self.TXT_DISABLED};
                }}
                
                QPushButton::enabled#BTN_LAUNCH, QPushButton::enabled#BTN_OK {{
                    color: {self.TXT_INV};
                    background-color: {self.PRIMARY};
                }}
                QPushButton#BTN_LAUNCH {{
                    min-width: 100px;
                }}
                QPushButton::hover#BTN_LAUNCH, QPushButton::hover#BTN_OK {{
                    color: {self.TXT_INV};
                    background-color: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 1,stop: 0 {modHex(self.PRIMARY, 50)}, stop: 1 {self.PRIMARY});
                }}
                
                QPushButton::hover#BTN_RESET {{
                    color: palette(highlighted-text);
                    background-color: {self.ERROR};
                }}
                QLabel#LBL_DBSTATUS {{
                    background-color: palette(base);
                    padding: 3px;
                    border-radius: {RADIUS};
                }}
                """
            return STYLESHEET
    
    
    def toggleMode(self, mode):
        """
        Sets and updates all application color palette
        """
        
        if not mode:
            """
            For Light Mode
            """
            self.PRIMARY = '#004B74'
            self.SECONDARY = '#008A9A'
            self.TERTIARY = '#F7EBC5'
            self.GOLD = '#FFA92D'
            self.WARN = '#D25900'
            self.ERROR = modHex('#9E1919', 15)
            self.BORDER = '#C0C0C0'
            self.BORDER_HIGHLIGHT = '#C0C0C0'
            self.BTN_DISABLED = '#B6B6B6'
            self.TXT_INV = '#FFFFFF'
            self.TXT_DISABLED = '#999999'
            self.STATUSBAR = '#CFCFCF'
            self.SCROLLBAR = '#A0A0A0'
            self.CARD = '#DDDDDD'
            self.CARDHOVER = '#EEEEEE'
            self.CTX_MENU = '#FFFFFF'
            PLT_LIGHT = QtGui.QPalette()
            PLT_LIGHT.setColor(QtGui.QPalette.Window, SYS.QCl('#F3F3F3'))
            PLT_LIGHT.setColor(QtGui.QPalette.WindowText, SYS.QCl('#202020'))
            PLT_LIGHT.setColor(QtGui.QPalette.Base, SYS.QCl('#CFCFCF'))
            PLT_LIGHT.setColor(QtGui.QPalette.AlternateBase, SYS.QCl('#2D2D2D'))
            PLT_LIGHT.setColor(QtGui.QPalette.ToolTipBase, SYS.QCl('#252525'))
            PLT_LIGHT.setColor(QtGui.QPalette.ToolTipText, SYS.QCl('#C5C5C5'))
            PLT_LIGHT.setColor(QtGui.QPalette.PlaceholderText, SYS.QCl('#999999'))
            PLT_LIGHT.setColor(QtGui.QPalette.HighlightedText, SYS.QCl('#EEEEEE'))
            PLT_LIGHT.setColor(QtGui.QPalette.Highlight, SYS.QCl(self.PRIMARY))
            PLT_LIGHT.setColor(QtGui.QPalette.Light, SYS.QCl('#D7D7D7'))
            PLT_LIGHT.setColor(QtGui.QPalette.Text, SYS.QCl('#202020'))
            PLT_LIGHT.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.Text, SYS.QCl('#434343')) ## <- Unused
            PLT_LIGHT.setColor(QtGui.QPalette.Midlight, SYS.QCl('#888888'))
            PLT_LIGHT.setColor(QtGui.QPalette.Mid, SYS.QCl('#D2D2D2'))
            PLT_LIGHT.setColor(QtGui.QPalette.Dark, SYS.QCl('#555555'))
            PLT_LIGHT.setColor(QtGui.QPalette.Button, SYS.QCl('#CCCCCC'))
            PLT_LIGHT.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.Button, SYS.QCl('#252525')) ## <- Unused
            PLT_LIGHT.setColor(QtGui.QPalette.ButtonText, SYS.QCl('#202020'))
            PLT_LIGHT.setColor(QtGui.QPalette.BrightText, SYS.QCl('#FFFFFF'))
            PLT_LIGHT.setColor(QtGui.QPalette.Link, SYS.QCl('#202020'))
            PLT_LIGHT.setColor(QtGui.QPalette.LinkVisited, SYS.QCl('#151515'))
            APP.setPalette(PLT_LIGHT)

        elif mode == 1:
            """
            For Dark Mode
            """
            self.PRIMARY = '#008A9A' ## #008A9A Original
            self.SECONDARY = '#004B74'
            self.TERTIARY = '#F7EBC5'
            self.GOLD = '#FFA92D'
            self.WARN = '#D25900'
            self.ERROR = '#9E1919'
            self.BORDER = '#303030'
            self.BORDER_HIGHLIGHT = '#505050'
            self.BTN_DISABLED = '#212121'
            self.TXT_INV = '#181818'
            self.TXT_DISABLED = '#434343'
            self.STATUSBAR = '#2A2A2A'
            self.SCROLLBAR = '#3B3B3B'
            self.CARD = '#2A2A2A'
            self.CARDHOVER = '#323232'
            self.CTX_MENU = '#1D1D1D'
            PLT_DARK = QtGui.QPalette()
            PLT_DARK.setColor(QtGui.QPalette.Window, SYS.QCl('#202020'))
            PLT_DARK.setColor(QtGui.QPalette.WindowText, SYS.QCl('#D5D5D5'))
            PLT_DARK.setColor(QtGui.QPalette.Base, SYS.QCl('#191919'))
            PLT_DARK.setColor(QtGui.QPalette.AlternateBase, SYS.QCl('#2D2D2D'))
            PLT_DARK.setColor(QtGui.QPalette.ToolTipBase, SYS.QCl('#252525'))
            PLT_DARK.setColor(QtGui.QPalette.ToolTipText, SYS.QCl('#C5C5C5'))
            PLT_DARK.setColor(QtGui.QPalette.PlaceholderText, SYS.QCl('#999999'))
            PLT_DARK.setColor(QtGui.QPalette.HighlightedText, SYS.QCl('#191919'))
            PLT_DARK.setColor(QtGui.QPalette.Highlight, SYS.QCl(self.PRIMARY))
            PLT_DARK.setColor(QtGui.QPalette.Light, SYS.QCl('#898989'))
            PLT_DARK.setColor(QtGui.QPalette.Text, SYS.QCl('#EFEFEF'))
            PLT_DARK.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.Text, SYS.QCl('#939393')) ## <- Unused
            PLT_DARK.setColor(QtGui.QPalette.Midlight, SYS.QCl('#888888'))
            PLT_DARK.setColor(QtGui.QPalette.Mid, SYS.QCl('#424242'))
            PLT_DARK.setColor(QtGui.QPalette.Dark, SYS.QCl('#555555'))
            PLT_DARK.setColor(QtGui.QPalette.Button, SYS.QCl('#353535'))
            PLT_DARK.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.Button, SYS.QCl('#252525')) ## <- Unused
            PLT_DARK.setColor(QtGui.QPalette.ButtonText, SYS.QCl('#EFEFEF'))
            PLT_DARK.setColor(QtGui.QPalette.BrightText, SYS.QCl('#FFFFFF'))
            PLT_DARK.setColor(QtGui.QPalette.Link, SYS.QCl('#D0D0D0'))
            PLT_DARK.setColor(QtGui.QPalette.LinkVisited, SYS.QCl('#CECECE'))
            APP.setPalette(PLT_DARK)

        elif mode == 2:
            """
            For Black Mode
            """
            self.PRIMARY = '#008A9A' ## #008A9A Original
            self.SECONDARY = '#004B74'
            self.TERTIARY = '#F7EBC5'
            self.GOLD = '#FFA92D'
            self.WARN = '#D25900'
            self.ERROR = '#9E1919'
            self.BORDER = '#202020'
            self.BORDER_HIGHLIGHT = '#404040'
            self.BTN_DISABLED = '#181818'
            self.TXT_INV = '#181818'
            self.TXT_DISABLED = '#434343'
            self.STATUSBAR = '#131313'
            self.SCROLLBAR = '#3B3B3B'
            self.CARD = '#151515'
            self.CARDHOVER = '#191919'
            self.CTX_MENU = '#1A1A1A'
            PLT_BLACK = QtGui.QPalette()
            PLT_BLACK.setColor(QtGui.QPalette.Window, SYS.QCl('#000000'))
            PLT_BLACK.setColor(QtGui.QPalette.WindowText, SYS.QCl('#D5D5D5'))
            PLT_BLACK.setColor(QtGui.QPalette.Base, SYS.QCl('#111111'))
            PLT_BLACK.setColor(QtGui.QPalette.AlternateBase, SYS.QCl('#2D2D2D'))
            PLT_BLACK.setColor(QtGui.QPalette.ToolTipBase, SYS.QCl('#252525'))
            PLT_BLACK.setColor(QtGui.QPalette.ToolTipText, SYS.QCl('#C5C5C5'))
            PLT_BLACK.setColor(QtGui.QPalette.PlaceholderText, SYS.QCl('#999999'))
            PLT_BLACK.setColor(QtGui.QPalette.HighlightedText, SYS.QCl('#FFFFFF'))
            PLT_BLACK.setColor(QtGui.QPalette.Highlight, SYS.QCl(self.PRIMARY))
            PLT_BLACK.setColor(QtGui.QPalette.Light, SYS.QCl('#505050'))
            PLT_BLACK.setColor(QtGui.QPalette.Text, SYS.QCl('#C7C7C7'))
            PLT_BLACK.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.Text, SYS.QCl('#434343')) ## <- Unused
            PLT_BLACK.setColor(QtGui.QPalette.Midlight, SYS.QCl('#888888'))
            PLT_BLACK.setColor(QtGui.QPalette.Mid, SYS.QCl('#424242'))
            PLT_BLACK.setColor(QtGui.QPalette.Dark, SYS.QCl('#555555'))
            PLT_BLACK.setColor(QtGui.QPalette.Button, SYS.QCl('#1F1F1F'))
            PLT_BLACK.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.Button, SYS.QCl('#252525')) ## <- Unused
            PLT_BLACK.setColor(QtGui.QPalette.ButtonText, SYS.QCl('#C5C5C5'))
            PLT_BLACK.setColor(QtGui.QPalette.BrightText, SYS.QCl('#FFFFFF'))
            PLT_BLACK.setColor(QtGui.QPalette.Link, SYS.QCl('#D0D0D0'))
            PLT_BLACK.setColor(QtGui.QPalette.LinkVisited, SYS.QCl('#CECECE'))
            APP.setPalette(PLT_BLACK)




class InputCore():
    """
    This class handles all data processing for Hymns including
    the presentation of details, time access, execution, statistic dumps, etc.
    """
    def __init__(self):
        pass
    

    def updateDetails(self, cmb=False):
        """
        Updates the detail area when the user typed something.
        
        The data from search box and recent list is called an ENTRY.
        The whole process only takes approximately ~1ms
        
        Flow:
            - Verify input if valid and in the HYMNAL
            - Identify the incoming entry source (Search Box | Recent List)
            - Get all information of the hymn number via HymnsDB
            - Add new record for query after 2 seconds delay of inactivity using the QUERY_TIMEOUT switch
            - Decides whether if the system is ready for executing the file or unavailable for execution

        This method is shared between the 2 user input objects and the behavior have slight differences:
            - Search Box (LNE_SEARCH); and
            - Recent List (CMB_RECENT)
        """
        UIZ.QUERY_TIMEOUT = True                                                                     ## Start Switch; Indicates that the query is set to hold and will wait for timeout fill before dumping the data
        UIZ.EXEC_POINTER = cmb                                                                       ## Execution Pointer; Indicates that the entry to be processed is either from Search Box or from Recent Lists

        if UIZ == UIA:
            if HDB.CPLTR_SEARCH.completionCount() > 1 and not cmb:
                UIA.STATUSBAR.showMessage(f'{HDB.CPLTR_SEARCH.completionCount()} suggestions')              ## Sets number of suggestions to be displayed in status bar
            if len(UIA.LNE_SEARCH.text()) <=0: UIA.STATUSBAR.showMessage('')                                ## TEMPORARY FIX for Status Bar Suggestions stucked

            if UIA.EXEC_POINTER == 0:
                UIA.CMB_RECENT.setCurrentIndex(-1)                                                          ## Deselects and clears current text from Recent List when user is using search box to reduce clutter and confusion.

        ENTRY = UIZ.LNE_SEARCH.text() if not cmb else UIZ.CMB_RECENT.currentText()            ## Substitutes the entry from Recent List when the pointer is set to 1(CMB)
        ENTRY_CMPLTN = HDB.CPLTR_SEARCH.currentCompletion()

        hN = ENTRY[:3]                                                                                      ## Parse the entry by filtering only the Hymn Numbers

        if not hN.isdigit() or not int(hN):                                                                 ## Prevents processing an incomplete user input
            if ENTRY != '' and ENTRY_CMPLTN[:3] not in HYMNAL.MISSING.list:                                 ## Instead of ignoring the whole entry,
                hN = ENTRY_CMPLTN[:3]                                                                       ## the current Completion will act as backup for entry
            else:                                                                                           ## Otherwise, update the UI with respect to the empty entry.
                self.updateButtons(0)
                self.updatePreviews(0)
                UIZ.QUERY_TIMEOUT = False                                                            ## Prevents dumping of query data when entry is invalid
                return

        HINFO = HDB.getStats(hN)                                                                            ## Retrieve Hymn Info/Stats
        self.CURRENT_HINFO = HINFO                                                                          ## Stores the Hymn Data as the program's current hymn.


        """ The block below contains several UI-related adjustments when updating details. """
        ## Clear the previews before setting new texts
        if UIZ == UIA:
            UIZ.LBL_PVW_TITLE.clear()
            UIZ.LBL_PVW_ACCSS.clear(); UIZ.LBL_PVW_ACCSS.setStatusTip('')
        elif UIZ == UIC:
            UIZ.LBL_PVW_TITLE.setText(SW.NAME)

        ## Title
        if ENTRY != f'{HINFO.num} {HINFO.title}':
            if UIZ == UIA: UIZ.LBL_PVW_TITLE.setText(HINFO.title)                             ## Displays hymn title when current search text does not match the title anymore. This is to let the user know the title even when the search box is incomplete
            if UIZ == UIC: UIZ.LBL_PVW_TITLE.setText(f'{HINFO.title} <font color="{QSS.TXT_DISABLED}">{f"({humanize.naturaltime(time.time()-HINFO.lastAccessed)})</font>" if HINFO.lastAccessed !=0 else ""}') 
        elif cmb and UIZ.CMB_RECENT.currentText() != f'{HINFO.num} {HINFO.title}':
            UIZ.LBL_PVW_TITLE.setText(HINFO.title)                                                   ## Similar behavior from above but differs in source since the entry is from recent list.                                              
        else:
            if UIZ != UIA: UIZ.LBL_PVW_TITLE.setText(f'{HINFO.eqTitle} <font color="{QSS.TXT_DISABLED}">{f"({humanize.naturaltime(time.time()-HINFO.lastAccessed)})</font>" if HINFO.lastAccessed !=0 else ""}')

        if UIZ == UIA:
            ## Equivalent Title
            UIA.LBL_PVW_EQUIV.setText(f'{HINFO.eqTitle} ')                                                  ## Displays the equivalent Hymn title. This statement does not need a placeholder text since the variable inserted here has it incase of missing equivalents
            
            ## Last Accessed
            if HINFO.lastAccessed > 0:
                UIA.LBL_PVW_ACCSS.setText(f"Last opened: {datetime.fromtimestamp(HINFO.lastAccessed).strftime('%b %d, %Y at %I:%M %p')} ({humanize.naturaltime(time.time()-HINFO.lastAccessed)})")
            
            ## Launched (Status Bar)
            if HINFO.launches !=0 and UIA.LBL_PVW_ACCSS.text() != '':
                UIA.LBL_PVW_ACCSS.setStatusTip(f'Launched {f"{HINFO.launches} times." if HINFO.launches > 1 else "once."}')

        ## Update Buttons
        self.updateButtons(1 if UIZ.EXEC_READY else -1)
    

    def updatePreviews(self, mode):
        """
        Handles all interface behaviors of the preview labels

        Modes:
        >>> -1 - Reserved
        >>> 0 - Not Available/Clear
        >>> 1 - Ready
        >>> 2 - Reserved
        """
        if mode == 0: ## Not Available
            if UIZ == UIA:                                                                           ## Full Size Mode
                UIZ.LBL_PVW_TITLE.setText('')
                UIZ.LBL_PVW_EQUIV.setText('')
                UIZ.LBL_PVW_ACCSS.setText('')
                UIZ.LBL_PVW_ACCSS.setStatusTip('')
            else:                                                                                           ## Compact Mode
                UIZ.LBL_PVW_TITLE.setText(SW.NAME)


    def updateButtons(self, mode):
        """
        Handles all interface behaviors of the buttons

        Modes:
        >>> -1 - No Available File
        >>> 0 - Waiting
        >>> 1 - Ready for launch
        >>> 2 - Launching
        >>> 3 - Launched
        """
        BTN0 = ['No file available', 'Insert a Hymn', 'Launch', 'Launching', 'Launched']

        if mode == -1: ## No Available File
            UIZ.BTN_LAUNCH.setEnabled(False); UIZ.BTN_LAUNCH.setText(BTN0[0])

        elif mode == 0: ## Waiting
            UIZ.BTN_LAUNCH.setEnabled(False); UIZ.BTN_LAUNCH.setText(BTN0[1])

        elif mode == 1: ## Ready
            UIZ.BTN_LAUNCH.setEnabled(True); UIZ.BTN_LAUNCH.setText(BTN0[2])

        elif mode == 2: ## Launching
            UIZ.BTN_LAUNCH.setEnabled(False); UIZ.BTN_LAUNCH.setText(BTN0[3])
            keyboard.send("esc")                                                                                # Sends escape key to hide completer for search box if it's active

        elif mode == 3: ## Launched
            UIZ.BTN_LAUNCH.setEnabled(False); UIZ.BTN_LAUNCH.setText(BTN0[4])


    def executeFile(self):
        """
        Opens the PowerPoint when the user clicked the 'Launch' button.

        This method is one of the most important part of this software as this executes the PowerPoint file.
        All Hymn Info is presumed to be valid here because there are currently no exception-catching in this method.
        """
        EXEC_TIME = time.time()                                                                                 ## Start logging the process time
        self.updateButtons(2)                                                                                   ## Flags all buttons into 'launching' state

        ## Extract the file
        EXEC_TIME = time.time()
        XDB = ZipFile(SYS.FILE_HYMNSDB, 'r')
        for FILE in XDB.namelist():
            spH = HDB.splitHymn(FILE)
            if self.CURRENT_HINFO.title in spH.title:
                FILEDIR = f'{SYS.DIR_TEMP}\\{FILE[3:]}'
                if not KPath.exists(FILEDIR):                                                                   ## Only extract if the file is non-existent in temp folder in order to save time for extraction
                    XDB.extract(FILE, SYS.DIR_TEMP)                                                             ## Extract to temporary directory set by System class
                    shutil.move(f'{SYS.DIR_TEMP}\\{FILE}', SYS.DIR_TEMP)                                        ## Relocate the file from a subfolder (EN/TL) to main temp directory
                break

        ## Other Properties
        SW = 'S' if CDATA['AutoSlideshow'] == 'True' else 'O'

        ## Execute the file by creating a subprocess thread
        PROCESS = Thread(target = lambda: subprocess.Popen(f'cd {SYS.PPT_EXEC} & start POWERPNT.exe /{SW} "{FILEDIR}"', shell=True))
        PROCESS.start()

        ## Check for overflow. Delete older files when the folder items reached the maximum recent files allowed.
        FMN.deleteRecent()

        ## Log to Statistic Data
        ARRAY = [self.CURRENT_HINFO.queries+1, self.CURRENT_HINFO.launches+1, time.time()]                              ## Issue: Might need to build a separate function for this array incase the array is expanded in the future
        SDATA['DATA'].update({str(self.CURRENT_HINFO.num): ARRAY}); SDB.dump()                                      ## Updates the main stat data (SDATA) and dumps into JSON using the Statistics Database class
        
        ## Update UI
        if UIZ == UIA:
            RECENT_ITEMS = [UIZ.CMB_RECENT.itemText(i) for i in range(int(CDATA['MaxAllowedRecent']))] 
            UIZ.updateRecentList(1 if f'{self.CURRENT_HINFO.num} {self.CURRENT_HINFO.title}' not in RECENT_ITEMS else -1)  ## Update Recent List so the recent launched file would show up
        self.updateDetails()
        self.updateButtons(3)

        ## Refocus to Main Window if the keep focus setting is enabled
        if CDATA['KeepFocusOnBrowser'] != 'True':
            time.sleep(.3)                                                                                          ## Places a 300ms delay for the program to simulate refocusing back to the program; This is also to minimize the possibility of glitches.
            if not UIB.isHidden(): UIB.show()                                                                       ## Also shows the Settings window
            UIZ.setFocus(True); UIZ.activateWindow(); UIZ.raise_(); UIZ.show()          ## Refocuses the active window mode (Full or Compact mode)
            keyboard.send("tab")                                                                                    ## Simulates tab keypress to automatically highlight the search bar which is on first index of the GUI layout

        ## Log
        PRS = Presentation(FILEDIR)
        SYS.CNT_SESSION_PRESN += 1
        LOG.info(f"Launched #{self.CURRENT_HINFO.num} {self.CURRENT_HINFO.title} ({len(PRS.slides)} Slide(s)) ({round((time.time()-EXEC_TIME)*1000)} ms) {getFileStat(FILEDIR, 'sz', szFmt='KB')} KB")


    def searchFill(self, event):
        """
        Handles Mouse Wheel Event for LNE_SEARCH
        This method helps the user scroll through search bar treating it similar to a combobox 
        """ 
        ## This whole block is not yet optimized.
        try: CHINFO = self.CURRENT_HINFO
        except AttributeError: CHINFO = HDB.getStats('001')
        NUM = int(CHINFO.num)

        DELTA = event.angleDelta().y()
        if DELTA == -120 and NUM <= SYS.HYMNS_MAX-1: ## Scroll Down
            UIZ.LNE_SEARCH.setText(f"{KString.toDigits(NUM+1, 3)} {HDB.getStats(KString.toDigits(NUM+1, 3)).title}")
        elif DELTA == 120 and NUM >= 2: ## Scroll Up
            if NUM > SYS.HYMNS_MAX: UIZ.LNE_SEARCH.setText(f"{KString.toDigits(SYS.HYMNS_MAX, 3)} {HDB.getStats(KString.toDigits(SYS.HYMNS_MAX, 3)).title}")
            else: UIZ.LNE_SEARCH.setText(f"{KString.toDigits(NUM-1, 3)} {HDB.getStats(KString.toDigits(NUM-1, 3)).title}")




class FileManager():
    """
    This class manages all external files covered by the software.
    
    This includes managing of recent files and temporary folder to maintain and prevent building up of unused data.
    """
    def __init__(self):
        self.deleteRecent()
        self.deleteLogs()


    def deleteRecent(self, deleteAll=False):
        """
        Check for overflow.
        """
        try: THRESHOLD = int(CDATA['MaxAllowedRecent'])                                                         ## Retrieves the amount of maximum allowed recent files
        except ValueError: THRESHOLD = SYS.RECENTS.DEFAULT
        if THRESHOLD < SYS.RECENTS.ALLOWEDMIN or THRESHOLD > SYS.RECENTS.ALLOWEDMAX: THRESHOLD = SYS.RECENTS.DEFAULT
        self.deleteOldest(self.getRecentFiles, THRESHOLD, deleteAll)

    def deleteLogs(self, deleteAll=False):
        self.deleteOldest(self.getLogFiles, SYS.LOG_FILE_LIMIT, deleteAll)
    
    def getRecentFiles(self):
        return [f'{SYS.DIR_TEMP}\\{i}' for i in os.listdir(SYS.DIR_TEMP) if not i.startswith('~$') and i.endswith('.pptx')]

    def getLogFiles(self):
        return [f'{SYS.DIR_LOG}\\{i}' for i in os.listdir(SYS.DIR_LOG) if i.endswith('.log')]


    def deleteOldest(self, fileList, threshold, deleteAll):
        """
        Delete older files when the folder items reached the maximum recent files allowed.
        """
        while len(fileList()) > (threshold if not deleteAll else 0):                                            ## Process code until the detected files are below threshold or when the switch is set to "Delete All" (0)
            try: os.remove(min(fileList(), key=os.path.getatime))                                               ## Eliminates the oldest accessed file
            except PermissionError as e: LOG.warn(e)                                                            ## Issue: There is no solution for this one yet. Administrator permissions could be used in future releases
            except FileNotFoundError as e: pass                                                                 ## Ignore when file is not there anymore
        



class ContextMenus():
    def __init__(self, UI):
        ## Structure 
        UI.CTX_MENU = QtWidgets.QMenu(UI)
        UI.ACT_MODE_SWITCH = UI.CTX_MENU.addAction("Compact Mode" if UI == UIA else "Full Size Mode"); UI.ACT_MODE_SWITCH.setShortcut('Ctrl+M')
        UI.CTX_MENU.addSeparator()
        UI.ACT_MINIMIZE = UI.CTX_MENU.addAction("Minimize"); UI.ACT_MINIMIZE.setShortcut(QtGui.QKeySequence('Alt+Down'))
        UI.CTX_MENU.addSeparator() 
        UI.ACT_SETTINGS = UI.CTX_MENU.addAction("Settings"); UI.ACT_SETTINGS.setShortcut(QtGui.QKeySequence('Alt+S'))
        UI.ACT_EXIT = UI.CTX_MENU.addAction("Exit")

        ## Connections
        QtWidgets.QShortcut(QtGui.QKeySequence('Ctrl+M'), UI).activated.connect(lambda: UI.switchMode())
        QtWidgets.QShortcut(QtGui.QKeySequence('Alt+Down'), UI).activated.connect(lambda: UI.showMinimized())
        QtWidgets.QShortcut(QtGui.QKeySequence('Alt+S'), UI).activated.connect(lambda: UIB.enterWindow())


    def forwardEvent(self, UI, event):
        ACT = UI.CTX_MENU.exec_(UI.mapToGlobal(event.pos()))
        if ACT == UI.ACT_EXIT:
            UI.close()
            UIB.close()
            # SYS.closeEvent(event)
        if ACT == UI.ACT_SETTINGS:
            UIB.enterWindow()
        if ACT == UI.ACT_MINIMIZE:
            UI.showMinimized()
        if ACT == UI.ACT_MODE_SWITCH:
            UI.switchMode()




class Animations():
    def __init__(self):
        self.BTN_FADE = 200
        self.DEFAULT_DURATION = 200

    ## Button
    class ButtonAnimation():
        def __init__(self, buttonObject, hexStart, hexEnd, textFg, textBg, duration=None):
            self.BUTTON = buttonObject
            self.HEX_START = self.HEX_FROM = hexStart
            self.HEX_END = self.HEX_TO = hexEnd
            self.DURATION = duration if duration is not None else ANM.DEFAULT_DURATION
            self.TEXT_FOREGROUND = textFg
            self.TEXT_BACKGROUND = textBg
            self.ANIMATION = QtCore.QVariantAnimation(
                startValue = SYS.QCl(hexStart),
                endValue = SYS.QCl(hexEnd),
                valueChanged = self.buttonValueChanged,
                duration = duration,
                )
            self.ANIMATION.setEasingCurve(QEasingCurve.InOutCubic)

        def connectEvents(self):
            ## Button State Checking

            def animate(cycle):
                self.ANIMATION.setDirection(QtCore.QAbstractAnimation.Forward) if cycle else self.ANIMATION.setDirection(QtCore.QAbstractAnimation.Backward)
                self.ANIMATION.start()
            self.BUTTON.enterEvent = lambda event: animate(0)
            self.BUTTON.leaveEvent = lambda event: animate(1)
            self.ANIMATION.start()


        def buttonValueChanged(self, color):
            self.BACKGROUND = color
            self.FOREGROUND = SYS.QCl(self.TEXT_FOREGROUND) if self.ANIMATION.direction() == QtCore.QAbstractAnimation.Forward else SYS.QCl(self.TEXT_BACKGROUND)
            self.buttonUpdateStylesheet()
            self.BUTTON.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))


        def buttonUpdateStylesheet(self):
            self.BUTTON.setStyleSheet(f"""
                color: rgba{self.FOREGROUND.getRgb()};
                background-color: rgba{self.BACKGROUND.getRgb()};
                border: 1px solid {modHex(SYS.RGBtoHEX(self.BACKGROUND.getRgb()), 20)};
                """
                )




class QWGT_BROWSER(QtWidgets.QMainWindow):
    """
    Main UI Class for the software.
    Uses PyQT5 for establishing the UI.
    """
    def __init__(self, parent = None):
        QtWidgets.QMainWindow.__init__(self, parent)
        self.ANIM = QtCore.QPropertyAnimation(self, b"pos")
        self.ANIM_2 = QtCore.QPropertyAnimation(self, b"size")
        self.ANIM_2.setStartValue(QtCore.QSize(0, 0))
        self.ANIM_2.setEndValue(QtCore.QSize(530, 270))
        self.ANIM.setEasingCurve(QtCore.QEasingCurve.OutCubic)
        self.ANIM_2.setEasingCurve(QtCore.QEasingCurve.OutCubic)

        # SCRS = QtWidgets.QApplication.desktop().availableGeometry(); x, y = int((SCRS.width()-530)/2), int((SCRS.height()-225)/2)
        # self.ANIM.setStartValue(QtCore.QPoint(int(APP.desktop().availableGeometry().center().x()-(530/2)), APP.desktop().availableGeometry().bottom()-200))
        # self.ANIM.setEndValue(QtCore.QPoint(x,y))
        # self.ANIM.setDuration(250)
        # self.ANIM_2.setDuration(450)
        # self.GRP_ANIM = QtCore.QParallelAnimationGroup()
        # self.GRP_ANIM.addAnimation(self.ANIM)
        # self.GRP_ANIM.addAnimation(self.ANIM_2)
        # self.GRP_ANIM.start()

        LOG.info("Initializing User Interface (UIA)")
        self._tsl = QtCore.QCoreApplication.translate
        self.QUERY_TIMEOUT = False                                              ## Delay for when the program will execute query dumping. Will wait for timeout fill to go zero.
        self.EXEC_READY = False                                                 ## Set to True for when all is ready for launch. Otherwise, False
        self.LAST_QUERY_ID = 0                                                  ## Indicator to prevent duplicate query dumps when the input is still the same
        self.QUERY_TIMEOUT_FILL = 2                                             ## Number of how many tries before dumping (1 second x 2 times, (2 second)
        self.EXEC_POINTER = 0                                                   ## 0 Means the program will read input from search box, 1 means from recent box. Default is 0
        self.POS_OLD = self.pos()

    # # Events
    # def paintEvent(self, event):
    #     # get current window size♪
    #     s = self.size()
    #     QP = QtGui.QPainter()
    #     QP.begin(self)
    #     QP.setRenderHint(QtGui.QPainter.Antialiasing, True)
    #     PEN = QtGui.QPen(SYS.QCl(QSS.BORDER), 2)
    #     QP.setPen(PEN)
    #     QP.setBrush(SYS.QCl(QSS.palette2Hex('window')))
    #     QP.drawRoundedRect(0, 0, s.width(), s.height(), 20, 20)
    #     QP.end(

    def moveEvent(self, event) -> None:
        pass
        # SYS.windowBlur(0)
        # KTime.sleep(0.013)


    def mousePressEvent(self, event):
        self.POS_OLD = event.globalPos()


    def mouseMoveEvent(self, event):
        delta = QtCore.QPoint(event.globalPos() - self.POS_OLD)
        self.move(self.x()+delta.x(), self.y()+delta.y())
        self.POS_OLD = event.globalPos()
        

    def mouseReleaseEvent(self, event):
        SYS.windowBlur(1)


    def setupUi(self):
        """
        Initiates most of the GUI's elements.
        """
        self.setObjectName("WIN_BROWSER")
        # self.setFixedSize(530, 270)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        # self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        if CDATA['AlwaysOnTop'] == "True": self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)

        self.WIDGET = QtWidgets.QWidget(self); self.WIDGET.setObjectName("QWGT_BROWSER")
        self.LYT_GRID = QtWidgets.QGridLayout(self.WIDGET); self.LYT_GRID.setObjectName("LYT_GRID")
        self.GRID_HYMN = QtWidgets.QGridLayout(); self.GRID_HYMN.setObjectName("GRID_HYMN")
        self.GRID_MAIN = QtWidgets.QGridLayout(); self.GRID_MAIN.setObjectName("GRID_MAIN")
        
        SPC0 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding); self.LYT_GRID.addItem(SPC0, 3, 1, 1, 1)
        SPC1 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum); self.LYT_GRID.addItem(SPC1, 2, 2, 1, 1)
        SPC2 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum); self.GRID_MAIN.addItem(SPC2, 4, 0, 1, 1)
        SPC3 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum); self.LYT_GRID.addItem(SPC3, 2, 0, 1, 1)
        SPC4 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding); self.LYT_GRID.addItem(SPC4, 1, 1, 1, 1)

        self.LBL_BROWSER = QtWidgets.QLabel(self.WIDGET); self.LBL_BROWSER.setObjectName("LBL_BROWSER");# self.LBL_BROWSER.setAlignment(QtCore.Qt.AlignCenter)
        self.LBL_BROWSERB = QtWidgets.QLabel(self.WIDGET); self.LBL_BROWSERB.setObjectName("LBL_BROWSERB");# self.LBL_BROWSER.setAlignment(QtCore.Qt.AlignCenter)
        self.LNE_SEARCH = QtWidgets.QLineEdit(self.WIDGET); self.LNE_SEARCH.setObjectName("LNE_SEARCH"); self.LNE_SEARCH.setMinimumSize(300, 0); self.LNE_SEARCH.setClearButtonEnabled(True)
        self.BTN_ADDQUEUE = QtWidgets.QPushButton(self.WIDGET); self.BTN_ADDQUEUE.setObjectName("BTN_ADDQUEUE")
        self.BTN_QUEUES = QtWidgets.QPushButton(self.WIDGET); self.BTN_QUEUES.setObjectName("BTN_QUEUES")
        self.BTN_MORE = QtWidgets.QPushButton(self.WIDGET); self.BTN_MORE.setObjectName("BTN_MORE")
        self.LBL_PVW_TITLE = QtWidgets.QLabel(self.WIDGET); self.LBL_PVW_TITLE.setObjectName("LBL_PVW_TITLE")
        self.LBL_PVW_EQUIV = QtWidgets.QLabel(self.WIDGET); self.LBL_PVW_EQUIV.setObjectName("LBL_PVW_EQUIV")
        self.LBL_PVW_ACCSS = QtWidgets.QLabel(self.WIDGET); self.LBL_PVW_ACCSS.setObjectName("LBL_PVW_ACCSS"); self.LBL_PVW_ACCSS.setEnabled(False)
        self.CMB_RECENT = QtWidgets.QComboBox(self.WIDGET); self.CMB_RECENT.setObjectName("CMB_RECENT"); self.CMB_RECENT.setEditable(True); self.CMB_RECENT.setMinimumWidth(300); self.CMB_RECENT.lineEdit().setReadOnly(1)
        self.CMB_RECENT.view().window().setWindowFlags(Qt.Popup | Qt.FramelessWindowHint); self.CMB_RECENT.view().window().setAttribute(Qt.WA_TranslucentBackground)
        self.BTN_LAUNCH = QtWidgets.QPushButton(self.WIDGET); self.BTN_LAUNCH.setObjectName("BTN_LAUNCH"); self.BTN_LAUNCH.setEnabled(False)
     
        ACT_RET = QtWidgets.QAction(self, triggered=self.BTN_LAUNCH.animateClick); ACT_RET.setShortcut(QtGui.QKeySequence("Return"))
        ACT_ENTER = QtWidgets.QAction(self, triggered=self.BTN_LAUNCH.animateClick); ACT_ENTER.setShortcut(QtGui.QKeySequence("Enter"))
        self.BTN_LAUNCH.addActions([ACT_RET, ACT_ENTER])

        SYS.windowBlur(1)
        self.CMN = ContextMenus(self)
        self.contextMenuEvent = lambda event: self.CMN.forwardEvent(self, event)

        # Stylesheets
        QtGui.QFontDatabase.addApplicationFont(f'.res/fonts/{SYS.RES_FONT_TITLE}')
        BASE_FONT = QFont("Segoe UI", 9); BASE_FONT.setStyleStrategy(QFont.PreferAntialias)
        TITLE_FONT = QFont('Integral CF', 14); TITLE_FONT.setStyleStrategy(QFont.PreferAntialias)
        APP.setFont(BASE_FONT)
        
        self.LBL_BROWSER.setFont(TITLE_FONT)
        self.LBL_BROWSERB.setFont(QFont("Segoe UI Bold", 7))

        SYS.CURR_THEME = int(CDATA['Theme'] if CDATA['Theme'] in list('012') else 0)

        self.GRID_HYMN.addWidget(self.LNE_SEARCH, 0, 1, 1, 1)
        self.GRID_HYMN.addWidget(self.BTN_ADDQUEUE, 0, 2, 1, 1) 
        self.GRID_HYMN.addWidget(self.BTN_QUEUES, 0, 3, 1, 1) 
        self.GRID_HYMN.addWidget(self.BTN_MORE, 0, 5, 1, 1)
        self.LYT_GRID.addLayout(self.GRID_MAIN, 2, 1, 1, 1)

        self.STATUSBAR = QtWidgets.QStatusBar(self); self.STATUSBAR.setObjectName("STATUSBAR")
        self.STATUSBAR.setSizeGripEnabled(0)

        self.GRID_BUTTONS = QtWidgets.QGridLayout()
        self.GRID_BUTTONS.addWidget(self.CMB_RECENT, 0, 0, 1, 1)
        self.GRID_BUTTONS.addWidget(self.BTN_LAUNCH, 0, 1, 1, 1)

        self.GRID_TITLES = QtWidgets.QGridLayout()
        self.GRID_TITLES.addWidget(self.LBL_BROWSER, 0, 0, 0, 0)
        self.GRID_TITLES.addWidget(self.LBL_BROWSERB, 0, 1, 0, 1)
        self.GRID_MAIN.addLayout(self.GRID_TITLES, 0, 0, 1, 1)
        self.GRID_MAIN.addLayout(self.GRID_HYMN, 1, 0, 1, 1)
        self.GRID_MAIN.addLayout(self.GRID_BUTTONS, 8, 0, 1, 1)
        self.GRID_MAIN.addWidget(self.LBL_PVW_TITLE, 2, 0, 1, 1)
        self.GRID_MAIN.addWidget(self.LBL_PVW_EQUIV, 3, 0, 1, 1)
        self.GRID_MAIN.addWidget(self.LBL_PVW_ACCSS, 4, 0, 1, 1)

        self.setCentralWidget(self.WIDGET)
        self.setStatusBar(self.STATUSBAR)
        QtCore.QMetaObject.connectSlotsByName(self)

        ## Connections and Slots
        self.INP = InputCore()
        self.LNE_SEARCH.textChanged.connect(lambda: self.INP.updateDetails())
        self.BTN_LAUNCH.clicked.connect(lambda:  self.INP.executeFile())
        self.BTN_MORE.clicked.connect(lambda: UIB.enterWindow())
        self.CMB_RECENT.currentIndexChanged.connect(lambda: self.INP.updateDetails(True))
        self.CMB_RECENT.currentTextChanged.connect(lambda: self.INP.updateDetails(True))
        self.LNE_SEARCH.wheelEvent = lambda event: self.INP.searchFill(event)


        ## Initial Functions
        self.retranslateUi()
        HDB.genSearchSuggestions(UIA)
        self.updateRecentList()
        QSS.initStylesheet()
        LOG.info("UIA initialized successfully")


    def mouseDoubleClickEvent(self, event):
        self.switchMode()


    def switchMode(self):
        UIC.show(); UIC.activateWindow()
        UIA.hide()
        UIC.LNE_SEARCH.setText(self.LNE_SEARCH.text())
        DELTA = UIA.geometry(); UIC.move(DELTA.x()+100, DELTA.y()+100) ## Calculate the change in position so the next window would appear near it.
        CDATA['CompactMode'] = 'True'; CFG.dump()
        global UIZ
        UIZ = UIC


    def retranslateUi(self):
        """ Mostly set texts for the GUI """
        self.setWindowTitle(SW.NAME)
        self.LBL_BROWSER.setText("Hymnal Browser")
        self.LBL_BROWSER.setToolTip(f"v{SW.VERSION} {SW.VERSION_NAME}\n\nMade for Seventh-day Adventist Church\n© {SW.PROD_YEAR} {SW.AUTHOR}")
        self.LBL_BROWSERB.setText(f"v{SW.VERSION} {SW.VERSION_NAME}")
        self.LNE_SEARCH.setPlaceholderText("Search")
        self.CMB_RECENT.lineEdit().setPlaceholderText("Browse Recent")
        self.BTN_LAUNCH.setText(f"Welcome!")
        self.STATUSBAR.showMessage(f"Welcome, {SYS.USER_NAME}!")

    
    def updateRecentList(self, whenLaunched=0):
        """
        Refreshes the list of recent files that are in the temp folder
        """

        ## Get Filenames
        FILES = [f'{SYS.DIR_TEMP}\\{x}' for x in os.listdir(SYS.DIR_TEMP) if not x.startswith('~$') and x.endswith('.pptx')]

        ## Sort the files based on their accessed time state
        SORTED = [f'{x[len(SYS.DIR_TEMP)+1:-5]}' for x in sorted(FILES, key=os.path.getatime, reverse=True)]

        """
        There are some circumstances where PowerPoint would not report the new launched file
        as "accessed" which would result to a delay in updating the combo box.
        The code block below is a workaround that would move and adjust all the items in sorted list generated
        from the previous of file stats assuming that the list is not updated due to some reasons of delay.
        Issued: 10/12/2021 2:37 PM
        """
        if whenLaunched:
            SORTED[SORTED.index(f'{self.INP.CURRENT_HINFO.num} {self.INP.CURRENT_HINFO.title}')] = '0'

            for i in reversed(range(int(CDATA['MaxAllowedRecent']))):
                if i == SORTED.index('0'):                                                          ## When the index reaches the index where we should do the adjustment
                    for j in range(i): SORTED[i-(j)] = SORTED[i-(j+1)]                              ## Move all remaining indexes (decreasing) by one step
                    SORTED[0] = f'{self.INP.CURRENT_HINFO.num} {self.INP.CURRENT_HINFO.title}'              ## Sets the recent launched file to index zero[0] so it would show up first
                    break

        ## Insert all sorted files into combo box
        self.CMB_RECENT.clear()
        for i in range(len(SORTED)): self.CMB_RECENT.addItem(SORTED[i])
        self.CMB_RECENT.setCurrentIndex(0 if whenLaunched else -1)




class QWGT_SETTINGS(QtWidgets.QMainWindow):
    def __init__(self, parent = None):
        self.HASCHANGES = 0 # Indicating that the settings window is in it's unchanged state
        self.ACTIVE_PAGE = (0,0)
        QtWidgets.QMainWindow.__init__(self, parent)


    def setupUi(self):
        self.setObjectName("WIN_SETTINGS")
        self.setWindowIcon(QtGui.QIcon(SYS.RES_LOGO))
        self.setFixedSize(QtCore.QSize(800, 480))
        self.setWindowFlags(Qt.WindowCloseButtonHint)
        # self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        if CDATA['AlwaysOnTop'] == "True": self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)

        self.WIDGET = QtWidgets.QWidget(self); self.WIDGET.setObjectName("QWGT_SETTINGS")

        self.LYT_GRID_SETTINGS = QtWidgets.QGridLayout(self.WIDGET); self.LYT_GRID_SETTINGS.setObjectName("LYT_GRID_SETTINGS")
        self.FORM_SETTINGS = QtWidgets.QGridLayout(); self.FORM_SETTINGS.setObjectName("FORM_SETTINGS")
        spacerItem2 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.LST_NORTH = QtWidgets.QListWidget(self.WIDGET); self.LST_NORTH.setObjectName("LST_NORTH")
        self.LST_NORTH.setFixedSize(QtCore.QSize(150, 600))
        self.LST_SOUTH = QtWidgets.QListWidget(self.WIDGET); self.LST_SOUTH.setObjectName("LST_SOUTH")
        self.LST_SOUTH.setFixedSize(QtCore.QSize(150, 50))
        self.LST_PANELS = [self.LST_NORTH, self.LST_SOUTH]

        self.GRID_FOOTER = QtWidgets.QGridLayout(); self.GRID_FOOTER.setObjectName("GRID_FOOTER")

        self.STACKED_ACS_NORTH = QtWidgets.QStackedWidget(self.WIDGET); self.STACKED_ACS_NORTH.setObjectName("STACKED_ACS_NORTH")
        self.STACKED_ACS_SOUTH = QtWidgets.QStackedWidget(self.WIDGET); self.STACKED_ACS_SOUTH.setObjectName("STACKED_ACS_SOUTH")
        self.STACKED_ACCESS = [self.STACKED_ACS_NORTH, self.STACKED_ACS_SOUTH]

        self.BTN_OK = QtWidgets.QPushButton(self.WIDGET); self.BTN_OK.setObjectName("BTN_OK"); self.BTN_OK.setMinimumWidth(80)
        self.BTN_RESET = QtWidgets.QPushButton(self.WIDGET); self.BTN_RESET.setObjectName("BTN_RESET"); self.BTN_RESET.setMinimumWidth(80)
        self.FORM_SETTINGS.addWidget(self.LST_NORTH, 0,0,1,1)
        self.FORM_SETTINGS.addWidget(self.LST_SOUTH, 1,0,1,1)
        self.GRID_FOOTER.addWidget(self.BTN_RESET, 0,0,1,1)
        self.GRID_FOOTER.addItem(spacerItem2, 0,1,1,1)
        self.GRID_FOOTER.addWidget(self.BTN_OK, 0,2,1,1)
        for i in range(2): self.FORM_SETTINGS.addWidget(self.STACKED_ACCESS[i], 0,1,2,1)
        self.LYT_GRID_SETTINGS.addLayout(self.FORM_SETTINGS, 0, 0, 1, 1)
        self.LYT_GRID_SETTINGS.addLayout(self.GRID_FOOTER, 2,0,1,1)
        QtCore.QMetaObject.connectSlotsByName(UIB)
        UIB.setCentralWidget(self.WIDGET)

        ## Dock Settings
        self.TBL_DOCK = QtWidgets.QDockWidget(UIB); self.TBL_DOCK.setObjectName("TBL_DOCK")
        self.TBL_DOCK.setFeatures(QtWidgets.QDockWidget.DockWidgetFloatable | QtWidgets.QDockWidget.DockWidgetMovable)
        self.TBL_DOCK.setAllowedAreas(Qt.RightDockWidgetArea)
        self.TBL_DOCK_CONTENTS = QtWidgets.QWidget(); self.TBL_DOCK.setObjectName("TBL_DOCK_CONTENTS")
        self.TBL_GRID = QtWidgets.QGridLayout(self.TBL_DOCK_CONTENTS); self.TBL_GRID.setObjectName("TBL_GRID")
        self.TBL_STATISTICS = QtWidgets.QTableWidget(self.TBL_DOCK_CONTENTS); self.TBL_STATISTICS.setObjectName("TBL_STATISTICS")
        self.TBL_STATISTICS.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.TBL_STATISTICS_LOADED = False
        self.BTN_REFRESH_STATS = QtWidgets.QPushButton(self.TBL_DOCK_CONTENTS); self.BTN_REFRESH_STATS.setObjectName("BTN_REFRESH_STATS")
        self.BTN_DOCK_EXPORTCSV = QtWidgets.QPushButton(self.TBL_DOCK_CONTENTS); self.BTN_DOCK_EXPORTCSV.setObjectName("BTN_DOCK_EXPORTCSV")
        SPCH = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.TBL_GRID.addWidget(self.TBL_STATISTICS, 0, 0, 1, 3)
        self.TBL_GRID.addItem(SPCH, 1, 0, 1, 1)
        self.TBL_GRID.addWidget(self.BTN_REFRESH_STATS, 1, 1, 1, 1)
        self.TBL_GRID.addWidget(self.BTN_DOCK_EXPORTCSV, 1, 2, 1, 1)
        self.TBL_DOCK.setWidget(self.TBL_DOCK_CONTENTS)
        self.TBL_DOCK.setWindowTitle('Hymnal Statistics')

        self.TBL_DOCK.visibilityChanged.connect(lambda: self.CLASS_STS.updateResizing())
        self.BTN_REFRESH_STATS.clicked.connect(lambda: self.CLASS_STS.forceRefreshStats())
        self.BTN_DOCK_EXPORTCSV.clicked.connect(lambda: self.CLASS_STS.exportStatsTable())

        self.BTN_REFRESH_STATS.setText("Refresh")
        self.BTN_DOCK_EXPORTCSV.setText("Export CSV")


    
        UIB.addDockWidget(Qt.RightDockWidgetArea, self.TBL_DOCK)


        ## Subclasses
        self.CLASS_GEN = self.General()
        self.CLASS_STS = self.Statistics()
        self.CLASS_FAV = self.Library()
        self.CLASS_LOG = self.ShowLog()
        self.CLASS_ABT = self.About()

        # Connections
        self.BTN_RESET.clicked.connect(lambda: self.triggerButton(0))
        self.BTN_OK.clicked.connect(lambda: self.triggerButton(1))
        self.LST_NORTH.currentItemChanged.connect(lambda: self.updateSettingItems(0))
        self.LST_SOUTH.itemClicked.connect(lambda: self.updateSettingItems(1))
        self.LST_SOUTH.itemPressed.connect(lambda: self.updateSettingItems(1))


        # Initial Functions
        QSS.initStylesheet()
        self.retranslateUi()
        self.generatePanels()


    def generatePanels(self):
        ## Generate Panels
        TSL_PANELS = ['General', 'Hymnal', 'Library', 'Debugging'], ['About']
        self.LST_PANEL_ITEMS = [], []
        self.DLG_LST_PANELS = QtWidgets.QStyledItemDelegate(); self.DLG_LST_PANELS.setObjectName('DLG_LST_PANELS')
        for i, panelObj in enumerate(self.LST_PANELS):
            panelObj.setItemDelegate(self.DLG_LST_PANELS)
            panelObj.setSortingEnabled(False)
            for j, name in enumerate(TSL_PANELS[i]):
                panelObj.addItem(QtWidgets.QListWidgetItem())
                panelObj.item(j).setText(name)
                self.LST_PANEL_ITEMS[i].append(self.LST_PANELS[i].item(j))
        QSS.initStylesheet() ## <- Placed to instantiate icons


    def retranslateUi(self):
        self.setWindowTitle("Settings")
        self.BTN_RESET.setText("Reset")
        self.BTN_OK.setText("OK")


    def enterWindow(self):
        if UIB.isHidden():
            ## Relocate Settings Window | Shows the settings UI to right side of main window 
            SYS.centerInsideWindow(self, UIZ)
            self.LST_PANELS[0].setCurrentRow(0)
            self.updateSettingItems(0) ## General Tab should be the default highlighted panel

        UIB.activateWindow()
        UIB.show()
        UIB.raise_()
        UIB.setWindowState(UIB.windowState() & ~QtCore.Qt.WindowMinimized | QtCore.Qt.WindowActive)


    def triggerButton(self, mode):
        """
        Handles events related to the left-bottom button
        """
        if not mode:
            if self.ACTIVE_PAGE != (0, 1):
                # Reset to defaults
                MSG_BOX = QtWidgets.QMessageBox()
                MSG_BOX.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
                MSG_BOX.setIcon(QtWidgets.QMessageBox.Warning)
                MSG_BOX.setText("This will erase all saved preferences.\nDo you want to continue?")
                MSG_BOX.setWindowTitle("Reset to Defaults")
                MSG_BOX.setEscapeButton(QtWidgets.QMessageBox.No)
                MSG_BOX.setDefaultButton(QtWidgets.QMessageBox.No)
                MSG_BOX.setWindowFlags(Qt.Drawer | Qt.WindowStaysOnTopHint)
            
                MSG_BOX.setStyleSheet(QSS.getStylesheet())
                MSG_BOX.setStyleSheet("QLabel{min-width:300px}")
                RET = MSG_BOX.exec_()
                if RET == QtWidgets.QMessageBox.Yes:
                    # INSERT RESTORATION TO DEFAULTS HERE
                    os.remove(SYS.FILE_CONFIG)
                    CFG.check()

                    ## < ISSUE > Dark mode doesn't apply after loading
                    SYS.CURR_THEME = int(CDATA['Theme']) if int(CDATA['Theme']) in [0,1] else 1
                    SYS.COLORS = SYS.CTHEMES[SYS.CURR_THEME]
                    QSS.initStylesheet()
                    return
                else:
                    return
            else:
                # Reset Stats
                MSG_BOX = QtWidgets.QMessageBox()
                MSG_BOX.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
                MSG_BOX.setIcon(QtWidgets.QMessageBox.Warning)
                MSG_BOX.setText("This will permanently erase all saved data for hymnal.\nDo you want to continue?")
                MSG_BOX.setWindowTitle("Reset Statistics")
                MSG_BOX.setEscapeButton(QtWidgets.QMessageBox.No)
                MSG_BOX.setDefaultButton(QtWidgets.QMessageBox.No)
                MSG_BOX.setWindowFlags(Qt.Drawer | Qt.WindowStaysOnTopHint)
            
                MSG_BOX.setStyleSheet(QSS.getStylesheet())
                MSG_BOX.setStyleSheet("QLabel{min-width:300px}")
                RET = MSG_BOX.exec_()
                if RET == QtWidgets.QMessageBox.Yes:
                    os.remove(SYS.FILE_DATA)
                    SDB.check()
                    global SDATA
                    SDATA = SDB.DATA
                    UIB.CLASS_STS.updateStatsTable()

                    MSG_BOX = QtWidgets.QMessageBox()
                    MSG_BOX.setStandardButtons(QtWidgets.QMessageBox.Ok)
                    MSG_BOX.setIcon(QtWidgets.QMessageBox.Information)
                    MSG_BOX.setText("Successfully resetted hymnal stats")
                    MSG_BOX.setWindowTitle("Reset Statistics")
                    MSG_BOX.setEscapeButton(QtWidgets.QMessageBox.Ok)
                    MSG_BOX.setDefaultButton(QtWidgets.QMessageBox.Ok)
                    MSG_BOX.setWindowFlags(Qt.Drawer | Qt.WindowStaysOnTopHint)
                    MSG_BOX.setStyleSheet(QSS.getStylesheet())
                    MSG_BOX.exec_()
                    return
                else:
                    return

        if mode == 1:
            # Save changes (OK)
            pass
            
        if mode == 2:
            # Discard changes (Cancel)
            pass

        if self.HASCHANGES:
            # Dialogue Box for confirmation

            ### CHANGE THIS
            pass
        else:
            # Close settings window
            UIB.hide()


    def updateSettingItems(self, idx):
        """
        Change the layout widget with respect to the left panel in settings
        """
        self.TBL_DOCK.hide()
        self.ACTIVE_PAGE = (idx, self.LST_PANELS[idx].currentIndex().row())
        
        ID = self.LST_PANELS[idx].currentRow()
        self.STACKED_ACCESS[invert(idx)].hide(); self.LST_PANELS[invert(idx)].clearSelection()
        self.STACKED_ACCESS[idx].show()
        self.STACKED_ACCESS[idx].setCurrentIndex(ID)


        ## Default Changes
        UIB.BTN_RESET.setText("Reset")

        if not idx:
            if ID == 0: ## General Page
                self.BTN_RESET.show()

                    
            if ID == 1: ## Hymnal Page
                self.TBL_DOCK.show()
                self.BTN_RESET.show(); UIB.BTN_RESET.setText("Reset Stats")

            if ID == 2: ## Library Page
                self.BTN_RESET.hide()
                
            if ID == 3: ## Debugging Page
                self.BTN_RESET.hide()
                self.CLASS_LOG.updateLogContents()
                # UIB.resize(800, 480)
                # UIB.setMinimumSize(QtCore.QSize(800, 480))
                # UIB.setMaximumSize(QtCore.QSize(800, 480))
        else:
            if ID == 0: ## About Page
                self.BTN_RESET.hide()
                self.CLASS_ABT.displayText()



    class General():
        def __init__(self):
            self.PAGE_AC_GEN = QtWidgets.QWidget();  self.PAGE_AC_GEN.setObjectName("PAGE_AC_GEN")
            
            self.LBL_DARKMODE = QtWidgets.QLabel(self.PAGE_AC_GEN); self.LBL_DARKMODE.setObjectName("LBL_DARKMODE")
            self.CBX_DARKMODE = QtWidgets.QComboBox(self.PAGE_AC_GEN); self.CBX_DARKMODE.setObjectName("CBX_DARKMODE"); self.CBX_DARKMODE.setEditable(0)
            self.CBX_DARKMODE.view().window().setWindowFlags(Qt.Popup | Qt.FramelessWindowHint); self.CBX_DARKMODE.view().window().setAttribute(Qt.WA_TranslucentBackground)
            self.CHK_ALWAYSONTOP = QtWidgets.QCheckBox(self.PAGE_AC_GEN); self.CHK_ALWAYSONTOP.setObjectName("CHK_ALWAYSONTOP")
            self.LBL_WINOPACITY = QtWidgets.QLabel(self.PAGE_AC_GEN); self.LBL_WINOPACITY.setObjectName("LBL_WINOPACITY")
            self.SLD_WINOPACITY = QtWidgets.QSlider(Qt.Horizontal, self.PAGE_AC_GEN); self.SLD_WINOPACITY.setObjectName("SLD_WINOPACITY"); self.SLD_WINOPACITY.setMaximumWidth(70); self.SLD_WINOPACITY.setMinimum(50); self.SLD_WINOPACITY.setMaximum(100)

            self.CHK_AUTOSLIDESHOW = QtWidgets.QCheckBox(self.PAGE_AC_GEN); self.CHK_AUTOSLIDESHOW.setObjectName("CHK_AUTOSLIDESHOW")
            self.CHK_KEEP_FOCUS_WHEN_LAUNCH = QtWidgets.QCheckBox(self.PAGE_AC_GEN); self.CHK_KEEP_FOCUS_WHEN_LAUNCH.setObjectName("CHK_KEEP_FOCUS_WHEN_LAUNCH")

            self.GBX_INTERFACE = QtWidgets.QGroupBox('Interface', self.PAGE_AC_GEN); self.GBX_INTERFACE.setObjectName("GBX_INTERFACE")
            self.GBX_EXECUTION = QtWidgets.QGroupBox('Execution', self.PAGE_AC_GEN); self.GBX_EXECUTION.setObjectName("GBX_EXECUTION")
            self.LYT_INTERFACE = QtWidgets.QGridLayout(self.GBX_INTERFACE); self.LYT_INTERFACE.setObjectName("LYT_INTERFACE")
            self.LYT_EXECUTION = QtWidgets.QGridLayout(self.GBX_EXECUTION); self.LYT_EXECUTION.setObjectName("LYT_EXECUTION")

            SPCV = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
            SPCH = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
            self.GRID_AC_GEN = QtWidgets.QGridLayout(self.PAGE_AC_GEN); self.GRID_AC_GEN.setObjectName("GRID_AC_GEN")
            self.GRID_AC_GEN.addItem(SPCV, 50, 0, 1, 1) # Vertical Spacer inside settings
            self.GRID_AC_GEN.addItem(SPCH, 1, 50, 1, 1) # Horizontal Spacer inside settings

            self.LYT_INTERFACE.addWidget(self.LBL_DARKMODE, 0, 0, 1, 1)
            self.LYT_INTERFACE.addWidget(self.CBX_DARKMODE, 0, 1, 1, 1)
            self.LYT_INTERFACE.addWidget(self.CHK_ALWAYSONTOP, 1, 0, 1, 1)
            self.LYT_INTERFACE.addWidget(self.LBL_WINOPACITY, 3, 0, 1, 1)
            self.LYT_INTERFACE.addWidget(self.SLD_WINOPACITY, 3, 1, 1, 1)

            self.LYT_EXECUTION.addWidget(self.CHK_AUTOSLIDESHOW, 0, 0, 1, 1)
            self.LYT_EXECUTION.addWidget(self.CHK_KEEP_FOCUS_WHEN_LAUNCH, 1, 0, 1, 1)

            self.GRID_AC_GEN.addWidget(self.GBX_INTERFACE, 0, 0, 1, 1)
            self.GRID_AC_GEN.addWidget(self.GBX_EXECUTION, 1, 0, 1, 1)

            UIB.STACKED_ACCESS[0].addWidget(self.PAGE_AC_GEN)

            ## Configurations
            self.CHK_ALWAYSONTOP.setChecked(True if CDATA['AlwaysOnTop'] == 'True' else False) 
            self.CHK_AUTOSLIDESHOW.setChecked(True if CDATA['AutoSlideshow'] == 'True' else False)
            self.SLD_WINOPACITY.setValue(int(CDATA['WindowOpacity'])); self.SLD_WINOPACITY.setTickPosition(QtWidgets.QSlider.TicksAbove)
            if self.SLD_WINOPACITY.value() < SYS.MIN_OPACITY : self.SLD_WINOPACITY.setValue(SYS.MIN_OPACITY) ## Floor
            UIA.setWindowOpacity(self.SLD_WINOPACITY.value()/100);UIC.setWindowOpacity(self.SLD_WINOPACITY.value()/100)
            for name in ['Light', 'Dark', 'Black']: self.CBX_DARKMODE.addItem(name)
            self.CBX_DARKMODE.setCurrentIndex(SYS.CURR_THEME)


            ## Connections
            self.CBX_DARKMODE.currentIndexChanged.connect(lambda: self.updateThemeSelect())
            self.CHK_ALWAYSONTOP.stateChanged.connect(lambda: self.updateCheckboxes(1))
            self.CHK_AUTOSLIDESHOW.stateChanged.connect(lambda: self.updateCheckboxes(2))
            self.CHK_KEEP_FOCUS_WHEN_LAUNCH.stateChanged.connect(lambda: self.updateCheckboxes(3))
            self.SLD_WINOPACITY.valueChanged.connect(lambda: self.updateWindowOpacity())

            
            ## Retranslate UI
            self.LBL_DARKMODE.setText("Theme:")
            self.CHK_ALWAYSONTOP.setText("Always on Top")
            self.LBL_WINOPACITY.setText(f"Opacity ({self.SLD_WINOPACITY.value()}%):")

            self.CHK_AUTOSLIDESHOW.setText("Always start in slideshow")
            self.CHK_KEEP_FOCUS_WHEN_LAUNCH.setText("Keep focus on browser when opening a file")
            

        def updateThemeSelect(self):
            SYS.CURR_THEME = self.CBX_DARKMODE.currentIndex()
            CDATA['Theme'] = str(SYS.CURR_THEME); CFG.dump()
            QSS.initStylesheet()

        def updateWindowOpacity(self):
            self.LBL_WINOPACITY.setText(f'Opacity ({self.SLD_WINOPACITY.value()}%):')
            UIA.setWindowOpacity((self.SLD_WINOPACITY.value())/100); UIC.setWindowOpacity((self.SLD_WINOPACITY.value())/100)
            CDATA['WindowOpacity'] = str(self.SLD_WINOPACITY.value()); CFG.dump()


        def updateCheckboxes(self, mode):

            if mode == 1:
                ## Always on top
                CDATA['AlwaysOnTop'] = str(invert(CDATA['AlwaysOnTop'], True)); CFG.dump()
                if CDATA['AlwaysOnTop'] == 'True':
                    UIA.setWindowFlags(UIA.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)
                    UIB.setWindowFlags(UIB.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)
                    UIC.setWindowFlags(UIC.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)
                else:
                    UIA.setWindowFlags(UIA.windowFlags() & ~QtCore.Qt.WindowStaysOnTopHint)
                    UIB.setWindowFlags(UIB.windowFlags() & ~QtCore.Qt.WindowStaysOnTopHint)
                    UIC.setWindowFlags(UIC.windowFlags() & ~QtCore.Qt.WindowStaysOnTopHint)
                UIA.show()
                UIB.show()
                
            if mode == 2:
                # Auto Slideshow
                CDATA['AutoSlideshow'] = 'False' if CDATA['AutoSlideshow'] == 'True' else 'True'; CFG.dump()

            if mode == 3:
                # Auto Slideshow
                CDATA['KeepFocusOnBrowser'] = 'False' if CDATA['KeepFocusOnBrowser'] == 'True' else 'True'; CFG.dump()



    class Statistics():
        def __init__(self):
            self.PAGE_AC_STS = QtWidgets.QWidget();  self.PAGE_AC_STS.setObjectName("PAGE_AC_STS")

            self.GBX_DATABASE = QtWidgets.QGroupBox('Database', self.PAGE_AC_STS); self.GBX_DATABASE.setObjectName("GBX_DATABASE")
            self.GBX_FILES = QtWidgets.QGroupBox('Recent Files', self.PAGE_AC_STS); self.GBX_FILES.setObjectName("GBX_FILES")
            self.GBX_STATISTICS = QtWidgets.QGroupBox('Statistics', self.PAGE_AC_STS); self.GBX_STATISTICS.setObjectName("GBX_STATISTICS")
            self.LYT_DATABASE = QtWidgets.QGridLayout(self.GBX_DATABASE); self.LYT_DATABASE.setObjectName("LYT_DATABASE")
            self.LYT_FILES = QtWidgets.QGridLayout(self.GBX_FILES); self.LYT_FILES.setObjectName("LYT_FILES")
            self.LYT_STATISTICS = QtWidgets.QGridLayout(self.GBX_STATISTICS); self.LYT_STATISTICS.setObjectName("LYT_STATISTICS")
            
            self.LBL_DBSTATUS = QtWidgets.QLabel(self.PAGE_AC_STS); self.LBL_DBSTATUS.setObjectName("LBL_DBSTATUS")
            self.LBL_STATS = QtWidgets.QLabel(self.PAGE_AC_STS); self.LBL_STATS.setObjectName("LBL_STATS")

            self.LBL_RECENT_DETAILS = QtWidgets.QLabel(self.PAGE_AC_STS); self.LBL_RECENT_DETAILS.setObjectName("LBL_RECENT_DETAILS")
            self.LBL_MAX_ALLOWED_RECENT = QtWidgets.QLabel(self.PAGE_AC_STS); self.LBL_MAX_ALLOWED_RECENT.setObjectName("LBL_MAX_ALLOWED_RECENT")
            self.SBX_MAX_ALLOWED_RECENT = QtWidgets.QSpinBox(self.GBX_FILES); self.SBX_MAX_ALLOWED_RECENT.setObjectName("SBX_MAX_ALLOWED_RECENT"); self.SBX_MAX_ALLOWED_RECENT.setMaximum(SYS.RECENTS.ALLOWEDMAX), self.SBX_MAX_ALLOWED_RECENT.setMinimum(SYS.RECENTS.ALLOWEDMIN)
            self.SBX_MAX_ALLOWED_RECENT.wheelEvent = lambda event: None ## Prevents scroll
            self.SBX_MAX_ALLOWED_RECENT.lineEdit().setContentsMargins(20, 2, 2, 2)
            self.BTN_CLEAR_RECENT = QtWidgets.QPushButton(self.GBX_FILES); self.BTN_CLEAR_RECENT.setObjectName("BTN_CLEAR_RECENT"); self.BTN_CLEAR_RECENT.setMaximumWidth(120)

            self.BTN_TARGET_PATH = QtWidgets.QPushButton(self.PAGE_AC_STS); self.BTN_TARGET_PATH.setObjectName("BTN_TARGET_PATH")
            self.LNE_TARGET_PATH = QtWidgets.QLineEdit(self.PAGE_AC_STS); self.LNE_TARGET_PATH.setObjectName("LNE_TARGET_PATH"); self.LNE_TARGET_PATH.setReadOnly(1)

            self.LBL_STATISTICS = QtWidgets.QLabel(self.PAGE_AC_STS); self.LBL_STATISTICS.setObjectName("LBL_STATISTICS")
            self.BTN_IMPORT = QtWidgets.QPushButton(self.PAGE_AC_STS); self.BTN_IMPORT.setObjectName("BTN_IMPORT")
            self.BTN_EXPORT = QtWidgets.QPushButton(self.PAGE_AC_STS); self.BTN_EXPORT.setObjectName("BTN_EXPORT")


            SPCV = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
            # SPCH = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)

            self.GRID_AC_STS = QtWidgets.QGridLayout(self.PAGE_AC_STS); self.GRID_AC_STS.setObjectName("GRID_AC_STS")
            # self.GRID_AC_STS.addItem(SPCH, 1, 50, 1, 1) # Horizontal Spacer inside settings

            self.LYT_DATABASE.addWidget(self.LBL_DBSTATUS, 0, 0, 1, 2)
            self.LYT_DATABASE.addWidget(self.LBL_STATS, 1, 0, 1, 2)
            self.LYT_DATABASE.addWidget(self.LNE_TARGET_PATH, 2, 0, 1, 1)
            self.LYT_DATABASE.addWidget(self.BTN_TARGET_PATH, 2, 1, 1, 1)

            self.LYT_FILES.addWidget(self.LBL_MAX_ALLOWED_RECENT, 0, 0, 1, 1)
            self.LYT_FILES.addWidget(self.SBX_MAX_ALLOWED_RECENT, 0, 1, 1, 1)
            self.LYT_FILES.addWidget(self.BTN_CLEAR_RECENT, 1, 0, 1, 1)
            self.LYT_FILES.addWidget(self.LBL_RECENT_DETAILS, 1, 1, 1, 1)

            self.LYT_STATISTICS.addWidget(self.LBL_STATISTICS, 0, 1, 2, 1)
            self.LYT_STATISTICS.addWidget(self.BTN_IMPORT, 0, 0, 1, 1)
            self.LYT_STATISTICS.addWidget(self.BTN_EXPORT, 1, 0, 1, 1)

            self.GRID_AC_STS.addWidget(self.GBX_DATABASE, 0, 0, 1, 1)
            self.GRID_AC_STS.addWidget(self.GBX_FILES, 1, 0, 1, 1)
            self.GRID_AC_STS.addWidget(self.GBX_STATISTICS, 2, 0, 1, 1)
            self.GRID_AC_STS.addItem(SPCV, 30, 0, 1, 1) ## Vertical Spacer inside settings

            UIB.STACKED_ACCESS[0].addWidget(self.PAGE_AC_STS)

            ## Connections
            self.SBX_MAX_ALLOWED_RECENT.valueChanged.connect(lambda: self.updateRecentFiles(0))
            self.BTN_CLEAR_RECENT.clicked.connect(lambda: self.updateRecentFiles(1))
            self.BTN_TARGET_PATH.clicked.connect(lambda: self.changeDatabase())

            ## Translations
            self.BTN_CLEAR_RECENT.setText('Clear All')
            self.SBX_MAX_ALLOWED_RECENT.setValue(int(CDATA['MaxAllowedRecent']) if CDATA['MaxAllowedRecent'].isdigit() else SYS.RECENTS.DEFAULT)
            self.LBL_MAX_ALLOWED_RECENT.setText('Maximum Files: ')
            self.BTN_TARGET_PATH.setText('Change')
            self.LBL_STATISTICS.setText('JSON Data: data.json')
            self.BTN_IMPORT.setText('Import')
            self.BTN_EXPORT.setText('Export')
            self.LNE_TARGET_PATH.setText(SYS.FILE_HYMNSDB)
            self.updateWidget()

        
        def changeDatabase(self):
            PKG_PATH = QtWidgets.QFileDialog.getOpenFileName(None, 'Browse for Hymnal Package', os.getcwd(), "SDA Package (*.sda)")
            if PKG_PATH[0] != '':
                HDB.updateDatabase(PKG_PATH)
                # self.CFG["DIRS"].update({"SCAN_RECENT": str(PKG_PATH)}); DATA.dump(self.CFG)

            pass


        def updateRecentDetails(self):
            """
            Shows data of recent files such as number of present files and total sizes
            """
            PATHS = FMN.getRecentFiles(); TOTAL = len(PATHS)
            SIZES = [getFileStat(PATH, 'sz', szFmt='MB', szRnd=2) for PATH in PATHS]
            self.LBL_RECENT_DETAILS.setText(f'{f"{TOTAL} File(s) " if TOTAL !=0 else ""} <font color={str(QSS.TXT_DISABLED)}>{f"({round(sum(SIZES), 2)} MB)" if TOTAL !=0 else "No Files"}</font>')


        def updateRecentFiles(self, mode):
            """
            URetrieves new value from user and then dumps into configuration file.

            Modes:
            >>> 0 - Updates the maximum allowed recent files
            >>> 1 - 
            """
            if not mode:
                VAL = self.SBX_MAX_ALLOWED_RECENT.value()
                CDATA['MaxAllowedRecent'] = str(VAL); CFG.dump()
                FMN.deleteRecent()
            elif mode:
                FMN.deleteRecent(True)

            if UIZ == UIA: UIA.updateRecentList()


        def updateWidget(self):
            """
            Handles update whenever the stats page is visited
            """
            STATS = SDB.getStats()
            SIDENOTE = 'Hymns'
            DBSIZE = convertDataUnit(getFileStat(SYS.FILE_HYMNSDB, 'sz'), 'B', 'MB')
            if HYMNAL.MISSING.length: SIDENOTE = f'(There are {HYMNAL.MISSING.length} missing hymns.)'
            self.LBL_STATS.setWordWrap(True)
            self.LBL_DBSTATUS.setWordWrap(True)
            self.LBL_DBSTATUS.setText(f"<strong>{HYMNAL.TOTAL.ALL} / {SYS.HYMNS_MAX}</strong> {SIDENOTE} <font color='{QSS.TXT_DISABLED}'>({round(DBSIZE.val,2)} {DBSIZE.unit})</font>") 
            self.LBL_STATS.setText(f"{STATS.queries:,.0f} Total Queries\n{STATS.launches:,.0f} Total Launches")


        def updateResizing(self):
            """
            Handles resizing event when dock widget visibility was changed
            """
            self.resizeStatsTable()

            if UIB.TBL_DOCK.isFloating():
                UIB.TBL_DOCK.setMinimumSize(0,0)
                UIB.TBL_DOCK.setMaximumWidth(700)
                UIB.TBL_STATISTICS.resizeColumnsToContents()
            else:
                UIB.TBL_DOCK.setMinimumSize(380,300)
                UIB.TBL_DOCK.setMaximumWidth(380)
                if UIB.LST_PANELS[0].currentRow() == 1: UIB.setFixedSize(QtCore.QSize(800, 480))


        def resizeStatsTable(self):

            """ Handles Resizing """
            if UIB.TBL_DOCK.isFloating(): 
                TABLE_SIZE = UIB.size().height(), UIB.TBL_STATISTICS.size().width()
                SIZES = [7,45,7,7,15]
                for i in range(UIB.TBL_STATISTICS.columnCount()):
                    UIB.TBL_STATISTICS.setColumnWidth(i, tP(SIZES[i], TABLE_SIZE[1]))
            else:
                UIB.TBL_STATISTICS.setColumnWidth(0, 5)
                UIB.TBL_STATISTICS.setColumnWidth(1, 110)
                UIB.TBL_STATISTICS.setColumnWidth(2, 5)
                UIB.TBL_STATISTICS.setColumnWidth(3, 5)
                UIB.TBL_STATISTICS.setColumnWidth(4, 85)


        def updateStatsTable(self, isNew=False):
            """
            Updates the statistic table data from json
            """
            if UIB.TBL_STATISTICS_LOADED: return
            UIB.TBL_STATISTICS.clear()
            UIB.TBL_STATISTICS.setRowCount(0)
            UIB.TBL_STATISTICS.setColumnCount(0)
            # UIB.TBL_STATISTICS.horizontalHeader().setDefaultAlignment(Qt.AlignHCenter)
            # UIB.TBL_STATISTICS.verticalHeader().setDefaultAlignment(Qt.AlignVCenter)

            ROWS = SYS.HYMNS_MAX
            COLS = SYS.TBL_STATS_COLUMNS
    
            UIB.TBL_STATISTICS.setSortingEnabled(True)
            UIB.TBL_STATISTICS.setRowCount(ROWS)
            UIB.TBL_STATISTICS.setColumnCount(COLS)
            UIB.TBL_STATISTICS.setMinimumHeight(230)
            
            ## COLUMN
            TSL = ['#', 'Hymn', 'Q', 'L', 'Last Opened']
            for i in range(len(TSL)):
                UIB.TBL_STATISTICS.setHorizontalHeaderItem(i, QtWidgets.QTableWidgetItem())
                UIB.TBL_STATISTICS.horizontalHeaderItem(i).setText(str(TSL[i]))

            ## ROW
            for i in range(ROWS):
                UIB.TBL_STATISTICS.setVerticalHeaderItem(i, QtWidgets.QTableWidgetItem())
                UIB.TBL_STATISTICS.verticalHeaderItem(i).setText(str(i+1))
                UIB.TBL_STATISTICS.verticalHeaderItem(i).setTextAlignment(Qt.AlignHCenter)
                UIB.TBL_STATISTICS.setRowHeight(i, 5)

                ## ITEMS
                for j in range(COLS):
                    UIB.TBL_STATISTICS.setItem(i, j, QtWidgets.QTableWidgetItem())
                    UIB.TBL_STATISTICS.item(i, j).setFlags(QtCore.Qt.ItemIsEnabled) ## Makes the table read-only


                ## Insert Cell Details
                try:
                    HYMN = HDB.getStats(KString.toDigits(i+1,3))
                    h = [UIB.TBL_STATISTICS.item(i, j) for j in range(COLS+1)]

                    h[0].setText(str(HYMN.num))
                    h[1].setText(HYMN.title)
                    h[2].setData(Qt.DisplayRole, HYMN.queries)
                    h[3].setData(Qt.DisplayRole, HYMN.launches)
                    h[4].setText(HYMN.lastAcssHumanize if HYMN.lastAccessed != 0 else '')

                    h[0].setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                    h[2].setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                    h[3].setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                except AttributeError as e: LOG.warn(e)
        
            self.resizeStatsTable()
            UIB.TBL_STATISTICS_LOADED = True
        
        
        def exportStatsTable(self):
            """
            Exports the table into CSV
            """
            ## This method needs further improvement as it's still a generic.
            ## Note that the user should have the ability to choose where to put the file
            with open(f'HDB Stats {datetime.now().strftime("%m-%d-%Y")}.csv', 'w') as stream:
                writer = csv.writer(stream)
                for ROW in range(UIB.TBL_STATISTICS.rowCount()):
                    RDATA = []
                    for COL in range(UIB.TBL_STATISTICS.columnCount()):
                        ITEM = UIB.TBL_STATISTICS.item(ROW, COL)
                        if ITEM is not None:
                            RDATA.append(RDATA.append(ITEM.text()))
                        else:
                            RDATA.append('')
                    writer.writerow(RDATA)
        
        
        def forceRefreshStats(self):
            """
            This method forces all statistic details to be refreshed
            """
            UIB.TBL_STATISTICS_LOADED = False
            self.updateWidget()
            self.updateStatsTable()



    class Library():
        def __init__(self):
            self.PAGE_AC_LIB = QtWidgets.QWidget();  self.PAGE_AC_LIB.setObjectName("PAGE_AC_LIB")
            self.LBL_STATUS = QtWidgets.QLabel(self.PAGE_AC_LIB); self.LBL_STATUS.setObjectName("LBL_FOOTER_VER"); self.LBL_STATUS.setEnabled(0)
            self.GRID_AC_FAV = QtWidgets.QGridLayout(self.PAGE_AC_LIB); self.GRID_AC_FAV.setObjectName("GRID_AC_FAV")
            self.GRID_AC_FAV.addWidget(self.LBL_STATUS, 0, 0, 1, 1)
            self.LBL_STATUS.setText('Coming soon.') ## Will add in future versions
            UIB.STACKED_ACCESS[0].addWidget(self.PAGE_AC_LIB)



    class ShowLog():
        def __init__(self):
            self.PAGE_AC_LOG = QtWidgets.QWidget();  self.PAGE_AC_LOG.setObjectName("PAGE_AC_LIB")
            self.CBX_LOG_LIST = QtWidgets.QComboBox(self.PAGE_AC_LOG); self.CBX_LOG_LIST.setObjectName("CBX_LOG_LIST"); self.CBX_LOG_LIST.setEditable(0)
            self.CBX_LOG_LIST.view().window().setWindowFlags(Qt.Popup | Qt.FramelessWindowHint); self.CBX_LOG_LIST.view().window().setAttribute(Qt.WA_TranslucentBackground)
            self.CHK_AUTO_SCROLL = QtWidgets.QCheckBox(self.PAGE_AC_LOG); self.CHK_AUTO_SCROLL.setObjectName("CHK_AUTO_SCROLL"); self.CHK_AUTO_SCROLL.setText('Auto Scroll')
            self.LBL_LOG_CONTENT = QtWidgets.QLabel(self.PAGE_AC_LOG); self.LBL_LOG_CONTENT.setObjectName("LBL_LOG_CONTENT"); self.LBL_LOG_CONTENT.setEnabled(0)
            self.PTE_LOG_CONTENT = QtWidgets.QPlainTextEdit(self.PAGE_AC_LOG); self.PTE_LOG_CONTENT.setObjectName("PTE_LOG_CONTENT")
            self.PTE_LOG_CONTENT.setReadOnly(1)
            self.PTE_LOG_CONTENT.setWordWrapMode(QtGui.QTextOption.WrapMode.NoWrap)
            self.PTE_LOG_CONTENT.setFont(QFont('Consolas'))
            self.GRID_AC_LOG = QtWidgets.QGridLayout(self.PAGE_AC_LOG); self.GRID_AC_LOG.setObjectName("GRID_AC_FAV")
            SPCH = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)

            self.GRID_AC_LOG.addWidget(self.CBX_LOG_LIST, 0, 0, 1, 1)
            self.GRID_AC_LOG.addWidget(self.CHK_AUTO_SCROLL, 0, 1, 1, 1)
            self.GRID_AC_LOG.addItem(SPCH, 0, 1, 1, 1)
            self.GRID_AC_LOG.addWidget(self.LBL_LOG_CONTENT, 0, 2, 1, 1)
            self.GRID_AC_LOG.addWidget(self.PTE_LOG_CONTENT, 1, 0, 1, 3)
            UIB.STACKED_ACCESS[0].addWidget(self.PAGE_AC_LOG)

            ## Load Config
            self.CHK_AUTO_SCROLL.setCheckState(0 if CDATA['AutoScroll'] == 'True' else 2)

            ## Generate Log File Items
            self.LOG_LIST = sorted([(int(getFilename(filepath, '\\').split()[1][:-4]), filepath) for filepath in FMN.getLogFiles()], reverse=True)
            for i, tms in enumerate([i for i, j in self.LOG_LIST]):
                dtTms = datetime.fromtimestamp(tms).strftime("%I:%M %p %x"), datetime.fromtimestamp(tms).strftime("%I:%M %p")
                if not i: self.CBX_LOG_LIST.addItem(f'Log #{i+1}: Current File')        
                else: self.CBX_LOG_LIST.addItem(f'Log #{i+1}: {humanize.naturaltime(time.time()-tms) if time.time()-tms < 86400 else dtTms[0]}')        

            self.CBX_LOG_LIST.currentIndexChanged.connect(lambda: self.updateLogContents())
            self.CHK_AUTO_SCROLL.stateChanged.connect(lambda: self.updateAutoScroll())

            self.LOCKED, self.OLD_HSCRPOS, self.OLD_VSCRPOS = False, 0, 0
            self.V_SCROLLBAR, self.H_SCROLLBAR = self.PTE_LOG_CONTENT.verticalScrollBar(), self.PTE_LOG_CONTENT.horizontalScrollBar()
            self.V_SCROLLBAR.valueChanged.connect(lambda: self.updateScrollbars())
            self.H_SCROLLBAR.valueChanged.connect(lambda: self.updateScrollbars())

        def updateScrollbars(self):
            if self.LOCKED: return
            self.OLD_VSCRPOS = self.V_SCROLLBAR.value()
            self.OLD_HSCRPOS = self.H_SCROLLBAR.value()
        

        def updateAutoScroll(self):
            STATE = (0, 'False') if not self.CHK_AUTO_SCROLL.isChecked() else (2, 'True')
            CDATA['AutoScroll'] = STATE[1]; CFG.dump()
            self.CHK_AUTO_SCROLL.setCheckState(STATE[0])


        def updateLogContents(self):
            self.LBL_LOG_CONTENT.setText(f'Log File | {datetime.now().strftime("%I:%M:%S %p")} (Usage: {round((SYS.PROCESS.memory_info().rss)/10**6, 2)} MB)')

            if not self.PTE_LOG_CONTENT.hasFocus(): ## Prevents auto-refresh and unintentional deselect when the cursor is focused.
                try:
                    with open(self.LOG_LIST[self.CBX_LOG_LIST.currentIndex()][1]) as FILE:
                        CONTENT = FILE.read()
                except FileNotFoundError as e: CONTENT = "Log unavailable. The file might be already removed."
                    
                self.LOCKED = True
                self.PTE_LOG_CONTENT.setPlainText(CONTENT)
                self.V_SCROLLBAR.setValue(self.OLD_VSCRPOS)
                self.H_SCROLLBAR.setValue(self.OLD_HSCRPOS)
                if self.CHK_AUTO_SCROLL.isChecked(): self.PTE_LOG_CONTENT.verticalScrollBar().setValue(self.PTE_LOG_CONTENT.verticalScrollBar().maximum()) ## Keep the log scrolled to bottom
                self.LOCKED = False



    class About():
        def __init__(self):
            self.PAGE_AC_ABT = QtWidgets.QWidget();  self.PAGE_AC_ABT.setObjectName("PAGE_AC_ABT")
            
            self.LBL_FOOTER = QtWidgets.QLabel(self.PAGE_AC_ABT); self.LBL_FOOTER.setObjectName("LBL_FOOTER")
            self.LBL_FOOTER_NOTE = QtWidgets.QLabel(self.PAGE_AC_ABT); self.LBL_FOOTER_NOTE.setObjectName("LBL_FOOTER_NOTE"); self.LBL_FOOTER_NOTE.setEnabled(1)
            self.LBL_FOOTER_VER = QtWidgets.QLabel(self.PAGE_AC_ABT); self.LBL_FOOTER_VER.setObjectName("LBL_FOOTER_VER"); self.LBL_FOOTER_VER.setEnabled(0)
            self.LBL_LOGO = QtWidgets.QLabel(self.PAGE_AC_ABT); self.LBL_FOOTER.setObjectName("LBL_LOGO")
            self.BTN_REPORT_FEEDBACK = QtWidgets.QPushButton(self.PAGE_AC_ABT); self.BTN_REPORT_FEEDBACK.setObjectName("BTN_REPORT_FEEDBACK"); self.BTN_REPORT_FEEDBACK.setMaximumWidth(120)
            self.BTN_REPORT_FEEDBACK.setText('Report a feedback')

            self.LBL_LOGO.setPixmap(QtGui.QPixmap(SYS.RES_LOGO).scaledToHeight(150, Qt.SmoothTransformation))

            SPCV = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
            SPCV2 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
            SPCH = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)

            self.GRID_AC_ABT = QtWidgets.QGridLayout(self.PAGE_AC_ABT); self.GRID_AC_ABT.setObjectName("GRID_AC_ABT")
            self.GRID_AC_ABT.addItem(SPCH, 1, 50, 1, 1) # Horizontal Spacer inside settings
            self.GRID_AC_ABT.addWidget(self.LBL_LOGO, 0, 0, 1, 1)
            self.GRID_AC_ABT.addWidget(self.LBL_FOOTER, 1, 0, 2, 1)
            self.GRID_AC_ABT.addItem(SPCV, 3, 0, 1, 2) #  Vertical Spacer inside settings
            self.GRID_AC_ABT.addWidget(self.BTN_REPORT_FEEDBACK, 3, 0, 1, 1)
            self.GRID_AC_ABT.addWidget(self.LBL_FOOTER_NOTE, 2, 1, 1, 1)
            self.GRID_AC_ABT.addItem(SPCV2, 1, 1, 1, 2) #  Vertical Spacer inside settings
            self.GRID_AC_ABT.addWidget(self.LBL_FOOTER_VER, 3, 1, 1, 1)

            UIB.STACKED_ACCESS[1].addWidget(self.PAGE_AC_ABT)
            self.BTN_REPORT_FEEDBACK.clicked.connect(lambda: UIFB.enterWindow())
            
            self.LBL_FOOTER.setWordWrap(True); self.LBL_FOOTER_NOTE.setWordWrap(True)
            self.LBL_FOOTER.setOpenExternalLinks(True); self.LBL_FOOTER.setTextInteractionFlags(self.LBL_FOOTER.textInteractionFlags() | Qt.LinksAccessibleByMouse)
            self.LBL_FOOTER_NOTE.setOpenExternalLinks(True); self.LBL_FOOTER_NOTE.setTextInteractionFlags(self.LBL_FOOTER_NOTE.textInteractionFlags() | Qt.LinksAccessibleByMouse)
            
    
        def displayText(self):
            self.LBL_FOOTER.setText(f"<font>{SW.PARENT_NAME} {SW.NAME}<br><br>A general hymnal browser for Seventh-day Adventist Church that supports up to {SYS.HYMNS_MAX} Hymns based on SDA Hymnal Philippine Edition.<br><br>This software is part of MSDAC System's collection of softwares<br>© {SW.PROD_YEAR} <a href=\"https://m.me/verdaderoken\">Ken Verdadero</a> <a href=\"https://m.me/reynald.ycong\">Reynald Ycong</a><br><br><a href=\"hhttps://github.com/msdacsystems/hymnalbrowser\">GitHub Repository</a><br><a href=\"https://discord.gg/Ymsa2BUhJp\">Discord Support Server</a></font>")
            self.LBL_FOOTER_NOTE.setText(f"<font size=-1 color={QSS.TXT_DISABLED}>This program is still under development. In case of any unexpected bugs and crashes, we encourage you to report the problem on our Discord support server or send us a feedback using the feedback button.<br><br>The Seventh-day Adventist logo is a trademark property of General Conference of Seventh-day Adventists®.</font>")
            self.LBL_FOOTER_VER.setText(f"<font size=-1>Made for Seventh-day Adventist Church<br>© MSDAC Systems {SW.PROD_YEAR} v{SW.VERSION} {SW.VERSION_NAME}</font>")
    




class QWGT_REPORT_FEEDBACK(QtWidgets.QMainWindow):
    def __init__(self, parent = None):
        QtWidgets.QMainWindow.__init__(self, parent)
        self.MIN_INPUT = 30 ## In characters
        self.MAX_INPUT = 4000 ## In characters.
        self.OVERFLOW = False


    def setupUi(self):
        self.setObjectName("WIN_REPORT_FEEDBACK")
        self.setWindowIcon(QtGui.QIcon(SYS.RES_LOGO))
        self.setFixedSize(500, 300)
        self.setWindowFlags(Qt.WindowCloseButtonHint)
        if CDATA['AlwaysOnTop'] == "True": self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)
        self.WGT_CENTRAL = QtWidgets.QWidget(self); self.WGT_CENTRAL.setObjectName("WGT_CENTRAL")
        self.LYT_MAIN = QtWidgets.QGridLayout(self.WGT_CENTRAL); self.LYT_MAIN.setObjectName("gridLayout_2")
        self.GRID_BODY = QtWidgets.QGridLayout(); self.GRID_BODY.setObjectName("GRID_BODY")
        self.LBL_INSTRUCTION = QtWidgets.QLabel(self.WGT_CENTRAL); self.LBL_INSTRUCTION.setObjectName("LBL_INSTRUCTION")
        self.LBL_CHAR_INDIC = QtWidgets.QLabel(self.WGT_CENTRAL); self.LBL_CHAR_INDIC.setObjectName("LBL_CHAR_INDIC"); self.LBL_CHAR_INDIC.setAlignment(Qt.AlignRight)
        self.PTE_FEEDBACK = QtWidgets.QPlainTextEdit(self.WGT_CENTRAL); self.PTE_FEEDBACK.setObjectName("PTE_FEEDBACK")
        self.LBL_AGREEMENT = QtWidgets.QLabel(self.WGT_CENTRAL); self.LBL_AGREEMENT.setObjectName("LBL_AGREEMENT"), self.LBL_AGREEMENT.setWordWrap(True)
        self.CBX_AGREE = QtWidgets.QCheckBox(self.WGT_CENTRAL); self.CBX_AGREE.setObjectName("CBX_AGREE")
        self.LYT_FOOTER = QtWidgets.QHBoxLayout(); self.LYT_FOOTER.setObjectName("LYT_FOOTER")
        SPC_H = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.BTN_SEND = QtWidgets.QPushButton(self.WGT_CENTRAL); self.BTN_SEND.setObjectName("BTN_SEND")
        self.GRID_BODY.addWidget(self.LBL_INSTRUCTION, 0, 0, 1, 1)
        self.GRID_BODY.addWidget(self.LBL_CHAR_INDIC, 0, 1, 1, 1)
        self.GRID_BODY.addWidget(self.PTE_FEEDBACK, 1, 0, 1, 2)
        self.GRID_BODY.addWidget(self.CBX_AGREE, 2, 0, 1, 2)
        self.GRID_BODY.addWidget(self.LBL_AGREEMENT, 3, 0, 1, 2)
        self.LYT_FOOTER.addItem(SPC_H)
        self.LYT_FOOTER.addWidget(self.BTN_SEND)
        self.LYT_MAIN.addLayout(self.GRID_BODY, 0, 0, 1, 1)
        self.LYT_MAIN.addLayout(self.LYT_FOOTER, 1, 0, 1, 1)
        self.setCentralWidget(self.WGT_CENTRAL)

        self.retranslateUi()

        ## Connections
        self.BTN_SEND.clicked.connect(lambda: self.sendFeedbackForm())
        self.CBX_AGREE.stateChanged.connect(lambda: self.updateAgreementCheck())
        self.PTE_FEEDBACK.textChanged.connect(lambda: self.checkForm())

    
    def enterWindow(self):
        if UIFB.isHidden():
            ## Reset fields when the window is initiated again
            self.updateAgreementCheck()
            self.BTN_SEND.setEnabled(0)
            self.LBL_AGREEMENT.setEnabled(0)
            self.PTE_FEEDBACK.setPlainText('')
            self.LBL_CHAR_INDIC.setText(f"0/{self.MAX_INPUT}")

        UIFB.activateWindow()
        SYS.centerInsideWindow(self, UIB)
        UIFB.show()

    
    def checkForm(self):
        LENGTH = len(self.PTE_FEEDBACK.toPlainText())
        self.LBL_CHAR_INDIC.setText(f"{LENGTH}/{self.MAX_INPUT}")
        if not self.OVERFLOW:
            if LENGTH >= self.MAX_INPUT: self.OVERFLOW = True; self.PTE_FEEDBACK.setPlainText(self.PTE_FEEDBACK.toPlainText()[:self.MAX_INPUT])
        else: self.OVERFLOW = False
        self.updateAgreementCheck() if LENGTH >= self.MIN_INPUT and LENGTH <= self.MAX_INPUT else self.BTN_SEND.setEnabled(0)


    def updateAgreementCheck(self):
        if len(self.PTE_FEEDBACK.toPlainText()) >= self.MIN_INPUT:
            self.BTN_SEND.setEnabled(1 if self.CBX_AGREE.checkState() == 2 else 0)


    def retranslateUi(self):
        self.setWindowTitle("Report a feedback")
        self.LBL_INSTRUCTION.setText("Please state all your concerns and feedbacks about the application.")
        self.LBL_AGREEMENT.setText("By clicking \"Send\", you agree to include additional details such as system info, crash logs, and other necessary data in the report for the improvements of this application.")
        self.CBX_AGREE.setText("I am sure that all information I have written is accurate")
        self.BTN_SEND.setText("Send")
        self.PTE_FEEDBACK.setPlaceholderText(f'Write something at least {self.MIN_INPUT} characters')


    def sendFeedbackForm(self):
        """
        Sends all the data in a form of JSON via MongoDB
        """
        self.TME_START = time.time()
        self.USER_TEXT = self.PTE_FEEDBACK.toPlainText()
        self.PTE_FEEDBACK.setPlainText("Please wait while we're sending your feedback.")
        ## Build Report
        LOG.info('Generating report...')
        MDB.REPORT_TASK = 1 ## Triggers thread to initiate the task
        self.close()

        MSG_BOX = QtWidgets.QMessageBox()
        MSG_BOX.setStandardButtons(QtWidgets.QMessageBox.Ok)
        MSG_BOX.setIcon(QtWidgets.QMessageBox.Information)
        MSG_BOX.setText(f'{"Your feedback was sent successfully. Thank you." if SYS.isOnline() else "It seems like you are not connected to the internet.{nL}Your feedback will be sent whenever available."}')
        MSG_BOX.setWindowTitle(SW.NAME)
        MSG_BOX.setWindowFlags(Qt.Drawer | Qt.WindowStaysOnTopHint)
        SYS.centerInsideWindow(MSG_BOX, self)
        
        MSG_BOX.setStyleSheet(QSS.getStylesheet())
        MSG_BOX.exec_()




class QWGT_COMPACT(QtWidgets.QMainWindow):
    def __init__(self, parent = None):
        QtWidgets.QMainWindow.__init__(self, parent)
        LOG.info("Initializing User Interface (UIC)")
        self.EXEC_READY = False
        self.QUERY_TIMEOUT = False
        self.LAST_QUERY_ID = 0
        self.QUERY_TIMEOUT_FILL = 2


    def center(self):
        qr = self.frameGeometry()
        cp = QtWidgets.QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())


    def mousePressEvent(self, event):
        self.POS_OLD = event.globalPos()


    def mouseMoveEvent(self, event):
        delta = QtCore.QPoint(event.globalPos() - self.POS_OLD)
        self.move(self.x()+delta.x(), self.y()+delta.y())
        self.POS_OLD = event.globalPos()
        # SYS.windowBlur(0)

    def setupUi(self):
        self.setObjectName("WIN_COMPACT")
        self.setFixedSize(QtCore.QSize(380, 70))
        self.setWindowIcon(QtGui.QIcon(SYS.RES_LOGO))
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        if CDATA['AlwaysOnTop'] == "True": self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)
        self.WGT_CENTRAL = QtWidgets.QWidget(self); self.WGT_CENTRAL.setObjectName("QWGT_COMPACT")
        self.GRID_MAIN = QtWidgets.QGridLayout(self.WGT_CENTRAL); self.GRID_MAIN.setObjectName("gridLayout_2")
        self.gridLayout = QtWidgets.QGridLayout(); self.gridLayout.setObjectName("gridLayout")
        self.LNE_SEARCH = QtWidgets.QLineEdit(self.WGT_CENTRAL); self.LNE_SEARCH.setObjectName("LNE_SEARCH"); self.LNE_SEARCH.setClearButtonEnabled(True)
        self.LBL_PVW_TITLE = QtWidgets.QLabel(self.WGT_CENTRAL); self.LBL_PVW_TITLE.setObjectName("LBL_PVW_TITLE")
        self.BTN_LAUNCH = QtWidgets.QPushButton(self.WGT_CENTRAL); self.BTN_LAUNCH.setObjectName("BTN_LAUNCH"); self.BTN_LAUNCH.setEnabled(0)
        self.gridLayout.addWidget(self.LBL_PVW_TITLE, 0, 0, 1, 1)
        self.GRID_MAIN.addWidget(self.LNE_SEARCH, 1, 0, 1, 1)
        self.GRID_MAIN.addWidget(self.BTN_LAUNCH, 1, 1, 1, 1)
        self.GRID_MAIN.addLayout(self.gridLayout, 0, 0, 1, 1)
        self.setCentralWidget(self.WGT_CENTRAL)

        ACT_RET = QtWidgets.QAction(self, triggered=self.BTN_LAUNCH.animateClick); ACT_RET.setShortcut(QtGui.QKeySequence("Return"))
        ACT_ENTER = QtWidgets.QAction(self, triggered=self.BTN_LAUNCH.animateClick); ACT_ENTER.setShortcut(QtGui.QKeySequence("Enter"))
        self.BTN_LAUNCH.addActions([ACT_RET, ACT_ENTER])
        
        SYS.windowBlur(1)
        self.CMN = ContextMenus(self)
        self.contextMenuEvent = lambda event: self.CMN.forwardEvent(self, event)
        self.retranslateUi()
        QtCore.QMetaObject.connectSlotsByName(self)

        ## Connections and Slots
        self.INP = InputCore()
        self.LNE_SEARCH.textChanged.connect(lambda: self.INP.updateDetails())
        self.BTN_LAUNCH.clicked.connect(lambda: self.INP.executeFile())
        self.LNE_SEARCH.wheelEvent = lambda event: self.INP.searchFill(event)
        
        ## Initial Functions
        HDB.genSearchSuggestions(UIC)
        QSS.initStylesheet()
        LOG.info("UIC Initialized Successfully")


    def retranslateUi(self):
        self.LBL_PVW_TITLE.setText("Hymnal Browser")
        self.LNE_SEARCH.setPlaceholderText('Search')
        self.BTN_LAUNCH.setText("Insert a Hymn")

    ## Events
    def moveEvent(self, event) -> None:
        return
        KTime.sleep(0.013)


    def resizeEvent(self, event) -> None:
        KTime.sleep(0.013)
        

    def contextMenuEvent(self, event):
        self.CMN.forwardEvent(self, event)

    
    def mouseDoubleClickEvent(self, event):
        self.switchMode()


    def switchMode(self):    
        UIA.show(); UIA.activateWindow()
        UIC.hide()
        UIA.LNE_SEARCH.setText(self.LNE_SEARCH.text())                                  ## Also passes the entry text to other window
        DELTA = UIC.geometry(); UIA.move(DELTA.x()-100, DELTA.y()-100)          ## Calculate the change in position so the next window would appear near it.
        CDATA['CompactMode'] = 'False'; CFG.dump()
        global UIZ
        UIZ = UIA
        UIZ.INP.updateDetails()                                                  ## Force updates
        




class ThreadBackground(QtCore.QThread):
    UPT = QtCore.pyqtSignal()
    def run(self):
        while True:
            """ 
            This method is visited whenever the thread is reporting back the progress
            """
            gc.collect()

            try:
                if not UIZ.LNE_SEARCH.hasFocus() and UIZ == UIA: UIZ.STATUSBAR.showMessage('')
                
                # UIZ.INP.updateDetails() # <--- This is for real-time update of details

                if UIZ.QUERY_TIMEOUT:
                    if not UIZ.QUERY_TIMEOUT_FILL and int(UIZ.INP.CURRENT_HINFO.num) in list(range(SYS.HYMNS_MAX+1)):
                        if UIZ.LAST_QUERY_ID != UIZ.INP.CURRENT_HINFO.num:                                                                        ## Check if the last query is the same or different
                            ARRAY = [UIZ.INP.CURRENT_HINFO.queries+1, UIZ.INP.CURRENT_HINFO.launches, UIZ.INP.CURRENT_HINFO.lastAccessed]
                            SDATA['DATA'].update({str(UIZ.INP.CURRENT_HINFO.num): ARRAY}); SDB.dump()                                                    ## Log to Statistic Data
                            UIZ.LAST_QUERY_ID = UIZ.INP.CURRENT_HINFO.num
                        UIZ.QUERY_TIMEOUT = False
                        UIZ.QUERY_TIMEOUT_FILL = 2
                    else: UIZ.QUERY_TIMEOUT_FILL -=1

                
                if not UIB.isHidden():
                    """
                    Force Update UI for Settings Window. This is helpful for monitoring live updates when something is performed while looking at statistics.
                    """
                    if UIB.ACTIVE_PAGE == (0, 0): pass
                    if UIB.ACTIVE_PAGE == (0, 1): ## Database Panel
                        UIB.CLASS_STS.updateWidget() ## <--- Temporarily Disabled due to lagging
                    if UIB.ACTIVE_PAGE == (0, 2): pass
                    if UIB.ACTIVE_PAGE == (0, 3): pass # Log Panel
                    if UIB.ACTIVE_PAGE == (0, 4): pass

                time.sleep(1)
                try: self.UPT.emit()
                except AttributeError: pass
            except RuntimeError: pass
            
        
    def loopFunction(self):
        if not UIB.isHidden():
            """
            Force Update UI for Settings Window. (This one is for light processes. Please Fix)
            """
            if UIB.ACTIVE_PAGE == (0, 0): pass
            if UIB.ACTIVE_PAGE == (0, 1): ## Database Panel
                UIB.CLASS_STS.updateRecentDetails()
            if UIB.ACTIVE_PAGE == (0, 2): pass
            if UIB.ACTIVE_PAGE == (0, 3): # Log Panel
                UIB.CLASS_LOG.updateLogContents()
            if UIB.ACTIVE_PAGE == (0, 4): pass



if __name__ == "__main__":
    """
    Initialization of Classes and UI
    """
    INIT_TIME = time.time()

    ## Create Application
    APP = QtWidgets.QApplication(sys.argv)

    ## Primary Software Initialization & Logging
    SW = KSoftware("Hymnal Browser", "0.8.8", "Ken Verdadero, Reynald Ycong", file=__file__, parentName="MSDAC Systems", prodYear=2022, versionName="BETA")
    LOG = KLog(System().DIR_LOG, __file__, SW.LOG_NAME_DATE(), SW.PY_NAME, SW.AUTHOR, cont=True, tms=True, delete_existing=True, tmsformat="%H:%M:%S.%f %m/%d/%y")

    ## System Initialization and Verification of Directories
    SYS = System()
    SYS.verifyDirectories()
    # LOG.setLogPath(SYS.DIR_LOG) ## <- Rebind Log file into the logging folder  ## <- Disabled due to complicated file transfers. Fixed temporarily but needs improvement as it loses the first fragment of the log when initiated for the first time.

    ## Mongo DB
    MDB = Mongo()

    ## Configuration Parsing
    CFG = Configuration()
    
    CDATA = CFG.CONFIG[CFG.HEADNAME]

    ## Hymnal Database
    HDB = HymnsDatabase()
    global HYMNAL
    HYMNAL = HDB.parseHymnDatabase()                                                                ## Scan the whole hymnal

    ## Statistical Database
    SDB = Data()
    global SDATA
    SDATA = SDB.DATA

    ## Application Stylesheet
    QSS = Stylesheet()
    
    ## File Manager
    FMN = FileManager()

    ## Animation Class
    ANM = Animations()


    LOG.info("Initiating Program")

    ## User Interface Initialization
    UIA = QWGT_BROWSER(); UIZ = UIA ## Placeholder First
    UIC = QWGT_COMPACT()
    UIB = QWGT_SETTINGS()
    UIFB = QWGT_REPORT_FEEDBACK()

    UIA.setupUi()
    UIC.setupUi()
    UIB.setupUi()
    UIFB.setupUi()
    
    UIA.closeEvent = lambda event: SYS.closeEvent(event)
    UIC.closeEvent = lambda event: SYS.closeEvent(event)
    
    SYS.verifyRequisites()                                                                          ## Check for MS Office PowerPoint availability
    SYS.checkInstances()                                                                            ## Check for duplicate instances
    SYS.startBackgroundTask()

    ## Check what window mode should be displayed first
    LOG.info("Executing window")
    UIZ = UIA if CDATA['CompactMode'] == 'False' else UIC                                           ## 0 = Full Size, 1 = Compact | UIZ = Active Window
    
    ## Load Stats 
    if UIB.TBL_STATISTICS_LOADED == False:
        UIB.CLASS_STS.updateStatsTable()
        UIB.CLASS_STS.updateWidget()
    # LOG.info(time.time()-INIT_TIME)

    UIZ.show()
    SYS.centerWindow(UIZ)

    SYS.STARTUP_TIME = time.time()-INIT_TIME
    LOG.info(f'Initialization completed in {round(SYS.STARTUP_TIME, 3)} seconds.')

    ## Background Tasks
    MDB.REPORT_INITIAL = 1
    gc.collect()
    
    sys.exit(APP.exec_())