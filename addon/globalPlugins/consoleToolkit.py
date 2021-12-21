# -*- coding: UTF-8 -*-
#A part of  Console Toolkit addon for NVDA
#Copyright (C) 2019-2020 Tony Malykh
#This file is covered by the GNU General Public License.
#See the file COPYING.txt for more details.

import addonHandler
import api
import bisect
import collections
import config
import controlTypes
import core
import copy
import ctypes
from ctypes import create_string_buffer, byref
import documentBase
import editableText
import globalPluginHandler
import gui
from gui import guiHelper, nvdaControls
from gui.settingsDialogs import SettingsPanel
import inputCore
import itertools
import json
import keyboardHandler
from logHandler import log
import math
import mouseHandler
import NVDAHelper
from NVDAObjects import behaviors, NVDAObject
from NVDAObjects.IAccessible import IAccessible
from NVDAObjects.UIA import UIA
from NVDAObjects.window import winword
import nvwave
import operator
import os
import re
from scriptHandler import script, willSayAllResume
import speech
import string
import struct
import subprocess
import tempfile
import textInfos
import threading
import time
import tones
import types
import ui
import watchdog
import wave
import winUser
import wx

winmm = ctypes.windll.winmm


debug = False
if debug:
    import threading
    LOG_FILE_NAME = "C:\\Users\\tony\\Dropbox\\1.txt"
    f = open(LOG_FILE_NAME, "w")
    f.close()
    LOG_MUTEX = threading.Lock()
    def mylog(s):
        with LOG_MUTEX:
            f = open(LOG_FILE_NAME, "a", encoding='utf-8')
            print(s, file=f)
            #f.write(s.encode('UTF-8'))
            #f.write('\n')
            f.close()
else:
    def mylog(*arg, **kwarg):
        pass

def myAssert(condition):
    if not condition:
        raise RuntimeError("Assertion failed")



module = "consoleToolkit"
def initConfiguration():
    confspec = {
        "consoleRealtime" : "boolean( default=True)",
        "consoleBeep" : "boolean( default=True)",
        "controlVInConsole" : "boolean( default=True)",
        "deletePromptMethod" : "integer( default=3, min=0, max=3)",
        "captureSuffix" : f"string( default='|less -c 2>&1')",
        "captureChimeVolume" : "integer( default=5, min=0, max=100)",
        "captureOpenOption" : "integer( default=0, min=0, max=3)",
        "captureTimeout" : "integer( default=60, min=0, max=1000000)",
    }
    config.conf.spec[module] = confspec

def getConfig(key):
    value = config.conf[module][key]
    return value
def setConfig(key, value):
    config.conf[module][key] = value


addonHandler.initTranslation()
initConfiguration()


class SettingsDialog(SettingsPanel):
    # Translators: Title for the settings dialog
    title = _("Console Toolkit settings")


    def makeSettings(self, settingsSizer):
        sHelper = gui.guiHelper.BoxSizerHelper(self, sizer=settingsSizer)

      # checkbox console realtime
        # Translators: Checkbox for realtime console
        label = _("Speak console output in realtime")
        self.consoleRealtimeCheckbox = sHelper.addItem(wx.CheckBox(self, label=label))
        self.consoleRealtimeCheckbox.Value = getConfig("consoleRealtime")
      # checkbox console beep
        # Translators: Checkbox for console beep on update
        label = _("Beep on update in consoles")
        self.consoleBeepCheckbox = sHelper.addItem(wx.CheckBox(self, label=label))
        self.consoleBeepCheckbox.Value = getConfig("consoleBeep")
      # checkbox enforce control+V in console
        # Translators: Checkbox for control+V enforcement in console
        label = _("Always enable Control+V in console (useful for SSH)")
        self.controlVInConsoleCheckbox = sHelper.addItem(wx.CheckBox(self, label=label))
        self.controlVInConsoleCheckbox.Value = getConfig("controlVInConsole")

      # Delete method Combo box
        # Translators: Label for delete line method for prompt editing combo box
        label = _("Method of deleting lines for prompt editing:")
        self.deleteMethodCombobox = sHelper.addLabeledControl(label, wx.Choice, choices=deleteMethodNames)
        index = getConfig("deletePromptMethod")
        self.deleteMethodCombobox.Selection = index
      # Capture suffix edit
        self.captureSuffixEdit = sHelper.addLabeledControl(_("Suffix to be appendedd to commands in output capturing mode."), wx.TextCtrl)
        self.captureSuffixEdit.Value = getConfig("captureSuffix")
      # capture open option combo box
        label = _("Open captured output in")
        self.captureOpenOptionCombobox = sHelper.addLabeledControl(label, wx.Choice, choices=captureOpenOptionNames)
        index = getConfig("captureOpenOption")
        self.captureOpenOptionCombobox.Selection = index
      # Capture timeout edit
        self.captureTimeoutEdit = sHelper.addLabeledControl(_("Capture timeout in seconds:"), wx.TextCtrl)
        self.captureTimeoutEdit.Value = str(getConfig("captureTimeout"))

      # Output capture chime  volume slider
        sizer=wx.BoxSizer(wx.HORIZONTAL)
        label=wx.StaticText(self,wx.ID_ANY,label=_("Volume of chime while capturing command output"))
        slider=wx.Slider(self, wx.NewId(), minValue=0,maxValue=100)
        slider.SetValue(getConfig("captureChimeVolume"))
        sizer.Add(label)
        sizer.Add(slider)
        settingsSizer.Add(sizer)
        self.captureChimeVolumeSlider = slider

    def onSave(self):
        try:
            if int(self.captureTimeoutEdit.Value) <= 0:
                raise Exception()
        except:
            self.captureTimeoutEdit.SetFocus()
            ui.message(_("Capture timeout must be a positive integer"))
            return
        setConfig("consoleRealtime", self.consoleRealtimeCheckbox.Value)
        setConfig("consoleBeep", self.consoleBeepCheckbox.Value)
        setConfig("controlVInConsole", self.controlVInConsoleCheckbox.Value)
        setConfig("deletePromptMethod", self.deleteMethodCombobox.Selection)
        setConfig("captureSuffix", self.captureSuffixEdit.Value)
        setConfig("captureOpenOption", self.captureOpenOptionCombobox.Selection)
        setConfig("captureTimeout", int(self.captureTimeoutEdit.Value))
        setConfig("captureChimeVolume", self.captureChimeVolumeSlider.Value)

class Memoize:
    def __init__(self, f):
        self.f = f
        self.memo = {}
    def __call__(self, *args):
        if not args in self.memo:
            self.memo[args] = self.f(*args)
        #Warning: You may wish to do a deepcopy here if returning objects
        return self.memo[args]

class Beeper:
    BASE_FREQ = speech.IDT_BASE_FREQUENCY
    def getPitch(self, indent):
        return self.BASE_FREQ*2**(indent/24.0) #24 quarter tones per octave.

    BEEP_LEN = 10 # millis
    PAUSE_LEN = 5 # millis
    MAX_CRACKLE_LEN = 400 # millis
    MAX_BEEP_COUNT = MAX_CRACKLE_LEN // (BEEP_LEN + PAUSE_LEN)

    def __init__(self):
        self.player = nvwave.WavePlayer(
            channels=2,
            samplesPerSec=int(tones.SAMPLE_RATE),
            bitsPerSample=16,
            outputDevice=config.conf["speech"]["outputDevice"],
            wantDucking=False
        )
        self.stopSignal = False



    def fancyCrackle(self, levels, volume):
        levels = self.uniformSample(levels, self.MAX_BEEP_COUNT )
        beepLen = self.BEEP_LEN
        pauseLen = self.PAUSE_LEN
        pauseBufSize = NVDAHelper.generateBeep(None,self.BASE_FREQ,pauseLen,0, 0)
        beepBufSizes = [NVDAHelper.generateBeep(None,self.getPitch(l), beepLen, volume, volume) for l in levels]
        bufSize = sum(beepBufSizes) + len(levels) * pauseBufSize
        buf = ctypes.create_string_buffer(bufSize)
        bufPtr = 0
        for l in levels:
            bufPtr += NVDAHelper.generateBeep(
                ctypes.cast(ctypes.byref(buf, bufPtr), ctypes.POINTER(ctypes.c_char)),
                self.getPitch(l), beepLen, volume, volume)
            bufPtr += pauseBufSize # add a short pause
        self.player.stop()
        self.player.feed(buf.raw)

    def simpleCrackle(self, n, volume):
        return self.fancyCrackle([0] * n, volume)


    NOTES = "A,B,H,C,C#,D,D#,E,F,F#,G,G#".split(",")
    NOTE_RE = re.compile("[A-H][#]?")
    BASE_FREQ = 220
    def getChordFrequencies(self, chord):
        myAssert(len(self.NOTES) == 12)
        prev = -1
        result = []
        for m in self.NOTE_RE.finditer(chord):
            s = m.group()
            i =self.NOTES.index(s)
            while i < prev:
                i += 12
            result.append(int(self.BASE_FREQ * (2 ** (i / 12.0))))
            prev = i
        return result

    @Memoize
    def prepareFancyBeep(self, chord, length, left=10, right=10):
        beepLen = length
        freqs = self.getChordFrequencies(chord)
        intSize = 8 # bytes
        bufSize = max([NVDAHelper.generateBeep(None,freq, beepLen, right, left) for freq in freqs])
        if bufSize % intSize != 0:
            bufSize += intSize
            bufSize -= (bufSize % intSize)
        bbs = []
        result = [0] * (bufSize//intSize)
        for freq in freqs:
            buf = ctypes.create_string_buffer(bufSize)
            NVDAHelper.generateBeep(buf, freq, beepLen, right, left)
            bytes = bytearray(buf)
            unpacked = struct.unpack("<%dQ" % (bufSize // intSize), bytes)
            result = map(operator.add, result, unpacked)
        maxInt = 1 << (8 * intSize)
        result = map(lambda x : x %maxInt, result)
        packed = struct.pack("<%dQ" % (bufSize // intSize), *result)
        return packed

    def fancyBeep(self, chord, length, left=10, right=10, repetitions=1 ):
        self.player.stop()
        buffer = self.prepareFancyBeep(self, chord, length, left, right)
        self.player.feed(buffer)
        repetitions -= 1
        if repetitions > 0:
            self.stopSignal = False
            # This is a crappy implementation of multithreading. It'll deadlock if you poke it.
            # Don't use for anything serious.
            def threadFunc(repetitions):
                for i in range(repetitions):
                    if self.stopSignal:
                        return
                    self.player.feed(buffer)
            t = threading.Thread(target=threadFunc, args=(repetitions,))
            t.start()

    def uniformSample(self, a, m):
        n = len(a)
        if n <= m:
            return a
        # Here assume n > m
        result = []
        for i in range(0, m*n, n):
            result.append(a[i  // m])
        return result
    def stop(self):
        self.stopSignal = True
        self.player.stop()


def executeAsynchronously(gen):
    """
    This function executes a generator-function in such a manner, that allows updates from the operating system to be processed during execution.
    For an example of such generator function, please see GlobalPlugin.script_editJupyter.
    Specifically, every time the generator function yilds a positive number,, the rest of the generator function will be executed
    from within wx.CallLater() call.
    If generator function yields a value of 0, then the rest of the generator function
    will be executed from within wx.CallAfter() call.
    This allows clear and simple expression of the logic inside the generator function, while still allowing NVDA to process update events from the operating system.
    Essentially the generator function will be paused every time it calls yield, then the updates will be processed by NVDA and then the remainder of generator function will continue executing.
    """
    if not isinstance(gen, types.GeneratorType):
        raise Exception("Generator function required")
    try:
        value = gen.__next__()
    except StopIteration:
        return
    l = lambda gen=gen: executeAsynchronously(gen)
    core.callLater(value, executeAsynchronously, gen)

class SpeechChunk:
    def __init__(self, text, now):
        self.text = text
        self.timestamp = now
        self.spoken = False
        self.nextChunk = None

    def speak(self):
        def callback():
            global currentSpeechChunk, latestSpeechChunk
            #mylog(f'callback, pre acquire "{self.text}"')
            with speechChunksLock:
                #mylog(f'callback lock acquired!')
                if self != currentSpeechChunk:
                    # This can happen when this callback has already been scheduled, but new speech has arrived
                    # and this chunk was cancelled due to timeout.
                    return
                myAssert(
                    (
                        currentSpeechChunk == latestSpeechChunk
                        and self.nextChunk is None
                    )
                    or (
                        currentSpeechChunk != latestSpeechChunk
                        and self.nextChunk is not  None
                    )
                )

                currentSpeechChunk = self.nextChunk
                if self.nextChunk is not None:
                    self.nextChunk.speak()
                else:
                    latestSpeechChunk = None
        speech.speak([
            self.text,
            speech.commands.CallbackCommand(callback),
        ])

currentSpeechChunk = None
latestSpeechChunk = None
speechChunksLock = threading.RLock()
originalReportNewText = None
originalSpeechSpeak = None
originalCancelSpeech = None
def newReportConsoleText(selfself, line, *args, **kwargs):
    global currentSpeechChunk, latestSpeechChunk
    if getConfig("consoleBeep"):
        tones.beep(100, 5)
    if not getConfig("consoleRealtime"):
        return originalReportNewText(selfself, line, *args, **kwargs)
    now = time.time()
    threshold = now - 1
    newChunk = SpeechChunk(line, now)
    #mylog(f'newReportConsoleText pre acquire line="{line}"')
    with speechChunksLock:
        #mylog(f'newReportConsoleText lock acquired!')
        myAssert((currentSpeechChunk is not None) == (latestSpeechChunk is not None))
        if latestSpeechChunk is not None:
            latestSpeechChunk.nextChunk = newChunk
            latestSpeechChunk = newChunk
            if currentSpeechChunk.timestamp < threshold:
                originalCancelSpeech()
                while currentSpeechChunk.timestamp < threshold:
                    currentSpeechChunk = currentSpeechChunk.nextChunk
                currentSpeechChunk.speak()
        else:
            currentSpeechChunk = latestSpeechChunk = newChunk
            newChunk.speak()
    #mylog(f'newReportConsoleText lock released!')

def newCancelSpeech(*args, **kwargs):
    global currentSpeechChunk, latestSpeechChunk
    #mylog(f'newCancelSpeech pre acquire')
    with speechChunksLock:
        #mylog(f'newCancelSpeech lock acquired!')
        currentSpeechChunk = latestSpeechChunk = None
    #mylog(f'newCancelSpeech lock released!')
    return originalCancelSpeech(*args, **kwargs)

class SingleLineEditTextDialog(wx.Dialog):
    # This is a single line text edit window.
    def __init__(self, parent, text, onTextComplete):
        self.tabValue = "    "
        title_string = _("Edit text")
        super(SingleLineEditTextDialog, self).__init__(parent, title=title_string)
        self.text = text
        self.onTextComplete = onTextComplete
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        sHelper = gui.guiHelper.BoxSizerHelper(self, orientation=wx.VERTICAL)

        self.textCtrl = wx.TextCtrl(self, style=wx.TE_DONTWRAP|wx.TE_PROCESS_ENTER)
        self.textCtrl.Bind(wx.EVT_CHAR, self.onChar)
        self.Bind(wx.EVT_CHAR_HOOK, self.OnKeyUP)
        sHelper.addItem(self.textCtrl)
        self.textCtrl.SetValue(text)
        self.SetFocus()
        self.Maximize(True)

    def onChar(self, event):
        control = event.ControlDown()
        shift = event.ShiftDown()
        alt = event.AltDown()
        keyCode = event.GetKeyCode()
        if event.GetKeyCode() in [10,winUser.VK_RETURN]:
            # 10 means Control+Enter
            modifiers = [
                control, shift, alt
            ]
            modifierNames = [
                "control",
                "shift",
                "alt",
            ]
            modifierTokens = [
                modifierNames[i]
                for i in range(len(modifiers))
                if modifiers[i]
            ]
            keystrokeName = "+".join(modifierTokens + ["Enter"])
            self.keystroke = fromNameSmart(keystrokeName)
            self.text = self.textCtrl.GetValue()
            self.temporarilySuspendTerminalTitleAnnouncement()
            self.EndModal(wx.ID_OK)
            wx.CallAfter(lambda: self.onTextComplete(wx.ID_OK, self.text, self.keystroke))
        elif event.GetKeyCode() == wx.WXK_TAB:
            if alt or control:
                event.Skip()
            elif not shift:
                # Just Tab
                self.textCtrl.WriteText(self.tabValue)
            else:
                # Shift+Tab
                curPos = self.textCtrl.GetInsertionPoint()
                lineNum = len(self.textCtrl.GetRange( 0, self.textCtrl.GetInsertionPoint() ).split("\n")) - 1
                priorText = self.textCtrl.GetRange( 0, self.textCtrl.GetInsertionPoint() )
                text = self.textCtrl.GetValue()
                postText = text[len(priorText):]
                if priorText.endswith(self.tabValue):
                    newText = priorText[:-len(self.tabValue)] + postText
                    self.textCtrl.SetValue(newText)
                    self.textCtrl.SetInsertionPoint(curPos - len(self.tabValue))
        elif event.GetKeyCode() == wx.WXK_CONTROL_A:
            self.textCtrl.SetSelection(-1,-1)
        elif event.GetKeyCode() == wx.WXK_HOME:
            if not any([control, shift, alt]):
                curPos = self.textCtrl.GetInsertionPoint()
                #lineNum = len(self.textCtrl.GetRange( 0, self.textCtrl.GetInsertionPoint() ).split("\n")) - 1
                #colNum = len(self.textCtrl.GetRange( 0, self.textCtrl.GetInsertionPoint() ).split("\n")[-1])
                _, colNum,lineNum = self.textCtrl.PositionToXY(self.textCtrl.GetInsertionPoint())
                lineText = self.textCtrl.GetLineText(lineNum)
                m = re.search("^\s*", lineText)
                if not m:
                    raise Exception("This regular expression must match always.")
                indent = len(m.group(0))
                if indent == colNum:
                    newColNum = 0
                else:
                    newColNum = indent
                newPos = self.textCtrl.XYToPosition(newColNum, lineNum)
                self.textCtrl.SetInsertionPoint(newPos)
            else:
                event.Skip()
        else:
            event.Skip()


    def OnKeyUP(self, event):
        keyCode = event.GetKeyCode()
        if keyCode == wx.WXK_ESCAPE:
            self.text = self.textCtrl.GetValue()
            self.temporarilySuspendTerminalTitleAnnouncement()
            self.EndModal(wx.ID_CANCEL)
            wx.CallAfter(lambda: self.onTextComplete(wx.ID_CANCEL, self.text, None))
        event.Skip()

    def temporarilySuspendTerminalTitleAnnouncement(self):
        global suppressTerminalTitleAnnouncement
        suppressTerminalTitleAnnouncement = True
        def reset():
            global suppressTerminalTitleAnnouncement
            suppressTerminalTitleAnnouncement = False
        core.callLater(1000, reset)

class MultilineEditTextDialog(wx.Dialog):
    def __init__(self, parent, text, onTextComplete):
        self.tabValue = "    "
        # Translators: Title of  dialog
        title_string = _("Command output")
        super(MultilineEditTextDialog, self).__init__(parent, title=title_string)
        self.text = text
        self.onTextComplete = onTextComplete
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        sHelper = gui.guiHelper.BoxSizerHelper(self, orientation=wx.VERTICAL)

        self.textCtrl = wx.TextCtrl(self, style=wx.TE_MULTILINE|wx.TE_DONTWRAP)
        self.textCtrl.Bind(wx.EVT_CHAR, self.onChar)
        self.Bind(wx.EVT_CHAR_HOOK, self.OnKeyUP)
        sHelper.addItem(self.textCtrl)
        self.textCtrl.SetValue(text)
        self.SetFocus()
        self.Maximize(True)

    def onChar(self, event):
        control = event.ControlDown()
        shift = event.ShiftDown()
        alt = event.AltDown()
        keyCode = event.GetKeyCode()
        if event.GetKeyCode() in [10, winUser.VK_RETURN]:
            # 10 means Control+Enter
            modifiers = [
                control, shift, alt
            ]
            if not any(modifiers):
                # Just pure enter without any modifiers
                # Perform Autoindent
                curPos = self.textCtrl.GetInsertionPoint
                lineNum = len(self.textCtrl.GetRange( 0, self.textCtrl.GetInsertionPoint() ).split("\n")) - 1
                lineText = self.textCtrl.GetLineText(lineNum)
                m = re.search("^\s*", lineText)
                if m:
                    self.textCtrl.WriteText("\n" + m.group(0))
                else:
                    self.textCtrl.WriteText("\n")
            else:
                modifierNames = [
                    "control",
                    "shift",
                    "alt",
                ]
                modifierTokens = [
                    modifierNames[i]
                    for i in range(len(modifiers))
                    if modifiers[i]
                ]
                keystrokeName = "+".join(modifierTokens + ["Enter"])
                self.keystroke = fromNameSmart(keystrokeName)
                self.text = self.textCtrl.GetValue()
                self.EndModal(wx.ID_OK)
                #wx.CallAfter(lambda: self.onTextComplete(wx.ID_OK, self.text, self.keystroke))
        elif event.GetKeyCode() == wx.WXK_TAB:
            if alt or control:
                event.Skip()
            elif not shift:
                # Just Tab
                self.textCtrl.WriteText(self.tabValue)
            else:
                # Shift+Tab
                curPos = self.textCtrl.GetInsertionPoint()
                lineNum = len(self.textCtrl.GetRange( 0, self.textCtrl.GetInsertionPoint() ).split("\n")) - 1
                priorText = self.textCtrl.GetRange( 0, self.textCtrl.GetInsertionPoint() )
                text = self.textCtrl.GetValue()
                postText = text[len(priorText):]
                if priorText.endswith(self.tabValue):
                    newText = priorText[:-len(self.tabValue)] + postText
                    self.textCtrl.SetValue(newText)
                    self.textCtrl.SetInsertionPoint(curPos - len(self.tabValue))
        elif event.GetKeyCode() == wx.WXK_CONTROL_A:
            self.textCtrl.SetSelection(-1,-1)
        elif event.GetKeyCode() == wx.WXK_HOME:
            if not any([control, shift, alt]):
                curPos = self.textCtrl.GetInsertionPoint()
                #lineNum = len(self.textCtrl.GetRange( 0, self.textCtrl.GetInsertionPoint() ).split("\n")) - 1
                #colNum = len(self.textCtrl.GetRange( 0, self.textCtrl.GetInsertionPoint() ).split("\n")[-1])
                _, colNum,lineNum = self.textCtrl.PositionToXY(self.textCtrl.GetInsertionPoint())
                lineText = self.textCtrl.GetLineText(lineNum)
                m = re.search("^\s*", lineText)
                if not m:
                    raise Exception("This regular expression must match always.")
                indent = len(m.group(0))
                if indent == colNum:
                    newColNum = 0
                else:
                    newColNum = indent
                newPos = self.textCtrl.XYToPosition(newColNum, lineNum)
                self.textCtrl.SetInsertionPoint(newPos)
            else:
                event.Skip()
        else:
            event.Skip()


    def OnKeyUP(self, event):
        keyCode = event.GetKeyCode()
        if keyCode == wx.WXK_ESCAPE:
            self.text = self.textCtrl.GetValue()
            self.EndModal(wx.ID_CANCEL)
            #wx.CallAfter(lambda: self.onTextComplete(wx.ID_CANCEL, self.text, None))
        event.Skip()

def popupEditTextDialog(text, onTextComplete):
    gui.mainFrame.prePopup()
    d = SingleLineEditTextDialog(gui.mainFrame, text, onTextComplete)
    result = d.Show()
    gui.mainFrame.postPopup()

# This function is a fixed version of fromName function.
# As of v2020.3 it doesn't work correctly for gestures containing letters when the default locale on the computer is set to non-Latin, such as Russian.
import vkCodes
en_us_input_Hkl = 1033 + (1033 << 16)
def fromNameEnglish(name):
    """Create an instance given a key name.
    @param name: The key name.
    @type name: str
    @return: A gesture for the specified key.
    @rtype: L{KeyboardInputGesture}
    """
    keyNames = name.split("+")
    keys = []
    for keyName in keyNames:
        if keyName == "plus":
            # A key name can't include "+" except as a separator.
            keyName = "+"
        if keyName == keyboardHandler.VK_WIN:
            vk = winUser.VK_LWIN
            ext = False
        elif keyName.lower() == keyboardHandler.VK_NVDA.lower():
            vk, ext = keyboardHandler.getNVDAModifierKeys()[0]
        elif len(keyName) == 1:
            ext = False
            requiredMods, vk = winUser.VkKeyScanEx(keyName, en_us_input_Hkl)
            if requiredMods & 1:
                keys.append((winUser.VK_SHIFT, False))
            if requiredMods & 2:
                keys.append((winUser.VK_CONTROL, False))
            if requiredMods & 4:
                keys.append((winUser.VK_MENU, False))
            # Not sure whether we need to support the Hankaku modifier (& 8).
        else:
            vk, ext = vkCodes.byName[keyName.lower()]
            if ext is None:
                ext = False
        keys.append((vk, ext))

    if not keys:
        raise ValueError

    return keyboardHandler.KeyboardInputGesture(keys[:-1], vk, 0, ext)

def fromNameSmart(name):
    try:
        return keyboardHandler.KeyboardInputGesture.fromName(name)
    except:
        log.error(f"Couldn't resolve {name} keystroke using system default locale.", exc_info=True)
    try:
        return fromNameEnglish(name)
    except:
        log.error(f"Couldn't resolve {name} keystroke using English default locale.", exc_info=True)
    return None

originalTerminalGainFocus = None
originalNVDAObjectFfocusEntered = None
suppressTerminalTitleAnnouncement = False
def terminalGainFocus(self):
    if suppressTerminalTitleAnnouncement:
        # We only skip super() call here
        self.startMonitoring()
    else:
        return originalTerminalGainFocus(self)
def nvdaObjectFfocusEntered(self):
    if suppressTerminalTitleAnnouncement:
        return
    return originalNVDAObjectFfocusEntered(self)
    
class BackupClipboard:
    def __init__(self, text):
        self.backup = api.getClipData()
        self.text = text
    def __enter__(self):
        api.copyToClip(self.text)
        return self
    def __exit__(self, *args, **kwargs):
        core.callLater(300, self.restore)
    def restore(self):
        api.copyToClip(self.backup)

def interruptAndSpeakMessage(message):
    speech.cancelSpeech()
    ui.message(message)

def verifyWindowUnderMousePointer(obj):
    _windowFromPoint = ctypes.windll.user32.WindowFromPoint
    p = winUser.getCursorPos()
    hwnd=_windowFromPoint(ctypes.wintypes.POINT(p[0],p[1]))
    pid,tid=winUser.getWindowThreadProcessID(hwnd)
    if pid != obj.processID:
        title=winUser.getWindowText(hwnd)
        am=appModuleHandler.AppModule(pid)
        api.q = {
            'title': title,
            'hwnd': hwnd,
            'appModule': am,
        }
        volume = 50
        Beeper().fancyBeep("HF", 100, volume, volume)
        # Here is what I think we should do:
        if False:
            # https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowpos
            HWND_BOTTOM = ctypes.wintypes.HWND(1)
            winUser.user32.SetWindowPos(
                api.q['hwnd'],
                HWND_BOTTOM,
                0, 0, 0, 0,
                winUser.SWP_NOACTIVATE | winUser.SWP_NOMOVE | winUser.SWP_NOSIZE
            )
        raise Exception("Wrong window under mouse pointer. Relevant debug info has been saved to api.q.")

def pastePuttyOld(obj):
    # This approach doesn't appear to work, since apparently Putty triggers context menu on right click. Middle button doesn't appear to do anything either.
    origX, origY = winUser.getCursorPos()
    (left,top,width,height) = obj.location
    x=left+(width//2)
    y=top+(height//2)
    winUser.setCursorPos(x,y)
    mouseHandler.executeMouseEvent(winUser.MOUSEEVENTF_RIGHTDOWN,0,0)
    mouseHandler.executeMouseEvent(winUser.MOUSEEVENTF_RIGHTUP,0,0)
    core.callLater(100, interruptAndSpeakMessage, _("Pasted"))
    core.callLater(300, winUser.setCursorPos, origX, origY)

class ReleaseControlModifier:
    AttachThreadInput = winUser.user32.AttachThreadInput
    GetKeyboardState = winUser.user32.GetKeyboardState
    SetKeyboardState = winUser.user32.SetKeyboardState
    
    def __init__(self, obj):
        self.obj = obj
    def __enter__(self):
        hwnd =  self.obj.windowHandle
        processID,ThreadId = winUser.getWindowThreadProcessID(hwnd)
        self.ThreadId = ThreadId
        self.AttachThreadInput(ctypes.windll.kernel32.GetCurrentThreadId(), ThreadId, True)
        PBYTE256 = ctypes.c_ubyte * 256
        pKeyBuffers = PBYTE256()
        
        pKeyBuffers_old = PBYTE256()
        self.GetKeyboardState( ctypes.byref(pKeyBuffers_old ))
        self.pKeyBuffers_old = pKeyBuffers_old
        
        self.SetKeyboardState( ctypes.byref(pKeyBuffers) )
        return self
    def __exit__(self, *args, **kwargs):
        self.SetKeyboardState( ctypes.byref(self.pKeyBuffers_old) )
        self.AttachThreadInput(ctypes.windll.kernel32.GetCurrentThreadId(), self.ThreadId, False)


def pastePutty(obj):
    tones.beep(500, 50)
    with ReleaseControlModifier(obj):
        fromNameSmart("Shift+Insert").send()

def pasteConsole(obj):
    # This sends WM_COMMAND message, with ID of Paste item of context menu of command prompt window.
    # Don't ask me how I figured out its ID...
    # https://stackoverflow.com/questions/34410697/how-to-capture-the-windows-message-that-is-sent-from-this-menu
    WM_COMMAND = 0x0111
    watchdog.cancellableSendMessage(obj.parent.windowHandle, WM_COMMAND, 0xfff1, 0)

def pasteTerminal(obj):
    if isinstance(obj, PuttyControlV):
        pastePutty(obj)
    elif isinstance(obj, ConsoleControlV):
        pasteConsole(obj)
    else:
        raise Exception(f"Unknown object of type f={type(obj)}")
# Just some random unicode character that is not likely to appear anywhere.
# This character is used for prompt editing automation
#controlCharacter = "➉" # U+2789, Dingbat circled sans-serif digit ten
controlCharacter = "⌂" # character code 127



def getVkLetter(keyName):
    en_us_input_Hkl = 1033 + (1033 << 16)
    requiredMods, vk = winUser.VkKeyScanEx(keyName, en_us_input_Hkl)
    return vk
def getVkCodes():
    d = {}
    d['home'] = (winUser.VK_HOME, True)
    d['end'] = (winUser.VK_END, True)
    d['delete'] = (winUser.VK_DELETE, True)
    d['backspace'] = (winUser.VK_BACK, False)
    return d

KEYEVENTF_EXTENDEDKEY = 0x0001
def makeVkInput(pairs):
    result = []
    if not isinstance(pairs, list):
        pairs = [pairs]
    for pair in pairs:
        try:
            vk, extended = pair
        except TypeError:
            vk = pair
            extended = False
        input = winUser.Input(type=winUser.INPUT_KEYBOARD)
        input.ii.ki.wVk = vk
        input.ii.ki.dwFlags = (KEYEVENTF_EXTENDEDKEY * extended)
        result.append(input)
    for pair in reversed(pairs):
        try:
            vk, extended = pair
        except TypeError:
            vk = pair
        input = winUser.Input(type=winUser.INPUT_KEYBOARD)
        input.ii.ki.wVk = vk
        input.ii.ki.dwFlags = (KEYEVENTF_EXTENDEDKEY * extended) | winUser.KEYEVENTF_KEYUP
        result.append(input)
    return result

def makeUnicodeInput(string):
    result = []
    for c in string:
        input = winUser.Input(type=winUser.INPUT_KEYBOARD)
        input.ii.ki.wScan = ord(c)
        input.ii.ki.dwFlags = winUser.KEYEVENTF_UNICODE
        result.append(input)
        input2 = winUser.Input(type=winUser.INPUT_KEYBOARD)
        input2.ii.ki.wScan = ord(c)
        input2.ii.ki.dwFlags = winUser.KEYEVENTF_UNICODE | winUser.KEYEVENTF_KEYUP
        result.append(input2)
    return result

def script_editPrompt(self, gesture):
    executeAsynchronously(editPrompt(self, gesture))
script_editPrompt.category = "Console toolkit"
script_editPrompt.__name__ = _("Edit prompt")
script_editPrompt.__doc__ = _("Opens accessible window that allows to edit current command line prompt.")

def script_captureOutput(self, gesture):
    executeAsynchronously(captureOutputAsync(self, gesture))
script_captureOutput.category = "Console toolkit"
script_captureOutput.__name__ = _("Capture command output")
script_captureOutput.__doc__ = _("Executes command, captures output and presents it in accessible window.")

def captureOutputAsync(self, gesture):
    global captureStopFlag
    # That's the worst multithreading ever. Hope that if there's another thread running, they'll see this flag before  this function finishes. I'm ashamed of this.
    captureStopFlag = True
    for delay in waitUntilModifiersReleased():
        yield delay
    captureSuffix = getConfig("captureSuffix")
    prompt = []
    for token in extractCurrentPrompt(self, prompt):
        yield token
    prompt = prompt[0]
    prompt = prompt.rstrip()
    inputs = []
    if not prompt.endswith(captureSuffix):
        d = getVkCodes()
        inputs+= makeVkInput(d['end'])
        inputs += makeUnicodeInput(captureSuffix)
    inputs += makeVkInput([winUser.VK_RETURN])
    with keyboardHandler.ignoreInjection():
        winUser.SendInput(inputs)
    captureStopFlag =False

    executeAsynchronously(captureAsync(self, None))

def extractCurrentPrompt(obj, promptResult):
    # promptResult must be an empty list, where we will write the output
    # Poor man's pass by reference
    # There is no other good way to return a value, since we're yielding timeouts
    UIAMode = isinstance(obj, UIA)
    text = obj.makeTextInfo(textInfos.POSITION_ALL).text
    if controlCharacter in text:
        ui.message(_("Control character found on the screen; clear window and try again."))
        return
    d = getVkCodes()

    inputs = []
    inputs.extend(makeVkInput(d['end']))
    inputs.extend(makeUnicodeInput(controlCharacter))
    inputs.extend(makeVkInput(d['home']))
    inputs.extend(makeUnicodeInput(controlCharacter))
    controlCharactersAtStart = 1
    with keyboardHandler.ignoreInjection():
        winUser.SendInput(inputs)

    try:
        timeoutSeconds = 1
        timeout = time.time() + timeoutSeconds
        found = False
        while time.time() < timeout:
            text = obj.makeTextInfo(textInfos.POSITION_ALL).text
            indices = [i for i,c in enumerate(text) if c == controlCharacter]
            if len(indices) >= 2:
                found = True
                break
            yield 10
        if not found:
            msg = _("Timed out while waiting for control characters to appear.")
            ui.message(msg)
            raise Exception(msg)
        if len(indices) > 2:
            raise Exception(f"Unexpected: encountered {len(indices)} control characters!")
        # now we are sure that there are only two indices
        text1 = text[indices[0] + 1 : indices[1]]
        if UIAMode:
            # In UIA mode, UIA conveniently enough removes all the trailing spaces.
            # On multiline prompts therefore we cannot tell whether the end of the first line should be glued to the second line with or without spaces.
            # So we print another control character in the beginning to shift everything again by one more character to be able to tell,
            # whetehr there is a space between first and second lines, or every pair of lines, or no space.
            # Note however, that it is impossible to figure out the number of spaces, therefore when multiple spaces are present, their count is not guaranteed to be preserved.
            inputs = []
            inputs.extend(makeVkInput(d['home']))
            inputs.extend(makeUnicodeInput(controlCharacter))
            with keyboardHandler.ignoreInjection():
                winUser.SendInput(inputs)
            controlCharactersAtStart += 1
            timeoutSeconds = 1
            timeout = time.time() + timeoutSeconds
            found = False
            while time.time() < timeout:
                text = obj.makeTextInfo(textInfos.POSITION_ALL).text
                indices = [i for i,c in enumerate(text) if c == controlCharacter]
                if len(indices) >= 3:
                    found = True
                    break
                yield 10
            if not found:
                msg = _("Timed out while waiting for control characters to appear.")
                ui.message(msg)
                raise Exception(msg)
            if len(indices) > 3:
                raise Exception(f"Unexpected: encountered {len(indices)} control characters on second iteration in UIA mode!")
            text2 = text[indices[1] + 1 : indices[2]]
    finally:
        inputs = []
        inputs.extend(makeVkInput(d['home']))
        for dummy in range(controlCharactersAtStart):
            inputs.extend(makeVkInput(d['delete']))
        inputs.extend(makeVkInput(d['end']))
        inputs.extend(makeVkInput(d['backspace']))
        with keyboardHandler.ignoreInjection():
            winUser.SendInput(inputs)
    if UIAMode:
        text1 = text1.replace("\n", "").replace("\r", "")
        text2 = text2.replace("\n", "").replace("\r", "")
        # text1 and text2 should be mostly identical, with the only difference being spaces possibly injected in at certain positions. near the end of lines.
        # Combine text1 and text2 into oldText while preserving those spaces.
        result = []
        n = len(text1)
        m = len(text2)
        i = j = 0
        def reportMatchingProblem():
            message = f"In UIA mode, error while matching text1 and text2. i={i}, j={j}, n={n}, m={m};\n{text1}\n{text2}"
            raise Exception(message)
        while True:
            if i >= n and j >= m:
                break
            if i >= n:
                if text2[j] == " ":
                    result.append(" ")
                    j += 1
                    continue
                else:
                    reportMatchingProblem()
            if j >= m:
                if text1[i] == " ":
                    result.append(" ")
                    i += 1
                    continue
                else:
                    reportMatchingProblem()
            # now both i and j are within bounds
            if text1[i] == text2[j]:
                result.append(text1[i])
                i += 1
                j += 1
            elif text1[i] == " ":
                result.append(" ")
                i += 1
            elif text2[j] == " ":
                result.append(" ")
                j += 1
            else:
                reportMatchingProblem()
        result = "".join(result)
        oldText = result
        mylog(f"text1 in UIA mode!:")
        mylog(f"{text1}")
        mylog(f"text2:")
        mylog(f"{text2}")
        mylog(f"oldText:")
        mylog(f"{oldText}")
    else:
        oldText = text1.replace("\n", "").replace("\r", "")
        mylog(f"text1:")
        mylog(f"{text1}")
        mylog(f"oldText:")
        mylog(f"{oldText}")
    promptResult.append(oldText)
def editPrompt(obj, gesture):
    global captureStopFlag
    captureStopFlag = True
    prompt = []
    for token in extractCurrentPrompt(obj, prompt):
        yield token
    prompt = prompt[0]
    # Strip capturing suffix if found
    oldText = prompt.rstrip()
    suffix = getConfig("captureSuffix").rstrip()
    if oldText.endswith(suffix):
        oldText = oldText[:-len(suffix)]
    onTextComplete = lambda result, newText, keystroke: executeAsynchronously(updatePrompt(result, newText, keystroke, oldText, obj))
    popupEditTextDialog(oldText, onTextComplete)


DELETE_METHOD_CONTROL_C = 0
DELETE_METHOD_ESCAPE = 1
DELETE_METHOD_CONTROL_K = 2
DELETE_METHOD_BACKSPACE = 3
deleteMethodNames = [
    _("Control+C: works in both cmd.exe and bash, but leaves previous prompt visible on the screen; doesn't work in emacs; sometimes unreliable on slow SSH connections"),
    _("Escape: works only in cmd.exe"),
    _("Control+A Control+K: works in bash and emacs; doesn't work in cmd.exe"),
    _("Backspace (recommended): works in all environments; however slower and may cause corruption if the length of the line has changed"),
]

def updatePrompt(result, text, keystroke, oldText, obj):
    yield from waitUntilModifiersReleased()
    doCapture = False
    rawCommand = text
    if result == wx.ID_OK:
        modifiers = keystroke.modifierNames
        mainKeyName = keystroke.mainKeyName
        if (
            modifiers == ["control"]
            and mainKeyName == "enter"
        ):
            text += getConfig("captureSuffix")
            doCapture = True

    obj.setFocus()
    yield 10 # if we don't capture output, we need NVDA to see current screen, so that the updates will be spoken correctly
    method = getConfig("deletePromptMethod")
    inputs = []
    if method == DELETE_METHOD_CONTROL_C:
        inputs.extend(makeVkInput([winUser.VK_LCONTROL, getVkLetter("C")]))
    elif method == DELETE_METHOD_ESCAPE:
        inputs.extend(makeVkInput(winUser.VK_ESCAPE))
    elif method == DELETE_METHOD_CONTROL_K:
        inputs.extend(makeVkInput([winUser.VK_LCONTROL, getVkLetter("A")]))
        inputs.extend(makeVkInput([winUser.VK_LCONTROL, getVkLetter("K")]))
    elif method == DELETE_METHOD_BACKSPACE:
        inputs.extend(makeVkInput(winUser.VK_END))
        for dummy in range(len(oldText)):
            inputs.extend(makeVkInput(winUser.VK_BACK))
    else:
        raise Exception(f"Unknown method {method}!")
    if isinstance(obj, PuttyControlV):
        inputs.extend(makeVkInput([winUser.VK_SHIFT, (winUser.VK_INSERT, True)]))
    with BackupClipboard(text):
        with keyboardHandler.ignoreInjection():
            winUser.SendInput(inputs)
        if isinstance(obj, ConsoleControlV):
            pasteConsole(obj)

    if doCapture:
        fromNameSmart("Enter").send()
        global captureStopFlag
        captureStopFlag = False
        executeAsynchronously(captureAsync(obj, rawCommand))
    elif result == wx.ID_OK:
        keystroke.send()


allModifiers = [
    winUser.VK_LCONTROL, winUser.VK_RCONTROL,
    winUser.VK_LSHIFT, winUser.VK_RSHIFT, winUser.VK_LMENU,
    winUser.VK_RMENU, winUser.VK_LWIN, winUser.VK_RWIN,
]

def waitUntilModifiersReleased():
    timeoutSeconds = 5
    timeout = time.time() + timeoutSeconds
    while time.time() < timeout:
        status = [
            winUser.getKeyState(k) & 32768
            for k in allModifiers
        ]
        if not any(status):
            return
        yield 10
    message = _("Timed out while waiting for modifiers to be released!")
    ui.message(message)
    raise Exception(message)

def injectKeystroke(hWnd, vkCode):
    # Here we use PostMessage() and WM_KEYDOWN event to inject keystroke into the terminal.
    # Alternative ways, such as WM_CHAR event, or using SendMessage can work in plain command prompt, but they don't appear to work in any falvours of ssh.
    # We don't use SendInput() function, since it can only send keystrokes to the active focused window,
    # and here we would like to be able to send keystrokes to console window regardless whether it is focused or not.
    mylog(f"injectKeystroke({vkCode}, {hWnd})")
    WM_KEYDOWN                      =0x0100
    WM_KEYUP                        =0x0101
    winUser.PostMessage(hWnd, WM_KEYDOWN, vkCode, 1)
    winUser.PostMessage(hWnd, WM_KEYUP, vkCode, 1 | (1<<30) | (1<<31))

captureBeeper = Beeper()
captureStopFlag = False
def captureAsync(obj, rawCommand):
    timeoutSeconds = getConfig("captureTimeout")
    timeout = time.time() + timeoutSeconds
    start = time.time()
    result = []
    if rawCommand is not None:
        result.append(f"$ {rawCommand}")
    previousLines = []
    previousLinesCounter = 0
    captureBeeper.fancyBeep("CDGA", length=5000, left=5, right=5, repetitions =int(math.ceil(timeoutSeconds / 5)) )
    try:
        while time.time() < timeout:
            t = time.time() - start
            mylog(f"{t:0.3}")
            if captureStopFlag:
                ui.message(_("Capture interrupted!"))
                return
            textInfo = obj.makeTextInfo(textInfos.POSITION_ALL)
            if isinstance(obj, UIA):
                lines = textInfo.text.split("\r\n")
            else:
                # Legacy winConsole support
                lines = list(textInfo.getTextInChunks(textInfos.UNIT_LINE))
            if lines == previousLines:
                mylog(f"Screen hasn't changed! counter={previousLinesCounter}")
                previousLinesCounter += 1
                if previousLinesCounter < 10:
                    yield 10
                    continue
                mylog("Current lines:")
                for line in lines:
                    line = line.rstrip("\r\n")
                    mylog(f"    {line}")
            previousLines = lines
            previousLinesCounter = 0
            lastLine = lines[-1].rstrip()
            pageComplete = lastLine == ":"
            fileComplete= lastLine == "(END)"
            mylog(f"pageComplete={pageComplete} fileComplete={fileComplete}")
            if fileComplete:
                index = len(lines) - 1
                while index > 0 and lines[index - 1].rstrip() == "~":
                    index -= 1
                lines = lines[:index]
                result += lines
                # Sending q letter to quit less command
                #watchdog.cancellableSendMessage(obj.windowHandle, WM_CHAR, 0x71, 0)
                injectKeystroke(obj.windowHandle, 0x51)
                presentCaptureResult(result)
                return
            elif pageComplete:
                result += lines[:-1]
                # Sending space key:
                #watchdog.cancellableSendMessage(obj.windowHandle, WM_CHAR, 0x20, 0)
                injectKeystroke(obj.windowHandle, 0x20)
            else:
                yield 1
    finally:
        captureBeeper.stop()
    message = _("Timed out while waiting for command output!")
    ui.message(message)
    raise Exception(message)

CAPTURE_COPY_TO_CLIPBOARD = 0
CAPTURE_OPEN_TEMP_WINDOW = 1
CAPTION_OPEN_NOTEPAD = 2
CAPTION_OPEN_NPP = 3
captureOpenOptionNames = [
    _("Copy to clipboard"),
    _("Open in temporary window"),
    _("Open in Notepad"),
    _("Open in Notepad++"),
]
def presentCaptureResult(lines):
    output = "\r\n".join(lines)
    option = getConfig("captureOpenOption")
    if option == CAPTURE_COPY_TO_CLIPBOARD:
        api.copyToClip(output)
        ui.message(_("Command output copied to clipboard"))
    elif option == CAPTURE_OPEN_TEMP_WINDOW:
        gui.mainFrame.prePopup()
        d = MultilineEditTextDialog(gui.mainFrame, output, None)
        result = d.Show()
        gui.mainFrame.postPopup()
    elif option in [CAPTION_OPEN_NOTEPAD, CAPTION_OPEN_NPP]:
        # Prepare temp file
        tf = tempfile.NamedTemporaryFile(delete=False, prefix="temp_")
        tf.write(output.encode('utf-8'))
        tf.close()
        if option == CAPTION_OPEN_NOTEPAD:
            subprocess.Popen(f"""notepad "{tf.name}" """)
        elif option ==         CAPTION_OPEN_NPP:
            os.system(f"""notepad++ "{tf.name}" """)
    else:
        raise Exception(f"Unknown option {option}")


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    scriptCategory = _("Console toolkit")

    def __init__(self, *args, **kwargs):
        super(GlobalPlugin, self).__init__(*args, **kwargs)
        self.createMenu()
        self.injectHooks()
        self.lastConsoleUpdateTime = 0
        self.beeper = Beeper()

    def chooseNVDAObjectOverlayClasses(self, obj, clsList):
        if getConfig("controlVInConsole") and obj.windowClassName == 'ConsoleWindowClass':
            clsList.insert(0, ConsoleControlV)
        if getConfig("controlVInConsole") and obj.windowClassName == 'PuTTY':
            clsList.insert(0, PuttyControlV)


    def createMenu(self):
        gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(SettingsDialog)


    def terminate(self):
        self.removeHooks()
        gui.settingsDialogs.NVDASettingsDialog.categoryClasses.remove(SettingsDialog)

    def injectHooks(self):
        global originalReportNewText, originalCancelSpeech, originalTerminalGainFocus, originalNVDAObjectFfocusEntered
        originalReportNewText = behaviors.LiveText._reportNewText
        behaviors.LiveText._reportNewText = newReportConsoleText
        originalCancelSpeech = speech.cancelSpeech
        speech.cancelSpeech = newCancelSpeech
        # Apparently we need to monkey patch in two places to avoid terminal title being spoken when we switch to it from edit prompt window.
        # behaviors.Terminal.event_gainFocus is needed for both legacy and UIA implementation,
        # but in legacy it speaks window title, while in UIA mode it speaks current line in the terminal
        # NVDAObject.event_focusEntered speaks window title in UIA mode.
        originalTerminalGainFocus = behaviors.Terminal.event_gainFocus
        behaviors.Terminal.event_gainFocus = terminalGainFocus
        originalNVDAObjectFfocusEntered = NVDAObject.event_focusEntered
        NVDAObject.event_focusEntered = nvdaObjectFfocusEntered
        behaviors.Terminal.script_editPrompt = script_editPrompt
        behaviors.Terminal.script_captureOutput = script_captureOutput
        try:
            behaviors.Terminal._Terminal__gestures
        except AttributeError:
            behaviors.Terminal._Terminal__gestures = {}
        behaviors.Terminal._Terminal__gestures["kb:NVDA+E"] = "editPrompt"
        behaviors.Terminal._Terminal__gestures["kb:Control+Enter"] = "captureOutput"

    def  removeHooks(self):
        behaviors.LiveText._reportNewText = originalReportNewText
        speech.cancelSpeech = originalCancelSpeech
        behaviors.Terminal.event_gainFocus = originalTerminalGainFocus
        NVDAObject.event_focusEntered = originalNVDAObjectFfocusEntered
        del behaviors.Terminal.script_editPrompt
        del behaviors.Terminal.script_captureOutput
        del behaviors.Terminal._Terminal__gestures["kb:NVDA+E"]
        del behaviors.Terminal._Terminal__gestures["kb:Control+Enter"]

    def preCalculateNewText(self, selfself, *args, **kwargs):
        oldLines = args[1]
        newLines = args[0]
        if oldLines == newLines:
            return []
        outLines =   self.originalCalculateNewText(selfself, *args, **kwargs)
        return outLines
        if len(outLines) == 1 and len(outLines[0].strip()) == 1:
            # Only a single character has changed - in this case NVDA thinks that's a typed character, so it is not spoken anyway. Con't interfere.
            return outLines
        if len(outLines) == 0:
            return outLines
        if getConfig("consoleBeep"):
            tones.beep(100, 5)
        if getConfig("consoleRealtime"):
            #if time.time() > self.lastConsoleUpdateTime + 0.5:
                #self.lastConsoleUpdateTime = time.time()
            speech.cancelSpeech()
        return outLines


class ConsoleControlV(NVDAObject):
    @script(description='Paste from clipboard', gestures=['kb:Control+V'])
    def script_paste(self, gesture):
        pasteConsole(self)


    #@script(description='Edit prompt', gestures=['kb:NVDA+E'])


class PuttyControlV(NVDAObject):
    @script(description='Paste from clipboard', gestures=['kb:Control+V'])
    def script_paste(self, gesture):
        pastePutty(self)
