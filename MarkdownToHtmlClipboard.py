import os
import sublime
import sublime_plugin
import subprocess

def err(theError):
    sublime.error_message("[Md To Clip: " + theError + "]")

class Md2rtfclipCommand(sublime_plugin.TextCommand):
  def run(self, edit):
    # self.view.insert(edit, 0, "Hello, World!")

    if self.convertMarkdownPreview():
      self.readClipboard()
    return

  def readClipboard(self):
    self.env = os.environ.copy()
    plat = sublime.platform()

    if plat == "osx":
      cmd = ['pbpaste', '|', 'textutil', '-stdin', '-format', 'html', '-convert', 'rtf', '-stdout', '|', 'pbcopy']

    if plat == "windows":
      if self.checkPowershellVersion(self.env):
        cmd = ['powershell', 'Get-Clipboard', ' -Format', 'Text', '|', 'Set-Clipboard', '-AsHtml']
      else:
        err("Please install Powershell 5.1")
        return

    if plat == "linux":
      cmd = ['xclip', '-selection', 'c', '-o', '|', 'xclip', '-t', 'text/html']

    try:
        subprocess.call(cmd, env=self.env)
    except Exception as e:
        err("Exception: " + str(e))
    return 0

  def convertMarkdownPreview(self):
    self.view.run_command("markdown_preview", {"target": "clipboard", "parser": "markdown"})
    return 1

  def checkPowershellVersion(self, env):
      cmd = ['powershell', '$PSVersionTable.PSVersion.Major']
      try:
          output = subprocess.check_output(cmd, env=env)
      except Exception as e:
          err("Exception: " + str(e))

      return output.startswith("5".encode("utf-8"))
