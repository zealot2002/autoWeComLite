import wx

from ui.main_frame import MainFrame

def main():
    app = wx.App(False)
    frame = MainFrame(None, title="autoWeComLite")
    app.MainLoop()

if __name__ == "__main__":
    main() 