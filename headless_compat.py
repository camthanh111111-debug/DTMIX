"""Headless replacements for tkinter dialogs used only by the legacy core."""
last_error = None
last_warning = None
last_info = None

def reset_messages():
    global last_error, last_warning, last_info
    last_error = last_warning = last_info = None

class _MessageBox:
    @staticmethod
    def showerror(title, message):
        global last_error
        last_error = f"{title}: {message}"
        return None
    @staticmethod
    def showwarning(title, message):
        global last_warning
        last_warning = f"{title}: {message}"
        return None
    @staticmethod
    def showinfo(title, message):
        global last_info
        last_info = f"{title}: {message}"
        return None

class _FileDialog:
    @staticmethod
    def askopenfilename(*args, **kwargs): return ""
    @staticmethod
    def askdirectory(*args, **kwargs): return ""

messagebox = _MessageBox()
filedialog = _FileDialog()
