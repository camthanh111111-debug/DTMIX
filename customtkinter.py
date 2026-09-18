"""Minimal headless compatibility shim for importing the legacy DTMIX core.
The Streamlit UI does not use CustomTkinter; only legacy algorithm methods are reused.
"""

def set_appearance_mode(*args, **kwargs):
    return None

class Variable:
    def __init__(self, value=None, *args, **kwargs): self._value = value
    def get(self): return self._value
    def set(self, value): self._value = value

class BooleanVar(Variable): pass
class StringVar(Variable): pass

class _Widget:
    def __init__(self, *args, **kwargs):
        self._value = kwargs.get('text', '')
        self._exists = True
    def get(self): return self._value
    def set(self, value): self._value = value
    def insert(self, index, value): self._value = str(value)
    def delete(self, *args, **kwargs): self._value = ''
    def configure(self, *args, **kwargs):
        if 'text' in kwargs: self._value = kwargs['text']
        return self
    config = configure
    def pack(self, *args, **kwargs): return self
    def pack_forget(self, *args, **kwargs): return self
    def grid(self, *args, **kwargs): return self
    def grid_forget(self, *args, **kwargs): return self
    def place(self, *args, **kwargs): return self
    def bind(self, *args, **kwargs): return self
    def destroy(self): self._exists = False
    def winfo_exists(self): return self._exists
    def winfo_children(self): return []
    def tag_config(self, *args, **kwargs): return None
    def tag_configure(self, *args, **kwargs): return None
    def select(self): self._value = True
    def deselect(self): self._value = False
    def icursor(self, *args, **kwargs): return None
    def index(self, *args, **kwargs): return 0
    def attributes(self, *args, **kwargs): return None
    def transient(self, *args, **kwargs): return None
    def grab_set(self, *args, **kwargs): return None
    def geometry(self, *args, **kwargs): return None
    def title(self, *args, **kwargs): return None

class CTk(_Widget):
    def after(self, delay, callback=None, *args):
        if callback: return callback(*args)
    def update(self): return None
    def update_idletasks(self): return None
    def mainloop(self): return None
    def state(self, *args, **kwargs): return None
    def grid_rowconfigure(self, *args, **kwargs): return None
    def grid_columnconfigure(self, *args, **kwargs): return None
    def configure(self, *args, **kwargs): return self
    def bind_all(self, *args, **kwargs): return None

CTkFrame = CTkLabel = CTkButton = CTkEntry = CTkTextbox = CTkScrollableFrame = CTkSwitch = CTkCheckBox = CTkComboBox = CTkProgressBar = CTkToplevel = CTk
class CTkImage:
    def __init__(self, *args, **kwargs): pass
