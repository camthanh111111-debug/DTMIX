"""Headless compatibility shim for the legacy DTMIX core.
The web engine never builds the CustomTkinter desktop UI.
"""
class CTk:
    pass
class BooleanVar:
    def __init__(self,value=False): self._v=value
    def get(self): return self._v
    def set(self,v): self._v=v

def set_appearance_mode(mode): pass
