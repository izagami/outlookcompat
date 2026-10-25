"""pythoncom stand-in (CoInitialize is a no-op here)"""
def CoInitialize(*a): pass
def CoUninitialize(*a): pass
