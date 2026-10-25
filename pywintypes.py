"""pywintypes stand-in: com_error and TimeType"""
import datetime
from win32com.client import com_error  # noqa: F401
TimeType = datetime.datetime
