"""outlookcompat: a stand-in for classic Outlook's COM object model on Linux (2026-10-25).

Code written for Windows -- `import win32com.client; ol = win32com.client.Dispatch("Outlook.Application")` --
runs unchanged when this directory's parent is on sys.path (PYTHONPATH=~/llm/outlookcompat). The mailbox is a
JSON file (env OUTLOOKCOMPAT_STORE, default ./mailbox.json); every change is written back at once, so a checker
can read what the code did (sent mail, moved items, flags, saved attachments).

Behaviour copied from the real object model, including the parts that trip people up:
  * collections are 1-based: Items.Item(1) / Items[1] is the first; Item(0) raises com_error
  * Items is LIVE: moving or deleting while iterating forward skips every other item (iterate backwards,
    or collect the items first)
  * Items.Restrict uses Jet syntax: "[UnRead] = True", "[SenderEmailAddress] = 'a@b.c'",
    "[ReceivedTime] >= '10/01/2026 08:00 AM'" (dates as MM/dd/yyyy hh:mm AM/PM), joined with AND / OR;
    "@SQL=" DASL filters support "urn:schemas:httpmail:subject" LIKE '%x%'. Bad filters raise com_error
  * Reply() / ReplyAll() / Forward() return a NEW unsent MailItem ("RE: " / "FW: " subject, quoted body)
  * Send() needs at least one recipient, moves the item to Sent Items, and the item can no longer be used
  * Move(folder) returns the moved item (the old reference is invalid afterwards)
  * ReceivedTime / Start / End are datetimes with tzinfo (like pywintypes.TimeType)
Constants: OlDefaultFolders (olFolderInbox = 6, ...), OlItemType (olMailItem = 0, olAppointmentItem = 1).
"""
import datetime as _dt
import json as _json
import os as _os
import re as _re
import uuid as _uuid

from . import constants  # noqa: F401  (win32com.client.constants)


class com_error(Exception):
    """like pywintypes.com_error: args = (hresult, text, (wcode, source, description, ...), argerr)"""

    def __init__(self, description, hresult=-2147352567):
        super().__init__(hresult, "Exception occurred.", (4096, "Microsoft Outlook", description, None, 0, -2147024809), None)
        self.description = description


JST = _dt.timezone(_dt.timedelta(hours=9))
FOLDER_IDS = {3: "Deleted Items", 4: "Outbox", 5: "Sent Items", 6: "Inbox", 9: "Calendar", 10: "Contacts", 16: "Drafts"}
JA_NAMES = {"Deleted Items": "削除済みアイテム", "Outbox": "送信トレイ", "Sent Items": "送信済みアイテム", "Inbox": "受信トレイ",
            "Calendar": "予定表", "Contacts": "連絡先", "Drafts": "下書き"}


# ---------------------------------------------------------------- store
class _Store:
    def __init__(self, path):
        self.path = path
        if _os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                self.data = _json.load(f)
        else:
            self.data = {"user": {"name": "テスト 太郎", "address": "taro@example.com"}, "folders": {}, "items": []}
        for name in FOLDER_IDS.values():
            self.data["folders"].setdefault(name, {"parent": None})
        self.save()

    def save(self):
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            _json.dump(self.data, f, ensure_ascii=False, indent=1, default=str)
        _os.replace(tmp, self.path)

    def items_in(self, folder):
        return [d for d in self.data["items"] if d["folder"] == folder and not d.get("_gone")]


_STORE = None


def _store():
    global _STORE
    if _STORE is None:
        _STORE = _Store(_os.environ.get("OUTLOOKCOMPAT_STORE", "mailbox.json"))
    return _STORE


def _parse_time(v):
    if isinstance(v, _dt.datetime):
        return v if v.tzinfo else v.replace(tzinfo=JST)
    return _dt.datetime.fromisoformat(str(v)).astimezone(JST) if "+" in str(v) else _dt.datetime.fromisoformat(str(v)).replace(tzinfo=JST)


# ---------------------------------------------------------------- items
class _Attachment:
    def __init__(self, item, d):
        self._item, self._d = item, d
        self.FileName = self.DisplayName = d["name"]
        self.Size = len(d.get("content", ""))

    def SaveAsFile(self, path):
        if _os.path.isdir(path):
            raise com_error("Cannot save the attachment. Path is a folder, not a file name.")
        with open(path, "w", encoding="utf-8") as f:
            f.write(self._d.get("content", ""))
        _store().data.setdefault("saved_attachments", []).append({"item": self._item._d["id"], "name": self._d["name"], "path": path})
        _store().save()


class _Collection:
    def __init__(self, objs):
        self._objs = objs

    @property
    def Count(self):
        return len(self._objs)

    def Item(self, i):
        if isinstance(i, str):
            for o in self._objs:
                if getattr(o, "Name", None) == i or getattr(o, "FileName", None) == i:
                    return o
            raise com_error(f"The attempted operation failed. An object could not be found. ({i!r})")
        if not 1 <= i <= len(self._objs):
            raise com_error("Array index out of bounds.")
        return self._objs[i - 1]

    __call__ = Item

    def __getitem__(self, i):
        return self.Item(i)

    def __iter__(self):
        return iter(list(self._objs))

    def __len__(self):
        return len(self._objs)


class _Attachments(_Collection):
    def __init__(self, item):
        self._item = item
        super().__init__([_Attachment(item, a) for a in item._d.setdefault("attachments", [])])

    def Add(self, source, *a):
        with open(source, encoding="utf-8", errors="replace") as f:
            d = {"name": _os.path.basename(source), "content": f.read()}
        self._item._d["attachments"].append(d)
        self._objs.append(_Attachment(self._item, d))
        _store().save()
        return self._objs[-1]


class _Recipients(_Collection):
    def __init__(self, item):
        self._item = item
        super().__init__([])

    def Add(self, address):
        cur = self._item._d.get("to", "")
        self._item._d["to"] = (cur + "; " if cur else "") + address
        return type("Recipient", (), {"Address": address, "Name": address, "Resolve": lambda s: True, "Type": 1})()

    def ResolveAll(self):
        return True


class _Item:
    _props = {}

    def __init__(self, d):
        object.__setattr__(self, "_d", d)

    def _check(self):
        if self._d.get("_gone"):
            raise com_error("The item has been moved or deleted.")

    def __getattr__(self, name):
        if name.startswith("_"):
            raise AttributeError(name)
        self._check()
        if name in self._props:
            key, kind = self._props[name]
            v = self._d.get(key, "" if kind == "s" else (False if kind == "b" else (1 if kind == "i" else None)))
            return _parse_time(v) if kind == "t" and v else v
        raise AttributeError(f"<unknown>.{name}")

    def __setattr__(self, name, value):
        self._check()
        if name not in self._props:
            raise AttributeError(f"Property '<unknown>.{name}' can not be set.")
        key, kind = self._props[name]
        self._d[key] = value.isoformat() if isinstance(value, _dt.datetime) else value
        _store().save()

    @property
    def EntryID(self):
        return self._d["id"]

    @property
    def Parent(self):
        return Folder(self._d["folder"])

    def Save(self):
        self._check()
        if self._d["folder"] == "_unsaved":
            self._d["folder"] = "Drafts" if self.Class == 43 else "Calendar"
        _store().save()

    def Delete(self):
        self._check()
        if self._d["folder"] == "Deleted Items":
            self._d["_gone"] = True
        else:
            self._d["folder"] = "Deleted Items"
            new = dict(self._d)
            self._d["_gone"] = True
            new["id"] = _uuid.uuid4().hex.upper()
            _store().data["items"].append(new)
        _store().save()

    def Move(self, folder):
        self._check()
        if not isinstance(folder, Folder):
            raise com_error("Could not complete the operation. One or more parameter values are not valid.")
        new = dict(self._d, folder=folder._name, id=_uuid.uuid4().hex.upper())
        self._d["_gone"] = True
        _store().data["items"].append(new)
        _store().save()
        return type(self)(new)

    def Display(self, *a):
        pass

    def Close(self, *a):
        pass


class MailItem(_Item):
    Class = 43
    MessageClass = "IPM.Note"
    _props = {"Subject": ("subject", "s"), "Body": ("body", "s"), "HTMLBody": ("html", "s"), "To": ("to", "s"),
              "CC": ("cc", "s"), "BCC": ("bcc", "s"), "SenderName": ("sender_name", "s"),
              "SenderEmailAddress": ("sender", "s"), "ReceivedTime": ("received", "t"), "SentOn": ("sent", "t"),
              "UnRead": ("unread", "b"), "Importance": ("importance", "i"), "Categories": ("categories", "s"),
              "FlagStatus": ("flag", "i"), "FlagRequest": ("flag_request", "s"), "ConversationTopic": ("subject", "s")}

    @property
    def Attachments(self):
        self._check()
        return _Attachments(self)

    @property
    def Recipients(self):
        self._check()
        return _Recipients(self)

    @property
    def Size(self):
        return len(self._d.get("body", "")) + 500

    def _reply(self, prefix, to, cc=""):
        self._check()
        sub = self._d.get("subject", "")
        if not _re.match(r"(?i)(re|fw|fwd):\s", sub) or not sub.upper().startswith(prefix.upper()):
            sub = f"{prefix} {sub}"
        quoted = ("\n\n-----Original Message-----\n"
                  f"From: {self._d.get('sender_name', '')} <{self._d.get('sender', '')}>\n"
                  f"Sent: {_parse_time(self._d['received']).strftime('%Y/%m/%d %H:%M') if self._d.get('received') else ''}\n"
                  f"To: {self._d.get('to', '')}\nSubject: {self._d.get('subject', '')}\n\n{self._d.get('body', '')}")
        d = {"id": _uuid.uuid4().hex.upper(), "folder": "_unsaved", "class": "mail", "subject": sub, "body": quoted,
             "to": to, "cc": cc, "unread": False, "in_reply_to": self._d["id"],
             "attachments": list(self._d.get("attachments", [])) if prefix == "FW:" else []}
        _store().data["items"].append(d)
        return MailItem(d)

    def Reply(self):
        return self._reply("RE:", self._d.get("sender", ""))

    def ReplyAll(self):
        me = _store().data["user"]["address"]
        others = [a.strip() for a in (self._d.get("to", "") + ";" + self._d.get("cc", "")).split(";") if a.strip() and a.strip() != me]
        return self._reply("RE:", self._d.get("sender", ""), "; ".join(others))

    def Forward(self):
        return self._reply("FW:", "")

    def Send(self):
        self._check()
        if not any(str(self._d.get(k) or "").strip() for k in ("to", "cc", "bcc")):
            raise com_error("Outlook does not recognize one or more names.")
        u = _store().data["user"]
        new = dict(self._d, folder="Sent Items", sent=_dt.datetime.now(JST).isoformat(), sender=u["address"],
                   sender_name=u["name"], unread=False)
        self._d["_gone"] = True        # the sent item object can't be used anymore (like the real one)
        _store().data["items"].append(new)
        _store().data.setdefault("sent_log", []).append(new["id"])
        _store().save()


class AppointmentItem(_Item):
    Class = 26
    MessageClass = "IPM.Appointment"
    _props = {"Subject": ("subject", "s"), "Body": ("body", "s"), "Location": ("location", "s"), "Start": ("start", "t"),
              "End": ("end", "t"), "AllDayEvent": ("all_day", "b"), "Organizer": ("organizer", "s"),
              "RequiredAttendees": ("to", "s"), "MeetingStatus": ("meeting_status", "i"), "Categories": ("categories", "s"),
              "ReminderMinutesBeforeStart": ("reminder", "i"), "BusyStatus": ("busy", "i")}

    @property
    def Duration(self):
        return int((self.End - self.Start).total_seconds() // 60)

    @property
    def Recipients(self):
        return _Recipients(self)

    def Send(self):
        self._check()
        self._d["folder"] = "Calendar"
        self._d["sent"] = _dt.datetime.now(JST).isoformat()
        _store().data.setdefault("sent_log", []).append(self._d["id"])
        _store().save()


def _wrap(d):
    return AppointmentItem(d) if d.get("class") == "appointment" else MailItem(d)


# ---------------------------------------------------------------- Restrict
_DASL = {"urn:schemas:httpmail:subject": "subject", "urn:schemas:httpmail:textdescription": "body",
         "urn:schemas:httpmail:fromemail": "sender", "urn:schemas:httpmail:read": "unread"}
_FIELDS = {"subject": "subject", "body": "body", "senderemailaddress": "sender", "sendername": "sender_name",
           "receivedtime": "received", "senton": "sent", "unread": "unread", "importance": "importance",
           "categories": "categories", "flagstatus": "flag", "start": "start", "end": "end", "location": "location",
           "to": "to", "cc": "cc"}


def _jet_time(s):
    for fmt in ("%m/%d/%Y %I:%M %p", "%m/%d/%Y %H:%M", "%m/%d/%Y", "%Y/%m/%d %H:%M", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return _dt.datetime.strptime(s.strip(), fmt).replace(tzinfo=JST)
        except ValueError:
            pass
    raise com_error(f"Cannot parse condition. Error at \"{s}\".")


def _cond(d, field, op, val):
    key = _FIELDS.get(field.lower())
    if key is None:
        raise com_error(f"The property \"{field}\" in the condition is not valid.")
    v = d.get(key)
    if key in ("received", "sent", "start", "end"):
        if not v:
            return False
        a, b = _parse_time(v), _jet_time(val)
    elif key == "unread":
        a, b = bool(v), val.lower() in ("true", "-1", "1")
    elif key in ("importance", "flag"):
        a, b = int(v or (1 if key == "importance" else 0)), int(val)
    else:
        a, b = str(v or "").lower(), val.lower()
    return {"=": a == b, "<>": a != b, ">": a > b, ">=": a >= b, "<": a < b, "<=": a <= b}[op]


def _restrict(dicts, flt):
    f = flt.strip()
    if f.lower().startswith("@sql="):
        m = _re.fullmatch(r'"?([^"\s]+)"?\s+(like|=)\s+\'(.*)\'', f[5:].strip(), _re.I)
        if not m or m.group(1) not in _DASL:
            raise com_error(f"Cannot parse condition. Error at \"{f[5:]}\".")
        key, op, pat = _DASL[m.group(1)], m.group(2).lower(), m.group(3)
        if op == "like":
            rx = "^" + _re.escape(pat).replace("%", ".*") + "$"
            return [d for d in dicts if _re.match(rx, str(d.get(key, "")), _re.I | _re.S)]
        return [d for d in dicts if str(d.get(key, "")).lower() == pat.lower()]
    tokens = _re.split(r"\s+(AND|OR)\s+", f, flags=_re.I)
    terms = tokens[0::2]
    ops = [t.upper() for t in tokens[1::2]]
    parsed = []
    for t in terms:
        m = _re.fullmatch(r"\(?\s*\[([^\]]+)\]\s*(<>|>=|<=|=|>|<)\s*(?:'([^']*)'|\"([^\"]*)\"|(\S+?))\s*\)?", t.strip())
        if not m:
            raise com_error(f"Cannot parse condition. Error at \"{t.strip()}\".")
        parsed.append((m.group(1), m.group(2), next(g for g in m.groups()[2:] if g is not None)))
    out = []
    for d in dicts:
        res = _cond(d, *parsed[0])
        for op, p in zip(ops, parsed[1:]):
            res = (res and _cond(d, *p)) if op == "AND" else (res or _cond(d, *p))
        if res:
            out.append(d)
    return out


class Items(_Collection):
    """live view of a folder's items (a Restrict()ed one is a snapshot of the matching items, also live for removal)"""

    def __init__(self, folder, dicts=None):
        self._folder = folder
        self._fixed = dicts
        self._sort = None
        self._pos = 0

    def _list(self):
        ds = [d for d in (self._fixed if self._fixed is not None else _store().items_in(self._folder)) if not d.get("_gone")
              and (self._fixed is None or d["folder"] == self._folder)]
        if self._sort:
            key, desc = self._sort
            ds = sorted(ds, key=lambda d: str(d.get(key) or ""), reverse=desc)
        return ds

    @property
    def _objs(self):
        return [_wrap(d) for d in self._list()]

    @property
    def Count(self):
        return len(self._list())

    def __iter__(self):
        # LIVE like COM: an index walks the current list, so removing the current item skips the next one
        i = 1
        while i <= self.Count:
            yield self.Item(i)
            i += 1

    def Sort(self, prop, descending=False):
        key = _FIELDS.get(prop.strip("[]").lower())
        if key is None:
            raise com_error(f"The property \"{prop}\" is unknown.")
        self._sort = (key, bool(descending))

    def Restrict(self, flt):
        r = Items(self._folder, _restrict(self._list(), flt))
        r._sort = self._sort
        return r

    def Find(self, flt):
        self._found = _restrict(self._list(), flt)
        self._fi = 0
        return self.FindNext()

    def FindNext(self):
        if getattr(self, "_fi", 0) < len(getattr(self, "_found", [])):
            self._fi += 1
            return _wrap(self._found[self._fi - 1])
        return None

    def GetFirst(self):
        self._pos = 0
        return self.GetNext()

    def GetNext(self):
        lst = self._list()
        if self._pos < len(lst):
            self._pos += 1
            return _wrap(lst[self._pos - 1])
        return None

    def GetLast(self):
        lst = self._list()
        self._pos = len(lst)
        return _wrap(lst[-1]) if lst else None

    def Add(self, kind=0):
        return Application().CreateItem(kind)


class Folders(_Collection):
    def __init__(self, parent):
        self._parent = parent

    @property
    def _objs(self):
        return [Folder(n) for n, f in _store().data["folders"].items() if f.get("parent") == self._parent]

    def Add(self, name, *a):
        if name in _store().data["folders"]:
            raise com_error("Cannot create the folder. A folder with that name already exists.")
        _store().data["folders"][name] = {"parent": self._parent}
        _store().save()
        return Folder(name)


class Folder:
    def __init__(self, name):
        if name not in _store().data["folders"]:
            raise com_error(f"The attempted operation failed. An object could not be found. ({name!r})")
        self._name = name

    @property
    def Name(self):
        return JA_NAMES.get(self._name, self._name) if _os.environ.get("OUTLOOKCOMPAT_LANG", "ja") == "ja" else self._name

    @property
    def Items(self):
        return Items(self._name)

    @property
    def Folders(self):
        return Folders(self._name)

    @property
    def UnReadItemCount(self):
        return sum(1 for d in _store().items_in(self._name) if d.get("unread"))

    @property
    def FolderPath(self):
        chain, n = [], self._name
        while n:
            chain.append(JA_NAMES.get(n, n))
            n = _store().data["folders"][n].get("parent")
        return "\\\\" + _store().data["user"]["address"] + "\\" + "\\".join(reversed(chain))

    @property
    def Parent(self):
        p = _store().data["folders"][self._name].get("parent")
        return Folder(p) if p else None

    def Delete(self):
        for d in _store().items_in(self._name):
            d["folder"] = "Deleted Items"
        del _store().data["folders"][self._name]
        _store().save()

    def __eq__(self, o):
        return isinstance(o, Folder) and o._name == self._name


class _Store_Root:
    def __init__(self):
        self.Name = self.DisplayName = _store().data["user"]["address"]

    @property
    def Folders(self):
        return Folders(None)


class Namespace:
    @property
    def CurrentUser(self):
        u = _store().data["user"]
        return type("Recipient", (), {"Name": u["name"], "Address": u["address"]})()

    def GetDefaultFolder(self, n):
        if n not in FOLDER_IDS:
            raise com_error("Could not complete the operation. One or more parameter values are not valid.")
        return Folder(FOLDER_IDS[n])

    @property
    def Folders(self):
        return _Collection([_Store_Root()])

    def GetItemFromID(self, entry_id, *a):
        for d in _store().data["items"]:
            if d["id"] == entry_id and not d.get("_gone"):
                return _wrap(d)
        raise com_error("The operation failed. An object could not be found.")

    def Logon(self, *a):
        pass


class Application:
    Name = "Outlook"
    Version = "16.0.0.0"

    def GetNamespace(self, kind):
        if str(kind).upper() != "MAPI":
            raise com_error("Could not complete the operation. One or more parameter values are not valid.")
        return Namespace()

    @property
    def Session(self):
        return Namespace()

    def CreateItem(self, kind):
        d = {"id": _uuid.uuid4().hex.upper(), "folder": "_unsaved", "class": "appointment" if kind == 1 else "mail",
             "unread": False}
        if kind not in (0, 1):
            raise com_error("CreateItem: only olMailItem (0) and olAppointmentItem (1) are emulated.")
        _store().data["items"].append(d)
        return _wrap(d)

    def Quit(self):
        _store().save()


def Dispatch(progid, *a, **k):
    if str(progid).lower() != "outlook.application":
        raise com_error(f"Invalid class string ({progid})", hresult=-2147221005)
    return Application()


DispatchEx = Dispatch
GetActiveObject = Dispatch


class gencache:
    @staticmethod
    def EnsureDispatch(progid):
        return Dispatch(progid)
