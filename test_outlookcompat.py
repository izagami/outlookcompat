#!/usr/bin/env python3
"""behaviour checks of the Outlook stand-in (run: python3 test_outlookcompat.py)"""
import json
import os
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
store = os.path.join(tempfile.mkdtemp(), "mailbox.json")
json.dump({"user": {"name": "テスト 太郎", "address": "taro@example.com"}, "folders": {}, "items": [
    {"id": "A1", "folder": "Inbox", "class": "mail", "subject": "会議の件", "body": "明日10時からです。", "sender": "sato@example.com",
     "sender_name": "佐藤 一郎", "to": "taro@example.com", "cc": "kato@example.com", "received": "2026-10-20T09:00:00+09:00", "unread": True},
    {"id": "A2", "folder": "Inbox", "class": "mail", "subject": "請求書 10月分", "body": "添付をご確認ください。", "sender": "keiri@example.com",
     "sender_name": "経理部", "to": "taro@example.com", "received": "2026-10-21T14:30:00+09:00", "unread": True,
     "attachments": [{"name": "invoice.csv", "content": "item,amount\nA,100\n"}]},
    {"id": "A3", "folder": "Inbox", "class": "mail", "subject": "【至急】サーバー停止", "body": "対応お願いします", "sender": "ops@example.com",
     "sender_name": "運用", "to": "taro@example.com", "received": "2026-10-22T08:15:00+09:00", "unread": False, "importance": 2},
    {"id": "A4", "folder": "Inbox", "class": "mail", "subject": "請求書 9月分", "body": "先月分です", "sender": "keiri@example.com",
     "sender_name": "経理部", "to": "taro@example.com", "received": "2026-09-21T10:00:00+09:00", "unread": False},
]}, open(store, "w"), ensure_ascii=False)
os.environ["OUTLOOKCOMPAT_STORE"] = store
import win32com.client as wc  # noqa: E402

ok = 0


def check(cond, what):
    global ok
    if not cond:
        raise SystemExit("FAIL: " + what)
    ok += 1


ol = wc.Dispatch("Outlook.Application")
ns = ol.GetNamespace("MAPI")
inbox = ns.GetDefaultFolder(6)
check(inbox.Name == "受信トレイ", "inbox name")
items = inbox.Items
check(items.Count == 4, "count")
try:
    items.Item(0)
    check(False, "Item(0) must fail")
except wc.com_error:
    ok += 1
check(items.Item(1).Subject == "会議の件", "1-based")
unread = items.Restrict("[UnRead] = True")
check(unread.Count == 2, "restrict unread")
check(items.Restrict("[SenderEmailAddress] = 'keiri@example.com' AND [ReceivedTime] >= '10/01/2026 12:00 AM'").Count == 1, "restrict and/date")
check(items.Restrict("@SQL=\"urn:schemas:httpmail:subject\" like '%請求書%'").Count == 2, "dasl like")
try:
    items.Restrict("UnRead = True")
    check(False, "bad filter must fail")
except wc.com_error:
    ok += 1
items.Sort("[ReceivedTime]", True)
check(items.Item(1).Subject == "【至急】サーバー停止", "sort desc")

# live collection: moving while iterating forward skips items
arch = inbox.Folders.Add("請求書")
moved = 0
for it in inbox.Items.Restrict("[SenderEmailAddress] = 'keiri@example.com'"):
    it.Move(arch)
    moved += 1
check(moved == 1 and arch.Items.Count == 1, f"forward-iteration skip reproduced (moved {moved})")
rest = inbox.Items.Restrict("[SenderEmailAddress] = 'keiri@example.com'")
for i in range(rest.Count, 0, -1):
    rest.Item(i).Move(arch)
check(arch.Items.Count == 2 and inbox.Items.Count == 2, "backward iteration moves all")

# reply / send
m = ns.GetItemFromID("A1")
r = m.Reply()
check(r.Subject == "RE: 会議の件" and r.To == "sato@example.com" and "明日10時" in r.Body, "reply draft")
r.Body = "承知しました。\n" + r.Body
r.Send()
try:
    r.Subject
    check(False, "sent item must be unusable")
except wc.com_error:
    ok += 1
sent = ns.GetDefaultFolder(5).Items
check(sent.Count == 1 and sent.Item(1).Body.startswith("承知しました"), "sent folder")
ra = m.ReplyAll()
check("kato@example.com" in ra.CC, "reply all keeps cc")
blank = ol.CreateItem(0)
blank.Subject = "x"
try:
    blank.Send()
    check(False, "send without recipient must fail")
except wc.com_error:
    ok += 1

# attachments
inv = arch.Items.Restrict("[Subject] = '請求書 10月分'").Item(1)
p = os.path.join(tempfile.mkdtemp(), "invoice.csv")
inv.Attachments.Item(1).SaveAsFile(p)
check(open(p).read().startswith("item,amount"), "save attachment")

# appointment
ap = ol.CreateItem(1)
ap.Subject = "定例"
import datetime as dt  # noqa: E402
ap.Start = dt.datetime(2026, 10, 27, 10, 0)
ap.End = dt.datetime(2026, 10, 27, 11, 0)
ap.Save()
check(ns.GetDefaultFolder(9).Items.Count == 1 and ap.Duration == 60, "appointment")
check(json.load(open(store))["sent_log"], "store written")
print(f"ok: {ok} checks")
