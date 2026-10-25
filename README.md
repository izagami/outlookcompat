# outlookcompat

A stand-in for classic Outlook's COM object model on Linux, so code written for Windows

```python
import win32com.client
ol = win32com.client.Dispatch("Outlook.Application")
inbox = ol.GetNamespace("MAPI").GetDefaultFolder(6)
for mail in inbox.Items.Restrict("[UnRead] = True"):
    print(mail.Subject)
```

runs unchanged against a local JSON mailbox. Built to develop and test AI agents that automate Outlook (triage,
replies, meeting requests) without a Windows machine or a real mailbox, and to score what they did.

## Use

```sh
export PYTHONPATH=/path/to/outlookcompat        # provides win32com.client, pywintypes, pythoncom
export OUTLOOKCOMPAT_STORE=/tmp/mailbox.json      # created on first use; every change is written back at once
python3 your_outlook_script.py
python3 test_outlookcompat.py                     # behaviour checks
```

The mailbox file is plain JSON (`user`, `folders`, `items`), so a test can prepare it and a checker can read what the
code did afterwards: sent mail (`sent_log`), moved/deleted items, flags, saved attachments.

## What is emulated

- `Dispatch("Outlook.Application")`, `GetNamespace("MAPI")`, `GetDefaultFolder` (Inbox, Sent Items, Drafts, Deleted
  Items, Outbox, Calendar, Contacts), sub-folders (`Folders.Add`), `GetItemFromID`, `CurrentUser`
- `Items`: `Count`, `Item(i)`, iteration, `Sort`, `Restrict` (Jet syntax with AND/OR and dates, `@SQL=` DASL
  `LIKE`), `Find`/`FindNext`, `GetFirst`/`GetNext`/`GetLast`
- `MailItem`: subject/body/recipients/sender/dates/unread/importance/categories/flags, `Reply`, `ReplyAll`,
  `Forward`, `Send`, `Save`, `Move`, `Delete`, `Attachments` (`SaveAsFile`, `Add`), `Recipients.Add`
- `AppointmentItem`: subject/start/end/location/body, `Duration`, `Save`, `Send`

## Behaviour copied from the real thing (the parts that trip people up)

- collections are 1-based; `Item(0)` raises `com_error`
- `Items` is live: moving or deleting while iterating forward skips every other item (iterate backwards)
- `Restrict` dates use `MM/dd/yyyy hh:mm AM/PM`; malformed filters raise `com_error`
- `Reply`/`Forward` return a new unsent item ("RE: " / "FW: ", quoted original)
- `Send` needs a recipient, moves the item to Sent Items, and the sent item object can no longer be used
- `Move` returns the moved item; the old reference becomes invalid
- folder names follow the Japanese UI by default (`OUTLOOKCOMPAT_LANG=ja`, e.g. 受信トレイ); IMAP-backed calendar and
  contacts folders carry the "(このコンピューターのみ)" suffix like classic Outlook with an IMAP account

An 18-step comparison script gives identical output on real classic Outlook (Windows 11, IMAP account) and on this
stand-in. Not emulated: Exchange-only features, rules, search folders, the Outlook UI and events.

## License

Apache License 2.0. See [LICENSE](LICENSE).
