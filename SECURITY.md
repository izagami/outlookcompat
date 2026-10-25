# Security policy

Please report vulnerabilities privately through GitHub's "Report a vulnerability" (Security tab) rather than in a
public issue. This project only reads and writes a local JSON file given by `OUTLOOKCOMPAT_STORE` and opens files
passed to `Attachments.Add` / `SaveAsFile`; treat mailbox files and attachment paths from untrusted sources as untrusted.
