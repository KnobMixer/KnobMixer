# KnobMixer Privacy Policy

*Applies to KnobMixer 3.0.1 and later. Last updated: 2026.*

KnobMixer is built to know as little about you as possible. This page lists
**everything** the app ever sends, word for word from how it is built.

## The daily ping (optional — the "Analytics" switch in the app)

Once per day, KnobMixer sends one small message so we can see that installed
copies are alive and working. It contains exactly six things:

1. a **random install ID** — generated on your PC, not derived from your
   hardware, name, or any account; we cannot connect it to you
2. the **app version** (e.g. 3.0.0)
3. the **internal build number**
4. the **install channel** (website or store)
5. your **Windows version number** (e.g. 10.0.26100 — just the number)
6. a **yes/no flag**: did you actually use KnobMixer since the last ping

That's the whole list. Never your app names, hotkeys, settings, volume
levels, files, or anything typed.

- After a successful update, one extra message says "the update worked"
  (same fields, no usage flag).
- If you turn Analytics **off**, the app sends one final "opting out"
  message, then **deletes the random ID** and never sends again. Turning it
  back on creates a brand-new ID — the old one cannot be linked.
- The uninstaller sends one final "uninstalled" message (ID and version
  only) so uninstalls don't look like broken installs. It respects the
  Analytics switch: if you opted out, it sends nothing.

## Report a problem (only when you press Send)

The in-app "Report a problem" form sends: your message, an email address
**only if you choose to enter one**, the app version, your Windows version
number, the random install ID, and — only if you leave the "Include log
files" box ticked — the most recent lines of the app's log, with your
Windows username automatically removed. The log can include a small
technical summary of the app's state (for example how many knobs you have
and which features are on). Nothing is sent until you press Send.

## The update check

Once per day the app asks GitHub whether a newer version exists. This is an
ordinary web request to github.com (GitHub sees it the way it sees any
website visit). Downloading an update only happens when you click it.

## That's everything

There is no other network traffic. KnobMixer never sends your app names,
hotkeys, settings, volume levels, keystrokes, or files anywhere,
automatically, ever.

## Where the data lives

Pings and reports are stored on the project's own server infrastructure
(hosted on Cloudflare). Reports are kept only as long as needed to handle
them. If you sent a report with your email and want it deleted, contact us
with that email address and we'll remove it.

## Contact

KnobMixer — https://github.com/KnobMixer/KnobMixer
