# Mailsorter
A lightweight, efficient utility designed to streamline and organize Outlook mailboxes and folder structures. Built on "Find, Don't Search": Instantly filter folders, file emails with a click & use recent history.

---

> ⚠️ **Disclaimer & IT Notice:** This tool interacts with Microsoft Outlook via VBA macros. Use it at your own risk. If you are working in a corporate environment or are unfamiliar with VBA macros, please consult your system administrator or IT department before installation.

## Stop Endless Scrolling: Why I Built This Tool

As someone who is not a professional developer, I build tools in my spare time to solve real-world productivity hurdles. This project started a long time ago as a simple, basic treeview to help sort emails. Over time it evolved into a feature-rich productivity utility.

### The Problem: The Chaos of Deep Folder Structures
Everyone who deals with a high volume of emails in Microsoft Outlook knows the struggle:
* **Endless Click-Chaos:** Navigating through hundreds of nested subfolders, projects, and archives to file an email takes way too many clicks and destroys your focus.
* **Sluggish Native Search:** The built-in Outlook search or manual tree browsing is often slow, cumbersome, and interrupts your workflow.
* **Shared Mailbox Overload:** In professional environments with multiple accounts or shared mailboxes, you are flooded with irrelevant folders.

### The Solution: "Find, Don't Search"
Following the philosophy of **"Find, Don't Search"**, this tool transforms the painful chore of email filing into a lightning-fast reflex:
* **Live Search Filter:** Type part of a folder name, press Enter, and the directory filters instantly in milliseconds.
* **Recent Folders History:** Automatically tracks your last 20 visited folders so you can file subsequent emails with a single click.
* **Root Folder Filtering:** Hide everything you don't need to declutter your workspace.

---

## Built for Speed: Features That Save You Thousands of Clicks

* **Instant Folder Caching:** Scans and caches your directory structure once on startup for blazing-fast navigation without lagging Outlook.
* **Live Directory Search:** Type directly to filter and find destination folders instantly.
* **Recent Folders History:** Keeps track of your most recently used folders so you can file emails with a single click.
* **Store Management:** Control which mailboxes (Stores) are active and define custom expansion levels or allowed root folders per account.
* **Safe Error Handling:** Robust connection checks to handle shared mailboxes and offline/archived stores smoothly (`store.IsOpen`).
* **Versatile Modes:** Move or copy selected emails, or use the tool purely as a quick navigation launcher ("Go To Folder").
* **Subfolder Creation:** Create new subfolders on the fly directly from the form.
* **Tested Compatibility:** Fully compatible with **Microsoft Outlook 2024 LTSC** on Windows 11 64-bit systems (and compatible desktop versions supporting VBA).

---

## Get Started in 5 Minutes: Setup & Security Guide

1. **Clone or Download** this repository.
2. **Configure Outlook Security Settings:** 
   * Go to **File** -> **Options** -> **Trust Center** -> **Trust Center Settings** -> **Macro Settings**.
   * Select **"Notifications for all macros"**.
3. Open Microsoft Outlook and press `Alt + F11` to open the VBA Editor.
4. **Import the Files:** Go to `File` -> `Import File...` and select the files from the `/src` folder (`.bas`, `.frm`, and `.frx` files).
5. **Enable TreeView Reference:** 
   * In the VBA Editor, go to **Tools** -> **References**.
   * Ensure **Microsoft Windows Common Controls 6.0 (SP6)** is checked.
6. **Macro Signature & Security Notice:**
   * Because this code is downloaded from GitHub, Outlook treats external macros with a safety warning upon first execution. Simply allow/enable it when prompted.
   * *Advanced (Self-Signing):* If required by your environment, you can use the built-in Windows tool `SELCRT.EXE` to create a personal certificate and sign the project via **Tools -> Digital Signature**.
7. **Assign Macro:** Assign the `RunFolderPicker` macro from the `modEmailSorter` module to a quick-access toolbar button or custom shortcut.

---

## How to Master "Find, Don't Search" in Daily Use

1. **Selecting Emails:** Select emails in Outlook and run the macro to open **Move/Copy mode**. Run it without selected emails to open **Navigation mode** ("Go To Folder").
2. **Navigating & Searching:** Expand the tree or click into the filter box to type part of the folder name for instant filtering.
3. **Filing Emails:** Double-click any folder or click the action buttons to move/copy instantly. Use the recent history panel for frequent destinations.

---

## Feedback, Bugs & Contributions

As this is a side project, feedback, bug reports, and optimization ideas are very welcome! Please open an **Issue** or submit a **Pull Request**.

---

## Support the Project

If this tool makes your daily work life easier and saves you time, I would be thrilled! If you'd like to support the ongoing maintenance and development of this project, you are welcome to buy me a coffee.

---

## License

This project is open-source and free to use for both **private and commercial purposes**, with one specific condition: **You may not commercialize this software, package it, or sell the idea/code in any form.**
